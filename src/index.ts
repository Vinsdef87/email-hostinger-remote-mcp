import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import { ImapFlow } from "imapflow";
import nodemailer from "nodemailer";
import { simpleParser } from "mailparser";

// ── Config ───────────────────────────────────────────────────────────────────
const IMAP_HOST = process.env.IMAP_HOST || "imap.hostinger.com";
const IMAP_PORT = parseInt(process.env.IMAP_PORT || "993", 10);
const IMAP_SECURE = (process.env.IMAP_SECURE ?? "true") !== "false";
const SMTP_HOST = process.env.SMTP_HOST || "smtp.hostinger.com";
const SMTP_PORT = parseInt(process.env.SMTP_PORT || "465", 10);
const SMTP_SECURE = (process.env.SMTP_SECURE ?? "true") !== "false";
const EMAIL_USER = process.env.EMAIL_USER || "";
const EMAIL_PASS = process.env.EMAIL_PASS || "";
const API_SECRET = process.env.API_SECRET || "";
const FROM_NAME = process.env.FROM_NAME || "";

if (!EMAIL_USER || !EMAIL_PASS) {
  console.error("Missing required env vars: EMAIL_USER, EMAIL_PASS");
  process.exit(1);
}

// ── Folder name mapping (provider-agnostic → IMAP actual) ────────────────────
// Clients (matching the Microsoft Graph connector) use friendly names like
// "inbox", "sentitems", "drafts", "deleteditems", "junkemail", "archive".
// On Hostinger (Roundcube/Postfix) typical folders are:
//   INBOX, INBOX.Sent, INBOX.Drafts, INBOX.Trash, INBOX.Junk, INBOX.Archive
// But some hosts use Sent/Drafts/Trash at root. We try the "best guess" first,
// then fall back by listing real folders on the server if the first doesn't exist.
const FOLDER_ALIASES: Record<string, string[]> = {
  inbox: ["INBOX"],
  sent: ["INBOX.Sent", "Sent", "Sent Items", "Sent Messages"],
  sentitems: ["INBOX.Sent", "Sent", "Sent Items", "Sent Messages"],
  drafts: ["INBOX.Drafts", "Drafts"],
  trash: ["INBOX.Trash", "Trash", "Deleted", "Deleted Items", "Deleted Messages"],
  deleteditems: ["INBOX.Trash", "Trash", "Deleted", "Deleted Items", "Deleted Messages"],
  junk: ["INBOX.Junk", "Junk", "Spam", "Junk E-mail"],
  junkemail: ["INBOX.Junk", "Junk", "Spam", "Junk E-mail"],
  spam: ["INBOX.Spam", "Spam", "INBOX.Junk", "Junk"],
  archive: ["INBOX.Archive", "Archive", "Archives"],
};

// ── IMAP connection helper ───────────────────────────────────────────────────
const IMAP_CONNECTION_TIMEOUT = parseInt(process.env.IMAP_CONNECTION_TIMEOUT || "15000", 10);
const IMAP_GREETING_TIMEOUT = parseInt(process.env.IMAP_GREETING_TIMEOUT || "15000", 10);
const IMAP_SOCKET_TIMEOUT = parseInt(process.env.IMAP_SOCKET_TIMEOUT || "60000", 10);

function newImapClient(): ImapFlow {
  const client = new ImapFlow({
    host: IMAP_HOST,
    port: IMAP_PORT,
    secure: IMAP_SECURE,
    auth: { user: EMAIL_USER, pass: EMAIL_PASS },
    logger: false,
    connectionTimeout: IMAP_CONNECTION_TIMEOUT,
    greetingTimeout: IMAP_GREETING_TIMEOUT,
    socketTimeout: IMAP_SOCKET_TIMEOUT,
  });
  // ImapFlow emits 'error' on socket problems (ECONNRESET, ETIMEOUT...).
  // Without a listener Node treats it as uncaught and kills the process.
  client.on("error", (err: Error) => {
    console.error("[imap] connection error:", (err as any)?.code || "", err?.message || err);
  });
  return client;
}

async function resolveFolder(client: ImapFlow, requested: string): Promise<string> {
  const key = (requested || "inbox").toLowerCase();
  const candidates = FOLDER_ALIASES[key] || [requested];
  // list once and pick the first match
  const tree = await client.list();
  const names = new Set(tree.map((m) => m.path));
  for (const c of candidates) {
    if (names.has(c)) return c;
  }
  // fallback: case-insensitive match or return INBOX
  for (const c of candidates) {
    const found = tree.find((m) => m.path.toLowerCase() === c.toLowerCase());
    if (found) return found.path;
  }
  // last resort: return the original (IMAP will error out if invalid)
  return requested || "INBOX";
}

async function withImap<T>(fn: (client: ImapFlow) => Promise<T>): Promise<T> {
  const client = newImapClient();
  try {
    await client.connect();
  } catch (err) {
    const e = err as any;
    try { client.close(); } catch { /* ignore */ }
    throw new Error(
      `IMAP connect to ${IMAP_HOST}:${IMAP_PORT} failed (${e?.code || e?.responseText || "unknown"}): ${e?.message || e}`
    );
  }
  try {
    return await fn(client);
  } finally {
    try { await client.logout(); } catch { try { client.close(); } catch { /* ignore */ } }
  }
}

// ── SMTP transporter (reused across requests) ────────────────────────────────
const smtpTransporter = nodemailer.createTransport({
  host: SMTP_HOST,
  port: SMTP_PORT,
  secure: SMTP_SECURE,
  auth: { user: EMAIL_USER, pass: EMAIL_PASS },
});

// ── MCP Server ───────────────────────────────────────────────────────────────
function createServer(): McpServer {
  const server = new McpServer({ name: "email-hostinger-mcp-server", version: "1.0.0" });

  // ── LIST MESSAGES ────────────────────────────────────────────────────────────
  server.registerTool("email_list_messages", {
    title: "List Emails",
    description: "List emails from a folder (inbox by default). Returns UID, subject, from, date, isRead, bodyPreview.",
    inputSchema: {
      folder: z.string().optional().default("inbox").describe("Folder: inbox, sent, drafts, trash, junk, archive (aliases supported)"),
      top: z.number().int().min(1).max(100).optional().default(10).describe("Number of emails to return"),
      unreadOnly: z.boolean().optional().default(false).describe("Return only unread emails"),
      search: z.string().optional().describe("Full-text search query (uses IMAP SEARCH TEXT)"),
    },
    annotations: { readOnlyHint: true, destructiveHint: false }
  }, async ({ folder, top, unreadOnly, search }) => {
    const data = await withImap(async (client) => {
      const mailbox = await resolveFolder(client, folder ?? "inbox");
      const lock = await client.getMailboxLock(mailbox);
      try {
        const searchCriteria: Record<string, unknown> = {};
        if (unreadOnly) searchCriteria.seen = false;
        if (search) searchCriteria.body = search;
        let uids: number[] | null = null;
        if (search) {
          // full-text search: has to scan the whole mailbox
          uids = (await client.search(searchCriteria, { uid: true })) || [];
        } else if (unreadOnly) {
          // UNSEEN over the whole mailbox is very slow on Hostinger for big
          // mailboxes (> 60s → socket timeout). Search a recent UID window and
          // widen it only if we did not find enough messages.
          const limitWanted = top ?? 10;
          const uidNext = (client.mailbox && typeof client.mailbox === "object" && (client.mailbox as any).uidNext) || 0;
          let window = Math.max(limitWanted * 20, 200);
          uids = [];
          for (let i = 0; i < 4; i++) {
            const fromUid = uidNext ? Math.max(1, uidNext - window) : 1;
            const found = (await client.search({ ...searchCriteria, uid: `${fromUid}:*` }, { uid: true })) || [];
            uids = found;
            if (found.length >= limitWanted || fromUid === 1 || !uidNext) break;
            window *= 4;
          }
        }

        const results: any[] = [];
        const limit = top ?? 10;

        if (uids) {
          // fetch specific uids, newest first
          const sliced = uids.slice(-limit).reverse();
          for await (const msg of client.fetch(sliced, { uid: true, envelope: true, flags: true, bodyStructure: true, source: false }, { uid: true })) {
            // fetch a preview
            let preview = "";
            try {
              const d = await client.download(String(msg.uid), undefined, { uid: true, maxBytes: 1024 });
              if (d && d.content) {
                const chunks: Buffer[] = [];
                for await (const c of d.content) chunks.push(c as Buffer);
                preview = Buffer.concat(chunks).toString("utf-8").slice(0, 200);
              }
            } catch { /* ignore preview errors */ }
            results.push({
              id: String(msg.uid),
              subject: msg.envelope?.subject ?? "",
              from: msg.envelope?.from?.[0] ? { name: msg.envelope.from[0].name, address: msg.envelope.from[0].address } : null,
              to: msg.envelope?.to?.map((a) => ({ name: a.name, address: a.address })) ?? [],
              date: msg.envelope?.date,
              isRead: msg.flags?.has("\\Seen") ?? false,
              bodyPreview: preview.replace(/\s+/g, " ").trim().slice(0, 150),
            });
          }
        } else {
          // walk from newest: use sequence range
          const status = await client.status(mailbox, { messages: true });
          const totalMsgs = status.messages ?? 0;
          if (totalMsgs === 0) return { folder: mailbox, count: 0, messages: [] };
          const first = Math.max(1, totalMsgs - limit + 1);
          const range = `${first}:${totalMsgs}`;
          const tmp: any[] = [];
          for await (const msg of client.fetch(range, { uid: true, envelope: true, flags: true })) {
            tmp.push(msg);
          }
          tmp.reverse(); // newest first
          for (const msg of tmp) {
            let preview = "";
            try {
              const d = await client.download(String(msg.uid), undefined, { uid: true, maxBytes: 1024 });
              if (d && d.content) {
                const chunks: Buffer[] = [];
                for await (const c of d.content) chunks.push(c as Buffer);
                preview = Buffer.concat(chunks).toString("utf-8").slice(0, 200);
              }
            } catch { /* ignore */ }
            results.push({
              id: String(msg.uid),
              subject: msg.envelope?.subject ?? "",
              from: msg.envelope?.from?.[0] ? { name: msg.envelope.from[0].name, address: msg.envelope.from[0].address } : null,
              to: msg.envelope?.to?.map((a) => ({ name: a.name, address: a.address })) ?? [],
              date: msg.envelope?.date,
              isRead: msg.flags?.has("\\Seen") ?? false,
              bodyPreview: preview.replace(/\s+/g, " ").trim().slice(0, 150),
            });
          }
        }
        return { folder: mailbox, count: results.length, messages: results };
      } finally {
        lock.release();
      }
    });
    return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
  });

  // ── GET MESSAGE ──────────────────────────────────────────────────────────────
  server.registerTool("email_get_message", {
    title: "Get Email",
    description: "Get full content of a specific email by UID. Returns headers, body (HTML + text), attachments metadata.",
    inputSchema: {
      messageId: z.string().describe("Email UID returned by email_list_messages"),
      folder: z.string().optional().default("inbox").describe("Folder where the message lives"),
    },
    annotations: { readOnlyHint: true, destructiveHint: false }
  }, async ({ messageId, folder }) => {
    const data = await withImap(async (client) => {
      const mailbox = await resolveFolder(client, folder ?? "inbox");
      const lock = await client.getMailboxLock(mailbox);
      try {
        const d = await client.download(messageId, undefined, { uid: true });
        if (!d || !d.content) return { error: "Message not found" };
        const chunks: Buffer[] = [];
        for await (const c of d.content) chunks.push(c as Buffer);
        const raw = Buffer.concat(chunks);
        const parsed = await simpleParser(raw);
        return {
          id: messageId,
          subject: parsed.subject,
          from: parsed.from?.value?.[0] ? { name: parsed.from.value[0].name, address: parsed.from.value[0].address } : null,
          to: parsed.to ? (Array.isArray(parsed.to) ? parsed.to : [parsed.to]).flatMap((t: any) => (t.value || []).map((a: any) => ({ name: a.name, address: a.address }))) : [],
          cc: parsed.cc ? (Array.isArray(parsed.cc) ? parsed.cc : [parsed.cc]).flatMap((t: any) => (t.value || []).map((a: any) => ({ name: a.name, address: a.address }))) : [],
          date: parsed.date,
          body: {
            text: parsed.text ?? "",
            html: parsed.html || null,
          },
          attachments: (parsed.attachments || []).map((a: any) => ({
            filename: a.filename,
            contentType: a.contentType,
            size: a.size,
            contentId: a.contentId,
          })),
          headers: {
            messageId: parsed.messageId,
            inReplyTo: parsed.inReplyTo,
            references: parsed.references,
          },
        };
      } finally {
        lock.release();
      }
    });
    return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
  });

  // ── SEND EMAIL ───────────────────────────────────────────────────────────────
  server.registerTool("email_send", {
    title: "Send Email",
    description: `Send an email from ${EMAIL_USER}`,
    inputSchema: {
      to: z.string().describe("Recipient email address (comma-separated for multiple)"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body in HTML or plain text"),
      isHtml: z.boolean().optional().default(false).describe("Set true if body is HTML"),
      cc: z.string().optional().describe("CC email address(es), comma-separated"),
      bcc: z.string().optional().describe("BCC email address(es), comma-separated"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false }
  }, async ({ to, subject, body, isHtml, cc, bcc }) => {
    const info = await smtpTransporter.sendMail({
      from: FROM_NAME ? `"${FROM_NAME}" <${EMAIL_USER}>` : EMAIL_USER,
      to,
      cc,
      bcc,
      subject,
      [isHtml ? "html" : "text"]: body,
    });
    // Append to Sent folder (SMTP doesn't do this automatically)
    try {
      await withImap(async (client) => {
        const sentFolder = await resolveFolder(client, "sent");
        const raw = [
          `From: ${FROM_NAME ? `"${FROM_NAME}" <${EMAIL_USER}>` : EMAIL_USER}`,
          `To: ${to}`,
          cc ? `Cc: ${cc}` : "",
          `Subject: ${subject}`,
          `Date: ${new Date().toUTCString()}`,
          `Message-ID: ${info.messageId}`,
          `MIME-Version: 1.0`,
          `Content-Type: ${isHtml ? "text/html" : "text/plain"}; charset=utf-8`,
          ``,
          body,
        ].filter(Boolean).join("\r\n");
        await client.append(sentFolder, raw, ["\\Seen"]);
      });
    } catch (err) {
      // don't fail the send if APPEND to Sent fails
      console.error("[email_send] append-to-sent failed:", (err as Error).message);
    }
    return { content: [{ type: "text", text: JSON.stringify({ success: true, messageId: info.messageId, to, subject }, null, 2) }] };
  });

  // ── REPLY TO EMAIL ───────────────────────────────────────────────────────────
  server.registerTool("email_reply", {
    title: "Reply to Email",
    description: "Reply to a specific email. Preserves threading (In-Reply-To, References) and quotes original.",
    inputSchema: {
      messageId: z.string().describe("Email UID to reply to"),
      body: z.string().describe("Reply body"),
      isHtml: z.boolean().optional().default(false).describe("Set true if body is HTML"),
      replyAll: z.boolean().optional().default(false).describe("Reply to all recipients (To + Cc)"),
      folder: z.string().optional().default("inbox").describe("Folder where the original message lives"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false }
  }, async ({ messageId, body, isHtml, replyAll, folder }) => {
    // Fetch original for headers
    const original = await withImap(async (client) => {
      const mailbox = await resolveFolder(client, folder ?? "inbox");
      const lock = await client.getMailboxLock(mailbox);
      try {
        const d = await client.download(messageId, undefined, { uid: true });
        if (!d || !d.content) return null;
        const chunks: Buffer[] = [];
        for await (const c of d.content) chunks.push(c as Buffer);
        const parsed = await simpleParser(Buffer.concat(chunks));
        return parsed;
      } finally {
        lock.release();
      }
    });
    if (!original) return { content: [{ type: "text", text: JSON.stringify({ error: "Original message not found" }) }] };

    const fromAddress = original.from?.value?.[0]?.address ?? "";
    const toAddresses: string[] = [fromAddress];
    let ccAddresses: string[] = [];
    if (replyAll) {
      const origTo = original.to ? (Array.isArray(original.to) ? original.to : [original.to]).flatMap((t: any) => (t.value || []).map((a: any) => a.address)) : [];
      const origCc = original.cc ? (Array.isArray(original.cc) ? original.cc : [original.cc]).flatMap((t: any) => (t.value || []).map((a: any) => a.address)) : [];
      // exclude ourselves from the reply list
      const others = [...origTo, ...origCc].filter((a) => a && a.toLowerCase() !== EMAIL_USER.toLowerCase());
      ccAddresses = others;
    }

    const subject = original.subject?.startsWith("Re:") ? original.subject : `Re: ${original.subject ?? ""}`;
    const inReplyTo = original.messageId;
    const references = (original.references ? (Array.isArray(original.references) ? original.references.join(" ") : original.references) + " " : "") + (original.messageId ?? "");

    const info = await smtpTransporter.sendMail({
      from: FROM_NAME ? `"${FROM_NAME}" <${EMAIL_USER}>` : EMAIL_USER,
      to: toAddresses.join(", "),
      cc: ccAddresses.length ? ccAddresses.join(", ") : undefined,
      subject,
      inReplyTo,
      references,
      [isHtml ? "html" : "text"]: body,
    });

    // Mark original as answered
    try {
      await withImap(async (client) => {
        const mailbox = await resolveFolder(client, folder ?? "inbox");
        const lock = await client.getMailboxLock(mailbox);
        try {
          await client.messageFlagsAdd(messageId, ["\\Answered"], { uid: true });
        } finally {
          lock.release();
        }
      });
    } catch { /* non-fatal */ }

    return { content: [{ type: "text", text: JSON.stringify({ success: true, messageId: info.messageId, to: toAddresses, cc: ccAddresses, subject }, null, 2) }] };
  });

  // ── MARK AS READ ─────────────────────────────────────────────────────────────
  server.registerTool("email_mark_read", {
    title: "Mark Email as Read",
    description: "Mark an email as read (adds \\Seen flag)",
    inputSchema: {
      messageId: z.string().describe("Email UID"),
      folder: z.string().optional().default("inbox").describe("Folder where the message lives"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false }
  }, async ({ messageId, folder }) => {
    await withImap(async (client) => {
      const mailbox = await resolveFolder(client, folder ?? "inbox");
      const lock = await client.getMailboxLock(mailbox);
      try {
        await client.messageFlagsAdd(messageId, ["\\Seen"], { uid: true });
      } finally {
        lock.release();
      }
    });
    return { content: [{ type: "text", text: "Email marked as read" }] };
  });

  // ── MOVE EMAIL ───────────────────────────────────────────────────────────────
  server.registerTool("email_move", {
    title: "Move Email",
    description: "Move an email to a different folder",
    inputSchema: {
      messageId: z.string().describe("Email UID"),
      destinationFolder: z.string().describe("Destination folder: inbox, sent, drafts, trash, junk, archive (aliases supported)"),
      folder: z.string().optional().default("inbox").describe("Current folder where the message lives"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false }
  }, async ({ messageId, destinationFolder, folder }) => {
    const result = await withImap(async (client) => {
      const src = await resolveFolder(client, folder ?? "inbox");
      const dst = await resolveFolder(client, destinationFolder);
      const lock = await client.getMailboxLock(src);
      try {
        const res = await client.messageMove(messageId, dst, { uid: true });
        return { from: src, to: dst, result: res };
      } finally {
        lock.release();
      }
    });
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  });

  // ── SEARCH EMAILS ────────────────────────────────────────────────────────────
  server.registerTool("email_search", {
    title: "Search Emails",
    description: "Search emails by keyword, sender, or subject. Supports prefixes: 'from:', 'subject:', 'body:' (default = body search).",
    inputSchema: {
      query: z.string().describe("Query like 'from:customer@example.com', 'subject:order', or free text"),
      folder: z.string().optional().default("inbox").describe("Folder to search in"),
      top: z.number().int().min(1).max(100).optional().default(10).describe("Number of results"),
    },
    annotations: { readOnlyHint: true, destructiveHint: false }
  }, async ({ query, folder, top }) => {
    const data = await withImap(async (client) => {
      const mailbox = await resolveFolder(client, folder ?? "inbox");
      const lock = await client.getMailboxLock(mailbox);
      try {
        const criteria: Record<string, unknown> = {};
        const q = query.trim();
        if (q.toLowerCase().startsWith("from:")) {
          criteria.from = q.slice(5).trim();
        } else if (q.toLowerCase().startsWith("subject:")) {
          criteria.subject = q.slice(8).trim();
        } else if (q.toLowerCase().startsWith("body:")) {
          criteria.body = q.slice(5).trim();
        } else {
          criteria.body = q;
        }
        const uids = await client.search(criteria, { uid: true });
        const limit = top ?? 10;
        const sliced = (uids || []).slice(-limit).reverse();
        const results: any[] = [];
        if (sliced.length > 0) {
          for await (const msg of client.fetch(sliced, { uid: true, envelope: true, flags: true }, { uid: true })) {
            results.push({
              id: String(msg.uid),
              subject: msg.envelope?.subject ?? "",
              from: msg.envelope?.from?.[0] ? { name: msg.envelope.from[0].name, address: msg.envelope.from[0].address } : null,
              date: msg.envelope?.date,
              isRead: msg.flags?.has("\\Seen") ?? false,
            });
          }
        }
        return { folder: mailbox, query: q, count: results.length, messages: results };
      } finally {
        lock.release();
      }
    });
    return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
  });

  return server;
}

// ── Express app ──────────────────────────────────────────────────────────────
const app = express();
app.use(express.json({ limit: "10mb" }));

app.use((req: Request, res: Response, next) => {
  if (req.path === "/health") return next();
  if (API_SECRET) {
    const token = req.headers.authorization?.replace("Bearer ", "");
    if (token !== API_SECRET) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
  }
  next();
});

app.get("/health", (_req: Request, res: Response) => {
  res.json({ status: "ok", service: "email-hostinger-mcp-server", version: "1.0.2", mailbox: EMAIL_USER });
});

// Diagnostic endpoint (bearer-protected): egress IP, DNS, raw TCP/TLS probe to IMAP/SMTP hosts.
app.get("/debug/net", async (req: Request, res: Response) => {
  const auth = req.headers.authorization || "";
  if (!API_SECRET || auth !== `Bearer ${API_SECRET}`) { res.status(401).json({ error: "unauthorized" }); return; }
  const dns = await import("node:dns/promises");
  const tls = await import("node:tls");
  const net = await import("node:net");
  const out: any = { node: process.version, imapHost: IMAP_HOST, imapPort: IMAP_PORT };
  try {
    const r = await fetch("https://api.ipify.org?format=json", { signal: AbortSignal.timeout(8000) });
    out.egressIp = (await r.json() as any).ip;
  } catch (e: any) { out.egressIp = `error: ${e?.message}`; }
  try {
    const r6 = await fetch("https://api64.ipify.org?format=json", { signal: AbortSignal.timeout(8000) });
    out.egressIp64 = (await r6.json() as any).ip;
  } catch (e: any) { out.egressIp64 = `error: ${e?.message}`; }
  const probe = (host: string, port: number, family: 4 | 6, useTls: boolean) => new Promise<any>((resolve) => {
    const t0 = Date.now();
    const done = (r: any) => resolve({ host, port, family, tls: useTls, ms: Date.now() - t0, ...r });
    const opts: any = { host, port, family, timeout: 10000 };
    const sock = useTls
      ? tls.connect({ ...opts, servername: host, rejectUnauthorized: false })
      : net.connect(opts);
    let greeting = "";
    sock.setTimeout(10000);
    sock.on(useTls ? "secureConnect" : "connect", () => { /* wait for greeting */ });
    sock.on("data", (d: Buffer) => { greeting += d.toString("utf8"); if (greeting.includes("\n")) { sock.destroy(); done({ ok: true, greeting: greeting.trim().slice(0, 120) }); } });
    sock.on("timeout", () => { sock.destroy(); done({ ok: false, error: "timeout", greeting: greeting.slice(0, 120) }); });
    sock.on("error", (e: any) => { done({ ok: false, error: e?.code || e?.message }); });
  });
  for (const h of [IMAP_HOST, SMTP_HOST]) {
    try { out[`dns_${h}`] = await dns.lookup(h, { all: true }); } catch (e: any) { out[`dns_${h}`] = `error: ${e?.code}`; }
  }
  out.probes = await Promise.all([
    probe(IMAP_HOST, IMAP_PORT, 4, true),
    probe(IMAP_HOST, IMAP_PORT, 6, true),
    probe(IMAP_HOST, 143, 4, false),
    probe(SMTP_HOST, SMTP_PORT, 4, true),
    probe(SMTP_HOST, 587, 4, false),
  ]);
  res.json(out);
});

// Diagnostic: step-by-step IMAP timings (bearer-protected)
app.get("/debug/imap", async (req: Request, res: Response) => {
  const auth = req.headers.authorization || "";
  if (!API_SECRET || auth !== `Bearer ${API_SECRET}`) { res.status(401).json({ error: "unauthorized" }); return; }
  const steps: any[] = [];
  const timed = async (name: string, fn: () => Promise<any>, guardMs = 25000) => {
    const t0 = Date.now();
    try {
      const r = await Promise.race([fn(), new Promise((_, rej) => setTimeout(() => rej(new Error("guard timeout")), guardMs))]);
      steps.push({ name, ms: Date.now() - t0, ok: true, result: r });
      return r;
    } catch (e: any) {
      steps.push({ name, ms: Date.now() - t0, ok: false, error: e?.code || e?.message });
      return undefined;
    }
  };
  const client = new ImapFlow({
    host: IMAP_HOST, port: IMAP_PORT, secure: IMAP_SECURE,
    auth: { user: EMAIL_USER, pass: EMAIL_PASS },
    logger: { debug: (o: any) => console.log("[imapdbg]", JSON.stringify(o).slice(0, 300)), info: () => {}, warn: (o: any) => console.warn("[imapdbg]", JSON.stringify(o).slice(0, 300)), error: (o: any) => console.error("[imapdbg]", JSON.stringify(o).slice(0, 300)) } as any,
    connectionTimeout: 15000, greetingTimeout: 15000, socketTimeout: 30000,
  });
  client.on("error", (e: Error) => steps.push({ event: "error", error: (e as any)?.code || e?.message }));
  await timed("connect", () => client.connect());
  const mb: any = await timed("mailboxOpen INBOX", async () => { const m: any = await client.mailboxOpen("INBOX"); return { exists: m.exists, uidNext: m.uidNext, uidValidity: String(m.uidValidity) }; });
  const uidNext = mb?.uidNext || 0;
  const w = Math.max(1, uidNext - 200);
  const u1: any = await timed(`search unseen uid ${w}:*`, async () => { const r = await client.search({ seen: false, uid: `${w}:*` }, { uid: true }); return { count: (r || []).length, last: (r || []).slice(-3) }; });
  const last = (u1?.last || []) as number[];
  if (last.length) {
    await timed("fetch envelope+flags", async () => { const out: any[] = []; for await (const m of client.fetch(last, { uid: true, envelope: true, flags: true }, { uid: true })) out.push(m.envelope?.subject); return out; });
    await timed("fetch +bodyStructure", async () => { const out: any[] = []; for await (const m of client.fetch(last, { uid: true, envelope: true, flags: true, bodyStructure: true }, { uid: true })) out.push(m.uid); return out; });
    await timed("download 1024B", async () => { const d = await client.download(String(last[last.length - 1]), undefined, { uid: true, maxBytes: 1024 }); let n = 0; if (d?.content) for await (const c of d.content) n += (c as Buffer).length; return { bytes: n }; });
  }
  await timed("search unseen FULL", async () => { const r = await client.search({ seen: false }, { uid: true }); return { count: (r || []).length }; }, 20000);
  try { await client.logout(); } catch { try { client.close(); } catch { /* ignore */ } }
  res.json({ steps });
});

app.all("/mcp", async (req: Request, res: Response) => {
  try {
    const server = createServer();
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error("[mcp] error:", (err as Error).message);
    if (!res.headersSent) res.status(500).json({ error: (err as Error).message });
  }
});

// Safety net: never let a stray socket error take the whole server down.
process.on("uncaughtException", (err) => {
  console.error("[process] uncaughtException:", (err as any)?.code || "", err?.message || err);
});
process.on("unhandledRejection", (reason) => {
  console.error("[process] unhandledRejection:", (reason as any)?.message || reason);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Email Hostinger MCP server v1.0.2 running on port ${PORT} (mailbox: ${EMAIL_USER})`);
});
