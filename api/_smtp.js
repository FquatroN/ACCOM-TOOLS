const net = require("node:net");
const tls = require("node:tls");
const { randomBytes } = require("node:crypto");
const { EventEmitter, once } = require("node:events");

function encodeHeader(value = "") {
  const text = String(value || "");
  if (!/[^\x20-\x7E]/.test(text)) return text;
  return `=?UTF-8?B?${Buffer.from(text, "utf8").toString("base64")}?=`;
}

function formatAddress(email, name = "") {
  const cleanEmail = String(email || "").trim();
  const cleanName = String(name || "").trim();
  return cleanName ? `${encodeHeader(cleanName)} <${cleanEmail}>` : cleanEmail;
}

function normalizeCrlf(value = "") {
  return String(value || "").replace(/\r?\n/g, "\r\n");
}

function escapeData(value = "") {
  return normalizeCrlf(value).replace(/^\./gm, "..");
}

function buildMimeMessage({ fromEmail, fromName, to = [], subject = "", html = "", text = "" }) {
  const boundary = `alt-${randomBytes(12).toString("hex")}`;
  const dateHeader = new Date().toUTCString();
  const domain = String(fromEmail || "").split("@")[1] || "localhost";
  const messageId = `<${Date.now()}.${randomBytes(6).toString("hex")}@${domain}>`;
  const headers = [
    `From: ${formatAddress(fromEmail, fromName)}`,
    `To: ${(Array.isArray(to) ? to : []).join(", ")}`,
    `Subject: ${encodeHeader(subject)}`,
    `Date: ${dateHeader}`,
    `Message-ID: ${messageId}`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/alternative; boundary="${boundary}"`,
  ];
  const body = [
    `--${boundary}`,
    'Content-Type: text/plain; charset="UTF-8"',
    "Content-Transfer-Encoding: 8bit",
    "",
    normalizeCrlf(text),
    `--${boundary}`,
    'Content-Type: text/html; charset="UTF-8"',
    "Content-Transfer-Encoding: 8bit",
    "",
    normalizeCrlf(html),
    `--${boundary}--`,
    "",
  ].join("\r\n");
  return `${headers.join("\r\n")}\r\n\r\n${body}`;
}

function createBufferedConnection(socket) {
  const emitter = new EventEmitter();
  let buffer = "";
  let activeSocket = socket;

  function tryExtractResponse() {
    let offset = 0;
    while (true) {
      const end = buffer.indexOf("\r\n", offset);
      if (end < 0) return null;
      const line = buffer.slice(offset, end);
      offset = end + 2;
      if (/^\d{3} /.test(line)) {
        const raw = buffer.slice(0, offset);
        buffer = buffer.slice(offset);
        return {
          code: Number(line.slice(0, 3)),
          lines: raw.trim().split(/\r\n/),
        };
      }
    }
  }

  function bindSocket(nextSocket) {
    activeSocket = nextSocket;
    activeSocket.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      emitter.emit("data");
    });
    activeSocket.on("error", (error) => emitter.emit("error", error));
    activeSocket.on("close", () => emitter.emit("close"));
  }

  bindSocket(socket);

  function waitForReadable() {
    return new Promise((resolve, reject) => {
      const onData = () => {
        cleanup();
        resolve();
      };
      const onError = (error) => {
        cleanup();
        reject(error);
      };
      const onClose = () => {
        cleanup();
        reject(new Error("SMTP connection closed unexpectedly."));
      };
      function cleanup() {
        emitter.off("data", onData);
        emitter.off("error", onError);
        emitter.off("close", onClose);
      }
      emitter.once("data", onData);
      emitter.once("error", onError);
      emitter.once("close", onClose);
    });
  }

  async function readResponse() {
    while (true) {
      const response = tryExtractResponse();
      if (response) return response;
      await waitForReadable();
    }
  }

  return {
    get socket() {
      return activeSocket;
    },
    set socket(nextSocket) {
      bindSocket(nextSocket);
    },
    async command(command, expectedCodes = [250]) {
      activeSocket.write(`${command}\r\n`);
      const response = await readResponse();
      if (!expectedCodes.includes(response.code)) {
        throw new Error(`SMTP error after "${command}": ${response.lines.join(" | ")}`);
      }
      return response;
    },
    async writeData(payload, expectedCodes = [250]) {
      activeSocket.write(`${escapeData(payload)}\r\n.\r\n`);
      const response = await readResponse();
      if (!expectedCodes.includes(response.code)) {
        throw new Error(`SMTP data rejected: ${response.lines.join(" | ")}`);
      }
      return response;
    },
    readResponse,
    end() {
      try {
        activeSocket.end();
      } catch {}
    },
  };
}

async function createSmtpSocket({ host, port, secure }) {
  const baseOptions = { host, port, servername: host };
  const socket = secure
    ? tls.connect(baseOptions)
    : net.connect(baseOptions);
  await Promise.race([
    once(socket, secure ? "secureConnect" : "connect"),
    once(socket, "error").then(([error]) => Promise.reject(error)),
  ]);
  return socket;
}

async function sendWithSmtp(config = {}, mail = {}) {
  const host = String(config.smtpHost || "").trim();
  const port = Number(config.smtpPort || 465);
  const secure = !!config.smtpSecure;
  const username = String(config.smtpUser || "").trim();
  const password = String(config.smtpPassword || "");
  const fromEmail = String(config.fromEmail || username).trim();
  const fromName = String(config.fromName || "").trim();
  const recipients = Array.isArray(mail.to) ? mail.to.filter(Boolean) : [];
  if (!host || !port || !username || !password || !fromEmail || !recipients.length) {
    throw new Error("SMTP configuration is incomplete.");
  }

  const socket = await createSmtpSocket({ host, port, secure });
  const connection = createBufferedConnection(socket);
  try {
    await connection.readResponse();
    await connection.command(`EHLO ${host}`);
    await connection.command("AUTH LOGIN", [334]);
    await connection.command(Buffer.from(username, "utf8").toString("base64"), [334]);
    await connection.command(Buffer.from(password, "utf8").toString("base64"), [235]);
    await connection.command(`MAIL FROM:<${fromEmail}>`);
    for (const recipient of recipients) {
      await connection.command(`RCPT TO:<${recipient}>`, [250, 251]);
    }
    await connection.command("DATA", [354]);
    const message = buildMimeMessage({
      fromEmail,
      fromName,
      to: recipients,
      subject: mail.subject,
      html: mail.html,
      text: mail.text,
    });
    await connection.writeData(message, [250]);
    await connection.command("QUIT", [221]);
    return { provider: "smtp", accepted: recipients };
  } finally {
    connection.end();
  }
}

module.exports = {
  sendWithSmtp,
};
