// The same contact form through a free relay instead of Microsoft 365: Brevo's transactional API
// (POST https://api.brevo.com/v3/smtp/email, header api-key, 201 with a messageId; Brevo API reference).
// WP Mail SMTP's plugin page says Brevo gives 300 free emails a day; check Brevo's own pricing page
// before you rely on it. The sender address must be verified in Brevo, and the owner's domain needs
// Brevo's SPF and DKIM records, or mail lands in spam. Brevo also offers plain SMTP with an SMTP key,
// so a plugin that only speaks SMTP can switch host and key with no code at all.
"use strict";
const { isEmail } = require("./graph-sendmail.js");

const esc = (s) =>
  s.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`).replace(/\n/g, "<br>");

// Pure. The Brevo JSON body. Brevo wants htmlContent when no template is used.
const brevoBody = (sender, { to, subject, text, replyTo }) => ({
  sender: { email: sender },
  to: [{ email: to }],
  subject,
  textContent: text,
  htmlContent: `<p>${esc(text)}</p>`,
  ...(replyTo ? { replyTo: { email: replyTo } } : {}),
});

function brevoSender({ apiKey, sender }, fetchImpl = globalThis.fetch) {
  if (!apiKey) throw new Error("no Brevo API key");
  if (!isEmail(sender)) throw new Error("bad sender");
  return async function send(msg) {
    if (!isEmail(msg.to)) throw new Error("bad recipient");
    const r = await fetchImpl("https://api.brevo.com/v3/smtp/email", {
      method: "POST",
      headers: {
        "api-key": apiKey,
        "content-type": "application/json",
        accept: "application/json",
      },
      body: JSON.stringify(brevoBody(sender, msg)),
      signal: AbortSignal.timeout(10000),
    });
    const j = await r.json().catch(() => ({}));
    if (r.status !== 201) throw new Error(`Brevo send failed: ${r.status}`);
    return {
      accepted: true,
      id: typeof j.messageId === "string" ? j.messageId : "",
    };
  };
}

module.exports = { brevoBody, brevoSender };
