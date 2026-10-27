// Send a website contact form through Microsoft Graph sendMail with OAuth client credentials,
// in place of smtp.office365.com with a mailbox password (Basic auth for SMTP AUTH is off by default
// for existing Microsoft 365 tenants at the end of December 2026; Exchange Team blog, Jan 27, 2026).
// Owner setup (in their own tenant, never ours): an app registration with the Graph Application
// permission Mail.Send and admin consent, and a client secret. Better: skip the tenant wide consent
// and grant "Application Mail.Send" with RBAC for Applications scoped to the one sending mailbox.
// API names from learn.microsoft.com: v2-oauth2-client-creds-grant-flow, graph/api/user-sendmail.
// Form input is untrusted: it never picks the recipient or the sender, only the text and Reply-To.
"use strict";

const EMAIL = /^[^\s@<>"',;:\\]{1,64}@[A-Za-z0-9.-]{1,253}\.[A-Za-z]{2,63}$/;
const TENANT =
  /^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|[A-Za-z0-9-]{1,63}(\.[A-Za-z0-9-]{1,63})+)$/i;
const isEmail = (s) =>
  typeof s === "string" && s.length <= 254 && EMAIL.test(s);
const oneLine = (s, n) =>
  String(s ?? "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .trim()
    .slice(0, n);

// Pure. Contact form fields -> one message to the owner. The visitor's address is only a Reply-To.
function formToMessage(form, owner) {
  if (!isEmail(owner)) throw new Error("owner address is not an email");
  const name = oneLine(form?.name, 100);
  const text = String(form?.message ?? "").slice(0, 10000);
  if (!text.trim()) throw new Error("empty message");
  return {
    to: owner,
    subject: oneLine(`Website form: ${name || "no name"}`, 200),
    text: `From: ${name}\nEmail: ${oneLine(form?.email, 254)}\n\n${text}`,
    replyTo: isEmail(form?.email) ? form.email : null,
  };
}

// Pure. The sendMail JSON body.
const graphBody = ({ to, subject, text, replyTo }) => ({
  message: {
    subject,
    body: { contentType: "Text", content: text },
    toRecipients: [{ emailAddress: { address: to } }],
    ...(replyTo ? { replyTo: [{ emailAddress: { address: replyTo } }] } : {}),
  },
  saveToSentItems: false,
});

function graphSender(
  { tenant, clientId, clientSecret, mailbox },
  fetchImpl = globalThis.fetch,
) {
  if (!TENANT.test(tenant || "")) throw new Error("bad tenant id");
  if (!/^[0-9a-f-]{36}$/i.test(clientId || ""))
    throw new Error("bad client id");
  if (!clientSecret) throw new Error("no client secret");
  if (!isEmail(mailbox)) throw new Error("bad sending mailbox");
  let token = null; // { value, until }
  async function getToken(now = Date.now()) {
    if (token && now < token.until) return token.value;
    const r = await fetchImpl(
      `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }).toString(),
        signal: AbortSignal.timeout(10000),
      },
    );
    const j = await r.json().catch(() => ({}));
    if (!r.ok || typeof j.access_token !== "string")
      throw new Error(
        `token request failed: ${r.status} ${oneLine(j.error, 80)}`,
      );
    const life = Math.min(Number(j.expires_in) || 0, 3600);
    token = {
      value: j.access_token,
      until: now + Math.max(life - 120, 0) * 1000,
    };
    return token.value;
  }
  return async function send(msg) {
    if (!isEmail(msg.to)) throw new Error("bad recipient");
    const r = await fetchImpl(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/sendMail`,
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${await getToken()}`,
          "content-type": "application/json",
        },
        body: JSON.stringify(graphBody(msg)),
        signal: AbortSignal.timeout(10000),
      },
    );
    // 202 means Graph accepted it, not that it was delivered.
    if (r.status !== 202) throw new Error(`sendMail failed: ${r.status}`);
    return { accepted: true };
  };
}

module.exports = { formToMessage, graphBody, graphSender, isEmail, oneLine };
