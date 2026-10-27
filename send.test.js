// Run: node --test --test-concurrency=1 send.test.js
// fetch is faked in every test: no tenant, no Brevo account, no network, no real mail.
"use strict";
const test = require("node:test");
const assert = require("node:assert");
const { formToMessage, graphSender } = require("./graph-sendmail.js");
const { brevoSender } = require("./brevo-send.js");

const CFG = {
  tenant: "contoso.onmicrosoft.com",
  clientId: "00001111-aaaa-2222-bbbb-3333cccc4444",
  clientSecret: "fake-secret",
  mailbox: "forms@contoso.com",
};
const res = (status, body) => ({
  ok: status >= 200 && status < 300,
  status,
  json: async () => body ?? Promise.reject(new Error("no body")),
});
function fakeFetch(replies) {
  const calls = [];
  const f = async (url, init) => {
    calls.push({ url, init });
    return replies.shift()(url, init);
  };
  f.calls = calls;
  return f;
}

test("form input never picks the recipient, and a header break in the name stays one line", () => {
  const m = formToMessage(
    {
      name: "Ann\r\nBcc: x@evil.test",
      email: "ann@example.com",
      message: "Hi",
    },
    "owner@contoso.com",
  );
  assert.strictEqual(m.to, "owner@contoso.com");
  assert.strictEqual(m.replyTo, "ann@example.com");
  assert.ok(!/[\r\n]/.test(m.subject), m.subject);
  const bad = formToMessage(
    { email: "a@b.com>, c@d.com", message: "x" },
    "owner@contoso.com",
  );
  assert.strictEqual(bad.replyTo, null);
  assert.throws(() => formToMessage({ message: " " }, "owner@contoso.com"));
});

test("Graph: client credentials token, then sendMail as the one mailbox, 202, token reused", async () => {
  const f = fakeFetch([
    () =>
      res(200, {
        token_type: "Bearer",
        expires_in: 3599,
        access_token: "tok1",
      }),
    () => res(202),
    () => res(202),
  ]);
  const send = graphSender(CFG, f);
  const msg = formToMessage(
    { name: "Ann", email: "ann@example.com", message: "Need a quote" },
    "owner@contoso.com",
  );
  assert.deepStrictEqual(await send(msg), { accepted: true });
  await send(msg);
  assert.strictEqual(f.calls.length, 3, "second send reuses the cached token");
  const [tok, mail] = f.calls;
  assert.strictEqual(
    tok.url,
    "https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0/token",
  );
  const form = new URLSearchParams(tok.init.body);
  assert.strictEqual(form.get("grant_type"), "client_credentials");
  assert.strictEqual(form.get("scope"), "https://graph.microsoft.com/.default");
  assert.strictEqual(
    mail.url,
    "https://graph.microsoft.com/v1.0/users/forms%40contoso.com/sendMail",
  );
  assert.strictEqual(mail.init.headers.authorization, "Bearer tok1");
  const body = JSON.parse(mail.init.body);
  assert.deepStrictEqual(body.message.toRecipients, [
    { emailAddress: { address: "owner@contoso.com" } },
  ]);
  assert.deepStrictEqual(body.message.replyTo, [
    { emailAddress: { address: "ann@example.com" } },
  ]);
  assert.strictEqual(body.message.body.contentType, "Text");
  assert.ok(!mail.init.body.includes("fake-secret"));
});

test("Graph fails closed: bad token reply, a non 202 send, a bad tenant", async () => {
  const msg = { to: "owner@contoso.com", subject: "s", text: "t" };
  await assert.rejects(
    graphSender(
      CFG,
      fakeFetch([() => res(401, { error: "invalid_client" })]),
    )(msg),
    /token request failed: 401 invalid_client/,
  );
  await assert.rejects(
    graphSender(
      CFG,
      fakeFetch([
        () => res(200, { access_token: "t", expires_in: 3599 }),
        () => res(403, {}),
      ]),
    )(msg),
    /sendMail failed: 403/,
  );
  assert.throws(() => graphSender({ ...CFG, tenant: "evil.test/../x" }));
});

test("Brevo: api-key header, 201 with a messageId, html escaped, errors fail closed", async () => {
  const f = fakeFetch([
    () => res(201, { messageId: "<1@relay.example>" }),
    () => res(401, { code: "unauthorized" }),
  ]);
  const send = brevoSender(
    { apiKey: "fake-key", sender: "forms@contoso.com" },
    f,
  );
  const msg = formToMessage(
    { name: "Ann", email: "ann@example.com", message: "<script>x</script>" },
    "owner@contoso.com",
  );
  assert.deepStrictEqual(await send(msg), {
    accepted: true,
    id: "<1@relay.example>",
  });
  const { url, init } = f.calls[0];
  assert.strictEqual(url, "https://api.brevo.com/v3/smtp/email");
  assert.strictEqual(init.headers["api-key"], "fake-key");
  const body = JSON.parse(init.body);
  assert.deepStrictEqual(body.to, [{ email: "owner@contoso.com" }]);
  assert.deepStrictEqual(body.replyTo, { email: "ann@example.com" });
  assert.ok(!body.htmlContent.includes("<script>"));
  await assert.rejects(send(msg), /Brevo send failed: 401/);
});
