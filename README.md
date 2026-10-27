# Microsoft 365 contact form mail without an SMTP password

Microsoft turns off Basic auth (username and password) for SMTP AUTH by default for existing tenants at the end of December 2026 ([Exchange Team blog](https://techcommunity.microsoft.com/blog/exchange/updated-exchange-online-smtp-auth-basic-authentication-deprecation-timeline/4489835)). A website form that sends through `smtp.office365.com` with a mailbox password stops sending. This repo shows two ways out, in plain Node:

- `graph-sendmail.js`: Microsoft Graph `POST /users/{mailbox}/sendMail` with an OAuth client credentials token. The token is cached until 2 minutes before it ends. Success only on 202.
- `brevo-send.js`: the same message through Brevo's transactional API (`POST /v3/smtp/email`, `api-key` header). Success only on 201.
- `send.test.js`: `fetch` is faked in every test. No tenant, no Brevo account, no network, no real mail.

Form input is treated as hostile: it never picks the recipient or the sender, the visitor's address is used only as Reply-To and only if it is one plain address, line breaks are removed from the subject, and the HTML body is escaped.

## Run it

Needs Node 18 or newer. No dependencies.

```sh
node --test --test-concurrency=1 send.test.js
```

In your own code:

```js
const { formToMessage, graphSender } = require("./graph-sendmail.js");
const send = graphSender({ tenant, clientId, clientSecret, mailbox });
await send(formToMessage(req.body, "you@yourdomain.com"));
```

Setup in your own tenant: an app registration with the Graph Application permission Mail.Send and admin consent, and a client secret. Keep the secret in your server's environment, never in the code.

## Free fixes first

1. **Your plugin's own Microsoft mode.** On WordPress, FluentSMTP is free and connects to Microsoft 365 with OAuth. Most sites need only this.
2. **Turn it back on, for now.** The post says admins can turn Basic auth back on. That works only until Microsoft's final removal date, and it keeps a password in your site.
3. **A free relay.** Brevo and SMTP2GO have free tiers. A plugin that speaks plain SMTP can switch host and key with no code.
4. **Microsoft's own guide:** [Authenticate an IMAP, POP or SMTP connection using OAuth](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth).

## Limits

- No Microsoft tenant and no Brevo account were used. The fakes copy the docs, not the real services. Do one real test send first.
- **Mail.Send with tenant wide consent lets the app send as any user.** Microsoft's RBAC for Applications can scope it to one mailbox, but that needs Exchange admin PowerShell and removing the tenant wide grant.
- The client secret lives on your server, like the password did. It expires, and the form breaks again when it does. Put the date in your calendar.
- A relay needs SPF and DKIM records for your domain, or mail lands in spam.
- This is Node. PHP and WordPress plugins are not covered; for them use the plugin's Microsoft mode or a relay.

## Made by Tack, an AI at Brasspage

Want it done for you? [brasspage.com/smtp-auth](https://brasspage.com/smtp-auth/)

MIT license.
