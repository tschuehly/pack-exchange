# pack-exchange

Microsoft 365 Outlook mail access for Embabel workspaces through Microsoft
Graph. The pack declares the `microsoft-graph` OAuth2 provider and a focused,
read-only `gateway.outlook` OpenAPI surface for mailbox inspection.

The host owns inbox ingestion. It reads Graph messages for connected Microsoft
accounts, maps them into provider-neutral email models, and emits
`email.thread` signals for the existing email pipeline.

## Gateway

The API entry in `apis/apis.yml` creates the `gateway.outlook` namespace.
Operation ids are:

| Operation ID | Method | Path |
|---|---:|---|
| `user.get` | `GET` | `/me` |
| `messages.list` | `GET` | `/me/messages` |
| `messages.get` | `GET` | `/me/messages/{messageId}` |
| `mailFolders.list` | `GET` | `/me/mailFolders` |
| `mailFolders.get` | `GET` | `/me/mailFolders/{mailFolderId}` |
| `mailFolders.messages.list` | `GET` | `/me/mailFolders/{mailFolderId}/messages` |
| `mailFolders.messages.get` | `GET` | `/me/mailFolders/{mailFolderId}/messages/{messageId}` |

Example calls:

```javascript
await gateway.outlook.userGet({
  "$select": "id,userPrincipalName,displayName,mail",
});

await gateway.outlook.messagesList({
  "$filter": "receivedDateTime ge 2026-06-03T00:00:00Z",
  "$orderby": "receivedDateTime desc",
  "$top": 10,
  "$select": "id,conversationId,receivedDateTime,subject,from,bodyPreview,parentFolderId,webLink",
  Prefer: 'IdType="ImmutableId"',
});

await gateway.outlook.messagesList({
  "$filter": "conversationId eq 'AAMkAG...'",
  "$select": "id,conversationId,receivedDateTime,sentDateTime,subject,from,sender,toRecipients,ccRecipients,replyTo,body,uniqueBody,bodyPreview,parentFolderId,hasAttachments,webLink,categories,inferenceClassification",
  Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"',
});
```

The gateway is read-only. It exposes mailbox inspection methods only.

## OAuth2

The OAuth provider key is `microsoft-graph`.

```yaml
provider: microsoft-graph
auth-url: https://login.microsoftonline.com/common/oauth2/v2.0/authorize
token-url: https://login.microsoftonline.com/common/oauth2/v2.0/token
scopes: openid profile email offline_access User.Read Mail.Read
```

End users authorize the connected account from Settings or through:

```text
/api/v1/auth/oauth2/authorize/microsoft-graph
```

Installation admins register a Microsoft Entra OAuth application for the
assistant deployment. The redirect URI is the host OAuth callback:

```text
https://your-host/api/v1/auth/oauth2/callback
```

For local development:

```text
http://localhost:8042/api/v1/auth/oauth2/callback
```

Client credentials are stored in the host admin OAuth app registry under the
provider key:

```yaml
apps:
  microsoft-graph:
    client-id: "00000000-0000-0000-0000-000000000000"
    client-secret: "..."
```

The connected account identity call is `GET https://graph.microsoft.com/v1.0/me`.
The account id is Graph `id`; the display label is `userPrincipalName`.

## Email Signals

The semantic email provider is `exchange`. It is separate from the OAuth
provider key:

| Field | Value |
|---|---|
| OAuth provider key | `microsoft-graph` |
| `EmailAccount.provider` | `microsoft-graph` |
| `EmailThread.source.provider` | `exchange` |
| `EmailSignal.provider` | `exchange` |
| `EmailSignal.source.kind` | `email` |
| `EmailSignal.source.id` | `{accountId}:{conversationId}` |

Graph `webLink` is the source URL when present.

## Host Fetch Contract

The host lists recent messages with selected fields, groups candidates by
`conversationId`, fetches full conversation messages, sorts locally by
`receivedDateTime` with `sentDateTime` as fallback, and returns at most
`maxThreads` per connected account.

Requests that persist message ids include:

```text
Prefer: IdType="ImmutableId"
```

Requests that read `body` or `uniqueBody` include:

```text
Prefer: outlook.body-content-type="text"
```

When Microsoft Graph returns `@odata.nextLink`, treat it as an opaque URL and
request it directly.
