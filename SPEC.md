# pack-exchange Specification

Status: draft.

## Purpose

`pack-exchange` provides Microsoft 365 Outlook mail access for Embabel
workspaces through Microsoft Graph. The pack supplies the OAuth2 provider
declaration and a focused, read-only OpenAPI surface for mailbox inspection.

The Embabel Me host owns inbox ingestion. The host reads Microsoft Graph
messages, maps them into provider-neutral email models, and emits
`email.thread` signals for the existing email ingestion pipeline.

## Boundary

`pack-exchange` is a declarative integration pack. It contains:

- `pack.yml`
- `apis/apis.yml`
- `apis/graph-mail-v1.json`
- `README.md`

The pack is responsible for:

- declaring the `microsoft-graph` OAuth2 provider
- declaring delegated Graph scopes needed for signed-in-user mail reads
- exposing typed `gateway.outlook.*` tools backed by a focused OpenAPI spec
- keeping the Graph gateway surface read-only

The host is responsible for:

- registering `ExchangeEmailService : EmailService`
- resolving Microsoft Graph OAuth credentials for each connected account
- calling Microsoft Graph with host HTTP infrastructure
- assembling Graph messages into `EmailThread` and `EmailMessage`
- normalizing message bodies to plain text
- setting provider-correct email signal metadata
- emitting `email.thread` signals through the existing ingestion source

The pack does not implement the ingestion loop, scoring, triage, persistence,
or email signal policies.

## Identity

| Concept | Value |
|---|---|
| Pack name | `exchange` |
| Gateway namespace | `gateway.outlook` |
| OAuth provider key | `microsoft-graph` |
| Semantic email provider | `exchange` |
| Display name | `Microsoft 365 Outlook` |
| Graph API base URL | `https://graph.microsoft.com/v1.0` |

The OAuth provider key and semantic email provider are separate. The
`microsoft-graph` value identifies the connected-account registry entry and
credential store slot. The `exchange` value identifies the email system that
produced an `email.thread` signal.

## OAuth2

The pack declares one OAuth2 provider:

```yaml
provider: microsoft-graph
auth-url: https://login.microsoftonline.com/common/oauth2/v2.0/authorize
token-url: https://login.microsoftonline.com/common/oauth2/v2.0/token
```

The first delegated scope set is:

```text
openid profile email offline_access User.Read Mail.Read
```

`Mail.Read` is required because the host needs full message bodies for thread
assembly and signal extraction. `Mail.ReadBasic` is not sufficient for this
workflow.

The identity endpoint is:

```text
GET https://graph.microsoft.com/v1.0/me
```

Connected account identity uses:

| Graph field | Connected account field |
|---|---|
| `id` | `accountId` |
| `userPrincipalName` | `displayName` |

The host authorization URL is:

```text
/api/v1/auth/oauth2/authorize/microsoft-graph
```

Installation admins register a Microsoft Entra OAuth application for the
assistant deployment and store the client credentials under provider key
`microsoft-graph`. Client credentials do not live in the pack.

## Gateway Surface

The initial gateway namespace is:

```text
gateway.outlook
```

The OpenAPI allowlist exposes only the read operations needed for mailbox
inspection and host email ingestion:

| Operation ID | Method | Path | Purpose |
|---|---:|---|---|
| `user.get` | `GET` | `/me` | Resolve the signed-in Graph user |
| `messages.list` | `GET` | `/me/messages` | List messages across the mailbox |
| `messages.get` | `GET` | `/me/messages/{messageId}` | Read one message |
| `mailFolders.list` | `GET` | `/me/mailFolders` | List top-level mail folders |
| `mailFolders.get` | `GET` | `/me/mailFolders/{mailFolderId}` | Read one folder by id or well-known name |
| `mailFolders.messages.list` | `GET` | `/me/mailFolders/{mailFolderId}/messages` | List messages in one folder |
| `mailFolders.messages.get` | `GET` | `/me/mailFolders/{mailFolderId}/messages/{messageId}` | Read one folder-scoped message |

The gateway does not expose write operations, MIME `$value` reads, attachment
download, delta sync, or webhook subscription management.

## Host Service

Embabel Me registers:

```kotlin
class ExchangeEmailService : EmailService
```

Service identity:

```kotlin
override val providerName: String = "microsoft-graph"
override val displayName: String = "Microsoft 365 Outlook"
```

The service implements:

```kotlin
recentThreads(
    period: Duration,
    maxThreads: Int,
    account: ConnectedAccount,
    workspace: Workspace,
): List<EmailThread>
```

Required behavior:

- resolve OAuth2 config for `microsoft-graph`
- obtain valid OAuth2 credentials for the connected account
- return an empty list when config, credentials, or provider calls are
  unavailable
- log provider failures without throwing through the email ingestion loop
- cap returned candidate threads by `maxThreads` per connected account
- tag returned threads with the source `EmailAccount`
- preserve provider-correct source metadata for `EmailSignal`

Disconnect behavior uses the OAuth provider key:

```text
microsoft-graph
```

## Email Metadata Contract

Outlook threads must carry enough metadata for provider-correct signal
construction.

Normative emitted signal fields:

```text
typeName = email.thread
provider = exchange
source.kind = email
source.id = {accountId}:{conversationId}
source.url = {Graph webLink when available}
```

The connected account tag uses the OAuth provider key:

```text
EmailAccount.provider = microsoft-graph
```

The email thread metadata uses the semantic provider:

```text
EmailThread.source.provider = exchange
EmailThread.source.url = {Graph webLink when available}
```

The host email model should represent thread source metadata explicitly rather
than deriving semantic provider from the OAuth provider string.

## Fetch Algorithm

For each connected Microsoft account, `ExchangeEmailService`:

1. Resolves a bearer token from the `microsoft-graph` credential slot.
2. Lists candidate messages within the requested time window.
3. Requests only the selected fields needed for candidate grouping.
4. Pages through `@odata.nextLink` until it has enough distinct conversations
   or the mailbox has no more candidate pages.
5. Groups candidate messages by `conversationId`.
6. Fetches all selected fields needed to assemble each conversation.
7. Sorts messages locally by `receivedDateTime`, falling back to
   `sentDateTime`.
8. Maps each conversation to `EmailThread`.
9. Returns at most `maxThreads` assembled threads for the account.

Every message list or get request that stores message ids includes:

```text
Prefer: IdType="ImmutableId"
```

Every message read that requests `body` or `uniqueBody` includes:

```text
Prefer: outlook.body-content-type="text"
```

The host treats `@odata.nextLink` as an opaque URL and requests it directly.

## Query Shape

Candidate listing:

```text
GET /me/messages
  ?$filter=receivedDateTime ge {instant}
  &$orderby=receivedDateTime desc
  &$top={pageSize}
  &$select=id,conversationId,subject,from,sender,receivedDateTime,sentDateTime,bodyPreview,parentFolderId,webLink,categories,inferenceClassification,hasAttachments
```

Conversation fetch:

```text
GET /me/messages
  ?$filter=conversationId eq '{conversationId}'
  &$select=id,conversationId,conversationIndex,internetMessageId,subject,from,sender,toRecipients,ccRecipients,replyTo,receivedDateTime,sentDateTime,body,uniqueBody,bodyPreview,parentFolderId,hasAttachments,importance,isRead,webLink,categories,inferenceClassification
```

When `$filter` and `$orderby` are combined, properties in `$orderby` must also
appear in `$filter` in the same order. The host should validate final query
forms against a Microsoft 365 mailbox before relying on them in CI fixtures.

## Message Mapping

Graph message fields map to host email values as follows:

| Graph field | Host value |
|---|---|
| `conversationId` | `EmailThread.threadId` |
| first non-blank `subject` | `EmailThread.subject` |
| `id` | `EmailMessage.messageId` |
| `from.emailAddress` or `sender.emailAddress` | `EmailMessage.from` |
| `toRecipients[].emailAddress` | `EmailMessage.to` |
| `ccRecipients[].emailAddress` | `EmailMessage.cc` |
| `receivedDateTime` or `sentDateTime` | `EmailMessage.receivedAt` |
| `body.content`, `uniqueBody.content`, or `bodyPreview` | `EmailMessage.body` |
| `hasAttachments` | `EmailMessage.hasAttachments` |
| `categories`, `inferenceClassification`, folder hints | `EmailMessage.labels` |
| latest message timestamp | `EmailThread.lastMessageAt` |
| `webLink` | `EmailThread.source.url` |

Participants are the de-duplicated sender, from, to, cc, and reply-to email
addresses across all messages in the thread.

Message bodies stored in `EmailMessage.body` are plain text. If Microsoft Graph
returns HTML despite the text preference, the assembler strips scripts, styles,
tags, and common entities, then collapses whitespace.

## Labels

Outlook classifications become provider-specific labels. The host must not
pretend Outlook categories or folders are Gmail labels.

Recommended label conventions:

```text
OUTLOOK_CATEGORY:{category}
OUTLOOK_INFERENCE:{focused|other}
OUTLOOK_FOLDER:{wellKnownFolderName}
OUTLOOK_FOLDER_ID:{parentFolderId}
```

For the first implementation, useful well-known folder names are:

```text
inbox
sentitems
junkemail
deleteditems
clutter
archive
drafts
```

Folder labels are scoring inputs. They do not replace account metadata or the
semantic signal provider.

## Source URLs

Source URLs prefer Graph `webLink`.

If `webLink` is absent, the service leaves the source URL empty. It does not
synthesize Outlook web URLs from incomplete mailbox state.

## Testing

The host implementation includes focused unit tests for:

- OAuth config lookup uses `microsoft-graph`
- connected accounts are read from provider `microsoft-graph`
- provider failures return an empty list
- `maxThreads` caps returned threads per account
- `conversationId` maps to `EmailThread.threadId`
- Graph message ids are fetched with immutable ids enabled
- Graph `webLink` maps to the email source URL
- emitted `EmailSignal.provider` is `exchange`
- emitted `EmailSignal.source.kind` is `email`
- emitted `EmailSignal.source.id` is `{accountId}:{conversationId}`
- text body content is preserved
- HTML body fallback is normalized to plain text
- senders and recipients become de-duplicated participants
- categories, inference classification, and folder hints become labels

Integration-style tests use deterministic Microsoft Graph fixtures. The test
harness must block unmocked Graph calls.

Fixture coverage:

- `/me` identity response
- recent messages with multiple conversations
- multiple messages with the same `conversationId`
- messages in inbox, sent items, junk email, deleted items, and archive
- text body responses
- HTML body fallback responses
- `@odata.nextLink` paging
- 401 unauthorized
- 403 insufficient permission
- 404 missing message
- 429 throttling with `Retry-After`
- 500 provider error

## Out Of Scope

The first version does not support:

- sending mail
- creating drafts
- replying or forwarding
- moving, deleting, archiving, or marking messages read
- downloading attachments
- reading MIME `$value`
- delta sync
- Graph change notifications
- webhook subscriptions
- application-permission mailbox access
- shared mailbox ingestion
- archive mailbox ingestion

## Validation Checklist

Before shipping the host service, validate these calls against a real Microsoft
365 mailbox or Graph Explorer:

```text
GET /me
```

```text
GET /me/messages?$top=5&$select=id,conversationId,receivedDateTime,subject,from,bodyPreview,webLink
```

```text
GET /me/messages?$filter=receivedDateTime ge {instant}&$orderby=receivedDateTime desc&$top=10&$select=id,conversationId,receivedDateTime,subject,from,bodyPreview,parentFolderId,webLink
```

```text
GET /me/messages?$filter=conversationId eq '{conversationId}'&$select=id,conversationId,receivedDateTime,sentDateTime,subject,from,sender,toRecipients,ccRecipients,replyTo,body,uniqueBody,bodyPreview,parentFolderId,hasAttachments,webLink,categories,inferenceClassification
```

Each validation request that reads or persists message ids includes:

```text
Prefer: IdType="ImmutableId"
```

Each validation request that reads message bodies includes:

```text
Prefer: outlook.body-content-type="text"
```

Validate:

- `conversationId eq` works for the target tenant
- sent messages are included when conversations contain user replies
- body content is available as text or can be normalized from HTML
- `webLink` opens the expected Outlook message for the signed-in account
- folder labels can be derived from `parentFolderId` and known folder lookups

## References

- Microsoft Graph mail API overview: https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0
- List messages: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0
- Message resource: https://learn.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0
- Mail folder resource: https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0
- Get immutable Outlook ids: https://learn.microsoft.com/en-us/graph/outlook-immutable-id
- Microsoft identity platform OAuth2 auth code flow: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow
- Microsoft Graph permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference
