# Microsoft Graph Email Adapter for Payload CMS

This [Payload CMS](https://payloadcms.com) email adapter allows you to send emails using the [Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/user-sendmail) REST API.

## Installation

```sh
npm install github:tokiwi/payload-email-msgraph
```

## Usage

- Register an app in Microsoft Entra ID
- Grant it the `Mail.Send` **application** permission (admin consent required)
- Create a client secret for the app
- (Optional) Scope the app to a single mailbox via an [ApplicationAccessPolicy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)
- Set the credentials as environment variables
- Configure your Payload config

```ts
// payload.config.ts
import { msGraphAdapter } from 'payload-email-msgraph'

export default buildConfig({
  email: msGraphAdapter({
    tenantId: process.env.MS_GRAPH_TENANT_ID || '',
    clientId: process.env.MS_GRAPH_CLIENT_ID || '',
    clientSecret: process.env.MS_GRAPH_CLIENT_SECRET || '',
    sender: process.env.MS_GRAPH_SENDER, // optional; defaults to defaultFromAddress
    defaultFromAddress: 'noreply@example.com',
    defaultFromName: 'My App',
  }),
})
```
