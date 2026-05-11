import type { EmailAdapter, SendEmailOptions } from 'payload'

import { APIError } from 'payload'

export type MSGraphAdapterArgs = {
  clientId: string
  clientSecret: string
  defaultFromAddress: string
  defaultFromName: string
  /**
   * Mailbox to send from (UPN or object id). Defaults to `defaultFromAddress`.
   * The app must have permission to send as this mailbox.
   */
  sender?: string
  tenantId: string
}

type MSGraphAdapter = EmailAdapter<MSGraphResponse>

type MSGraphError = {
  error: {
    code: string
    message: string
  }
}

type MSGraphResponse = { ok: true } | MSGraphError

/**
 * Email adapter for [Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/user-sendmail) REST API.
 *
 * Uses the OAuth2 client-credentials flow (app-only). The Entra ID app
 * registration needs the `Mail.Send` application permission (admin consent
 * required); optionally scope it to a single mailbox via an
 * ApplicationAccessPolicy.
 */
export const msGraphAdapter = (args: MSGraphAdapterArgs): MSGraphAdapter => {
  const { clientId, clientSecret, defaultFromAddress, defaultFromName, sender, tenantId } = args
  const mailbox = sender || defaultFromAddress
  const getToken = createTokenProvider({ clientId, clientSecret, tenantId })

  const adapter: MSGraphAdapter = () => ({
    name: 'ms-graph-rest',
    defaultFromAddress,
    defaultFromName,
    sendEmail: async (message) => {
      // Map the Payload email options to Microsoft Graph sendMail options
      const sendEmailOptions = mapPayloadEmailToGraphMessage(
        message,
        defaultFromAddress,
        defaultFromName,
      )

      const accessToken = await getToken()

      const res = await fetch(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/sendMail`,
        {
          body: JSON.stringify(sendEmailOptions),
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
          method: 'POST',
        },
      )

      if (res.status === 202) {
        return { ok: true }
      }

      const data = (await res.json().catch(() => ({}))) as Partial<MSGraphError>
      let formattedError = `Error sending email: ${res.status}`
      if (data.error?.code && data.error?.message) {
        formattedError += ` ${data.error.code} - ${data.error.message}`
      }

      throw new APIError(formattedError, res.status)
    },
  })

  return adapter
}

function createTokenProvider(args: {
  clientId: string
  clientSecret: string
  tenantId: string
}): () => Promise<string> {
  const { clientId, clientSecret, tenantId } = args
  let cached: { expiresAt: number; token: string } | null = null

  return async () => {
    const now = Date.now()
    if (cached && cached.expiresAt - 60_000 > now) {
      return cached.token
    }

    const body = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'client_credentials',
      scope: 'https://graph.microsoft.com/.default',
    })

    const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      body,
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      method: 'POST',
    })

    const data = (await res.json().catch(() => ({}))) as {
      access_token?: string
      error?: string
      error_description?: string
      expires_in?: number
    }

    if (!res.ok || !data.access_token) {
      const detail = data.error_description || data.error || `HTTP ${res.status}`
      throw new APIError(`Failed to acquire MS Graph access token: ${detail}`, res.status || 500)
    }

    cached = {
      expiresAt: now + (data.expires_in ?? 3600) * 1000,
      token: data.access_token,
    }
    return cached.token
  }
}

function mapPayloadEmailToGraphMessage(
  message: SendEmailOptions,
  defaultFromAddress: string,
  defaultFromName: string,
): GraphSendMailOptions {
  const html = message.html?.toString()
  const text = message.text?.toString()

  return {
    message: {
      attachments: mapAttachments(message.attachments),
      bccRecipients: mapAddresses(message.bcc),
      body: html
        ? { content: html, contentType: 'HTML' }
        : { content: text ?? '', contentType: 'Text' },
      ccRecipients: mapAddresses(message.cc),
      from: mapFromAddress(message.from, defaultFromName, defaultFromAddress),
      replyTo: mapAddresses(message.replyTo),
      subject: message.subject ?? '',
      toRecipients: mapAddresses(message.to),
    },
    saveToSentItems: false,
  }
}

function mapFromAddress(
  address: SendEmailOptions['from'],
  defaultFromName: string,
  defaultFromAddress: string,
): GraphRecipient {
  if (!address) {
    return { emailAddress: { address: defaultFromAddress, name: defaultFromName } }
  }

  if (typeof address === 'string') {
    return { emailAddress: { address } }
  }

  // Graph's `from` is a single recipient — take the first if an array was passed
  const single = Array.isArray(address) ? address[0] : address
  if (!single) {
    return { emailAddress: { address: defaultFromAddress, name: defaultFromName } }
  }
  if (typeof single === 'string') {
    return { emailAddress: { address: single } }
  }
  return { emailAddress: { address: single.address, name: single.name } }
}

function mapAddresses(addresses: SendEmailOptions['to']): GraphRecipient[] {
  if (!addresses) {
    return []
  }

  if (typeof addresses === 'string') {
    return [{ emailAddress: { address: addresses } }]
  }

  if (Array.isArray(addresses)) {
    return addresses.map((address) =>
      typeof address === 'string'
        ? { emailAddress: { address } }
        : { emailAddress: { address: address.address, name: address.name } },
    )
  }

  return [{ emailAddress: { address: addresses.address, name: addresses.name } }]
}

function mapAttachments(
  attachments: SendEmailOptions['attachments'],
): GraphFileAttachment[] | undefined {
  if (!attachments) {
    return undefined
  }

  return attachments.map((attachment) => {
    if (!attachment.filename || !attachment.content) {
      throw new APIError('Attachment is missing filename or content', 400)
    }

    let contentBytes: string
    if (typeof attachment.content === 'string') {
      contentBytes = Buffer.from(attachment.content).toString('base64')
    } else if (attachment.content instanceof Buffer) {
      contentBytes = attachment.content.toString('base64')
    } else {
      throw new APIError('Attachment content must be a string or a buffer', 400)
    }

    return {
      '@odata.type': '#microsoft.graph.fileAttachment',
      contentBytes,
      contentType: attachment.contentType || 'application/octet-stream',
      name: attachment.filename,
    }
  })
}

type GraphRecipient = {
  emailAddress: {
    address: string
    name?: string
  }
}

type GraphFileAttachment = {
  '@odata.type': '#microsoft.graph.fileAttachment'
  contentBytes: string
  contentType: string
  name: string
}

type GraphMessage = {
  attachments?: GraphFileAttachment[]
  bccRecipients: GraphRecipient[]
  body: {
    content: string
    contentType: 'HTML' | 'Text'
  }
  ccRecipients: GraphRecipient[]
  from: GraphRecipient
  replyTo: GraphRecipient[]
  subject: string
  toRecipients: GraphRecipient[]
}

type GraphSendMailOptions = {
  message: GraphMessage
  saveToSentItems: boolean
}
