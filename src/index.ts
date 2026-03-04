#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { FastmailAuth, FastmailConfig } from './auth.js';
import { JmapClient, JmapRequest } from './jmap-client.js';
import { ContactsCalendarClient } from './contacts-calendar.js';
import {
  ListMailboxesSchema, ListEmailsSchema, GetEmailSchema, SendEmailSchema,
  SearchEmailsSchema, ListContactsSchema, GetContactSchema, SearchContactsSchema,
  ListCalendarsSchema, CreateCalendarSchema, ListCalendarEventsSchema,
  GetCalendarEventSchema, CreateCalendarEventSchema, UpdateCalendarEventSchema,
  DeleteCalendarEventSchema, ListIdentitiesSchema, GetRecentEmailsSchema,
  MarkEmailReadSchema, DeleteEmailSchema, MoveEmailSchema, GetEmailAttachmentsSchema,
  DownloadAttachmentSchema, AdvancedSearchSchema, GetThreadSchema,
  GetMailboxStatsSchema, GetAccountSummarySchema, BulkMarkReadSchema,
  BulkMoveSchema, BulkDeleteSchema, CheckFunctionAvailabilitySchema,
  TestBulkOperationsSchema,
} from './schemas.js';

const server = new McpServer({
  name: 'fastmail-mcp',
  version: '1.8.0',
});

let jmapClient: JmapClient | null = null;
let contactsCalendarClient: ContactsCalendarClient | null = null;

function resolveEnvValue(...keys: string[]): string | undefined {
  const isPlaceholder = (val: string) => /\$\{[^}]+\}/.test(val.trim());
  for (const key of keys) {
    const raw = process.env[key];
    if (typeof raw === 'string' && raw.trim().length > 0 && !isPlaceholder(raw)) {
      return raw.trim();
    }
  }
  return undefined;
}

function findEnvValue(keys: string[]): { value?: string; key?: string; wasPlaceholder: boolean } {
  const isPlaceholder = (val: string) => /\$\{[^}]+\}/.test(val.trim());
  for (const key of keys) {
    const raw = process.env[key];
    if (typeof raw === 'string' && raw.trim().length > 0) {
      if (isPlaceholder(raw)) {
        return { value: undefined, key, wasPlaceholder: true };
      }
      return { value: raw.trim(), key, wasPlaceholder: false };
    }
  }
  return { value: undefined, key: undefined, wasPlaceholder: false };
}

function initializeClient(): JmapClient {
  if (jmapClient) {
    return jmapClient;
  }

  const tokenInfo = findEnvValue([
    'FASTMAIL_API_TOKEN',
    'USER_CONFIG_FASTMAIL_API_TOKEN',
    'USER_CONFIG_fastmail_api_token',
    'fastmail_api_token',
  ]);
  const apiToken = tokenInfo.value;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }

  const baseInfo = findEnvValue([
    'FASTMAIL_BASE_URL',
    'USER_CONFIG_FASTMAIL_BASE_URL',
    'USER_CONFIG_fastmail_base_url',
    'fastmail_base_url',
  ]);

  const config: FastmailConfig = {
    apiToken,
    baseUrl: baseInfo.value
  };

  const auth = new FastmailAuth(config);
  jmapClient = new JmapClient(auth);
  return jmapClient;
}

function initializeContactsCalendarClient(): ContactsCalendarClient {
  if (contactsCalendarClient) {
    return contactsCalendarClient;
  }

  const tokenInfo = findEnvValue([
    'FASTMAIL_API_TOKEN',
    'USER_CONFIG_FASTMAIL_API_TOKEN',
    'USER_CONFIG_fastmail_api_token',
    'fastmail_api_token',
  ]);
  const apiToken = tokenInfo.value;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }

  const baseInfo = findEnvValue([
    'FASTMAIL_BASE_URL',
    'USER_CONFIG_FASTMAIL_BASE_URL',
    'USER_CONFIG_fastmail_base_url',
    'fastmail_base_url',
  ]);

  const config: FastmailConfig = {
    apiToken,
    baseUrl: baseInfo.value
  };

  const auth = new FastmailAuth(config);
  contactsCalendarClient = new ContactsCalendarClient(auth);
  return contactsCalendarClient;
}

// Helper to wrap handler bodies with consistent error handling
function textResult(text: string) {
  return { content: [{ type: 'text' as const, text }] };
}

function jsonResult(data: unknown) {
  return textResult(JSON.stringify(data, null, 2));
}

// --- Email tools ---

server.tool('list_mailboxes', 'List all mailboxes in the Fastmail account', ListMailboxesSchema.shape, async () => {
  const client = initializeClient();
  const mailboxes = await client.getMailboxes();
  return jsonResult(mailboxes);
});

server.tool('list_emails', 'List emails from a mailbox', ListEmailsSchema.shape, async ({ mailboxId, limit = 20 }) => {
  const client = initializeClient();
  const emails = await client.getEmails(mailboxId, limit);
  return jsonResult(emails);
});

server.tool('get_email', 'Get a specific email by ID', GetEmailSchema.shape, async ({ emailId }) => {
  const client = initializeClient();
  const email = await client.getEmailById(emailId);
  return jsonResult(email);
});

server.tool('send_email', 'Send an email', SendEmailSchema.shape, async (args) => {
  const { to, cc, bcc, from, mailboxId, subject, textBody, htmlBody } = args;
  if (!textBody && !htmlBody) {
    throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
  }
  const client = initializeClient();
  const submissionId = await client.sendEmail({
    to, cc, bcc, from, mailboxId, subject, textBody, htmlBody,
  });
  return textResult(`Email sent successfully. Submission ID: ${submissionId}`);
});

server.tool('search_emails', 'Search emails by subject or content', SearchEmailsSchema.shape, async ({ query, limit = 20 }) => {
  const client = initializeClient();
  const session = await client.getSession();
  const request: JmapRequest = {
    using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
    methodCalls: [
      ['Email/query', {
        accountId: session.accountId,
        filter: { text: query },
        sort: [{ property: 'receivedAt', isAscending: false }],
        limit
      }, 'query'],
      ['Email/get', {
        accountId: session.accountId,
        '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
        properties: ['id', 'subject', 'from', 'to', 'receivedAt', 'preview', 'hasAttachment']
      }, 'emails']
    ]
  };
  const response = await client.makeRequest(request);
  const emails = response.methodResponses[1][1].list;
  return jsonResult(emails);
});

// --- Contact tools ---

server.tool('list_contacts', 'List contacts from the address book', ListContactsSchema.shape, async ({ limit = 50 }) => {
  const contactsClient = initializeContactsCalendarClient();
  const contacts = await contactsClient.getContacts(limit);
  return jsonResult(contacts);
});

server.tool('get_contact', 'Get a specific contact by ID', GetContactSchema.shape, async ({ contactId }) => {
  const contactsClient = initializeContactsCalendarClient();
  const contact = await contactsClient.getContactById(contactId);
  return jsonResult(contact);
});

server.tool('search_contacts', 'Search contacts by name or email', SearchContactsSchema.shape, async ({ query, limit = 20 }) => {
  const contactsClient = initializeContactsCalendarClient();
  const contacts = await contactsClient.searchContacts(query, limit);
  return jsonResult(contacts);
});

// --- Calendar tools ---

server.tool('list_calendars', 'List all calendars', ListCalendarsSchema.shape, async () => {
  const contactsClient = initializeContactsCalendarClient();
  const calendars = await contactsClient.getCalendars();
  return jsonResult(calendars);
});

server.tool('create_calendar', 'Create a new calendar', CreateCalendarSchema.shape, async ({ name, color, isVisible, isSubscribed }) => {
  const contactsClient = initializeContactsCalendarClient();
  const calendarId = await contactsClient.createCalendar({
    name, color, isVisible, isSubscribed,
  });
  return textResult(`Calendar "${name}" created successfully. Calendar ID: ${calendarId}`);
});

server.tool('list_calendar_events', 'List events from a calendar, optionally filtered by date range', ListCalendarEventsSchema.shape, async ({ calendarId, limit = 50, after, before, expandRecurrences }) => {
  const contactsClient = initializeContactsCalendarClient();
  const events = await contactsClient.getCalendarEvents(calendarId, limit, {
    after, before, expandRecurrences,
  });
  return jsonResult(events);
});

server.tool('get_calendar_event', 'Get a specific calendar event by ID', GetCalendarEventSchema.shape, async ({ eventId }) => {
  const contactsClient = initializeContactsCalendarClient();
  const event = await contactsClient.getCalendarEventById(eventId);
  return jsonResult(event);
});

server.tool('create_calendar_event', 'Create a new calendar event', CreateCalendarEventSchema.shape, async (args) => {
  const { calendarId, title, description, start, end, duration,
          location, participants, timeZone, showWithoutTime,
          freeBusyStatus, recurrenceRules, alerts } = args;
  if (!end && !duration) {
    throw new McpError(ErrorCode.InvalidParams, 'Either end or duration must be provided');
  }
  const contactsClient = initializeContactsCalendarClient();
  const eventId = await contactsClient.createCalendarEvent({
    calendarId, title, description, start, end, duration,
    location, participants, timeZone, showWithoutTime,
    freeBusyStatus, recurrenceRules, alerts,
  });
  return textResult(`Calendar event created successfully. Event ID: ${eventId}`);
});

server.tool('update_calendar_event', 'Update an existing calendar event. Only provided fields are changed.', UpdateCalendarEventSchema.shape, async ({ eventId, ...updates }) => {
  const contactsClient = initializeContactsCalendarClient();
  await contactsClient.updateCalendarEvent(eventId, updates);
  return textResult(`Calendar event ${eventId} updated successfully.`);
});

server.tool('delete_calendar_event', 'Delete a calendar event', DeleteCalendarEventSchema.shape, async ({ eventId }) => {
  const contactsClient = initializeContactsCalendarClient();
  await contactsClient.deleteCalendarEvent(eventId);
  return textResult(`Calendar event ${eventId} deleted successfully.`);
});

// --- Identity tools ---

server.tool('list_identities', 'List sending identities (email addresses that can be used for sending)', ListIdentitiesSchema.shape, async () => {
  const client = initializeClient();
  const identities = await client.getIdentities();
  return jsonResult(identities);
});

// --- Email management tools ---

server.tool('get_recent_emails', 'Get the most recent emails from inbox (like top-ten)', GetRecentEmailsSchema.shape, async ({ limit = 10, mailboxName = 'inbox' }) => {
  const client = initializeClient();
  const emails = await client.getRecentEmails(limit, mailboxName);
  return jsonResult(emails);
});

server.tool('mark_email_read', 'Mark an email as read or unread', MarkEmailReadSchema.shape, async ({ emailId, read = true }) => {
  const client = initializeClient();
  await client.markEmailRead(emailId, read);
  return textResult(`Email ${read ? 'marked as read' : 'marked as unread'} successfully`);
});

server.tool('delete_email', 'Delete an email (move to trash)', DeleteEmailSchema.shape, async ({ emailId }) => {
  const client = initializeClient();
  await client.deleteEmail(emailId);
  return textResult('Email deleted successfully (moved to trash)');
});

server.tool('move_email', 'Move an email to a different mailbox', MoveEmailSchema.shape, async ({ emailId, targetMailboxId }) => {
  const client = initializeClient();
  await client.moveEmail(emailId, targetMailboxId);
  return textResult('Email moved successfully');
});

server.tool('get_email_attachments', 'Get list of attachments for an email', GetEmailAttachmentsSchema.shape, async ({ emailId }) => {
  const client = initializeClient();
  const attachments = await client.getEmailAttachments(emailId);
  return jsonResult(attachments);
});

server.tool('download_attachment', 'Download an email attachment', DownloadAttachmentSchema.shape, async ({ emailId, attachmentId }) => {
  const client = initializeClient();
  try {
    const downloadUrl = await client.downloadAttachment(emailId, attachmentId);
    return textResult(`Download URL: ${downloadUrl}`);
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      'Attachment download failed. Verify emailId and attachmentId and try again.'
    );
  }
});

server.tool('advanced_search', 'Advanced email search with multiple criteria', AdvancedSearchSchema.shape, async ({ query, from, to, subject, hasAttachment, isUnread, mailboxId, after, before, limit }) => {
  const client = initializeClient();
  const emails = await client.advancedSearch({
    query, from, to, subject, hasAttachment, isUnread, mailboxId, after, before, limit
  });
  return jsonResult(emails);
});

server.tool('get_thread', 'Get all emails in a conversation thread', GetThreadSchema.shape, async ({ threadId }) => {
  const client = initializeClient();
  try {
    const thread = await client.getThread(threadId);
    return jsonResult(thread);
  } catch (error) {
    throw new McpError(ErrorCode.InternalError, `Thread access failed: ${error instanceof Error ? error.message : String(error)}`);
  }
});

server.tool('get_mailbox_stats', 'Get statistics for a mailbox (unread count, total emails, etc.)', GetMailboxStatsSchema.shape, async ({ mailboxId }) => {
  const client = initializeClient();
  const stats = await client.getMailboxStats(mailboxId);
  return jsonResult(stats);
});

server.tool('get_account_summary', 'Get overall account summary with statistics', GetAccountSummarySchema.shape, async () => {
  const client = initializeClient();
  const summary = await client.getAccountSummary();
  return jsonResult(summary);
});

// --- Bulk tools ---

server.tool('bulk_mark_read', 'Mark multiple emails as read/unread', BulkMarkReadSchema.shape, async ({ emailIds, read = true }) => {
  const client = initializeClient();
  await client.bulkMarkRead(emailIds, read);
  return textResult(`${emailIds.length} emails ${read ? 'marked as read' : 'marked as unread'} successfully`);
});

server.tool('bulk_move', 'Move multiple emails to a mailbox', BulkMoveSchema.shape, async ({ emailIds, targetMailboxId }) => {
  const client = initializeClient();
  await client.bulkMove(emailIds, targetMailboxId);
  return textResult(`${emailIds.length} emails moved successfully`);
});

server.tool('bulk_delete', 'Delete multiple emails (move to trash)', BulkDeleteSchema.shape, async ({ emailIds }) => {
  const client = initializeClient();
  await client.bulkDelete(emailIds);
  return textResult(`${emailIds.length} emails deleted successfully (moved to trash)`);
});

// --- Utility tools ---

server.tool('check_function_availability', 'Check which MCP functions are available based on account permissions', CheckFunctionAvailabilitySchema.shape, async () => {
  const client = initializeClient();
  const session = await client.getSession();

  const availability = {
    email: {
      available: true,
      functions: [
        'list_mailboxes', 'list_emails', 'get_email', 'send_email', 'search_emails',
        'get_recent_emails', 'mark_email_read', 'delete_email', 'move_email',
        'get_email_attachments', 'download_attachment', 'advanced_search', 'get_thread',
        'get_mailbox_stats', 'get_account_summary', 'bulk_mark_read', 'bulk_move', 'bulk_delete'
      ]
    },
    identity: {
      available: true,
      functions: ['list_identities']
    },
    contacts: {
      available: !!session.capabilities['urn:ietf:params:jmap:contacts'],
      functions: ['list_contacts', 'get_contact', 'search_contacts'],
      note: session.capabilities['urn:ietf:params:jmap:contacts'] ?
        'Contacts are available' :
        'Contacts access not available - may require enabling in Fastmail account settings',
      enablementGuide: session.capabilities['urn:ietf:params:jmap:contacts'] ? null : {
        steps: [
          '1. Log into Fastmail web interface',
          '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
          '3. Check if contacts scope is enabled for your API token',
          '4. If not available, you may need to upgrade your Fastmail plan or contact support'
        ],
        documentation: 'https://www.fastmail.com/help/technical/jmap-api.html'
      }
    },
    calendar: {
      available: !!session.capabilities['urn:ietf:params:jmap:calendars'],
      functions: ['list_calendars', 'list_calendar_events', 'get_calendar_event', 'create_calendar_event'],
      note: session.capabilities['urn:ietf:params:jmap:calendars'] ?
        'Calendar is available' :
        'Calendar access not available - may require enabling in Fastmail account settings',
      enablementGuide: session.capabilities['urn:ietf:params:jmap:calendars'] ? null : {
        steps: [
          '1. Log into Fastmail web interface',
          '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
          '3. Check if calendar scope is enabled for your API token',
          '4. If not available, you may need to upgrade your Fastmail plan or contact support'
        ],
        documentation: 'https://www.fastmail.com/help/technical/jmap-api.html'
      }
    },
    capabilities: Object.keys(session.capabilities)
  };

  return jsonResult(availability);
});

server.tool('test_bulk_operations', 'Test bulk operations by finding recent emails and performing safe operations (mark read/unread)', TestBulkOperationsSchema.shape, async ({ dryRun = true, limit = 3 }) => {
  const client = initializeClient();

  const testLimit = Math.min(Math.max(limit, 1), 10);
  const emails = await client.getRecentEmails(testLimit, 'inbox');

  if (emails.length === 0) {
    return textResult('No emails found for bulk operation testing. Try sending yourself a test email first.');
  }

  const emailIds = emails.slice(0, testLimit).map((email: any) => email.id);
  const operations = [
    {
      name: 'bulk_mark_read',
      description: `Mark ${emailIds.length} emails as read`,
      parameters: { emailIds, read: true }
    },
    {
      name: 'bulk_mark_read (undo)',
      description: `Mark ${emailIds.length} emails as unread (undo previous)`,
      parameters: { emailIds, read: false }
    }
  ];

  const results: { testEmails: any[]; operations: any[] } = {
    testEmails: emails.map((email: any) => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.[0]?.email || 'unknown',
      receivedAt: email.receivedAt
    })),
    operations: []
  };

  if (dryRun) {
    results.operations = operations.map(op => ({
      ...op,
      status: 'DRY RUN - Would execute but not actually performed',
      executed: false
    }));

    return textResult(`BULK OPERATIONS TEST (DRY RUN)\n\n${JSON.stringify(results, null, 2)}\n\nTo actually execute the test, set dryRun: false`);
  } else {
    for (const operation of operations) {
      try {
        await client.bulkMarkRead(operation.parameters.emailIds, operation.parameters.read);
        results.operations.push({
          ...operation,
          status: 'SUCCESS',
          executed: true,
          timestamp: new Date().toISOString()
        });

        await new Promise(resolve => setTimeout(resolve, 500));
      } catch (error) {
        results.operations.push({
          ...operation,
          status: 'FAILED',
          executed: false,
          error: error instanceof Error ? error.message : String(error),
          timestamp: new Date().toISOString()
        });
      }
    }

    return textResult(`BULK OPERATIONS TEST (EXECUTED)\n\n${JSON.stringify(results, null, 2)}`);
  }
});

async function runServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Fastmail MCP server running on stdio');
}

runServer().catch(() => {
  console.error('Fastmail MCP server failed to start');
  process.exit(1);
});
