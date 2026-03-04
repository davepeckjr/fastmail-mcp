import { z } from 'zod';

// --- Shared helpers ---

const emailAddress = z.string().email().describe('Email address');
const emailAddressArray = z.array(emailAddress).min(1);
const limitDefault = (def: number) =>
  z.number().int().min(1).max(100).default(def);
const bulkEmailIds = z.array(z.string().min(1)).min(1).max(500)
  .describe('Array of email IDs');

// --- Email tools ---

export const ListMailboxesSchema = z.object({});

export const ListEmailsSchema = z.object({
  mailboxId: z.string().min(1)
    .describe('ID of the mailbox to list emails from (optional, defaults to all)')
    .optional(),
  limit: limitDefault(20)
    .describe('Maximum number of emails to return (default: 20)')
    .optional(),
});

export const GetEmailSchema = z.object({
  emailId: z.string().min(1).describe('ID of the email to retrieve'),
});

export const SendEmailSchema = z.object({
  to: emailAddressArray.describe('Recipient email addresses'),
  cc: z.array(emailAddress).describe('CC email addresses (optional)').optional(),
  bcc: z.array(emailAddress).describe('BCC email addresses (optional)').optional(),
  from: emailAddress.describe('Sender email address (optional, defaults to account primary email)').optional(),
  mailboxId: z.string().min(1).describe('Mailbox ID to save the email to (optional, defaults to Drafts folder)').optional(),
  subject: z.string().describe('Email subject'),
  textBody: z.string().describe('Plain text body (optional)').optional(),
  htmlBody: z.string().describe('HTML body (optional)').optional(),
});

export const SearchEmailsSchema = z.object({
  query: z.string().min(1).describe('Search query string'),
  limit: limitDefault(20)
    .describe('Maximum number of results (default: 20)')
    .optional(),
});

// --- Contact tools ---

export const ListContactsSchema = z.object({
  limit: limitDefault(50)
    .describe('Maximum number of contacts to return (default: 50)')
    .optional(),
});

export const GetContactSchema = z.object({
  contactId: z.string().min(1).describe('ID of the contact to retrieve'),
});

export const SearchContactsSchema = z.object({
  query: z.string().min(1).describe('Search query string'),
  limit: limitDefault(20)
    .describe('Maximum number of results (default: 20)')
    .optional(),
});

// --- Calendar tools ---

export const ListCalendarsSchema = z.object({});

export const CreateCalendarSchema = z.object({
  name: z.string().min(1).describe('Name of the calendar (e.g. "Work", "Personal")'),
  color: z.string().describe('Calendar color as hex string (e.g. "#FF5733"). Optional.').optional(),
  isVisible: z.boolean().describe('Whether the calendar is visible by default (default: true)').optional(),
  isSubscribed: z.boolean().describe('Whether the user is subscribed to this calendar (default: true)').optional(),
});

export const ListCalendarEventsSchema = z.object({
  calendarId: z.string().min(1)
    .describe('ID of the calendar (optional, defaults to all calendars)')
    .optional(),
  limit: limitDefault(50)
    .describe('Maximum number of events to return (default: 50)')
    .optional(),
  after: z.string().describe('Only return events starting at or after this UTC datetime (ISO 8601)').optional(),
  before: z.string().describe('Only return events starting before this UTC datetime (ISO 8601)').optional(),
  expandRecurrences: z.boolean().describe('Expand recurring events into individual occurrences (default: false)').optional(),
});

export const GetCalendarEventSchema = z.object({
  eventId: z.string().min(1).describe('ID of the event to retrieve'),
});

const participantSchema = z.object({
  email: z.string().email(),
  name: z.string().optional(),
});

const freeBusyStatusEnum = z.enum(['free', 'busy', 'tentative', 'unavailable'])
  .describe('Free/busy status: free, busy, tentative, unavailable');

export const CreateCalendarEventSchema = z.object({
  calendarId: z.string().min(1).describe('ID of the calendar to create the event in'),
  title: z.string().min(1).describe('Event title'),
  description: z.string().describe('Event description (optional)').optional(),
  start: z.string().min(1).describe('Start time in ISO 8601 format (e.g. 2025-06-15T10:00:00)'),
  end: z.string().describe('End time in ISO 8601 format. Provide end or duration (at least one required)').optional(),
  duration: z.string().describe('ISO 8601 duration (e.g. PT1H, PT30M, P1D). Alternative to end').optional(),
  location: z.string().describe('Event location (optional)').optional(),
  participants: z.array(participantSchema).describe('Event participants (optional)').optional(),
  timeZone: z.string().describe('IANA time zone (e.g. America/New_York). Optional.').optional(),
  showWithoutTime: z.boolean().describe('All-day event flag (default: false)').optional(),
  freeBusyStatus: freeBusyStatusEnum.optional(),
  recurrenceRules: z.array(z.record(z.unknown()))
    .describe('JMAP RecurrenceRule objects (e.g. [{"@type":"RecurrenceRule","frequency":"weekly"}])')
    .optional(),
  alerts: z.record(z.unknown())
    .describe('JMAP Alert map keyed by alert id')
    .optional(),
});

export const UpdateCalendarEventSchema = z.object({
  eventId: z.string().min(1).describe('ID of the event to update'),
  title: z.string().describe('New title').optional(),
  description: z.string().describe('New description').optional(),
  start: z.string().describe('New start time (ISO 8601)').optional(),
  end: z.string().describe('New end time (ISO 8601). Converted to duration.').optional(),
  duration: z.string().describe('New ISO 8601 duration').optional(),
  location: z.string().describe('New location').optional(),
  participants: z.array(participantSchema).describe('New participants list (replaces existing)').optional(),
  calendarId: z.string().min(1).describe('Move event to this calendar').optional(),
  timeZone: z.string().describe('New IANA time zone').optional(),
  showWithoutTime: z.boolean().describe('All-day event flag').optional(),
  freeBusyStatus: freeBusyStatusEnum.optional(),
  recurrenceRules: z.array(z.record(z.unknown())).describe('New recurrence rules').optional(),
  alerts: z.record(z.unknown()).describe('New alerts map').optional(),
});

export const DeleteCalendarEventSchema = z.object({
  eventId: z.string().min(1).describe('ID of the event to delete'),
});

// --- Identity tools ---

export const ListIdentitiesSchema = z.object({});

// --- Email management tools ---

export const GetRecentEmailsSchema = z.object({
  limit: z.number().int().min(1).max(50).default(10)
    .describe('Number of recent emails to retrieve (default: 10, max: 50)')
    .optional(),
  mailboxName: z.string().default('inbox')
    .describe('Mailbox to search (default: inbox)')
    .optional(),
});

export const MarkEmailReadSchema = z.object({
  emailId: z.string().min(1).describe('ID of the email to mark'),
  read: z.boolean().default(true)
    .describe('true to mark as read, false to mark as unread')
    .optional(),
});

export const DeleteEmailSchema = z.object({
  emailId: z.string().min(1).describe('ID of the email to delete'),
});

export const MoveEmailSchema = z.object({
  emailId: z.string().min(1).describe('ID of the email to move'),
  targetMailboxId: z.string().min(1).describe('ID of the target mailbox'),
});

export const GetEmailAttachmentsSchema = z.object({
  emailId: z.string().min(1).describe('ID of the email'),
});

export const DownloadAttachmentSchema = z.object({
  emailId: z.string().min(1).describe('ID of the email'),
  attachmentId: z.string().min(1).describe('ID of the attachment'),
});

export const AdvancedSearchSchema = z.object({
  query: z.string().describe('Text to search for in subject/body').optional(),
  from: z.string().describe('Filter by sender email').optional(),
  to: z.string().describe('Filter by recipient email').optional(),
  subject: z.string().describe('Filter by subject').optional(),
  hasAttachment: z.boolean().describe('Filter emails with attachments').optional(),
  isUnread: z.boolean().describe('Filter unread emails').optional(),
  mailboxId: z.string().min(1).describe('Search within specific mailbox').optional(),
  after: z.string().describe('Emails after this date (ISO 8601)').optional(),
  before: z.string().describe('Emails before this date (ISO 8601)').optional(),
  limit: limitDefault(50)
    .describe('Maximum results (default: 50)')
    .optional(),
});

export const GetThreadSchema = z.object({
  threadId: z.string().min(1).describe('ID of the thread/conversation'),
});

export const GetMailboxStatsSchema = z.object({
  mailboxId: z.string().min(1)
    .describe('ID of the mailbox (optional, defaults to all mailboxes)')
    .optional(),
});

export const GetAccountSummarySchema = z.object({});

// --- Bulk tools ---

export const BulkMarkReadSchema = z.object({
  emailIds: bulkEmailIds,
  read: z.boolean().default(true)
    .describe('true to mark as read, false as unread')
    .optional(),
});

export const BulkMoveSchema = z.object({
  emailIds: bulkEmailIds,
  targetMailboxId: z.string().min(1).describe('ID of target mailbox'),
});

export const BulkDeleteSchema = z.object({
  emailIds: bulkEmailIds,
});

export const CheckFunctionAvailabilitySchema = z.object({});

export const TestBulkOperationsSchema = z.object({
  dryRun: z.boolean().default(true)
    .describe('If true, only shows what would be done without making changes (default: true)')
    .optional(),
  limit: z.number().int().min(1).max(10).default(3)
    .describe('Number of emails to test with (default: 3, max: 10)')
    .optional(),
});
