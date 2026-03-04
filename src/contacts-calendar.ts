import { JmapClient, JmapRequest } from './jmap-client.js';

// Expanded properties list for calendar event queries
const CALENDAR_EVENT_PROPERTIES = [
  'id', 'title', 'description', 'start', 'duration', 'timeZone',
  'showWithoutTime', 'freeBusyStatus', 'location', 'participants',
  'calendarIds', 'recurrenceRules', 'recurrenceOverrides', 'alerts',
];

/**
 * Compute an ISO 8601 duration string from start and end timestamps.
 * E.g. "PT1H30M", "P1DT2H", "PT45M"
 */
function computeDuration(start: string, end: string): string {
  const diffMs = new Date(end).getTime() - new Date(start).getTime();
  if (diffMs <= 0) return 'PT0S';

  const totalMinutes = Math.floor(diffMs / 60000);
  const days = Math.floor(totalMinutes / 1440);
  const hours = Math.floor((totalMinutes % 1440) / 60);
  const minutes = totalMinutes % 60;

  let dur = 'P';
  if (days) dur += `${days}D`;
  if (hours || minutes) {
    dur += 'T';
    if (hours) dur += `${hours}H`;
    if (minutes) dur += `${minutes}M`;
  }
  // Edge case: exact day boundary with no remainder
  if (dur === 'P') dur = 'PT0S';
  return dur;
}

/**
 * Convert a simple participants array to the JMAP participants map.
 */
function convertParticipantsToMap(
  participants: Array<{ email: string; name?: string }>
): Record<string, any> {
  const map: Record<string, any> = {};
  for (const p of participants) {
    map[p.email] = {
      '@type': 'Participant',
      name: p.name || '',
      roles: { attendee: true },
      sendTo: { imip: `mailto:${p.email}` },
    };
  }
  return map;
}

export class ContactsCalendarClient extends JmapClient {

  private async checkContactsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:contacts'];
  }

  private async checkCalendarsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:calendars'];
  }

  async getContacts(limit: number = 50): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();

    // Try CardDAV namespace first, then Fastmail specific
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/query', {
          accountId: session.accountId,
          limit
        }, 'query'],
        ['Contact/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Contact/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[1][1].list;
    } catch (error) {
      // Fallback: try to get contacts using AddressBook methods
      const fallbackRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
        methodCalls: [
          ['AddressBook/get', {
            accountId: session.accountId
          }, 'addressbooks']
        ]
      };

      try {
        const fallbackResponse = await this.makeRequest(fallbackRequest);
        return fallbackResponse.methodResponses[0][1].list || [];
      } catch (fallbackError) {
        throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
      }
    }
  }

  async getContactById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].list[0];
    } catch (error) {
      throw new Error(`Contact access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  async searchContacts(query: string, limit: number = 20): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/query', {
          accountId: session.accountId,
          filter: { text: query },
          limit
        }, 'query'],
        ['Contact/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Contact/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[1][1].list;
    } catch (error) {
      throw new Error(`Contact search not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  async getCalendars(): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['Calendar/get', {
          accountId: session.accountId
        }, 'calendars']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].list;
    } catch (error) {
      // Calendar access might require special permissions
      throw new Error(`Calendar access not supported or requires additional permissions. This may be due to account settings or JMAP scope limitations: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEvents(
    calendarId?: string,
    limit: number = 50,
    options?: { after?: string; before?: string; expandRecurrences?: boolean }
  ): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    const filter: Record<string, any> = {};
    if (calendarId) filter.inCalendar = calendarId;
    if (options?.after) filter.after = options.after;
    if (options?.before) filter.before = options.before;

    const queryArgs: Record<string, any> = {
      accountId: session.accountId,
      filter,
      sort: [{ property: 'start', isAscending: true }],
      limit,
    };
    if (options?.expandRecurrences) {
      queryArgs.expandRecurrences = true;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/query', queryArgs, 'query'],
        ['CalendarEvent/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'CalendarEvent/query', path: '/ids' },
          properties: CALENDAR_EVENT_PROPERTIES
        }, 'events']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[1][1].list;
    } catch (error) {
      throw new Error(`Calendar events access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEventById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/get', {
          accountId: session.accountId,
          ids: [id],
          properties: CALENDAR_EVENT_PROPERTIES
        }, 'event']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].list[0];
    } catch (error) {
      throw new Error(`Calendar event access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string;
    end?: string;
    duration?: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    timeZone?: string;
    showWithoutTime?: boolean;
    freeBusyStatus?: string;
    recurrenceRules?: any[];
    alerts?: Record<string, any>;
  }): Promise<string> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    // Compute duration from end if not provided directly
    const duration = event.duration || (event.end ? computeDuration(event.start, event.end) : 'PT1H');

    const eventObject: Record<string, any> = {
      calendarIds: { [event.calendarId]: true },
      title: event.title,
      start: event.start,
      duration,
    };

    if (event.description) eventObject.description = event.description;
    if (event.location) eventObject.location = event.location;
    if (event.timeZone) eventObject.timeZone = event.timeZone;
    if (event.showWithoutTime !== undefined) eventObject.showWithoutTime = event.showWithoutTime;
    if (event.freeBusyStatus) eventObject.freeBusyStatus = event.freeBusyStatus;
    if (event.recurrenceRules) eventObject.recurrenceRules = event.recurrenceRules;
    if (event.alerts) eventObject.alerts = event.alerts;
    if (event.participants?.length) {
      eventObject.participants = convertParticipantsToMap(event.participants);
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          create: { newEvent: eventObject }
        }, 'createEvent']
      ]
    };

    const response = await this.makeRequest(request);
    const result = response.methodResponses[0][1];

    if (result.notCreated && result.notCreated.newEvent) {
      const err = result.notCreated.newEvent;
      throw new Error(`Failed to create calendar event: ${err.description || JSON.stringify(err)}`);
    }

    return result.created.newEvent.id;
  }

  async updateCalendarEvent(
    eventId: string,
    updates: {
      title?: string;
      description?: string;
      start?: string;
      end?: string;
      duration?: string;
      location?: string;
      participants?: Array<{ email: string; name?: string }>;
      calendarId?: string;
      timeZone?: string;
      showWithoutTime?: boolean;
      freeBusyStatus?: string;
      recurrenceRules?: any[];
      alerts?: Record<string, any>;
    }
  ): Promise<void> {
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled.');
    }

    const session = await this.getSession();

    const patch: Record<string, any> = {};
    if (updates.title !== undefined) patch.title = updates.title;
    if (updates.description !== undefined) patch.description = updates.description;
    if (updates.start !== undefined) patch.start = updates.start;
    if (updates.location !== undefined) patch.location = updates.location;
    if (updates.timeZone !== undefined) patch.timeZone = updates.timeZone;
    if (updates.showWithoutTime !== undefined) patch.showWithoutTime = updates.showWithoutTime;
    if (updates.freeBusyStatus !== undefined) patch.freeBusyStatus = updates.freeBusyStatus;
    if (updates.recurrenceRules !== undefined) patch.recurrenceRules = updates.recurrenceRules;
    if (updates.alerts !== undefined) patch.alerts = updates.alerts;
    if (updates.participants !== undefined) {
      patch.participants = updates.participants.length
        ? convertParticipantsToMap(updates.participants)
        : null;
    }
    if (updates.calendarId !== undefined) {
      patch.calendarIds = { [updates.calendarId]: true };
    }
    // Duration: explicit duration wins, else compute from start+end
    if (updates.duration !== undefined) {
      patch.duration = updates.duration;
    } else if (updates.end !== undefined && updates.start !== undefined) {
      patch.duration = computeDuration(updates.start, updates.end);
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          update: { [eventId]: patch }
        }, 'updateEvent']
      ]
    };

    const response = await this.makeRequest(request);
    const result = response.methodResponses[0][1];

    if (result.notUpdated && result.notUpdated[eventId]) {
      const err = result.notUpdated[eventId];
      throw new Error(`Failed to update calendar event: ${err.description || JSON.stringify(err)}`);
    }
  }

  async deleteCalendarEvent(eventId: string): Promise<void> {
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          destroy: [eventId]
        }, 'deleteEvent']
      ]
    };

    const response = await this.makeRequest(request);
    const result = response.methodResponses[0][1];

    if (result.notDestroyed && result.notDestroyed[eventId]) {
      const err = result.notDestroyed[eventId];
      throw new Error(`Failed to delete calendar event: ${err.description || JSON.stringify(err)}`);
    }
  }
}
