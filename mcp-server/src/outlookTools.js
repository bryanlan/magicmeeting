/**
 * outlookTools.js - Defines MCP tools for Outlook calendar operations
 */

import {
  listEvents,
  createEvent,
  findFreeSlots,
  getAttendeeStatus,
  getCalendars,
  deleteEvent,
  updateEvent,
  getFreeBusy,
  findAvailableRooms,
  sendEmail,
  searchInbox,
  getEmailContent,
  resolveRecipient,
  expandDistributionList,
  addAttendee,
  removeAttendee
} from './scriptRunner.js';

/**
 * Generates current date context for scheduling operations
 * Helps LLM understand temporal relationships accurately
 * @returns {Object} - Date context with today, tomorrow, and relative dates
 */
function getDateContext() {
  const now = new Date();
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  const formatDate = (date) => {
    const dayName = days[date.getDay()];
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    return `${dayName}, ${month}/${day}/${year}`;
  };

  const addDays = (date, daysToAdd) => {
    const result = new Date(date);
    result.setDate(result.getDate() + daysToAdd);
    return result;
  };

  // Find next Monday
  const daysUntilMonday = (8 - now.getDay()) % 7 || 7;
  const nextMonday = addDays(now, daysUntilMonday);

  return {
    today: formatDate(now),
    tomorrow: formatDate(addDays(now, 1)),
    nextMonday: formatDate(nextMonday),
    dayOfWeek: days[now.getDay()]
  };
}

/**
 * Defines the MCP tools for Outlook calendar operations
 * @returns {Object} - Object containing tool definitions
 */
export function defineOutlookTools() {
  return {
    // List calendar events
    list_events: {
      name: 'list_events',
      description: 'List calendar events within a specified date range with pagination support',
      inputSchema: {
        type: 'object',
        properties: {
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          },
          limit: {
            type: 'number',
            description: 'Maximum number of events to return (optional, default 50, max 200)'
          },
          offset: {
            type: 'number',
            description: 'Number of events to skip for pagination (optional, default 0)'
          },
          compact: {
            type: 'boolean',
            description: 'Return compact format with fewer fields (optional, default false)'
          },
          subjectContains: {
            type: 'string',
            description: 'Filter events by subject containing this text (case-insensitive, optional)'
          },
          attendeeEmail: {
            type: 'string',
            description: 'Filter events by attendee email address (partial match, optional)'
          },
          locationContains: {
            type: 'string',
            description: 'Filter events by location containing this text (case-insensitive, optional)'
          }
        },
        required: ['startDate', 'endDate']
      },
      handler: async ({ startDate, endDate, calendar, limit, offset, compact, subjectContains, attendeeEmail, locationContains }) => {
        try {
          const events = await listEvents(startDate, endDate, calendar, limit, offset, compact, subjectContains, attendeeEmail, locationContains);
          const dateCtx = getDateContext();
          return {
            content: [
              {
                type: 'text',
                text: `⏰ DATE CONTEXT: Today is ${dateCtx.today} (${dateCtx.dayOfWeek}), tomorrow is ${dateCtx.tomorrow}, next Monday is ${dateCtx.nextMonday}`
              },
              {
                type: 'text',
                text: JSON.stringify(events, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error listing events: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Create calendar event
    create_event: {
      name: 'create_event',
      description: 'Create a new calendar event or meeting with optional conference room and Teams link',
      inputSchema: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'Event subject/title'
          },
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          startTime: {
            type: 'string',
            description: 'Start time in HH:MM AM/PM format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional, defaults to start date)'
          },
          endTime: {
            type: 'string',
            description: 'End time in HH:MM AM/PM format (optional, defaults to 30 minutes after start time)'
          },
          location: {
            type: 'string',
            description: 'Event location (optional, overridden by room if provided)'
          },
          body: {
            type: 'string',
            description: 'Event description/body (optional)'
          },
          isMeeting: {
            type: 'boolean',
            description: 'Whether this is a meeting with attendees (optional, defaults to false)'
          },
          attendees: {
            type: 'string',
            description: 'Semicolon-separated list of attendee email addresses (optional)'
          },
          room: {
            type: 'string',
            description: 'Conference room email address to book (optional, e.g., cfh2235@microsoft.com)'
          },
          teamsMeeting: {
            type: 'boolean',
            description: 'Whether to create a Teams meeting link (optional, defaults to false)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['subject', 'startDate', 'startTime']
      },
      handler: async (eventDetails) => {
        try {
          const result = await createEvent(eventDetails);
          // Validate that we got a real event ID
          if (!result.eventId || !result.eventId.trim()) {
            return {
              content: [
                {
                  type: 'text',
                  text: 'Error creating event: Event was not created (empty ID returned)'
                }
              ],
              isError: true
            };
          }
          let message = `Event created successfully with ID: ${result.eventId}`;
          if (result.room) {
            message += `\nRoom: ${result.room} (${result.roomEmail})`;
          }
          if (result.teamsMeeting) {
            message += `\nTeams meeting: enabled`;
          }
          return {
            content: [
              {
                type: 'text',
                text: message
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error creating event: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Find free time slots
    find_free_slots: {
      name: 'find_free_slots',
      description: 'Find available time slots in the calendar',
      inputSchema: {
        type: 'object',
        properties: {
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional, defaults to 7 days from start date)'
          },
          duration: {
            type: 'number',
            description: 'Duration in minutes (optional, defaults to 30)'
          },
          workDayStart: {
            type: 'number',
            description: 'Work day start hour (0-23) (optional, defaults to 9)'
          },
          workDayEnd: {
            type: 'number',
            description: 'Work day end hour (0-23) (optional, defaults to 17)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['startDate']
      },
      handler: async ({ startDate, endDate, duration, workDayStart, workDayEnd, calendar }) => {
        try {
          const freeSlots = await findFreeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendar);
          const dateCtx = getDateContext();
          return {
            content: [
              {
                type: 'text',
                text: `⏰ DATE CONTEXT: Today is ${dateCtx.today} (${dateCtx.dayOfWeek}), tomorrow is ${dateCtx.tomorrow}, next Monday is ${dateCtx.nextMonday}`
              },
              {
                type: 'text',
                text: JSON.stringify(freeSlots, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error finding free slots: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Get attendee status
    get_attendee_status: {
      name: 'get_attendee_status',
      description: 'Check the response status of meeting attendees',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['eventId']
      },
      handler: async ({ eventId, calendar }) => {
        try {
          const attendeeStatus = await getAttendeeStatus(eventId, calendar);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(attendeeStatus, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error getting attendee status: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Delete calendar event
    delete_event: {
      name: 'delete_event',
      description: 'Delete a calendar event by its ID',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID to delete'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          }
        },
        required: ['eventId']
      },
      handler: async ({ eventId, calendar }) => {
        try {
          const result = await deleteEvent(eventId, calendar);
          return {
            content: [
              {
                type: 'text',
                text: result.success 
                  ? `Event deleted successfully` 
                  : `Failed to delete event`
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error deleting event: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Update calendar event
    update_event: {
      name: 'update_event',
      description: 'Update an existing calendar event',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID to update'
          },
          subject: {
            type: 'string',
            description: 'New event subject/title (optional)'
          },
          startDate: {
            type: 'string',
            description: 'New start date in MM/DD/YYYY format (optional)'
          },
          startTime: {
            type: 'string',
            description: 'New start time in HH:MM AM/PM format (optional)'
          },
          endDate: {
            type: 'string',
            description: 'New end date in MM/DD/YYYY format (optional)'
          },
          endTime: {
            type: 'string',
            description: 'New end time in HH:MM AM/PM format (optional)'
          },
          location: {
            type: 'string',
            description: 'New event location (optional)'
          },
          body: {
            type: 'string',
            description: 'New event description/body (optional)'
          },
          calendar: {
            type: 'string',
            description: 'Calendar name (optional)'
          },
          originalStart: {
            type: 'string',
            description: 'For recurring meetings: the original start date/time of the occurrence to modify (format: MM/DD/YYYY HH:MM AM/PM). Required when modifying a single instance of a recurring series.'
          },
          sendUpdate: {
            type: 'boolean',
            description: 'Whether to send meeting update to attendees (optional, defaults to false)'
          },
          updateSeries: {
            type: 'boolean',
            description: 'For recurring meetings: set to true to update the ENTIRE series (all occurrences), not just one instance. When true, originalStart is not needed.'
          }
        },
        required: ['eventId']
      },
      handler: async (args) => {
        const { eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar, originalStart, sendUpdate, updateSeries } = args;

        try {
          const result = await updateEvent(eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar, originalStart, sendUpdate, updateSeries);
          return {
            content: [
              {
                type: 'text',
                text: result.success
                  ? `Event updated successfully`
                  : `Failed to update event`
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error updating event: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Get calendars
    get_calendars: {
      name: 'get_calendars',
      description: 'List available calendars',
      inputSchema: {
        type: 'object',
        properties: {}
      },
      handler: async () => {
        try {
          const calendars = await getCalendars();
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(calendars, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error getting calendars: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Get free/busy for multiple attendees
    get_free_busy: {
      name: 'get_free_busy',
      description: 'Get free/busy information for multiple attendees and find common available time slots',
      inputSchema: {
        type: 'object',
        properties: {
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional, defaults to 7 days from start date)'
          },
          attendees: {
            type: 'string',
            description: 'Semicolon-separated list of attendee email addresses'
          },
          duration: {
            type: 'number',
            description: 'Meeting duration in minutes (optional, defaults to 30)'
          }
        },
        required: ['startDate', 'attendees']
      },
      handler: async ({ startDate, endDate, attendees, duration }) => {
        try {
          const freeBusyInfo = await getFreeBusy(startDate, endDate, attendees, duration);
          const dateCtx = getDateContext();
          return {
            content: [
              {
                type: 'text',
                text: `⏰ DATE CONTEXT: Today is ${dateCtx.today} (${dateCtx.dayOfWeek}), tomorrow is ${dateCtx.tomorrow}, next Monday is ${dateCtx.nextMonday}`
              },
              {
                type: 'text',
                text: JSON.stringify(freeBusyInfo, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error getting free/busy information: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Find available conference rooms
    find_available_rooms: {
      name: 'find_available_rooms',
      description: 'Find available conference rooms by building, floor, time range, and capacity',
      inputSchema: {
        type: 'object',
        properties: {
          building: {
            type: 'string',
            description: 'Building name (e.g., "STUDIO E", "50", "Building 40")'
          },
          floor: {
            type: 'number',
            description: 'Floor number (optional, derived from 1st digit of room number)'
          },
          startDate: {
            type: 'string',
            description: 'Start date in MM/DD/YYYY format'
          },
          startTime: {
            type: 'string',
            description: 'Start time in HH:MM AM/PM format'
          },
          endDate: {
            type: 'string',
            description: 'End date in MM/DD/YYYY format (optional, defaults to startDate)'
          },
          endTime: {
            type: 'string',
            description: 'End time in HH:MM AM/PM format'
          },
          capacity: {
            type: 'number',
            description: 'Minimum room capacity (number of people)'
          }
        },
        required: ['building', 'startDate', 'startTime', 'endTime', 'capacity']
      },
      handler: async ({ building, floor, startDate, startTime, endDate, endTime, capacity }) => {
        try {
          const rooms = await findAvailableRooms(building, floor, startDate, startTime, endDate, endTime, capacity);
          const dateCtx = getDateContext();
          return {
            content: [
              {
                type: 'text',
                text: `⏰ DATE CONTEXT: Today is ${dateCtx.today} (${dateCtx.dayOfWeek}), tomorrow is ${dateCtx.tomorrow}, next Monday is ${dateCtx.nextMonday}`
              },
              {
                type: 'text',
                text: JSON.stringify(rooms, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error finding available rooms: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Send email
    send_email: {
      name: 'send_email',
      description: 'Send an email via Outlook with optional HTML formatting',
      inputSchema: {
        type: 'object',
        properties: {
          to: {
            type: 'string',
            description: 'Recipient email address(es), semicolon-separated for multiple'
          },
          cc: {
            type: 'string',
            description: 'CC email address(es), semicolon-separated (optional)'
          },
          subject: {
            type: 'string',
            description: 'Email subject line'
          },
          body: {
            type: 'string',
            description: 'Email body content (HTML or plain text)'
          },
          isHtml: {
            type: 'boolean',
            description: 'Whether body is HTML (optional, defaults to true)'
          }
        },
        required: ['to', 'subject']
      },
      handler: async ({ to, cc, subject, body, isHtml }) => {
        try {
          const result = await sendEmail(to, cc, subject, body, isHtml);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error sending email: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Search inbox
    search_inbox: {
      name: 'search_inbox',
      description: 'Search emails matching criteria (subject, sender, body, date range). Supports partial matching.',
      inputSchema: {
        type: 'object',
        properties: {
          subjectContains: {
            type: 'string',
            description: 'Search for emails with subject containing this text (partial match)'
          },
          fromAddresses: {
            type: 'string',
            description: 'Filter by sender name or email (partial match), semicolon-separated for multiple'
          },
          toAddresses: {
            type: 'string',
            description: 'Filter by recipient name or email (partial match), semicolon-separated for multiple. Useful for searching sent folder.'
          },
          bodyContains: {
            type: 'string',
            description: 'Search for emails with body containing these words (space-separated, all must match)'
          },
          folder: {
            type: 'string',
            description: 'Folder to search: "inbox" (default), "sent", "drafts", or subfolder name'
          },
          receivedAfter: {
            type: 'string',
            description: 'Only emails received after this date (MM/DD/YYYY format)'
          },
          receivedBefore: {
            type: 'string',
            description: 'Only emails received before this date (MM/DD/YYYY format)'
          },
          limit: {
            type: 'number',
            description: 'Maximum number of emails to return (default 50, max 200)'
          },
          includeBody: {
            type: 'boolean',
            description: 'Include email body in results (default false, auto-true if bodyContains used)'
          }
        },
        required: []
      },
      handler: async ({ subjectContains, fromAddresses, toAddresses, bodyContains, folder, receivedAfter, receivedBefore, limit, includeBody }) => {
        try {
          const result = await searchInbox(subjectContains, fromAddresses, toAddresses, bodyContains, folder, receivedAfter, receivedBefore, limit, includeBody);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error searching inbox: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Get email content
    get_email_content: {
      name: 'get_email_content',
      description: 'Get full email content by its EntryID',
      inputSchema: {
        type: 'object',
        properties: {
          emailId: {
            type: 'string',
            description: 'Email EntryID (from search_inbox results)'
          }
        },
        required: ['emailId']
      },
      handler: async ({ emailId }) => {
        try {
          const result = await getEmailContent(emailId);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error getting email content: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Resolve recipient from GAL
    resolve_recipient: {
      name: 'resolve_recipient',
      description: 'Resolve a name, alias, or email address to a GAL (Global Address List) entry. Returns details including type (user vs distribution list), email, and alias.',
      inputSchema: {
        type: 'object',
        properties: {
          query: {
            type: 'string',
            description: 'Name, alias, or email address to resolve (e.g., "John Smith", "jsmith", or "jsmith@company.com")'
          }
        },
        required: ['query']
      },
      handler: async ({ query }) => {
        try {
          const result = await resolveRecipient(query);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error resolving recipient: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Expand distribution list
    expand_distribution_list: {
      name: 'expand_distribution_list',
      description: 'Expand a distribution list (DL) to get its members. Can optionally expand nested DLs recursively.',
      inputSchema: {
        type: 'object',
        properties: {
          name: {
            type: 'string',
            description: 'Name, alias, or email of the distribution list to expand'
          },
          recursive: {
            type: 'boolean',
            description: 'Whether to recursively expand nested distribution lists (default: false)'
          },
          maxDepth: {
            type: 'number',
            description: 'Maximum recursion depth when recursive is true (default: 3)'
          }
        },
        required: ['name']
      },
      handler: async ({ name, recursive, maxDepth }) => {
        try {
          const result = await expandDistributionList(name, recursive, maxDepth);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error expanding distribution list: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Add attendee to existing meeting
    add_attendee: {
      name: 'add_attendee',
      description: 'Add an attendee to an existing meeting and send them an invite. Only works if you are the meeting organizer.',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID of the meeting to add attendee to'
          },
          attendee: {
            type: 'string',
            description: 'Email address of the attendee to add'
          },
          type: {
            type: 'string',
            description: 'Attendee type: "required" (default), "optional", or "resource"'
          },
          sendUpdate: {
            type: 'boolean',
            description: 'Whether to send meeting invite to the new attendee (default: true)'
          }
        },
        required: ['eventId', 'attendee']
      },
      handler: async ({ eventId, attendee, type, sendUpdate }) => {
        try {
          const result = await addAttendee(eventId, attendee, type, sendUpdate);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error adding attendee: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    },

    // Remove attendee from meeting
    remove_attendee: {
      name: 'remove_attendee',
      description: 'Remove an attendee from an existing meeting. Only works if you are the meeting organizer. For recurring meetings, use originalStart to target a specific occurrence, or updateSeries=true to update all occurrences.',
      inputSchema: {
        type: 'object',
        properties: {
          eventId: {
            type: 'string',
            description: 'Event ID of the meeting to remove attendee from'
          },
          attendee: {
            type: 'string',
            description: 'Email address or name of the attendee to remove (partial match supported)'
          },
          sendUpdate: {
            type: 'boolean',
            description: 'Whether to send meeting update to attendees (default: true)'
          },
          originalStart: {
            type: 'string',
            description: 'For recurring meetings: the original start date/time of the occurrence to modify (format: MM/DD/YYYY HH:MM AM/PM)'
          },
          updateSeries: {
            type: 'boolean',
            description: 'For recurring meetings: set to true to update the ENTIRE series (all occurrences)'
          }
        },
        required: ['eventId', 'attendee']
      },
      handler: async ({ eventId, attendee, sendUpdate, originalStart, updateSeries }) => {
        try {
          const result = await removeAttendee(eventId, attendee, sendUpdate, originalStart, updateSeries);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(result, null, 2)
              }
            ]
          };
        } catch (error) {
          return {
            content: [
              {
                type: 'text',
                text: `Error removing attendee: ${error.message}`
              }
            ],
            isError: true
          };
        }
      }
    }
  };
}
