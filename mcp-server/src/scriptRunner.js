/**
 * scriptRunner.js - Handles execution of VBScript files and processes their output
 */

import { spawn } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Constants
const SCRIPTS_DIR = path.resolve(__dirname, '../scripts');
const SUCCESS_PREFIX = 'SUCCESS:';
const ERROR_PREFIX = 'ERROR:';

/**
 * Executes a VBScript file with the given parameters
 * Uses spawn with argument array to avoid cmd.exe shell parsing issues.
 * Multi-line content (body field) is passed via temp file to handle newlines safely
 * (stdin doesn't work reliably with windowsHide:true for cscript.exe).
 * @param {string} scriptName - Name of the script file (without .vbs extension)
 * @param {Object} params - Parameters to pass to the script
 * @returns {Promise<Object>} - Promise that resolves with the script output
 */
export async function executeScript(scriptName, params = {}) {
  const scriptPath = path.join(SCRIPTS_DIR, `${scriptName}.vbs`);

  // Build argument array (avoids cmd.exe shell parsing issues)
  const args = ['//NoLogo', scriptPath];
  let bodyTempFile = null;

  // Add parameters as command-line arguments
  for (const [key, value] of Object.entries(params)) {
    // Include parameter if it has a meaningful value
    const shouldInclude = (typeof value === 'boolean')
      ? value === true
      : (value !== undefined && value !== null && value !== '');

    if (shouldInclude) {
      // Pass 'body' via temp file to handle newlines safely
      if (key === 'body') {
        bodyTempFile = path.join(os.tmpdir(), `mcp_body_${crypto.randomUUID()}.txt`);
        fs.writeFileSync(bodyTempFile, value.toString(), 'utf8');
        args.push(`/bodyFile:${bodyTempFile}`);
      } else {
        // Regular parameters go on command line
        args.push(`/${key}:${value.toString()}`);
      }
    }
  }

  return new Promise((resolve, reject) => {
    // Debug: log the arguments being passed to a file
    const debugPath = path.join(os.tmpdir(), 'mcp_debug.log');
    fs.appendFileSync(debugPath, `[${new Date().toISOString()}] Script: ${scriptName}, Args: ${JSON.stringify(args)}\n`);

    // Spawn cscript directly (no shell)
    const child = spawn('cscript.exe', args, {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';

    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (data) => { stdout += data; });
    child.stderr.on('data', (data) => { stderr += data; });

    child.on('error', (error) => {
      // Clean up temp file on error
      if (bodyTempFile) { try { fs.unlinkSync(bodyTempFile); } catch {} }
      reject(new Error(`Script execution failed: ${error.message}`));
    });

    child.on('close', (code) => {
      // Clean up temp file
      if (bodyTempFile) { try { fs.unlinkSync(bodyTempFile); } catch {} }

      // Check for script errors
      if (stdout.includes(ERROR_PREFIX)) {
        const errorMessage = stdout.substring(stdout.indexOf(ERROR_PREFIX) + ERROR_PREFIX.length).trim();
        return reject(new Error(`Script error: ${errorMessage}`));
      }

      // Process successful output
      if (stdout.includes(SUCCESS_PREFIX)) {
        try {
          const jsonStr = stdout.substring(stdout.indexOf(SUCCESS_PREFIX) + SUCCESS_PREFIX.length).trim();
          const result = JSON.parse(jsonStr);
          return resolve(result);
        } catch (parseError) {
          return reject(new Error(`Failed to parse script output: ${parseError.message}`));
        }
      }

      // If we get here, something unexpected happened
      if (code !== 0) {
        reject(new Error(`Script exited with code ${code}\n${stderr}`));
      } else {
        reject(new Error(`Unexpected script output: ${stdout}`));
      }
    });

    // Close stdin immediately (not using it)
    child.stdin.end();
  });
}

/**
 * Lists calendar events within a specified date range
 * @param {string} startDate - Start date in MM/DD/YYYY format
 * @param {string} endDate - End date in MM/DD/YYYY format (optional)
 * @param {string} calendar - Calendar name (optional)
 * @param {number} limit - Maximum number of events to return (optional, default 50, max 200)
 * @param {number} offset - Number of events to skip (optional, default 0)
 * @param {boolean} compact - Return compact format with fewer fields (optional, default false)
 * @param {string} subjectContains - Filter by subject text (optional)
 * @param {string} attendeeEmail - Filter by attendee email (optional)
 * @param {string} locationContains - Filter by location text (optional)
 * @returns {Promise<Object>} - Promise that resolves with paginated events
 */
export async function listEvents(startDate, endDate, calendar, limit, offset, compact, subjectContains, attendeeEmail, locationContains) {
  return executeScript('listEvents', { startDate, endDate, calendar, limit, offset, compact, subjectContains, attendeeEmail, locationContains });
}

/**
 * Creates a new calendar event
 * @param {Object} eventDetails - Event details including optional recurrence
 * @param {string} eventDetails.subject - Event subject
 * @param {string} eventDetails.startDate - Start date MM/DD/YYYY
 * @param {string} eventDetails.startTime - Start time HH:MM AM/PM
 * @param {string} eventDetails.endDate - End date (optional)
 * @param {string} eventDetails.endTime - End time (optional)
 * @param {string} eventDetails.location - Location (optional)
 * @param {string} eventDetails.body - Description (optional)
 * @param {boolean} eventDetails.isMeeting - Is meeting with attendees
 * @param {string} eventDetails.attendees - Semicolon-separated emails
 * @param {string} eventDetails.room - Room email (optional)
 * @param {boolean} eventDetails.teamsMeeting - Create Teams link
 * @param {string} eventDetails.recurrenceType - none/daily/weekly/monthly/yearly
 * @param {number} eventDetails.recurrenceInterval - Every N periods
 * @param {string} eventDetails.recurrenceDays - Comma-separated days for weekly (e.g., "monday,friday")
 * @param {string} eventDetails.recurrenceEndDate - End date for series MM/DD/YYYY
 * @param {number} eventDetails.recurrenceOccurrences - Number of occurrences
 * @returns {Promise<Object>} - Promise that resolves with the created event ID
 */
export async function createEvent(eventDetails) {
  return executeScript('createEvent', eventDetails);
}

/**
 * Finds free time slots in the calendar
 * @param {string} startDate - Start date in MM/DD/YYYY format
 * @param {string} endDate - End date in MM/DD/YYYY format (optional)
 * @param {number} duration - Duration in minutes (optional)
 * @param {number} workDayStart - Work day start hour (0-23) (optional)
 * @param {number} workDayEnd - Work day end hour (0-23) (optional)
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Array>} - Promise that resolves with an array of free time slots
 */
export async function findFreeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendar) {
  return executeScript('findFreeSlots', {
    startDate,
    endDate,
    duration,
    workDayStart,
    workDayEnd,
    calendar
  });
}

/**
 * Gets the response status of meeting attendees
 * @param {string} eventId - Event ID
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Object>} - Promise that resolves with meeting details and attendee status
 */
export async function getAttendeeStatus(eventId, calendar) {
  return executeScript('getAttendeeStatus', { eventId, calendar });
}

/**
 * Deletes a calendar event by its ID
 * @param {string} eventId - Event ID
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Object>} - Promise that resolves with the deletion result
 */
export async function deleteEvent(eventId, calendar) {
  return executeScript('deleteEvent', { eventId, calendar });
}

/**
 * Cancels a meeting with an optional custom cancellation message
 * Only works if you are the meeting organizer
 * For recurring meetings: use occurrenceStart for one instance, or cancelSeries for entire series
 * @param {string} eventId - Event ID (series master ID for recurring meetings)
 * @param {string} occurrenceStart - For recurring: occurrence start datetime (MM/DD/YYYY HH:MM AM/PM)
 * @param {boolean} cancelSeries - For recurring: cancel entire series (not just one instance)
 * @param {string} comment - Custom cancellation message (optional)
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Object>} - Promise that resolves with the cancellation result
 */
export async function cancelEvent(eventId, occurrenceStart, cancelSeries, comment, calendar) {
  return executeScript('cancelEvent', { eventId, occurrenceStart, cancelSeries, body: comment, calendar });
}

/**
 * Updates an existing calendar event
 * @param {string} eventId - Event ID to update
 * @param {string} subject - New subject (optional)
 * @param {string} startDate - New start date in MM/DD/YYYY format (optional)
 * @param {string} startTime - New start time in HH:MM AM/PM format (optional)
 * @param {string} endDate - New end date in MM/DD/YYYY format (optional)
 * @param {string} endTime - New end time in HH:MM AM/PM format (optional)
 * @param {string} location - New location (optional)
 * @param {string} body - New body/description (optional)
 * @param {string} calendar - Calendar name (optional)
 * @param {string} originalStart - For recurring: original occurrence start (MM/DD/YYYY HH:MM AM/PM)
 * @param {boolean} sendUpdate - Whether to send meeting update to attendees
 * @param {boolean} updateSeries - For recurring: update entire series instead of single instance
 * @returns {Promise<Object>} - Promise that resolves with the update result
 */
export async function updateEvent(eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar, originalStart, sendUpdate, updateSeries) {
  return executeScript('updateEvent', {
    eventId,
    subject,
    startDate,
    startTime,
    endDate,
    endTime,
    location,
    body,
    calendar,
    originalStart,
    sendUpdate,
    updateSeries
  });
}

/**
 * Lists available calendars
 * @returns {Promise<Array>} - Promise that resolves with an array of calendars
 */
export async function getCalendars() {
  return executeScript('getCalendars');
}

/**
 * Gets free/busy information for multiple attendees and finds common free slots
 * @param {string} startDate - Start date in MM/DD/YYYY format
 * @param {string} endDate - End date in MM/DD/YYYY format (optional)
 * @param {string} attendees - Semicolon-separated list of attendee email addresses
 * @param {number} duration - Duration in minutes (optional, defaults to 30)
 * @returns {Promise<Object>} - Promise that resolves with free/busy info and common slots
 */
export async function getFreeBusy(startDate, endDate, attendees, duration) {
  return executeScript('getFreeBusy', {
    startDate,
    endDate,
    attendees,
    duration
  });
}

/**
 * Finds available conference rooms matching criteria
 * @param {string} building - Building name (e.g., "STUDIO E", "50")
 * @param {number} floor - Floor number (optional, derived from 1st digit of room#)
 * @param {string} startDate - Start date in MM/DD/YYYY format
 * @param {string} startTime - Start time in HH:MM AM/PM format
 * @param {string} endDate - End date in MM/DD/YYYY format (optional, defaults to startDate)
 * @param {string} endTime - End time in HH:MM AM/PM format
 * @param {number} capacity - Minimum room capacity required
 * @returns {Promise<Object>} - Promise that resolves with available rooms
 */
export async function findAvailableRooms(building, floor, startDate, startTime, endDate, endTime, capacity) {
  return executeScript('findAvailableRooms', {
    building,
    floor,
    startDate,
    startTime,
    endDate,
    endTime,
    capacity
  });
}

/**
 * Sends an email via Outlook
 * @param {string} to - Recipient email address(es), semicolon-separated
 * @param {string} cc - CC email address(es), semicolon-separated (optional)
 * @param {string} subject - Email subject line
 * @param {string} body - Email body content
 * @param {boolean} isHtml - Whether body is HTML (defaults to true)
 * @returns {Promise<Object>} - Promise that resolves with send result
 */
export async function sendEmail(to, cc, subject, body, isHtml) {
  // Body is passed via stdin by executeScript to handle newlines safely
  return executeScript('sendEmail', {
    to,
    cc,
    subject,
    body,
    isHtml
  });
}

/**
 * Searches emails for matching criteria
 * @param {string} subjectContains - Search for subject containing this text
 * @param {string} fromAddresses - Filter by sender addresses (semicolon-separated)
 * @param {string} bodyContains - Search body for these words (space-separated, all must match)
 * @param {string} folder - Folder to search: inbox, sent, drafts, or subfolder name
 * @param {string} receivedAfter - Only emails after this date (MM/DD/YYYY)
 * @param {string} receivedBefore - Only emails before this date (MM/DD/YYYY)
 * @param {number} limit - Maximum number of emails to return (default 50)
 * @param {boolean} includeBody - Include email body in results (default false)
 * @returns {Promise<Object>} - Promise that resolves with search results
 */
export async function searchInbox(subjectContains, fromAddresses, toAddresses, bodyContains, folder, receivedAfter, receivedBefore, limit, includeBody) {
  return executeScript('searchInbox', {
    subjectContains,
    fromAddresses,
    toAddresses,
    bodyContains,
    folder,
    receivedAfter,
    receivedBefore,
    limit,
    includeBody
  });
}

/**
 * Gets full email content by EntryID
 * @param {string} emailId - Email EntryID
 * @returns {Promise<Object>} - Promise that resolves with email content
 */
export async function getEmailContent(emailId) {
  return executeScript('getEmailContent', { emailId });
}

/**
 * Resolves a recipient name/alias/email to GAL entry details
 * @param {string} query - Name, alias, or email to resolve
 * @returns {Promise<Object>} - Promise that resolves with recipient details
 */
export async function resolveRecipient(query) {
  return executeScript('resolveRecipient', { query });
}

/**
 * Expands a distribution list to get its members
 * @param {string} name - Name, alias, or email of the distribution list
 * @param {boolean} recursive - Whether to expand nested DLs
 * @param {number} maxDepth - Maximum recursion depth (default 3)
 * @returns {Promise<Object>} - Promise that resolves with DL members
 */
export async function expandDistributionList(name, recursive, maxDepth) {
  return executeScript('expandDistributionList', { name, recursive, maxDepth });
}

/**
 * Adds an attendee to an existing meeting and sends update
 * @param {string} eventId - Event ID of the meeting
 * @param {string} attendee - Email address of the attendee to add
 * @param {string} type - Attendee type: required, optional, or resource
 * @param {boolean} sendUpdate - Whether to send meeting update to the new attendee
 * @returns {Promise<Object>} - Promise that resolves with result
 */
export async function addAttendee(eventId, attendee, type, sendUpdate) {
  return executeScript('addAttendee', { eventId, attendee, type, sendUpdate });
}

/**
 * Removes an attendee from an existing meeting
 * @param {string} eventId - Event ID of the meeting
 * @param {string} attendee - Email address or name of the attendee to remove
 * @param {boolean} sendUpdate - Whether to send meeting update to attendees
 * @param {string} originalStart - For recurring: original occurrence start (MM/DD/YYYY HH:MM AM/PM)
 * @param {boolean} updateSeries - For recurring: update entire series instead of single instance
 * @returns {Promise<Object>} - Promise that resolves with result
 */
export async function removeAttendee(eventId, attendee, sendUpdate, originalStart, updateSeries) {
  return executeScript('removeAttendee', { eventId, attendee, sendUpdate, originalStart, updateSeries });
}
