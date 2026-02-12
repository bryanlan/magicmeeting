#!/usr/bin/env node
/**
 * pollMonitor.js - Background poller for availability poll responses
 *
 * Runs standalone (not as MCP server) via Windows Task Scheduler every 15 minutes.
 * Checks inbox for poll responses and updates the state file automatically.
 */

import { exec } from 'child_process';
import path from 'path';
import fs from 'fs';
import os from 'os';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Constants
const SCRIPTS_DIR = path.resolve(__dirname, '../scripts');
const SUCCESS_PREFIX = 'SUCCESS:';
const ERROR_PREFIX = 'ERROR:';
const CLAUDE_DIR = path.join(os.homedir(), '.claude');
const POLLS_FILE = path.join(CLAUDE_DIR, 'outlook-polls.json');
const LOG_FILE = path.join(CLAUDE_DIR, 'poll-monitor.log');

/**
 * Log message to file and console
 */
function log(message) {
  const timestamp = new Date().toISOString();
  const logLine = `[${timestamp}] ${message}`;
  console.log(logLine);

  try {
    fs.appendFileSync(LOG_FILE, logLine + '\n');
  } catch (err) {
    console.error('Failed to write to log file:', err.message);
  }
}

/**
 * Execute a VBScript file
 */
async function executeScript(scriptName, params = {}) {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(SCRIPTS_DIR, `${scriptName}.vbs`);
    let command = `chcp 65001 >nul 2>&1 && cscript //NoLogo "${scriptPath}"`;

    for (const [key, value] of Object.entries(params)) {
      if (value !== undefined && value !== null && value !== '') {
        const escapedValue = value.toString().replace(/"/g, '\\"');
        command += ` /${key}:"${escapedValue}"`;
      }
    }

    exec(command, { encoding: 'utf8' }, (error, stdout, stderr) => {
      if (error && !stdout.includes(SUCCESS_PREFIX)) {
        return reject(new Error(`Script execution failed: ${error.message}`));
      }

      if (stdout.includes(ERROR_PREFIX)) {
        const errorMessage = stdout.substring(stdout.indexOf(ERROR_PREFIX) + ERROR_PREFIX.length).trim();
        return reject(new Error(`Script error: ${errorMessage}`));
      }

      if (stdout.includes(SUCCESS_PREFIX)) {
        try {
          const jsonStr = stdout.substring(stdout.indexOf(SUCCESS_PREFIX) + SUCCESS_PREFIX.length).trim();
          const result = JSON.parse(jsonStr);
          return resolve(result);
        } catch (parseError) {
          return reject(new Error(`Failed to parse script output: ${parseError.message}`));
        }
      }

      reject(new Error(`Unexpected script output: ${stdout}`));
    });
  });
}

/**
 * Load polls state from file
 */
function loadPolls() {
  try {
    if (!fs.existsSync(POLLS_FILE)) {
      return { version: '1.0', polls: [] };
    }
    const data = fs.readFileSync(POLLS_FILE, 'utf8');
    return JSON.parse(data);
  } catch (err) {
    log(`Error loading polls file: ${err.message}`);
    return { version: '1.0', polls: [] };
  }
}

/**
 * Save polls state to file
 */
function savePolls(pollsData) {
  try {
    // Ensure .claude directory exists
    if (!fs.existsSync(CLAUDE_DIR)) {
      fs.mkdirSync(CLAUDE_DIR, { recursive: true });
    }
    fs.writeFileSync(POLLS_FILE, JSON.stringify(pollsData, null, 2));
    log('Polls file saved successfully');
  } catch (err) {
    log(`Error saving polls file: ${err.message}`);
  }
}

/**
 * Parse slot numbers from response text
 * Handles formats like:
 * - "Available: 1, 3, 4"
 * - "1, 2, 3"
 * - "Slots 1 and 3 work"
 * - "all" = all slots
 * - "none" = no slots
 */
function parseSlotNumbers(responseText, totalSlots) {
  const text = responseText.toLowerCase();

  // Check for "all" keyword
  if (/\ball\b/.test(text) && !/\bnot all\b/.test(text)) {
    return Array.from({ length: totalSlots }, (_, i) => i);
  }

  // Check for "none" keyword
  if (/\bnone\b/.test(text) || /\bno slots?\b/.test(text)) {
    return [];
  }

  // Extract numbers from text
  const numbers = [];
  const numberMatches = text.match(/\d+/g);

  if (numberMatches) {
    for (const match of numberMatches) {
      const num = parseInt(match, 10);
      // Slot numbers are 1-indexed in the email, convert to 0-indexed
      if (num >= 1 && num <= totalSlots) {
        numbers.push(num - 1);
      }
    }
  }

  // Remove duplicates and sort
  return [...new Set(numbers)].sort((a, b) => a - b);
}

/**
 * Search for poll responses in inbox
 */
async function searchForPollResponses(poll, sentAtDate) {
  try {
    // Search for replies to the poll by subject
    // Extract the key part of the subject to search for
    const subjectToSearch = poll.subject.replace(/^\[.*?\]\s*/, ''); // Remove leading [TAG]

    const result = await executeScript('searchInbox', {
      subjectContains: subjectToSearch,
      limit: 100,
      includeBody: true
    });

    // Filter emails to only those received after poll was sent
    const emails = (result.emails || []).filter(email => {
      // Parse the received date (format: "2/2/2026 03:17 PM")
      const receivedDate = new Date(email.received);
      return receivedDate >= sentAtDate;
    });

    return emails;
  } catch (err) {
    log(`Error searching inbox for poll ${poll.pollId}: ${err.message}`);
    return [];
  }
}

/**
 * Process a single poll
 */
async function processPoll(poll) {
  if (poll.status !== 'pending' && poll.status !== 'partial') {
    return poll;
  }

  log(`Processing poll: ${poll.pollId} (${poll.meetingSubject})`);

  // Convert sentAt to Date object
  const sentDate = new Date(poll.sentAt);

  // Search for responses
  const emails = await searchForPollResponses(poll, sentDate);
  log(`Found ${emails.length} potential responses for poll ${poll.pollId}`);

  let updatedCount = 0;

  for (const email of emails) {
    // Find matching attendee by email
    const senderEmail = email.fromEmail.toLowerCase();
    const attendee = poll.attendees.find(a =>
      a.email.toLowerCase() === senderEmail ||
      senderEmail.includes(a.email.toLowerCase().split('@')[0])
    );

    if (!attendee) {
      log(`No matching attendee found for sender: ${email.fromEmail}`);
      continue;
    }

    if (attendee.responded) {
      // Already processed this attendee's response
      continue;
    }

    // Extract only the reply text (before quoted original)
    // Split on common reply markers
    let replyText = email.body;
    const splitMarkers = [
      '\n\nFrom:',
      '\n\n_____',
      '\n\nOn ',
      '\n\n>',
      '\n\n________________________________'
    ];

    for (const marker of splitMarkers) {
      const idx = replyText.indexOf(marker);
      if (idx > 0) {
        replyText = replyText.substring(0, idx);
        break;
      }
    }

    // Parse slot numbers from response
    const availableSlots = parseSlotNumbers(replyText, poll.proposedSlots.length);

    log(`Attendee ${attendee.email} available for slots: ${availableSlots.map(s => s + 1).join(', ') || 'none'}`);

    // Update attendee
    attendee.responded = true;
    attendee.availableSlots = availableSlots;
    attendee.respondedAt = email.received;
    updatedCount++;
  }

  // Recalculate slot availability counts
  for (const slot of poll.proposedSlots) {
    slot.availableCount = poll.attendees.filter(a =>
      a.responded && a.availableSlots.includes(slot.index)
    ).length;
  }

  // Update poll status
  const respondedCount = poll.attendees.filter(a => a.responded).length;
  if (respondedCount === poll.attendees.length) {
    poll.status = 'complete';
    log(`Poll ${poll.pollId} is now complete - all attendees responded`);
  } else if (respondedCount > 0) {
    poll.status = 'partial';
    log(`Poll ${poll.pollId} has ${respondedCount}/${poll.attendees.length} responses`);
  }

  if (updatedCount > 0) {
    log(`Updated ${updatedCount} attendee responses for poll ${poll.pollId}`);
  }

  return poll;
}

/**
 * Main function
 */
async function main() {
  log('=== Poll Monitor Started ===');

  // Load polls
  const pollsData = loadPolls();

  if (pollsData.polls.length === 0) {
    log('No polls found');
    log('=== Poll Monitor Finished ===');
    return;
  }

  // Filter to active polls
  const activePolls = pollsData.polls.filter(p =>
    p.status === 'pending' || p.status === 'partial'
  );

  if (activePolls.length === 0) {
    log('No active polls to process');
    log('=== Poll Monitor Finished ===');
    return;
  }

  log(`Processing ${activePolls.length} active poll(s)`);

  // Process each active poll
  for (let i = 0; i < pollsData.polls.length; i++) {
    const poll = pollsData.polls[i];
    if (poll.status === 'pending' || poll.status === 'partial') {
      pollsData.polls[i] = await processPoll(poll);
    }
  }

  // Save updated polls
  savePolls(pollsData);

  log('=== Poll Monitor Finished ===');
}

// Run main function
main().catch(err => {
  log(`Fatal error: ${err.message}`);
  process.exit(1);
});
