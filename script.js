// ============================================================================
// PLACEMENT EMAIL TO CALENDAR AUTOMATION SCRIPT
// ============================================================================
// This Google Apps Script automatically processes placement-related emails
// and creates calendar events for interviews, assessments, and deadlines.
//
// SETUP INSTRUCTIONS:
// 1. Fill in your personal details in the CONFIG object below
// 2. Add your college placement email addresses to ALLOWED_SENDERS
// 3. (Optional) Create a dedicated Google Calendar and add its ID to CALENDAR_ID
// 4. Deploy this script in Google Apps Script (script.google.com)
// 5. Set up a time-based trigger to run myFunction() every 10-30 minutes
//
// PREREQUISITES:
// - Enable Gmail API, Calendar API, and Drive API in Apps Script
// ============================================================================

const CONFIG = {
  // ============================
  // YOUR PERSONAL INFORMATION
  // ============================
  // TODO: Replace with your full name exactly as it appears in emails
  MY_NAME: "Your Full Name",
  
  // TODO: Replace with your college registration number
  MY_REG: "YOUR_REG_NUMBER",
  
  // TODO: Replace with your branch (e.g., "Computer Science Engineering")
  MY_BRANCH: "Your Branch Name",
  
  // TODO: Replace with your current CGPA (e.g., 8.5)
  MY_CGPA: 0.0,
  
  // TODO: Replace with your current percentage (e.g., 85.0)
  MY_PERCENTAGE: 0.0,
  
  // TODO: Replace with your 10th grade percentage (e.g., 85)
  MY_10TH: 0,
  
  // TODO: Replace with your 12th grade percentage (e.g., 87)
  MY_12TH: 0,
  
  // TODO: Replace with number of active backlogs (0 if none)
  MY_BACKLOGS: 0,
  
  // TODO: Replace with your college email address
  MY_EMAIL: "your.email@college.edu",

  // ============================
  // BRANCH ALIASES
  // ============================
  // Maps common abbreviations to full branch names
  // Add more aliases if needed for your college
  ALLOWED_BRANCH_ALIASES: {
    "cse": "Computer Science Engineering",
    "computer science": "Computer Science Engineering",
    "cs": "Computer Science Engineering",
    "btech cse": "Computer Science Engineering",
    "b.tech cse": "Computer Science Engineering",
    "computer science and engineering": "Computer Science Engineering",
    "computer science & engineering": "Computer Science Engineering",
    "it": "Information Technology",
    "information technology": "Information Technology",
    "ece": "Electronics and Communication Engineering",
    "electronics and communication": "Electronics and Communication Engineering",
    "eee": "Electrical and Electronics Engineering",
    "electrical": "Electrical and Electronics Engineering",
    "mech": "Mechanical Engineering",
    "mechanical": "Mechanical Engineering",
    "civil": "Civil Engineering",
    "biomedical": "Biomedical Engineering",
    "bio": "Biomedical Engineering",
    "aids": "Artificial Intelligence and Data Science",
    "aiml": "Artificial Intelligence and Machine Learning",
    "all branches": "All",
    "all": "All",
    "any branch": "All",
    "open to all": "All"
  },

  // ============================
  // ALLOWED EMAIL SENDERS
  // ============================
  // TODO: Add your college's placement/CDC email addresses
  // Examples: "placement@college.edu", "cdc@college.edu", "careers@college.edu"
  ALLOWED_SENDERS: [
    "placement@college.edu",  // Replace with your college placement email
    "cdc@college.edu",        // Replace with your CDC email
    "careers@college.edu",    // Replace with careers email
    "placement",              // Matches any email containing "placement"
    "cdc",                    // Matches any email containing "cdc"
    "career",
    "recruitment",
    "hr"
  ],

  // ============================
  // DETECTION KEYWORDS
  // ============================
  // Keywords that identify placement-related emails
  KEYWORDS: [
    "talk", "test", "process", "online", "assessment", "exam", 
    "ppt", "placement", "interview", "screening", "recruitment", 
    "pre-placement", "registration", "opportunity", "drive", "hiring",
    "offer", "shortlist", "selected", "campus", "aptitude", "coding",
    "technical", "hr round", "group discussion", "gd", "resume",
    "application", "internship", "job", "role", "position", "opening",
    "pool campus", "off campus", "on campus", "walk-in", "company visit",
    "next round", "round 2", "round 3", "final round"
  ],

  // Keywords that indicate event is open to all students
  OPEN_TO_ALL_KEYWORDS: [
    "all students",
    "all applied",
    "all opted",
    "all registered",
    "all candidates",
    "all participants",
    "all shortlisted",
    "all eligible",
    "open to all",
    "all branches",
    "any branch",
    "everyone who",
    "those who applied",
    "those who registered",
    "all who opted"
  ],

  // ============================
  // CALENDAR SETTINGS
  // ============================
  // TODO (Optional): Create a dedicated calendar and paste its ID here
  // Leave empty ("") to use your default calendar
  // To find Calendar ID: Calendar Settings > Integrate calendar > Calendar ID
  CALENDAR_ID: "",

  // ============================
  // ADVANCED SETTINGS
  // ============================
  DEFAULT_EVENT_DURATION_MINUTES: 60,  // Default event length
  RETENTION_DAYS: 60,                   // Days to keep processed email records
  SEARCH_WINDOW_MINUTES: 30,            // How far back to search for new emails
  LABEL_NAME: "AutoCalendarProcessed",  // Gmail label for processed emails
  ELIGIBILITY_TOLERANCE: 0.3,           // CGPA tolerance (e.g., 7.5 required, 7.2 accepted)
  SEND_SUMMARY_EMAIL: false             // Set to true to receive summary emails
};

// ============================================================================
// MAIN FUNCTION
// ============================================================================
function myFunction() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      Logger.log("Another instance is running");
      return;
    }
  } catch (e) {
    Logger.log("Could not obtain lock: " + e.toString());
    return;
  }

  try {
    const stats = {
      processed: 0,
      created: 0,
      skipped: [],
      errors: []
    };

    cleanOldPersistedState();
    const label = getOrCreateLabel(CONFIG.LABEL_NAME);
    
    const now = new Date();
    const timeAgo = new Date(now.getTime() - CONFIG.SEARCH_WINDOW_MINUTES * 60 * 1000);
    
    const senderQuery = CONFIG.ALLOWED_SENDERS.map(s => {
      if (s.includes("@")) return `from:"${s}"`;
      return `from:${s}`;
    }).join(" OR ");
    
    const query = `(${senderQuery} OR to:"${CONFIG.MY_EMAIL}") is:unread newer_than:2d`;
    const threads = GmailApp.search(query);
    const recentThreads = threads.filter(thread => thread.getLastMessageDate() >= timeAgo);

    if (recentThreads.length === 0) {
      lock.releaseLock();
      return;
    }

    for (let thread of recentThreads) {
      try {
        const messages = thread.getMessages();
        for (let msg of messages) {
          const msgId = msg.getId();
          if (hasProcessedId(msgId)) continue;

          stats.processed++;
          const subject = msg.getSubject();
          const body = msg.getPlainBody();
          const htmlBody = msg.getBody();
          const sender = msg.getFrom();

          const hasKeyword = CONFIG.KEYWORDS.some(k => 
            subject.toLowerCase().includes(k) || 
            body.toLowerCase().includes(k) ||
            htmlBody.toLowerCase().includes(k)
          );

          if (!hasKeyword) {
            stats.skipped.push({msg: subject, reason: "No relevant keywords"});
            persistProcessedId(msgId);
            continue;
          }

          const fullText = subject + " " + body + " " + htmlBody;
          const eligibility = checkEligibility(fullText);
          
          if (eligibility.strictlyIneligible) {
            stats.skipped.push({msg: subject, reason: eligibility.reason});
            persistProcessedId(msgId);
            continue;
          }

          const attachments = msg.getAttachments();
          let foundInList = false;
          let venue = "TBD";
          let hasIcs = false;

          // Check attachments for name/reg
          for (let attachment of attachments) {
            const fileName = attachment.getName().toLowerCase();
            const contentType = attachment.getContentType();

            if (fileName.endsWith(".ics") || contentType.includes("calendar")) {
              hasIcs = true;
              const icsResult = handleIcsAttachment(attachment, msg, stats);
              if (icsResult) foundInList = true;
              continue;
            }

            if (contentType.includes("sheet") || fileName.match(/\.(xlsx|xls|csv)$/)) {
              const sheetResult = scanSheetForNameOrBranch(attachment);
              if (sheetResult.found) {
                foundInList = true;
                if (sheetResult.venue) venue = sheetResult.venue;
                Logger.log(`✓ Found your name/reg in Excel: ${fileName}`);
              }
            } else if (contentType.includes("pdf")) {
              const pdfText = convertPdfToText(attachment);
              if (pdfText) {
                const pdfCheck = checkNameInText(pdfText);
                if (pdfCheck.found) {
                  foundInList = true;
                  if (pdfCheck.venue) venue = pdfCheck.venue;
                  Logger.log(`✓ Found your name/reg in PDF: ${fileName}`);
                }
              }
            }
          }

          // Check email body for name/reg
          const bodyCheck = checkNameInText(fullText);
          if (bodyCheck.found) {
            foundInList = true;
            if (bodyCheck.venue && venue === "TBD") venue = bodyCheck.venue;
            Logger.log(`✓ Found your name/reg in email body`);
          }

          // Check if it's open to all using enhanced detection
          const isOpenToAll = checkIfOpenToAll(fullText);
          
          // CRITICAL LOGIC: Only proceed if name found OR explicitly open to all
          if (!foundInList && !isOpenToAll) {
            stats.skipped.push({
              msg: subject, 
              reason: "❌ SKIPPED: Your name/reg number NOT found in candidate list, and NOT open to all students"
            });
            Logger.log(`⚠️ SKIPPING: ${subject} - Name not found and not open to all`);
            persistProcessedId(msgId);
            continue;
          }

          // Log why we're proceeding
          if (foundInList) {
            Logger.log(`✅ PROCEEDING: Your name/reg found - creating event`);
          } else if (isOpenToAll) {
            Logger.log(`✅ PROCEEDING: Open to all students - creating event`);
          }

          if (hasIcs) {
            persistProcessedId(msgId);
            continue;
          }

          if (venue === "TBD") {
            venue = extractVenueFromText(fullText, attachments);
          }

          const regLink = findRegistrationLink(fullText);
          const dateTimes = parseDateTimes(fullText);

          let eventCreated = false;

          if (dateTimes.registrationDeadline) {
            const eventKey = generateEventKey(subject, dateTimes.registrationDeadline, thread.getId());
            if (!hasProcessedEventKey(eventKey) && !calendarEventExists(`Register: ${subject}`, dateTimes.registrationDeadline)) {
              createCalendarEvent(
                `Register: ${subject}`,
                dateTimes.registrationDeadline,
                new Date(dateTimes.registrationDeadline.getTime() + 30 * 60 * 1000),
                venue,
                msg,
                regLink,
                [1440, 120],
                eligibility,
                foundInList,
                isOpenToAll
              );
              persistEventKey(eventKey);
              stats.created++;
              eventCreated = true;
            }
          }

          if (dateTimes.eventDateTime) {
            const eventKey = generateEventKey(subject, dateTimes.eventDateTime, thread.getId());
            if (!hasProcessedEventKey(eventKey) && !calendarEventExists(subject, dateTimes.eventDateTime)) {
              const duration = dateTimes.duration || CONFIG.DEFAULT_EVENT_DURATION_MINUTES;
              const endTime = new Date(dateTimes.eventDateTime.getTime() + duration * 60 * 1000);
              createCalendarEvent(
                subject, 
                dateTimes.eventDateTime, 
                endTime, 
                venue, 
                msg, 
                regLink, 
                [60, 15], 
                eligibility,
                foundInList,
                isOpenToAll
              );
              persistEventKey(eventKey);
              stats.created++;
              eventCreated = true;
            }
          } else if (!dateTimes.registrationDeadline && (foundInList || isOpenToAll)) {
            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            tomorrow.setHours(10, 0, 0, 0);
            const eventKey = generateEventKey(subject, tomorrow, thread.getId());
            if (!hasProcessedEventKey(eventKey) && !calendarEventExists(subject, tomorrow)) {
              const endTime = new Date(tomorrow.getTime() + CONFIG.DEFAULT_EVENT_DURATION_MINUTES * 60 * 1000);
              createCalendarEvent(
                subject, 
                tomorrow, 
                endTime, 
                venue, 
                msg, 
                regLink, 
                [60, 15], 
                eligibility,
                foundInList,
                isOpenToAll
              );
              persistEventKey(eventKey);
              stats.created++;
              eventCreated = true;
            }
          }

          if (!eventCreated && (foundInList || isOpenToAll)) {
            const defaultDate = new Date();
            defaultDate.setDate(defaultDate.getDate() + 2);
            defaultDate.setHours(10, 0, 0, 0);
            const eventKey = generateEventKey(subject, defaultDate, thread.getId());
            if (!hasProcessedEventKey(eventKey) && !calendarEventExists(subject, defaultDate)) {
              const endTime = new Date(defaultDate.getTime() + CONFIG.DEFAULT_EVENT_DURATION_MINUTES * 60 * 1000);
              createCalendarEvent(
                subject, 
                defaultDate, 
                endTime, 
                venue, 
                msg, 
                regLink, 
                [1440, 60], 
                eligibility,
                foundInList,
                isOpenToAll
              );
              persistEventKey(eventKey);
              stats.created++;
            }
          }

          persistProcessedId(msgId);
        }
        thread.addLabel(label);
      } catch (e) {
        stats.errors.push({thread: thread.getFirstMessageSubject(), error: e.toString()});
      }
    }

    // Only send email if enabled in config
    if (CONFIG.SEND_SUMMARY_EMAIL) {
      sendSummaryEmail(stats);
    } else {
      Logger.log(`Summary: ${stats.created} events created, ${stats.skipped.length} skipped, ${stats.errors.length} errors`);
    }
  } finally {
    lock.releaseLock();
  }
}

function checkIfOpenToAll(text) {
  const lowerText = text.toLowerCase();
  
  // Check for explicit "open to all" keywords
  for (let keyword of CONFIG.OPEN_TO_ALL_KEYWORDS) {
    if (lowerText.includes(keyword)) {
      Logger.log(`✓ Detected open to all: "${keyword}"`);
      return true;
    }
  }
  
  // Check for patterns like "all students who applied/registered"
  const openPatterns = [
    /all\s+(?:students?|candidates?|participants?)\s+(?:who|that)\s+(?:applied|registered|opted|enrolled)/gi,
    /(?:every|each)\s+(?:student|candidate|participant)\s+who\s+(?:applied|registered|opted)/gi,
    /open\s+for\s+all/gi,
    /everyone\s+is\s+(?:invited|welcome|eligible)/gi
  ];
  
  for (let pattern of openPatterns) {
    if (pattern.test(text)) {
      Logger.log(`✓ Detected open to all pattern`);
      return true;
    }
  }
  
  return false;
}

function checkEligibility(text) {
  const result = {
    strictlyIneligible: false,
    reason: "",
    requiresNameCheck: false,
    isOpenToAll: false,
    eligibilityCriteria: []
  };

  const lowerText = text.toLowerCase();

  const branchPatterns = [
    /eligible\s+(?:for\s+)?branch(?:es)?[:\s]*([^\n\r\.;]+)/gi,
    /branch(?:es)?\s*(?:allowed|eligible|accepted)[:\s]*([^\n\r\.;]+)/gi,
    /(?:open|available)\s+(?:to|for)[:\s]*([^\n\r\.;]+?)(?:branch|student|only)/gi,
    /only\s+for[:\s]*([^\n\r\.;]+?)(?:branch|student)/gi,
    /(?:course|stream|department)[:\s]*([^\n\r\.;]+)/gi
  ];

  let foundBranches = [];
  let hasBranchRestriction = false;

  for (let pattern of branchPatterns) {
    let match;
    const regex = new RegExp(pattern);
    while ((match = regex.exec(text)) !== null) {
      hasBranchRestriction = true;
      const branchText = match[1].toLowerCase();
      
      for (let [alias, canonical] of Object.entries(CONFIG.ALLOWED_BRANCH_ALIASES)) {
        if (branchText.includes(alias)) {
          foundBranches.push(canonical);
        }
      }
    }
  }

  if (hasBranchRestriction && foundBranches.length > 0) {
    const isEligible = foundBranches.some(b => 
      b === CONFIG.MY_BRANCH || 
      b === "All" || 
      b.includes("Computer") ||
      b.includes("Information")
    );

    if (!isEligible) {
      result.strictlyIneligible = true;
      result.reason = `Branch restricted to: ${foundBranches.join(", ")}`;
      return result;
    }
    result.eligibilityCriteria.push(`Branch: ${foundBranches.join(", ")}`);
  }

  const cgpaPatterns = [
    /(?:minimum|min|required|atleast|at least)\s+(?:cgpa|gpa)[:\s]*(\d+\.?\d*)/gi,
    /cgpa[:\s]*(?:>=|above|minimum|min|should be|must be)?[:\s]*(\d+\.?\d*)/gi,
    /(\d+\.?\d*)\s+(?:cgpa|gpa)\s+(?:and above|or above|or more|or higher)/gi,
    /(?:cgpa|gpa)\s+(?:of|:)\s*(\d+\.?\d*)/gi
  ];

  let requiredCgpa = null;
  for (let pattern of cgpaPatterns) {
    const match = text.match(pattern);
    if (match) {
      const cgpaMatch = match[0].match(/(\d+\.?\d*)/);
      if (cgpaMatch) {
        requiredCgpa = parseFloat(cgpaMatch[1]);
        if (requiredCgpa > 10) requiredCgpa = requiredCgpa / 10;
        break;
      }
    }
  }

  if (requiredCgpa !== null) {
    const tolerance = CONFIG.ELIGIBILITY_TOLERANCE;
    if (CONFIG.MY_CGPA < (requiredCgpa - tolerance)) {
      result.strictlyIneligible = true;
      result.reason = `CGPA too low: requires ${requiredCgpa}, you have ${CONFIG.MY_CGPA}`;
      return result;
    }
    result.eligibilityCriteria.push(`CGPA: ${requiredCgpa}+`);
  }

  const percentagePatterns = [
    /(?:minimum|min|required|atleast|at least)\s+(?:percentage|%)[:\s]*(\d+\.?\d*)/gi,
    /(\d+\.?\d*)\s*%\s+(?:and above|or above|or more)/gi,
    /(?:10th|tenth|sslc)[:\s]*(\d+\.?\d*)\s*%/gi,
    /(?:12th|twelfth|hsc|intermediate)[:\s]*(\d+\.?\d*)\s*%/gi
  ];

  for (let pattern of percentagePatterns) {
    const match = text.match(pattern);
    if (match) {
      const percMatch = match[0].match(/(\d+\.?\d*)/);
      if (percMatch) {
        const required = parseFloat(percMatch[1]);
        const is10th = match[0].toLowerCase().includes("10") || match[0].toLowerCase().includes("tenth") || match[0].toLowerCase().includes("sslc");
        const is12th = match[0].toLowerCase().includes("12") || match[0].toLowerCase().includes("twelfth") || match[0].toLowerCase().includes("hsc");
        
        if (is10th && CONFIG.MY_10TH < (required - 2)) {
          result.strictlyIneligible = true;
          result.reason = `10th percentage too low: requires ${required}%, you have ${CONFIG.MY_10TH}%`;
          return result;
        } else if (is12th && CONFIG.MY_12TH < (required - 2)) {
          result.strictlyIneligible = true;
          result.reason = `12th percentage too low: requires ${required}%, you have ${CONFIG.MY_12TH}%`;
          return result;
        } else if (!is10th && !is12th && CONFIG.MY_PERCENTAGE < (required - 2)) {
          result.strictlyIneligible = true;
          result.reason = `Percentage too low: requires ${required}%, you have ${CONFIG.MY_PERCENTAGE}%`;
          return result;
        }
        result.eligibilityCriteria.push(`Percentage: ${required}%+`);
      }
    }
  }

  const backlogPatterns = [
    /no\s+(?:active\s+)?(?:standing\s+)?backlogs?/gi,
    /(?:0|zero|nil)\s+backlogs?/gi,
    /without\s+backlogs?/gi,
    /backlog\s*(?:should be|:)\s*(?:0|zero|nil)/gi
  ];

  for (let pattern of backlogPatterns) {
    if (pattern.test(text)) {
      if (CONFIG.MY_BACKLOGS > 0) {
        result.strictlyIneligible = true;
        result.reason = "No backlogs required, you have backlogs";
        return result;
      }
      result.eligibilityCriteria.push("No backlogs");
      break;
    }
  }

  result.isOpenToAll = checkIfOpenToAll(text);

  if (!result.isOpenToAll && (hasBranchRestriction || requiredCgpa !== null || result.eligibilityCriteria.length > 0)) {
    result.requiresNameCheck = true;
  }

  return result;
}

function checkNameInText(text) {
  const result = {found: false, venue: null};
  const lowerText = text.toLowerCase();
  
  if (lowerText.includes(CONFIG.MY_NAME.toLowerCase()) || lowerText.includes(CONFIG.MY_REG.toLowerCase())) {
    result.found = true;
    
    const venuePatterns = [
      new RegExp(`${CONFIG.MY_NAME}[^\\n]{0,100}(?:venue|location|room|lab)[:\\s]*([^\\n\\.;]+)`, 'gi'),
      new RegExp(`${CONFIG.MY_REG}[^\\n]{0,100}(?:venue|location|room|lab)[:\\s]*([^\\n\\.;]+)`, 'gi'),
      new RegExp(`(?:venue|location|room|lab)[:\\s]*([^\\n\\.;]+)[^\\n]{0,100}${CONFIG.MY_NAME}`, 'gi'),
      new RegExp(`(?:venue|location|room|lab)[:\\s]*([^\\n\\.;]+)[^\\n]{0,100}${CONFIG.MY_REG}`, 'gi')
    ];
    
    for (let pattern of venuePatterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        result.venue = match[1].trim();
        break;
      }
    }
  }
  
  return result;
}

function findRegistrationLink(text) {
  const urlPattern = /(https?:\/\/[^\s<>"{}|\\^`\[\]]+)/gi;
  const matches = text.match(urlPattern);
  
  if (!matches) {
    if (text.toLowerCase().includes("neopat")) {
      return "NEOPAT Portal";
    }
    return null;
  }

  const priorityKeywords = ["register", "registration", "signup", "sign-up", "apply", "application", "form", "neopat", "portal"];
  
  for (let url of matches) {
    const lower = url.toLowerCase();
    for (let keyword of priorityKeywords) {
      if (lower.includes(keyword)) {
        return url;
      }
    }
  }

  return matches[0];
}

function parseDateTimes(text) {
  const result = {
    eventDateTime: null,
    registrationDeadline: null,
    duration: null
  };

  const regDeadlinePatterns = [
    /(?:last\s+date|deadline|register\s+by|registration\s+(?:closes|ends?|before|till|until))[:\s]*(\d{1,2}[-\/\.\s]+\d{1,2}[-\/\.\s]+\d{2,4})(?:[,\s]*(?:by|at|before)?[,\s]*(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?)?/gi,
    /(?:last\s+date|deadline|register\s+by)[:\s]*(\d{1,2}(?:st|nd|rd|th)?\s+\w+\s+\d{2,4})(?:[,\s]*(?:by|at)?[,\s]*(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?)?/gi,
    /registration[:\s]+(\d{1,2}[-\/\.]\d{1,2}[-\/\.]\d{2,4})(?:\s+(?:by|at|till)\s+)?(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?/gi
  ];

  for (let pattern of regDeadlinePatterns) {
    let match;
    const regex = new RegExp(pattern);
    while ((match = regex.exec(text)) !== null) {
      const parsed = extractDateTimeFromMatch(match);
      if (parsed) {
        result.registrationDeadline = parsed;
        break;
      }
    }
    if (result.registrationDeadline) break;
  }

  const eventPatterns = [
    /(?:date|on|scheduled)[:\s]*(\d{1,2}[-\/\.\s]+\d{1,2}[-\/\.\s]+\d{2,4})(?:[,\s]*(?:at|@|time)?[,\s]*(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?)?/gi,
    /(?:interview|test|assessment|exam|screening|round)[:\s]*(\d{1,2}(?:st|nd|rd|th)?\s+\w+\s+\d{2,4})(?:[,\s]*(?:at|@)?[,\s]*(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?)?/gi,
    /(\d{1,2}[-\/\.]\d{1,2}[-\/\.]\d{2,4})[,\s]+(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)/gi,
    /(?:will be conducted on|scheduled on)[:\s]*(\d{1,2}(?:st|nd|rd|th)?\s+\w+\s+\d{2,4})(?:[,\s]*(?:at)?[,\s]*(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?)?/gi,
    /(\d{1,2}(?:st|nd|rd|th)?\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{2,4})[,\s]+(\d{1,2}):?(\d{2})?\s*([ap]\.?m\.?)?/gi
  ];

  for (let pattern of eventPatterns) {
    let match;
    const regex = new RegExp(pattern);
    while ((match = regex.exec(text)) !== null) {
      if (result.registrationDeadline) {
        const matchText = match[0].toLowerCase();
        if (matchText.includes("register") || matchText.includes("deadline") || matchText.includes("last date")) {
          continue;
        }
      }
      const parsed = extractDateTimeFromMatch(match);
      if (parsed) {
        result.eventDateTime = parsed;
        break;
      }
    }
    if (result.eventDateTime) break;
  }

  const durationPatterns = [
    /duration[:\s]*(\d+)\s*(?:hour|hr|hrs)/gi,
    /(\d+)\s*(?:hour|hr|hrs)\s*(?:duration|long|test|exam)/gi,
    /(\d{1,2}):(\d{2})\s*(?:-|to|till)\s*(\d{1,2}):(\d{2})/gi
  ];

  for (let pattern of durationPatterns) {
    const match = text.match(pattern);
    if (match) {
      if (match[1] && match[2] && match[3] && match[4]) {
        const startMinutes = parseInt(match[1]) * 60 + parseInt(match[2]);
        const endMinutes = parseInt(match[3]) * 60 + parseInt(match[4]);
        result.duration = endMinutes - startMinutes;
      } else {
        const hours = parseInt(match[1]);
        result.duration = hours * 60;
      }
      break;
    }
  }

  return result;
}

function extractDateTimeFromMatch(match) {
  try {
    const dateStr = match[1];
    const hour = match[2];
    const minute = match[3] || "00";
    const ampm = match[4] || "";

    const extractedDate = extractDate(dateStr);
    if (!extractedDate) return null;

    let timeStr = "10:00 AM";
    if (hour) {
      timeStr = `${hour}:${minute}`;
      if (ampm) {
        timeStr += ` ${ampm.replace(/\./g, '').toUpperCase()}`;
      } else if (parseInt(hour) < 12 && !ampm) {
        timeStr += " AM";
      }
    }

    return new Date(`${extractedDate} ${timeStr}`);
  } catch (e) {
    Logger.log("Error in extractDateTimeFromMatch: " + e.toString());
    return null;
  }
}

function extractDate(text) {
  const monthMap = {
    'jan': 'January', 'feb': 'February', 'mar': 'March', 'apr': 'April',
    'may': 'May', 'jun': 'June', 'jul': 'July', 'aug': 'August',
    'sep': 'September', 'oct': 'October', 'nov': 'November', 'dec': 'December',
    'january': 'January', 'february': 'February', 'march': 'March', 'april': 'April',
    'june': 'June', 'july': 'July', 'august': 'August', 'september': 'September',
    'october': 'October', 'november': 'November', 'december': 'December'
  };

  const patterns = [
    /(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{2,4})/,
    /(\d{1,2})(?:st|nd|rd|th)?\s+(\w+)\s+(\d{2,4})/i,
    /(\w+)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{2,4})/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match) {
      let year = match[3];
      if (year.length === 2) {
        year = "20" + year;
      }
      
      if (i === 0) {
        return `${match[2]}/${match[1]}/${year}`;
      } else if (i === 1) {
        const month = monthMap[match[2].toLowerCase()];
        if (month) return `${month} ${match[1]} ${year}`;
      } else if (i === 2) {
        const month = monthMap[match[1].toLowerCase()];
        if (month) return `${month} ${match[2]} ${year}`;
      }
    }
  }

  return null;
}

function extractTime(text) {
  const timePatterns = [
    /(\d{1,2}):(\d{2})\s*([ap]\.?m\.?)/i,
    /(\d{1,2})\s*([ap]\.?m\.?)/i,
    /(\d{4})\s*(?:hrs|hours)/i
  ];

  for (let pattern of timePatterns) {
    const match = text.match(pattern);
    if (match) {
      if (match[3]) {
        const hour = match[1];
        const minute = match[2] || "00";
        const ampm = match[3].replace(/\./g, '').toUpperCase();
        return `${hour}:${minute} ${ampm}`;
      } else if (match[2]) {
        const hour = match[1];
        const ampm = match[2].replace(/\./g, '').toUpperCase();
        return `${hour}:00 ${ampm}`;
      } else if (match[1].length === 4) {
        const hour = match[1].substring(0, 2);
        const minute = match[1].substring(2, 4);
        const hourInt = parseInt(hour);
        const ampm = hourInt >= 12 ? "PM" : "AM";
        const hour12 = hourInt > 12 ? hourInt - 12 : (hourInt === 0 ? 12 : hourInt);
        return `${hour12}:${minute} ${ampm}`;
      }
    }
  }

  return null;
}

function extractVenueFromText(text, attachments) {
  const venuePatterns = [
    /venue[:\s]+([^\n\r\.\;]{3,50})/gi,
    /location[:\s]+([^\n\r\.\;]{3,50})/gi,
    /room[:\s]+([^\n\r\.\;]{3,30})/gi,
    /lab[:\s]+([^\n\r\.\;]{3,30})/gi,
    /block[:\s]+([^\n\r\.\;]{3,30})/gi,
    /hall[:\s]+([^\n\r\.\;]{3,30})/gi
  ];

  for (let pattern of venuePatterns) {
    let match;
    const regex = new RegExp(pattern);
    while ((match = regex.exec(text)) !== null) {
      if (match[1]) {
        const venue = match[1].trim();
        if (venue.length > 3 && !venue.toLowerCase().includes("http")) {
          return venue;
        }
      }
    }
  }

  if (text.toLowerCase().includes("own location") || text.toLowerCase().includes("your location")) {
    return "Own Location";
  }

  if (text.toLowerCase().includes("online") || text.toLowerCase().includes("virtual")) {
    return "Online";
  }

  for (let attachment of attachments) {
    const result = scanSheetForNameOrBranch(attachment);
    if (result.venue) return result.venue;
  }

  return "TBD";
}

function scanSheetForNameOrBranch(attachment) {
  let tempFile = null;
  try {
    const fileName = attachment.getName().toLowerCase();
    const contentType = attachment.getContentType();

    if (!contentType.includes("sheet") && !fileName.match(/\.(xlsx|xls|csv)$/)) {
      return {found: false, venue: null};
    }

    const blob = attachment.copyBlob();
    const fileResource = {
      name: attachment.getName().replace(/\.(xlsx|xls|csv)$/, ''),
      mimeType: 'application/vnd.google-apps.spreadsheet'
    };

    tempFile = Drive.Files.create(fileResource, blob);
    const spreadsheet = SpreadsheetApp.openById(tempFile.id);
    
    for (let sheet of spreadsheet.getSheets()) {
      const data = sheet.getDataRange().getValues();
      if (data.length === 0) continue;

      const headers = data[0].map(h => h.toString().toLowerCase());
      const venueIndex = headers.findIndex(h => 
        h.includes("venue") || 
        h.includes("location") || 
        h.includes("room") || 
        h.includes("lab")
      );
      
      for (let i = 1; i < data.length; i++) {
        const rowText = data[i].join(' ').toLowerCase();
        const nameMatch = rowText.includes(CONFIG.MY_NAME.toLowerCase());
        const regMatch = rowText.includes(CONFIG.MY_REG.toLowerCase());
        
        if (nameMatch || regMatch) {
          const venue = venueIndex !== -1 && data[i][venueIndex] ? data[i][venueIndex].toString() : null;
          return {found: true, venue: venue};
        }
      }
    }

    return {found: false, venue: null};
  } catch (e) {
    Logger.log("Error scanning sheet: " + e.toString());
    return {found: false, venue: null};
  } finally {
    if (tempFile) {
      try {
        DriveApp.getFileById(tempFile.id).setTrashed(true);
      } catch (e) {
        Logger.log("Error deleting temp file: " + e.toString());
      }
    }
  }
}

function convertPdfToText(attachment) {
  let tempFile = null;
  try {
    const blob = attachment.copyBlob();
    const fileResource = {
      name: attachment.getName(),
      mimeType: 'application/vnd.google-apps.document'
    };

    tempFile = Drive.Files.create(fileResource, blob);
    const doc = DocumentApp.openById(tempFile.id);
    const text = doc.getBody().getText();
    return text;
  } catch (e) {
    Logger.log("Error converting PDF: " + e.toString());
    return null;
  } finally {
    if (tempFile) {
      try {
        DriveApp.getFileById(tempFile.id).setTrashed(true);
      } catch (e) {}
    }
  }
}

function handleIcsAttachment(attachment, msg, stats) {
  try {
    const icsContent = attachment.getDataAsString();
    const eventData = parseIcsContent(icsContent);
    
    if (!eventData) return false;

    const eventKey = generateEventKey(eventData.summary, eventData.start, msg.getThread().getId());
    if (hasProcessedEventKey(eventKey) || calendarEventExists(eventData.summary, eventData.start)) {
      return true;
    }

    const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    calendar.createEvent(eventData.summary, eventData.start, eventData.end, {
      location: eventData.location || "TBD",
      description: `${eventData.description || ''}\n\n📧 Original Email: ${msg.getThread().getPermalink()}`
    });

    persistEventKey(eventKey);
    stats.created++;
    return true;
  } catch (e) {
    Logger.log("Error handling ICS: " + e.toString());
    return false;
  }
}

function parseIcsContent(icsContent) {
  try {
    const summaryMatch = icsContent.match(/SUMMARY:([^\r\n]+)/);
    const startMatch = icsContent.match(/DTSTART[;:]([^\r\n]+)/);
    const endMatch = icsContent.match(/DTEND[;:]([^\r\n]+)/);
    const locationMatch = icsContent.match(/LOCATION:([^\r\n]+)/);
    const descMatch = icsContent.match(/DESCRIPTION:([^\r\n]+)/);

    if (!summaryMatch || !startMatch) return null;

    const startDate = parseIcsDateTime(startMatch[1]);
    const endDate = endMatch ? parseIcsDateTime(endMatch[1]) : new Date(startDate.getTime() + 3600000);

    return {
      summary: summaryMatch[1],
      start: startDate,
      end: endDate,
      location: locationMatch ? locationMatch[1] : null,
      description: descMatch ? descMatch[1] : null
    };
  } catch (e) {
    Logger.log("Error parsing ICS: " + e.toString());
    return null;
  }
}

function parseIcsDateTime(dateStr) {
  const cleaned = dateStr.replace(/[TZ:-]/g, '');
  const year = parseInt(cleaned.substr(0, 4));
  const month = parseInt(cleaned.substr(4, 2)) - 1;
  const day = parseInt(cleaned.substr(6, 2));
  const hour = parseInt(cleaned.substr(8, 2)) || 0;
  const minute = parseInt(cleaned.substr(10, 2)) || 0;
  
  return new Date(year, month, day, hour, minute);
}

function createCalendarEvent(title, start, end, venue, msg, regLink, reminderMinutes, eligibility, foundInList, isOpenToAll) {
  try {
    const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    
    let description = `📍 Venue: ${venue}\n⏱️ Duration: ${Math.round((end.getTime() - start.getTime()) / 60000)} minutes\n\n`;
    
    // Add selection status
    if (foundInList) {
      description += `✅ YOU ARE SELECTED/SHORTLISTED\n`;
      description += `   Your name/reg number found in candidate list\n\n`;
    } else if (isOpenToAll) {
      description += `🔓 OPEN TO ALL STUDENTS\n`;
      description += `   Anyone can attend/register\n\n`;
    }
    
    if (eligibility && eligibility.eligibilityCriteria.length > 0) {
      description += `📋 Eligibility:\n`;
      eligibility.eligibilityCriteria.forEach(criteria => {
        description += `   • ${criteria}\n`;
      });
      description += `\n`;
    }
    
    if (regLink) {
      if (regLink === "NEOPAT Portal") {
        description += `📝 Register via NEOPAT portal\n   Visit: https://neopat.vit.ac.in/\n\n`;
      } else {
        description += `📝 Register here: ${regLink}\n\n`;
      }
    }
    
    description += `📧 Original Email: ${msg.getThread().getPermalink()}\n\n`;
    description += `💡 Checklist:\n`;
    description += `   ✓ Bring ID card\n`;
    description += `   ✓ Arrive 15-20 mins early\n`;
    description += `   ✓ Carry resume copies\n`;
    description += `   ✓ Check dress code if mentioned\n`;
    
    if (venue === "TBD" || venue === "Online") {
      description += `\n⚠️ ${venue === "TBD" ? "Venue not confirmed" : "Mode: Online"} - check original email for details.`;
    }

    calendar.createEvent(title, start, end, {
      location: venue,
      description: description,
      reminders: {
        useDefault: false,
        overrides: reminderMinutes.map(m => ({method: 'popup', minutes: m}))
      }
    });
    
    Logger.log(`✓ Created event: ${title} at ${start}`);
  } catch (e) {
    Logger.log("Error creating calendar event: " + e.toString());
    throw e;
  }
}

function calendarEventExists(title, startTime) {
  try {
    const calendar = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    const start = new Date(startTime.getTime() - 300000);
    const end = new Date(startTime.getTime() + 300000);
    const events = calendar.getEvents(start, end);
    
    return events.some(e => e.getTitle().toLowerCase() === title.toLowerCase());
  } catch (e) {
    Logger.log("Error checking calendar: " + e.toString());
    return false;
  }
}

function generateEventKey(subject, datetime, threadId) {
  const normalized = subject.toLowerCase().replace(/[^a-z0-9]/g, '').substring(0, 50);
  return `${normalized}|${datetime.getTime()}|${threadId}`;
}

function persistProcessedId(id) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`processed_${id}`, Date.now().toString());
  } catch (e) {
    Logger.log("Error persisting ID: " + e.toString());
  }
}

function hasProcessedId(id) {
  try {
    const props = PropertiesService.getScriptProperties();
    return props.getProperty(`processed_${id}`) !== null;
  } catch (e) {
    Logger.log("Error checking processed ID: " + e.toString());
    return false;
  }
}

function persistEventKey(key) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`event_${key}`, Date.now().toString());
  } catch (e) {
    Logger.log("Error persisting event key: " + e.toString());
  }
}

function hasProcessedEventKey(key) {
  try {
    const props = PropertiesService.getScriptProperties();
    return props.getProperty(`event_${key}`) !== null;
  } catch (e) {
    Logger.log("Error checking event key: " + e.toString());
    return false;
  }
}

function cleanOldPersistedState() {
  try {
    const props = PropertiesService.getScriptProperties();
    const all = props.getProperties();
    const cutoff = Date.now() - (CONFIG.RETENTION_DAYS * 24 * 60 * 60 * 1000);
    
    let deleteCount = 0;
    for (let key in all) {
      if ((key.startsWith('processed_') || key.startsWith('event_'))) {
        const timestamp = parseInt(all[key]);
        if (timestamp < cutoff) {
          props.deleteProperty(key);
          deleteCount++;
        }
      }
    }
    
    if (deleteCount > 0) {
      Logger.log(`Cleaned ${deleteCount} old entries from persistent storage`);
    }
  } catch (e) {
    Logger.log("Error cleaning old state: " + e.toString());
  }
}

function getOrCreateLabel(labelName) {
  try {
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
      Logger.log(`Created new label: ${labelName}`);
    }
    return label;
  } catch (e) {
    Logger.log("Error getting/creating label: " + e.toString());
    return null;
  }
}

function sendSummaryEmail(stats) {
  try {
    const timestamp = new Date().toLocaleString('en-IN', {timeZone: 'Asia/Kolkata'});
    
    let body = `=== Placement Calendar Automation Summary ===\n`;
    body += `Run Time: ${timestamp}\n\n`;
    body += `📊 Statistics:\n`;
    body += `   • Emails Processed: ${stats.processed}\n`;
    body += `   • Events Created: ${stats.created}\n`;
    body += `   • Opportunities Skipped: ${stats.skipped.length}\n`;
    body += `   • Errors Encountered: ${stats.errors.length}\n\n`;
    
    if (stats.skipped.length > 0) {
      body += `⏭️ Skipped Opportunities:\n`;
      stats.skipped.forEach((s, idx) => {
        body += `   ${idx + 1}. ${s.msg}\n`;
        body += `      ${s.reason}\n\n`;
      });
    }
    
    if (stats.errors.length > 0) {
      body += `❌ Errors:\n`;
      stats.errors.forEach((e, idx) => {
        body += `   ${idx + 1}. ${e.thread}\n`;
        body += `      Error: ${e.error}\n\n`;
      });
    }
    
    if (stats.created === 0 && stats.skipped.length === 0 && stats.errors.length === 0) {
      body += `✅ No new placement opportunities found in this run.\n`;
    } else if (stats.created > 0) {
      body += `✅ Successfully created ${stats.created} calendar event(s).\n`;
      body += `   Check your Google Calendar for details.\n`;
    }
    
    body += `\n---\n`;
    body += `This is an automated summary from your Placement Calendar Script.\n`;
    body += `Script Version: 3.0 (Strict Name/Reg Verification)\n`;
    body += `\n⚠️ IMPORTANT: Events are ONLY created if:\n`;
    body += `   1. Your name/reg number is found in attachments/email body, OR\n`;
    body += `   2. The opportunity is explicitly open to all students\n`;
    
    MailApp.sendEmail({
      to: CONFIG.MY_EMAIL,
      subject: `📅 Placement Calendar - ${stats.created} Event(s) Created`,
      body: body
    });
    
    Logger.log("Summary email sent successfully");
  } catch (e) {
    Logger.log("Error sending summary email: " + e.toString());
  }
}

// ============================================================================
// TESTING FUNCTIONS (Optional - for debugging)
// ============================================================================

function testEligibilityCheck() {
  const testCases = [
    "Eligible Branches: CSE, IT, ECE. Minimum CGPA: 7.5",
    "Open to all branches with 8.0 CGPA",
    "Only for Biomedical and Bio students",
    "Required: 10th - 80%, 12th - 85%, No backlogs",
    "All students who applied are invited",
    "Next round for all registered candidates"
  ];
  
  testCases.forEach((test, idx) => {
    Logger.log(`\nTest Case ${idx + 1}: ${test}`);
    const result = checkEligibility(test);
    Logger.log(`Strictly Ineligible: ${result.strictlyIneligible}`);
    Logger.log(`Is Open To All: ${result.isOpenToAll}`);
    Logger.log(`Reason: ${result.reason || "N/A"}`);
    Logger.log(`Eligibility Criteria: ${result.eligibilityCriteria.join(", ") || "None"}`);
  });
}

function testDateTimeParsing() {
  const testDates = [
    "Interview on 1-10-2025 at 3:00 PM",
    "Last date for registration: 30/09/2025 (5:00 pm)",
    "Assessment on 15th October 2025 by 2:30 PM",
    "Scheduled for Oct 20, 2025 at 10 AM",
    "Registration deadline: 1st Nov 2025"
  ];
  
  testDates.forEach((test, idx) => {
    Logger.log(`\nTest ${idx + 1}: ${test}`);
    const result = parseDateTimes(test);
    Logger.log(`Event DateTime: ${result.eventDateTime}`);
    Logger.log(`Registration Deadline: ${result.registrationDeadline}`);
  });
}
