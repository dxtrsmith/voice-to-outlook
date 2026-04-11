import Anthropic from "@anthropic-ai/sdk";
import http from "http";

const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const ALLOWED_USER_ID = parseInt(process.env.ALLOWED_USER_ID);
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const PORT = process.env.PORT || 3000;

const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

let accessToken = null;
let tokenExpiry = null;
let currentRefreshToken = process.env.MS_REFRESH_TOKEN;
const pendingLookups = {};

function buildDateContext() {
  const days = ["zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"];
  const now = new Date();
  const lines = ["Today is " + now.toISOString().slice(0, 10) + " (" + days[now.getDay()] + ")."];
  lines.push("Use these exact dates for relative references:");
  for (let i = 0; i <= 14; i++) {
    const d = new Date(now);
    d.setDate(now.getDate() + i);
    const label = i === 0 ? "vandaag" : i === 1 ? "morgen" : days[d.getDay()] + (i > 7 ? " (volgende week)" : "");
    lines.push("- " + label + " = " + d.toISOString().slice(0, 10));
  }
  return lines.join("\n");
}

async function getAccessToken() {
  if (accessToken && tokenExpiry && Date.now() < tokenExpiry) {
    return accessToken;
  }
  const url = "https://login.microsoftonline.com/" + AZURE_TENANT_ID + "/oauth2/v2.0/token";
  const body = new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    grant_type: "refresh_token",
    refresh_token: currentRefreshToken,
    scope: "Mail.ReadWrite Calendars.ReadWrite Tasks.ReadWrite People.Read offline_access",
  });
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const data = await res.json();
  if (!data.access_token) {
    throw new Error("Token refresh error: " + JSON.stringify(data));
  }
  accessToken = data.access_token;
  tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  if (data.refresh_token) {
    currentRefreshToken = data.refresh_token;
  }
  return accessToken;
}

async function graphRequest(method, path, body) {
  const token = await getAccessToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    method: method,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error("Graph API error " + res.status + ": " + err);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function lookupRecipient(name) {
  if (!name) return null;

  const encoded = encodeURIComponent(name);
  const data = await graphRequest("GET", "/me/people?$search=" + encoded + "&$top=10");

  // Only real people with email addresses
  const people = (data.value || []).filter(function(p) {
    return p.scoredEmailAddresses &&
      p.scoredEmailAddresses.length > 0 &&
      p.personType &&
      p.personType.class === "Person";
  });

  // Filter by name words: keep only candidates whose display name contains
  // the first AND last word of the spoken name (case insensitive)
  const words = name.trim().toLowerCase().split(/\s+/);
  const firstName = words[0];
  const lastName = words[words.length - 1];

  const matched = people.filter(function(p) {
    const dn = (p.displayName || "").toLowerCase();
    if (words.length === 1) {
      return dn.includes(firstName);
    }
    return dn.includes(firstName) && dn.includes(lastName);
  });

  return matched;
}

async function createDraft(subject, body, toName, toEmail) {
  const toRecipients = toEmail
    ? [{ emailAddress: { name: toName || toEmail, address: toEmail } }]
    : toName
    ? [{ emailAddress: { name: toName, address: "" } }]
    : [];
  await graphRequest("POST", "/me/messages", {
    subject: subject,
    body: { contentType: "Text", content: body },
    toRecipients: toRecipients,
  });
}

async function createCalendarEvent(title, details, startDateTime, endDateTime) {
  const start = startDateTime || new Date(Date.now() + 86400000).toISOString().slice(0, 16);
  const end = endDateTime || new Date(Date.now() + 90000000).toISOString().slice(0, 16);
  await graphRequest("POST", "/me/events", {
    subject: title,
    body: { contentType: "Text", content: details },
    start: { dateTime: start, timeZone: "Europe/Amsterdam" },
    end: { dateTime: end, timeZone: "Europe/Amsterdam" },
    isOnlineMeeting: false,
    attendees: [],
  });
}

async function getOrCreateToDoList() {
  const lists = await graphRequest("GET", "/me/todo/lists");
  const existing = lists.value.find(function(l) { return l.displayName === "Voice Assistant"; });
  if (existing) return existing.id;
  const newList = await graphRequest("POST", "/me/todo/lists", { displayName: "Voice Assistant" });
  return newList.id;
}

async function createTask(title, notes, dueDate) {
  const listId = await getOrCreateToDoList();
  const task = {
    title: title,
    body: { contentType: "text", content: notes || "" },
  };
  if (dueDate) {
    task.dueDateTime = { dateTime: dueDate + "T00:00:00", timeZone: "Europe/Amsterdam" };
  }
  await graphRequest("POST", "/me/todo/lists/" + listId + "/tasks", task);
}

async function processWithClaude(transcript) {
  const dateContext = buildDateContext();
  const system = "You are Dexter's personal voice-to-Outlook assistant. Dexter Smith is a Product Director at IceMobile, working with clients like Kruidvat (AS Watson) and Albert Heijn.\n\n" + dateContext + "\n\nTASK: Analyze the Dutch transcript, detect the intent(s), and produce structured JSON output.\n\nINTENT DETECTION:\n- email: Dexter wants to send an email\n- meeting: Dexter wants to schedule a meeting or calendar event\n- task: Dexter wants to add a to-do item\n\nSTRICT RULES:\n- Never auto-send, never invite attendees. Emails go to Drafts only.\n- Only email who Dexter explicitly mentions. Never assume recipients.\n- Default language is ALWAYS Dutch. Only switch to English if Dexter explicitly says so.\n- No em-dashes. Ever.\n\nDUTCH EMAIL STYLE:\n- Opening: always \"Hi [name],\" or \"Hi,\" - never \"Beste\" or \"Geachte\"\n- Sign-off: \"Groetjes,\\nDexter\"\n- Tone: direct, friendly, no padding. Point in first sentence.\n- Short declarative sentences. Bullet points for multiple items.\n\nENGLISH EMAIL STYLE (only when Dexter explicitly requests):\n- Opening: \"Hi [name]\" or \"Hi all\"\n- Sign-off: \"Regards,\\nDexter\"\n- Confident, clear, action-oriented.\n\nDATE/TIME PARSING:\n- Use ONLY the exact dates listed above. Do not calculate dates yourself.\n- For meetings: extract start datetime and calculate end (default 1 hour unless specified)\n- Format datetimes as: YYYY-MM-DDTHH:mm:00 (24h, no timezone suffix)\n- For tasks: extract due date as YYYY-MM-DD\n- If no date/time mentioned, use null\n\nRespond ONLY with valid JSON, no markdown, no backticks:\n{\n  \"outputs\": [\n    {\n      \"type\": \"email\",\n      \"to_name\": \"Full name as spoken, or empty string\",\n      \"subject\": \"Short email subject\",\n      \"body\": \"Full email body text\"\n    },\n    {\n      \"type\": \"meeting\",\n      \"title\": \"Short meeting title\",\n      \"details\": \"Agenda and context as plain text\",\n      \"start_datetime\": \"YYYY-MM-DDTHH:mm:00 or null\",\n      \"end_datetime\": \"YYYY-MM-DDTHH:mm:00 or null\"\n    },\n    {\n      \"type\": \"task\",\n      \"title\": \"Clear action item\",\n      \"notes\": \"Any relevant context or empty string\",\n      \"due_date\": \"YYYY-MM-DD or null\"\n    }\n  ]\n}";

  const message = await anthropic.messages.create({
    model: "claude-sonnet-4-20250514",
    max_tokens: 1024,
    system: system,
    messages: [{ role: "user", content: transcript }],
  });
  const text = message.content[0].text;
  const clean = text.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

async function sendTelegramMessage(chatId, text) {
  await fetch("https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text: text }),
  });
}

async function handleUpdate(update) {
  const message = update.message;
  if (!message || !message.text) return;

  const chatId = message.chat.id;
  const userId = message.from.id;
  const text = message.text.trim();

  if (userId !== ALLOWED_USER_ID) return;

  if (pendingLookups[chatId]) {
    const pending = pendingLookups[chatId];
    const choice = parseInt(text);
    if (!isNaN(choice) && choice >= 1 && choice <= pending.candidates.length) {
      const chosen = pending.candidates[choice - 1];
      const email = chosen.scoredEmailAddresses[0].address;
      delete pendingLookups[chatId];
      await sendTelegramMessage(chatId, "Even verwerken...");
      try {
        await createDraft(pending.subject, pending.body, chosen.displayName, email);
        await sendTelegramMessage(chatId, "Email opgeslagen in Drafts\nAan: " + chosen.displayName + " (" + email + ")\nOnderwerp: " + pending.subject);
      } catch (err) {
        await sendTelegramMessage(chatId, "Er ging iets mis: " + err.message);
      }
    } else {
      await sendTelegramMessage(chatId, "Stuur een getal tussen 1 en " + pending.candidates.length + ".");
    }
    return;
  }

  if (text === "/start") {
    await sendTelegramMessage(chatId, "Stuur je transcript en ik verwerk het direct in Outlook.");
    return;
  }

  await sendTelegramMessage(chatId, "Even verwerken...");

  try {
    const parsed = await processWithClaude(text);
    const results = [];

    for (const output of parsed.outputs) {
      if (output.type === "email") {
        let toEmail = null;
        let toDisplayName = output.to_name;

        if (output.to_name) {
          const candidates = await lookupRecipient(output.to_name);
          if (candidates && candidates.length === 1) {
            toEmail = candidates[0].scoredEmailAddresses[0].address;
            toDisplayName = candidates[0].displayName;
          } else if (candidates && candidates.length > 1) {
            pendingLookups[chatId] = {
              candidates: candidates,
              subject: output.subject,
              body: output.body,
            };
            const options = candidates.map(function(p, i) {
              return (i + 1) + ". " + p.displayName + " (" + p.scoredEmailAddresses[0].address + ")";
            }).join("\n");
            await sendTelegramMessage(chatId, "Welke " + output.to_name + " bedoel je?\n\n" + options);
            continue;
          }
        }

        await createDraft(output.subject, output.body, toDisplayName, toEmail);
        const toLine = toEmail
          ? "Aan: " + toDisplayName + " (" + toEmail + ")"
          : "Aan: " + (toDisplayName || "onbekend") + " - vul e-mailadres in";
        results.push("Email opgeslagen in Drafts\n" + toLine + "\nOnderwerp: " + output.subject);

      } else if (output.type === "meeting") {
        await createCalendarEvent(output.title, output.details, output.start_datetime, output.end_datetime);
        const timeInfo = output.start_datetime
          ? "\nTijd: " + output.start_datetime.replace("T", " om ").slice(0, 16)
          : "\nTijd: nog in te vullen";
        results.push("Meeting toegevoegd aan je agenda\nTitel: " + output.title + timeInfo);

      } else if (output.type === "task") {
        await createTask(output.title, output.notes, output.due_date);
        const dueInfo = output.due_date ? "\nDeadline: " + output.due_date : "";
        results.push("Taak aangemaakt in To Do\nTaak: " + output.title + dueInfo);
      }
    }

    if (results.length > 0) {
      await sendTelegramMessage(chatId, results.join("\n\n---\n\n"));
    }

  } catch (err) {
    console.error("Error:", err);
    await sendTelegramMessage(chatId, "Er ging iets mis: " + err.message);
  }
}

const server = http.createServer(async function(req, res) {
  if (req.method === "POST") {
    let body = "";
    req.on("data", function(chunk) { body += chunk; });
    req.on("end", async function() {
      try {
        const update = JSON.parse(body);
        await handleUpdate(update);
      } catch (err) {
        console.error("Webhook error:", err);
      }
      res.writeHead(200);
      res.end("OK");
    });
  } else {
    res.writeHead(200);
    res.end("Voice-to-Outlook bot is running.");
  }
});

server.listen(PORT, function() {
  console.log("Bot running on port " + PORT);
});
