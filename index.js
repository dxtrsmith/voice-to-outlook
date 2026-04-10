import Anthropic from "@anthropic-ai/sdk";
import http from "http";

const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const ALLOWED_USER_ID = parseInt(process.env.ALLOWED_USER_ID);
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const USER_EMAIL = process.env.USER_EMAIL;
const PORT = process.env.PORT || 3000;

const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

let accessToken = null;
let tokenExpiry = null;

async function getAccessToken() {
  if (accessToken && tokenExpiry && Date.now() < tokenExpiry) {
    return accessToken;
  }
  const url = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const data = await res.json();
  if (!data.access_token) {
    throw new Error(`Token error: ${JSON.stringify(data)}`);
  }
  accessToken = data.access_token;
  tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  return accessToken;
}

async function graphRequest(method, path, body) {
  const token = await getAccessToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph API error ${res.status}: ${err}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function createDraft(subject, body, toName) {
  const message = {
    subject,
    body: { contentType: "Text", content: body },
    toRecipients: toName
      ? [{ emailAddress: { name: toName, address: "" } }]
      : [],
  };
  await graphRequest("POST", `/users/${USER_EMAIL}/messages`, message);
}

async function createCalendarEvent(title, details) {
  const event = {
    subject: title,
    body: { contentType: "Text", content: details },
    start: { dateTime: new Date(Date.now() + 86400000).toISOString(), timeZone: "Europe/Amsterdam" },
    end: { dateTime: new Date(Date.now() + 90000000).toISOString(), timeZone: "Europe/Amsterdam" },
    isOnlineMeeting: false,
    attendees: [],
  };
  await graphRequest("POST", `/users/${USER_EMAIL}/events`, event);
}

async function getOrCreateToDoList() {
  const lists = await graphRequest("GET", `/users/${USER_EMAIL}/todo/lists`);
  const existing = lists.value.find((l) => l.displayName === "Voice Assistant");
  if (existing) return existing.id;
  const newList = await graphRequest("POST", `/users/${USER_EMAIL}/todo/lists`, {
    displayName: "Voice Assistant",
  });
  return newList.id;
}

async function createTask(title, notes) {
  const listId = await getOrCreateToDoList();
  await graphRequest("POST", `/users/${USER_EMAIL}/todo/lists/${listId}/tasks`, {
    title,
    body: { contentType: "text", content: notes || "" },
  });
}

const STYLE_PROMPT = `You are Dexter's personal voice-to-Outlook assistant. Dexter Smith is a Product Director at IceMobile, working with clients like Kruidvat (AS Watson) and Albert Heijn.

TASK: Analyze the Dutch transcript, detect the intent(s), and produce structured JSON output.

INTENT DETECTION — detect one or more of:
- email: Dexter wants to send an email
- meeting: Dexter wants to schedule a meeting or calendar event
- task: Dexter wants to add a to-do item

STRICT RULES — NEVER BREAK THESE:
- Never auto-send, never invite attendees, never notify anyone. Emails go to Drafts only.
- Only email who Dexter explicitly mentions. Never assume recipients.
- Default language is ALWAYS Dutch. Only switch to English if Dexter explicitly says so.
- No em-dashes. Ever.

DUTCH EMAIL STYLE:
- Opening: always "Hi [name]," or "Hi," — never "Beste" or "Geachte"
- Sign-off: "Groetjes,\nDexter"
- Tone: direct, friendly, no padding. Point in first sentence.
- Sentences: short and declarative.
- Lists: bullet points for multiple items
- No fluff unless responding to direct personal question

ENGLISH EMAIL STYLE (only when Dexter explicitly requests it):
- Opening: "Hi [name]" or "Hi all"
- Sign-off: "Regards,\nDexter"
- Tone: confident, clear, action-oriented. No hedging.

Respond ONLY with a valid JSON object, no markdown, no backticks:
{
  "outputs": [
    {
      "type": "email",
      "to_name": "First name of recipient or empty string",
      "subject": "Short email subject",
      "body": "Full email body text"
    },
    {
      "type": "meeting",
      "title": "Short meeting title",
      "details": "Agenda and any other details as plain text"
    },
    {
      "type": "task",
      "title": "Clear action item",
      "notes": "Any relevant context or empty string"
    }
  ]
}`;

async function processWithClaude(transcript) {
  const message = await anthropic.messages.create({
    model: "claude-sonnet-4-20250514",
    max_tokens: 1024,
    system: STYLE_PROMPT,
    messages: [{ role: "user", content: transcript }],
  });
  const text = message.content[0].text;
  const clean = text.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

async function sendTelegramMessage(chatId, text) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text }),
  });
}

async function handleUpdate(update) {
  const message = update.message;
  if (!message || !message.text) return;

  const chatId = message.chat.id;
  const userId = message.from.id;
  const text = message.text.trim();

  if (userId !== ALLOWED_USER_ID) {
    console.log(`Blocked unauthorized user: ${userId}`);
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
        await createDraft(output.subject, output.body, output.to_name);
        results.push(`Email opgeslagen in Drafts\nOnderwerp: ${output.subject}`);
      } else if (output.type === "meeting") {
        await createCalendarEvent(output.title, output.details);
        results.push(`Meeting toegevoegd aan je agenda\nTitel: ${output.title}`);
      } else if (output.type === "task") {
        await createTask(output.title, output.notes);
        results.push(`Taak aangemaakt in To Do\nTaak: ${output.title}`);
      }
    }

    await sendTelegramMessage(chatId, results.join("\n\n---\n\n"));
  } catch (err) {
    console.error("Error:", err);
    await sendTelegramMessage(chatId, `Er ging iets mis: ${err.message}`);
  }
}

const server = http.createServer(async (req, res) => {
  if (req.method === "POST") {
    let body = "";
    req.on("data", (chunk) => { body += chunk; });
    req.on("end", async () => {
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

server.listen(PORT, () => {
  console.log(`Bot running on port ${PORT}`);
});
