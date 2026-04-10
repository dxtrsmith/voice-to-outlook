import Anthropic from "@anthropic-ai/sdk";
import http from "http";

const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const PORT = process.env.PORT || 3000;

const anthropic = new Anthropic({ apiKey: ANTHROPIC_API_KEY });

const STYLE_PROMPT = `You are Dexter's personal voice-to-Outlook assistant. Dexter Smith is a Product Director at IceMobile, working with clients like Kruidvat (AS Watson) and Albert Heijn.

TASK: Analyze the Dutch transcript, detect the intent(s), and produce the correct formatted output(s).

INTENT DETECTION — detect one or more of:
- EMAIL: Dexter wants to send an email to someone
- MEETING: Dexter wants to schedule a meeting or calendar event
- TASK: Dexter wants to add a to-do item
- MIXED: Multiple intents — produce multiple outputs

STRICT RULES — NEVER BREAK THESE:
- Never auto-send, never invite attendees, never notify anyone. Emails go to Drafts only.
- Only email who Dexter explicitly mentions. If he says "mail Dennis", only Dennis. Never assume others.
- Default language is ALWAYS Dutch. Only switch to English if Dexter explicitly says so (e.g. "zet het in het Engels").
- Never assume English based on the recipient's name or company.
- No em-dashes. Ever.

DUTCH EMAIL STYLE:
- Opening: always "Hi [name]," or "Hi," — never "Beste" or "Geachte"
- Sign-off: "Groetjes, Dexter"
- Tone: direct, friendly, no padding. Point in first sentence.
- Sentences: short and declarative. No long subordinate clauses.
- Lists: bullet points for multiple items, never prose
- Key phrases: "Zie bijgevoegd", "Zie hier", "Ik hoor graag", "Helder en thanks", "Groetjes"
- Mixes languages casually: "thanks" or "copy changes" inside Dutch sentences is fine
- No fluff: no "Ik hoop dat je een fijn weekend hebt gehad" unless responding to that question
- Short confirmations: "Goed hoor!", "Klopt!", "Helder!"

ENGLISH EMAIL STYLE (only when Dexter explicitly requests it):
- Opening: "Hi [name]" or "Hi all"
- Sign-off: "Regards, Dexter"
- Tone: confident, clear, action-oriented. No hedging.
- Replies: 1-3 sentences, answers then closes
- Closers: "Let me know!", "Thanks.", "Let me know! Thanks."
- No over-explanation
- Longer emails: use bullet points, never walls of text

MEETING OUTPUT FORMAT:
Title: [short descriptive title]
Date/time: [as mentioned, or "to be confirmed"]
Duration: [as mentioned, or "to be confirmed"]
Location/link: [as mentioned, or "to be confirmed"]
Attendees: [only who Dexter explicitly mentioned]
Agenda:
- [bullet points of topics]
Note: Add to Dexter's calendar only. Do not invite anyone.

TASK OUTPUT FORMAT:
Task: [clear action item]
Due: [if mentioned, otherwise "no due date"]
Notes: [any relevant context]

FORMAT YOUR RESPONSE as plain text with clear section headers if there are multiple outputs. Use --- to separate multiple outputs. Keep it clean and copy-paste ready.`;

async function sendTelegramMessage(chatId, text) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text }),
  });
  return res.json();
}

async function processWithClaude(transcript) {
  const message = await anthropic.messages.create({
    model: "claude-sonnet-4-20250514",
    max_tokens: 1024,
    system: STYLE_PROMPT,
    messages: [{ role: "user", content: transcript }],
  });
  return message.content[0].text;
}

async function handleUpdate(update) {
  const message = update.message;
  if (!message || !message.text) return;

  const chatId = message.chat.id;
  const text = message.text.trim();

  if (text === "/start") {
    await sendTelegramMessage(
      chatId,
      "Stuur je transcript en ik maak er een email, meeting of taak van."
    );
    return;
  }

  await sendTelegramMessage(chatId, "Even verwerken...");

  try {
    const result = await processWithClaude(text);
    await sendTelegramMessage(chatId, result);
  } catch (err) {
    console.error("Error:", err);
    await sendTelegramMessage(
      chatId,
      "Er ging iets mis. Probeer het opnieuw."
    );
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
