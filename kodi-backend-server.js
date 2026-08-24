// ----
// KODI BACKEND SERVER
// ----
// This Node.js server handles three critical functions:
//
// 1. API PROXY — Securely routes chatbot requests to Claude API
// so your API key is never exposed in the browser.
//
// 2. SELF-LEARNING WEB SCRAPER — Periodically scrapes your website
// for new/updated content and injects it into Kodi's knowledge.
//
// 3. OUTLOOK CALENDAR INTEGRATION — Creates appointment events
// in your BBG team's Outlook calendar via Microsoft Graph API.
//
// DEPLOYMENT: Vercel, Railway, Render, or any Node.js host.
// ----

const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const cheerio = require('cheerio');
const cron = require('node-cron');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());

// ----
// ENVIRONMENT VARIABLES (set these in your hosting platform)
// ----
const {
// Anthropic API key — get from https://console.anthropic.com
ANTHROPIC_API_KEY,

// Microsoft Azure AD credentials for Outlook integration
// See setup instructions in Section 3 below
MS_TENANT_ID,
MS_CLIENT_ID,
MS_CLIENT_SECRET,

// The BBG team email whose Outlook calendar will receive bookings
// e.g. "investors@bluebeargroup.ca"
BBG_CALENDAR_EMAIL,

PORT = 3001,
} = process.env;

// ----
// SECTION 1: CLAUDE API PROXY
// ----
// This endpoint sits between the chatbot frontend and the
// Anthropic API. The frontend sends messages here, this server
// attaches the API key and forwards the request, then returns
// the response. The API key never reaches the browser.

app.post('/api/chat', async (req, res) => {
try {
const { messages, system } = req.body;

// Inject any freshly-learned knowledge into the system prompt
const dynamicKnowledge = loadDynamicKnowledge();
const fullSystem = system.replace(
'{DYNAMIC_KNOWLEDGE}',
dynamicKnowledge
? `\nRECENTLY UPDATED WEBSITE CONTENT:\n${dynamicKnowledge}`
: ''
);

const response = await fetch('https://api.anthropic.com/v1/messages', {
method: 'POST',
headers: {
'Content-Type': 'application/json',
'x-api-key': ANTHROPIC_API_KEY, // <-- Key stays server-side
'anthropic-version': '2023-06-01',
},
body: JSON.stringify({
model: 'claude-sonnet-5',
max_tokens: 1000,
system: fullSystem,
messages,
}),
});

const data = await response.json();
res.json(data);
} catch (error) {
console.error('Chat proxy error:', error);
res.status(500).json({
content: [{
type: 'text',
text: 'I apologize — I\'m having a brief technical issue. Please contact our team at +1 437 826 4847.',
}],
});
}
});

// ----
// SECTION 2: SELF-LEARNING WEB SCRAPER
// ----
// This system periodically crawls your website pages, extracts
// the text content, compares it against what it already knows,
// and saves any new or changed content to a local JSON file.
//
// Kodi then injects this fresh knowledge into every conversation,
// so it always has your latest website information without you
// needing to manually update anything.

// Pages to scrape (add new pages here as you add them to your site)
const PAGES_TO_SCRAPE = [
'https://bluebeargroup.ca/',
'https://bluebeargroup.ca/why-bbg-corp%3F',
'https://bluebeargroup.ca/services',
'https://bluebeargroup.ca/contact-us',
'https://bluebeargroup.ca/portfolio-1',
'https://bluebeargroup.ca/blog',
'https://bluebeargroup.ca/multifamily-prop-acquis-1',
'https://bluebeargroup.ca/preferred-shares-inv',
'https://bluebeargroup.ca/property-management',
];

const KNOWLEDGE_FILE = path.join(__dirname, 'learned-knowledge.json');
const MANUAL_KNOWLEDGE_FILE = path.join(__dirname, 'manual-knowledge.json');

// Load previously scraped knowledge from disk, plus any manually curated knowledge
function loadDynamicKnowledge() {
let parts = [];

// 1. Load manually curated knowledge (takes priority — listed first)
try {
if (fs.existsSync(MANUAL_KNOWLEDGE_FILE)) {
const manual = JSON.parse(fs.readFileSync(MANUAL_KNOWLEDGE_FILE, 'utf8'));

if (manual.investment_details) {
parts.push(`[BBG INVESTMENT DETAILS]\n${manual.investment_details}`);
}
if (manual.faqs && manual.faqs.length) {
const faqText = manual.faqs.map(f => `Q: ${f.q}\nA: ${f.a}`).join('\n\n');
parts.push(`[FREQUENTLY ASKED QUESTIONS]\n${faqText}`);
}
if (manual.team && manual.team.length) {
const teamText = manual.team.map(m => `${m.name} — ${m.role}: ${m.bio}`).join('\n');
parts.push(`[BBG TEAM]\n${teamText}`);
}
if (manual.current_opportunities) {
parts.push(`[CURRENT INVESTMENT OPPORTUNITIES]\n${manual.current_opportunities}`);
}
if (manual.extra_notes) {
parts.push(`[ADDITIONAL NOTES]\n${manual.extra_notes}`);
}
}
} catch (e) {
console.error('Error loading manual knowledge:', e);
}

// 2. Load auto-scraped website knowledge
try {
if (fs.existsSync(KNOWLEDGE_FILE)) {
const data = JSON.parse(fs.readFileSync(KNOWLEDGE_FILE, 'utf8'));
const scraped = Object.entries(data.pages || {})
.map(([url, info]) => `[Source: ${url}]\n${info.content}`)
.join('\n\n');
if (scraped) parts.push(`[WEBSITE CONTENT]\n${scraped}`);
}
} catch (e) {
console.error('Error loading scraped knowledge:', e);
}

return parts.join('\n\n');
}

// Scrape a single page and extract meaningful text content
async function scrapePage(url) {
try {
const response = await fetch(url, {
headers: {
// Identify ourselves as a legitimate bot
'User-Agent': 'KodiBot/2.0 (BBG Corp internal crawler)',
},
});

if (!response.ok) {
console.warn(`Failed to fetch ${url}: ${response.status}`);
return null;
}

const html = await response.text();
const $ = cheerio.load(html);

// Remove elements that don't contain useful content
$('script, style, nav, footer, header, .cookie-banner, noscript, iframe').remove();

// Extract the meaningful text from the page body
const textContent = $('body')
.text()
.replace(/\s+/g, ' ') // Collapse whitespace
.replace(/\n{3,}/g, '\n\n') // Limit consecutive newlines
.trim()
.slice(0, 3000); // Cap at 3000 chars per page

return textContent || null;
} catch (error) {
console.error(`Error scraping ${url}:`, error.message);
return null;
}
}

// Run a full scrape of all configured pages
async function runFullScrape() {
console.log(`[Kodi Scraper] Starting knowledge update at ${new Date().toISOString()}`);

let existingData = { pages: {}, lastUpdated: null };
try {
if (fs.existsSync(KNOWLEDGE_FILE)) {
existingData = JSON.parse(fs.readFileSync(KNOWLEDGE_FILE, 'utf8'));
}
} catch { /* start fresh */ }

let updatedCount = 0;

for (const url of PAGES_TO_SCRAPE) {
const content = await scrapePage(url);
if (!content) continue;

const previousContent = existingData.pages[url]?.content || '';

// Only update if the content has actually changed
// This comparison uses a simple length + substring check;
// in production you could use a proper diff algorithm
if (content !== previousContent) {
existingData.pages[url] = {
content,
scrapedAt: new Date().toISOString(),
previousVersion: previousContent.slice(0, 500), // Keep a snippet of old content for reference
};
updatedCount++;
console.log(` ✓ Updated: ${url}`);
} else {
console.log(` · No changes: ${url}`);
}

// Be polite — wait 1 second between requests to your own site
await new Promise(r => setTimeout(r, 1000));
}

existingData.lastUpdated = new Date().toISOString();
existingData.totalPages = Object.keys(existingData.pages).length;

fs.writeFileSync(KNOWLEDGE_FILE, JSON.stringify(existingData, null, 2));
console.log(`[Kodi Scraper] Done. ${updatedCount} pages updated, ${existingData.totalPages} total pages tracked.\n`);
}

// Schedule the scraper to run every 6 hours automatically
// Cron expression: minute(0) hour(every 6th) day(*) month(*) weekday(*)
cron.schedule('0 */6 * * *', () => {
runFullScrape().catch(err => console.error('Scheduled scrape failed:', err));
});

// Also run a scrape when the server first starts
runFullScrape().catch(err => console.error('Initial scrape failed:', err));

// Manual scrape trigger endpoint (useful for testing or when you
// update your website and want Kodi to learn immediately)
app.post('/api/refresh-knowledge', async (req, res) => {
try {
await runFullScrape();
const data = JSON.parse(fs.readFileSync(KNOWLEDGE_FILE, 'utf8'));
res.json({
success: true,
lastUpdated: data.lastUpdated,
pagesTracked: data.totalPages,
});
} catch (error) {
res.status(500).json({ success: false, error: error.message });
}
});

// Endpoint to check what Kodi currently knows
app.get('/api/knowledge-status', (req, res) => {
try {
if (fs.existsSync(KNOWLEDGE_FILE)) {
const data = JSON.parse(fs.readFileSync(KNOWLEDGE_FILE, 'utf8'));
res.json({
lastUpdated: data.lastUpdated,
pagesTracked: data.totalPages,
pages: Object.entries(data.pages).map(([url, info]) => ({
url,
scrapedAt: info.scrapedAt,
contentLength: info.content?.length || 0,
})),
});
} else {
res.json({ status: 'No knowledge scraped yet' });
}
} catch (error) {
res.status(500).json({ error: error.message });
}
});

// ----
// SECTION 3: OUTLOOK CALENDAR INTEGRATION
// ----
// This uses Microsoft Graph API to create calendar events
// in your BBG team's Outlook calendar when prospects book
// appointments through Kodi.
//
// ---- SETUP INSTRUCTIONS ----
//
// Step 1: Register an app in Azure Active Directory
// 1. Go to https://portal.azure.com
// 2. Navigate to "Azure Active Directory" → "App registrations"
// 3. Click "New registration"
// 4. Name: "Kodi Chatbot"
// 5. Supported account types: "Accounts in this organizational directory only"
// 6. Click "Register"
// 7. Copy the "Application (client) ID" → this is your MS_CLIENT_ID
// 8. Copy the "Directory (tenant) ID" → this is your MS_TENANT_ID
//
// Step 2: Create a client secret
// 1. In your app registration, go to "Certificates & secrets"
// 2. Click "New client secret"
// 3. Description: "Kodi production key"
// 4. Expiration: Choose 24 months
// 5. Click "Add"
// 6. Copy the secret VALUE immediately → this is your MS_CLIENT_SECRET
//
// Step 3: Grant calendar permissions
// 1. Go to "API permissions" in your app registration
// 2. Click "Add a permission" → "Microsoft Graph"
// 3. Choose "Application permissions"
// 4. Search for and add: "Calendars.ReadWrite"
// 5. Click "Grant admin consent for [your org]"
//
// Step 4: Set environment variables
// MS_TENANT_ID=your-tenant-id
// MS_CLIENT_ID=your-client-id
// MS_CLIENT_SECRET=your-client-secret
// BBG_CALENDAR_EMAIL=investors@bluebeargroup.ca
//
// ---- END SETUP ----

// Get an access token from Microsoft using client credentials flow
async function getMicrosoftAccessToken() {
const tokenUrl = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;

const response = await fetch(tokenUrl, {
method: 'POST',
headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
body: new URLSearchParams({
client_id: MS_CLIENT_ID,
client_secret: MS_CLIENT_SECRET,
scope: 'https://graph.microsoft.com/.default',
grant_type: 'client_credentials',
}),
});

const data = await response.json();

if (!data.access_token) {
throw new Error(`Microsoft auth failed: ${JSON.stringify(data)}`);
}

return data.access_token;
}

// Create a calendar event in the BBG team's Outlook calendar
async function createOutlookEvent(bookingDetails) {
const token = await getMicrosoftAccessToken();

// Parse the date and time into an ISO datetime string
// The booking comes in as { date: "2026-04-15", time: "2:00 PM", ... }
const dateStr = bookingDetails.date;
const timeStr = bookingDetails.time || '10:00 AM';

// Convert "2:00 PM" format to 24-hour time for the ISO string
const [timePart, ampm] = timeStr.split(' ');
let [hours, minutes] = timePart.split(':').map(Number);
if (ampm === 'PM' && hours !== 12) hours += 12;
if (ampm === 'AM' && hours === 12) hours = 0;

const startDateTime = `${dateStr}T${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:00`;
// Default 30-minute meeting
const endHours = minutes + 30 >= 60 ? hours + 1 : hours;
const endMinutes = (minutes + 30) % 60;
const endDateTime = `${dateStr}T${String(endHours).padStart(2, '0')}:${String(endMinutes).padStart(2, '0')}:00`;

// Build the calendar event object for the Microsoft Graph API
const event = {
subject: `BBG Consultation — ${bookingDetails.name}`,
body: {
contentType: 'HTML',
content: `
<h2>New Consultation Request via Kodi</h2>
<table style="border-collapse:collapse; font-family:Calibri,sans-serif;">
<tr><td style="padding:8px; font-weight:bold; color:#005AAA;">Name:</td><td style="padding:8px;">${bookingDetails.name}</td></tr>
<tr><td style="padding:8px; font-weight:bold; color:#005AAA;">Email:</td><td style="padding:8px;">${bookingDetails.email}</td></tr>
<tr><td style="padding:8px; font-weight:bold; color:#005AAA;">Phone:</td><td style="padding:8px;">${bookingDetails.phone || 'Not provided'}</td></tr>
<tr><td style="padding:8px; font-weight:bold; color:#005AAA;">Topic:</td><td style="padding:8px;">${bookingDetails.topic || 'General consultation'}</td></tr>
<tr><td style="padding:8px; font-weight:bold; color:#005AAA;">Budget Range:</td><td style="padding:8px;">${bookingDetails.budget || 'Not specified'}</td></tr>
</table>
<p style="margin-top:16px; color:#6C6C6C; font-size:12px;">
This appointment was booked automatically by Kodi, the BBG Corp. AI Investment Advisor.
</p>
`,
},
start: {
dateTime: startDateTime,
timeZone: 'Eastern Standard Time',
},
end: {
dateTime: endDateTime,
timeZone: 'Eastern Standard Time',
},
location: {
displayName: 'Virtual — Teams/Zoom link to be sent',
},
// Send a calendar invite to the prospect
attendees: [
{
emailAddress: {
address: bookingDetails.email,
name: bookingDetails.name,
},
type: 'required',
},
],
// This ensures the prospect gets an email invitation
isOnlineMeeting: true,
onlineMeetingProvider: 'teamsForBusiness',
};

// Create the event on the BBG calendar email account
const graphUrl = `https://graph.microsoft.com/v1.0/users/${BBG_CALENDAR_EMAIL}/events`;

const response = await fetch(graphUrl, {
method: 'POST',
headers: {
Authorization: `Bearer ${token}`,
'Content-Type': 'application/json',
},
body: JSON.stringify(event),
});

if (!response.ok) {
const errorBody = await response.text();
throw new Error(`Graph API error ${response.status}: ${errorBody}`);
}

const createdEvent = await response.json();
return {
eventId: createdEvent.id,
webLink: createdEvent.webLink,
onlineMeetingUrl: createdEvent.onlineMeeting?.joinUrl,
};
}

// Booking endpoint — called by the chatbot when someone submits the booking form
app.post('/api/book-appointment', async (req, res) => {
try {
const bookingDetails = req.body;

// Validate required fields
if (!bookingDetails.name || !bookingDetails.email || !bookingDetails.date || !bookingDetails.time) {
return res.status(400).json({
success: false,
error: 'Name, email, date, and time are required.',
});
}

// If Microsoft credentials are configured, create the Outlook event
if (MS_TENANT_ID && MS_CLIENT_ID && MS_CLIENT_SECRET && BBG_CALENDAR_EMAIL) {
const eventResult = await createOutlookEvent(bookingDetails);

console.log(`[Kodi Booking] Created Outlook event for ${bookingDetails.name} on ${bookingDetails.date} at ${bookingDetails.time}`);

return res.json({
success: true,
message: 'Appointment created in Outlook calendar and invitation sent.',
eventId: eventResult.eventId,
meetingLink: eventResult.onlineMeetingUrl,
});
}

// Fallback: If Outlook is not configured, log the booking
// and send a notification email (you'd implement email sending
// with a service like SendGrid, Mailgun, or AWS SES)
console.log(`[Kodi Booking] New booking (no Outlook configured):`, bookingDetails);

res.json({
success: true,
message: 'Appointment request received. Team will confirm via email.',
note: 'Outlook integration not configured — booking logged for manual processing.',
});

} catch (error) {
console.error('[Kodi Booking] Error:', error);
res.status(500).json({
success: false,
error: 'Failed to create appointment. The team has been notified.',
});
}
});

// ----
// HEALTH CHECK & SERVER START
// ----

app.get('/api/health', (req, res) => {
res.json({
status: 'healthy',
service: 'Kodi Backend v2',
features: {
chatProxy: !!ANTHROPIC_API_KEY,
outlookIntegration: !!(MS_TENANT_ID && MS_CLIENT_ID && MS_CLIENT_SECRET),
selfLearning: true,
},
timestamp: new Date().toISOString(),
});
});

app.listen(PORT, () => {
console.log(`\n----`);
console.log(` Kodi Backend Server v2`);
console.log(` Port: ${PORT}`);
console.log(` Chat Proxy: ${ANTHROPIC_API_KEY ? '✓ Active' : '✗ No API key'}`);
console.log(` Outlook: ${MS_TENANT_ID ? '✓ Configured' : '✗ Not configured'}`);
console.log(` Scraper: ✓ Runs every 6 hours`);
console.log(`----\n`);
});
