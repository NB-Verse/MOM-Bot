# MOM Bot: an AI notetaker that joins your meetings under your name

MOM Bot sends an AI notetaker into your team meetings on Microsoft Teams, Zoom, Google Meet and Webex. The bot shows up under your name (for example "Nishchal Bafna"). It records the meeting, transcribes it, and writes the Minutes of Meeting (MOM): a summary, the discussion by topic, decisions, an action-item table with owner, due date and priority, risks, open questions and the next meeting.

There are two ways to use it:

1. **Paste a link.** Open the app, paste the meeting link, and press **Join now**.
2. **Connect your calendar.** The bot joins every meeting in your calendar that has a Teams, Zoom, Meet or Webex link. You don't need to paste anything. You can switch off any meeting you don't want it in.

Everything lives in one folder with no subfolders. There is nothing to install: the app has no dependencies and needs Node.js 20 or newer.

| File | What it is |
|---|---|
| `server.js` | The whole backend: sign-in and invites, the bot, calendar sync, webhooks, transcripts and AI minutes |
| `index.html` | The web app |
| `.env.example` | Every setting, explained |
| `render.yaml` | One-click deploy to Render |
| `Dockerfile` | Optional, for other hosts |

## Quick start

### 1. Put the code on GitHub
Create a repository and upload every file in this folder. Never upload `.env`; it is already excluded.

> GitHub Pages can't run this app, because the bot needs a server that is always on. Keep the code on GitHub and deploy it to Render (step 3). Every push to GitHub then redeploys it automatically.

### 2. Get two API keys
- **Recall.ai** runs the bot that actually joins the meeting. Sign up and choose a region (pay-as-you-go accounts use `us-west-2`). Under **Developers → API Keys & Secrets**, create:
  - an **API key**
  - a **workspace verification secret** (it starts with `whsec_`)
- **Anthropic** writes the minutes. Create a key at console.anthropic.com.

### 3. Deploy on Render
1. In Render, choose **New + → Blueprint** and pick your repository.
2. Fill in these values:
   - `PUBLIC_BASE_URL`: your app address, such as `https://mom-bot.onrender.com`
   - `ADMIN_EMAIL`, `ADMIN_PASSWORD` (at least 10 characters) and `ADMIN_NAME`. Use your real name (for example `Nishchal Bafna`), because the bot uses it.
   - `RECALL_API_KEY`, `RECALL_VERIFICATION_SECRET` and `ANTHROPIC_API_KEY`
3. Render attaches a small disk so that meetings survive restarts. This needs a paid instance.

### 4. Connect Recall.ai to your app
In the Recall dashboard, open **Webhooks → Add Endpoint** and enter:
- **URL:** `https://YOUR-APP/webhooks/recall`
- **Events:** `bot.*`, `recording.done`, `recording.failed`, `transcript.done`, `transcript.failed`, `calendar.update`, `calendar.sync_events`

### 5. Use it
Sign in with your admin email and go to **Calendar & bot**:
- Set the name the bot shows in meetings.
- Connect your calendar (see below).

To add teammates, open **Team & invites** and create invite links. Each teammate gets their own bot name and can connect their own calendar.

## Connecting a calendar

### Easiest: paste your calendar's private link (no extra setup)
- **Google Calendar:** Settings → click your calendar → **Secret address in iCal format** → copy the link.
- **Outlook / Microsoft 365:** Outlook on the web → Settings → Calendar → **Shared calendars** → **Publish a calendar** → choose **Can view all details** → copy the **ICS** link.

Paste the link in **Calendar & bot**. The app checks your calendar every 5 minutes and books the bot up to 24 hours ahead. If a meeting moves, the booking moves with it. If a meeting is cancelled or you decline it, the bot is cancelled too.

What to know about this method:
- Outlook can take a while to refresh a published calendar, so a meeting added at the last minute may be missed. For those, just paste the link on the home page.
- Some companies switch off "Publish a calendar". If yours has, use the sign-in method below.

### Most reliable: "Connect Google Calendar" / "Connect Outlook" (one-time admin setup)
With this method, calendar changes arrive within seconds and Recall.ai manages the bookings. An admin has to register an app once:

- **Google:**
  1. In Google Cloud Console, create an OAuth client (type "Web application") and enable the **Google Calendar API**.
  2. Add `https://YOUR-APP/oauth/callback` as a redirect URI.
  3. Set `GOOGLE_CLIENT_ID` and `GOOGLE_CLIENT_SECRET` in Render.
- **Outlook / Microsoft 365:**
  1. In Azure Portal, go to **App registrations → New registration**.
  2. Add the redirect URI `https://YOUR-APP/oauth/callback` (Web), then create a client secret.
  3. Add these delegated permissions: `Calendars.Read`, `offline_access`, `openid`, `email`.
  4. Set `MICROSOFT_CLIENT_ID` and `MICROSOFT_CLIENT_SECRET` (and `MICROSOFT_TENANT` if you want to limit it to your company).

  Corporate Microsoft 365 accounts often need IT to approve the app once.

Once these are set, the **Connect Google Calendar** and **Connect Outlook** buttons appear in the app. Recall's own guides for this setup are at docs.recall.ai (Calendar V2, Google Calendar and Microsoft Outlook).

### Teammates in the same meeting
If two teammates have the same meeting in their calendars, only one bot joins. Both of them see the transcript and the minutes.

## In the meeting
- Paste the full join link, including any passcode. Outlook Safe Links are unwrapped automatically.
- **Admit the bot from the lobby.** It waits up to 20 minutes. Teams removes lobby guests after 30 minutes.
- Teams must allow external or anonymous guests. Google Meet meetings set to "signed-in users only" will refuse the bot.
- When the bot joins, it posts a pinned chat message saying it is your AI notetaker and is recording. This matters most when the bot uses your exact name, because people may otherwise think it is you.
- The bot leaves when everyone leaves. You can also press **Remove bot from meeting**.
- The minutes usually arrive a few minutes after the call. If anything stalls, the app retries by itself every 30 minutes, or you can press **Sync with bot**.

If your company blocks outside participants entirely, use **Import**: upload the Teams or Zoom transcript file and get the same minutes.

## How future-proof is it?
No meeting tool can promise this forever, because Microsoft, Zoom and Google keep changing their apps. This design keeps the risk low:
- **The hard part is outsourced.** Joining Teams, Zoom and Meet is Recall.ai's full-time job. When those apps change, Recall updates its bots and this app keeps working.
- **No packages.** The app uses only built-in Node.js features, so there are no libraries that go out of date or break. It runs on Node 20, 22 and later.
- **Swappable AI.** Changing the AI model is one setting (`ANTHROPIC_MODEL`), or set `LLM_PROVIDER=openai` to switch providers.
- **Self-repairing.** Every webhook is verified, and repeats are ignored. Failed calls are retried with back-off. Meetings that get stuck are re-checked automatically.
- **Safe data storage.** Each save is complete or not written at all, and a backup copy is kept.

What can still need your attention: your Recall.ai and Anthropic accounts (keys, billing, region), your company's Teams policies, and Render keeping the disk attached.

## Running it on your own computer
```bash
cp .env.example .env      # fill in the values
node server.js            # then open http://localhost:3000
```
Recall.ai can't reach `localhost`. To test the bot, use a fixed public tunnel (for example a static ngrok domain) as `PUBLIC_BASE_URL`.

## Privacy
Recording people requires their awareness. Tell participants about the notetaker, and follow your company's policy and the law (for example India's DPDP Act). Deleting a meeting in the app also deletes the recording stored at Recall.ai.

## Security
- Accounts are invite-only. Invite links work once and expire, and you can limit sign-ups to your email domain.
- Passwords are hashed, sign-in cookies are signed, and repeated failed sign-ins are slowed down.
- Admins can disable a user, which signs them out immediately.
- Every message from Recall.ai is checked with a secret signature.
- Calendar access is read-only.
- Calendar links that point to private network addresses are refused.

## Costs
Recall.ai charges per hour of bot time plus transcription. Anthropic charges per use. Render's paid instance with a disk is a small monthly fee. Check each provider's current pricing before rolling out to a big team.
