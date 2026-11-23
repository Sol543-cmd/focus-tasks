# Focus Tasks for your phone: setup

**Keep your own config.js in this folder** (the one with your App ID), so you never have to look it up again.

This version adds: an Undated tab, a Projects list (tap one to see just that project, by section), the
new shortcuts (@label, #Project or its shortcut, /Section), a small calendar (the 📅 icon) with your
meetings and plan, and three themes (under ⋯).

(Hardened version details below.)

## New: opens instantly, syncs in the background, and a Size setting

- **Opens instantly.** The app shows your last copy straight away, then syncs with OneDrive behind it.
- **Changes save instantly.** Adding, finishing or editing a task shows at once and is kept on the phone,
  then sent to OneDrive in the background (on the newest copy, so your computers' changes aren't lost).
  Offline? Changes wait on the phone and go when you're back online, even if you close the app.
- **No surprise trips to Microsoft.** About once a day Microsoft wants a quick sign-in check (its rule for
  web apps). Instead of jumping away, the app keeps working and shows a **Reconnect** button: one tap,
  a moment on Microsoft's page, and you're back.
- **Size** (⋯ menu, at the top): Small, Normal, Large, Larger, Largest. Scales text, buttons and spacing
  everywhere, Training included.

## New: Training (the 📖 button at the top)

Your Training from Focus, on your phone and Boox: today's lesson, Books practice, Flashcards, Timeline,
Greek words, Alphabet, Greek concepts, the Concept quiz, your library (with reading and highlights) and this
week's numbers, all on one dashboard.

- **Same progress everywhere.** This phone saves to its own file, `Focus/training/progress-web-….json`, the
  same way each computer has its own. Your computers read it, so a lesson finished here is finished there
  (and the other way round). Nothing here can overwrite a computer's progress.
- **Offline.** Your books, chapter packs and pictures are kept on the phone. Study without a connection;
  your answers save to OneDrive when you're back online (the dot at the top says Synced, Waiting to save or Offline).
- **Reading** uses Flint's page turning: tap the right side or swipe to turn, tap the middle for the bar,
  **Aa** for font, size, page color (white, cream, green, dark, black) and margins. Select text to highlight.
- **Boox / e-ink:** Training settings (⚙), E-ink screen: On (Auto turns it on for Boox). No animations,
  black on white, instant page turns.
- Lesson days, days off and adding books stay on your computer.

**Upload these files** with the rest (all at once): `training.css`, `training-data.js`,
`training-reader.js`, `training-ui.js`, and the **`fonts` folder** (drag the folder in), plus the updated
`index.html` and `sw.js`. Keep your own `config.js`.

The first time you open Training it downloads your books (do it on Wi-Fi: pictures can be large).


Ten files (plus your config.js) go into your GitHub repository (upload them all at once):

| File | What it is |
|---|---|
| index.html | The page |
| app.js | The app (screens, sign-in, saving) |
| logic.js | Task logic, the same as Focus on your computer |
| msal-browser.min.js | Microsoft's official sign-in library (v2.38.3), kept here instead of loaded from the internet |
| config.js | Your settings: the only file you edit |
| manifest.json, sw.js | Let your phone install the page as an app |
| icon-192.png, icon-512.png, icon-maskable-512.png, apple-touch-icon.png | The app icon |

## What makes it hardened

- **No outside code.** Every script comes from your own repository, and the page's security
  rules forbid loading code from anywhere else.
- **Tamper checks.** The browser checks that the page's code files are exactly the ones you
  were given, and refuses to run them if anything has changed.
- **Talks only to Microsoft.** The page can only connect to Microsoft sign-in and OneDrive.
- **Only Focus's files.** The page refuses to read or write anything except Focus's task files in the
  Focus folder. Training may read Focus/training and the Greek files, and may write only this phone's own
  progress file and your Greek practice (vocab.json, alphabet.json, concepts.json).
- **Stays signed in.** Your sign-in is remembered on this phone only (keep a screen lock on it).
  Every so often Microsoft does a quick check that passes through its page and straight back,
  usually without typing. Tap ⎋ to sign out and forget it.

## Setup

1. **Turn on two-step verification on GitHub** (your picture at the top right, then Settings,
   then Password and authentication, then Enable two-factor authentication). This protects the
   page itself.
2. **Put your ID in config.js.** Open it in Notepad and replace PASTE-YOUR-CLIENT-ID-HERE with
   your Application (client) ID. Save.
3. **Upload all ten files.** In your repository, click Add file, then Upload files, drag in all
   ten at once, and click Commit changes. (If you already uploaded the earlier index.html, this
   replaces it.)
4. **GitHub Pages** stays as you set it up (main, /(root)). Your link doesn't change.
5. **The Microsoft app registration** stays the same too (same link, Files.ReadWrite).
6. Wait a minute, open your link on your phone, and sign in.
7. **Install it as an app.** Android (Chrome): tap the menu (three dots), then **Install app**
   (or accept the "Install" banner). iPhone (Safari): tap Share, then **Add to Home Screen**.
   It opens full screen with its own icon, like any other app.

## Updating later

When you get new files, upload everything except config.js (keep yours). This update changes
app.js, logic.js, index.html and sw.js. Never edit app.js,
logic.js or msal-browser.min.js by hand: the tamper check will stop the page from running if
they change.

## Removing access any time

Sign out with the ⎋ button, remove the app at myapps.microsoft.com (or account.live.com for a
personal account), or delete the app registration.


## Updating to the version with Personal tasks
Upload these five files to your GitHub repository (keep your config.js): index.html, app.js, logic.js, sw.js, manifest.json.
Then close the app on your phone completely and open it again (twice, if it still looks the same).
Swipe sideways to switch between Tasks and Personal; the dots under the title show where you are.
The ⋯ menu sets your theme, and Personal's own theme and accent color.

## The red bar at the bottom of the screen
That bar is Android's own navigation bar; an installed app takes its color from the app's settings, which Android
only reads when the app is installed. After uploading this version: remove Focus from your home screen, open your
Focus link in Chrome, and choose Install app (or Add to Home screen) again. Your tasks are in OneDrive, so nothing is lost.

## New icon and look
The icon is now round, so it looks clean in your phone's circle. Phones only pick up a new icon when the app is
installed again: remove Focus from your home screen, open your Focus link in Chrome, and install it again.

## Google Calendar for Personal (two-way)
Personal's calendar is your Google Calendar: you see your events there (Day or Month), and events you add, change or
delete in the app change in Google too. The Tasks page keeps your work calendar from Focus. One-time setup (about 10 minutes):

1. Go to https://console.cloud.google.com and sign in with the Google account whose calendar you want.
2. Create a project (top bar, "Select a project", then "New project"); call it Focus.
3. Open "APIs & Services", then "Library", search for "Google Calendar API", and click Enable.
4. Open "APIs & Services", then "OAuth consent screen" (or "Google Auth Platform"): choose External, name it Focus,
   use your email address, and under "Audience" (or "Test users") add your own Google address.
5. Open "Credentials", then "Create credentials", then "OAuth client ID". Choose "Web application".
   - Authorized JavaScript origins: your site, like https://yourname.github.io
   - Authorized redirect URIs: the exact address you open the app at, like https://yourname.github.io/focus-tasks/
6. Copy the Client ID it shows (it ends in .apps.googleusercontent.com).
7. In the app: swipe to Personal, tap the three dots, paste the Client ID under Google Calendar, and tap Connect.
   Google may say the app isn't verified (it's your own): tap Continue.

Google's sign-in lasts about an hour; when it runs out, the calendar shows a Connect button (usually one tap).
If you'd rather keep the Client ID in config.js, add:  googleClientId: "your-id.apps.googleusercontent.com",
