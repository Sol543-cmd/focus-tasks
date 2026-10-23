# Focus Tasks for your phone: setup

**Keep your own config.js in this folder** (the one with your App ID), so you never have to look it up again.

This version adds: an Undated tab, a Projects list (tap one to see just that project, by section), the
new shortcuts (@label, #Project or its shortcut, /Section), a small calendar (the 📅 icon) with your
meetings and plan, and three themes (under ⋯).

(Hardened version details below.)

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
- **Only Focus's files.** The page refuses to read or write anything except Focus's three task
  files (task-bin.json, resolved-tasks.json, projects.json) in the Focus folder.
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
