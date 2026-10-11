# Focus Tasks for your phone (hardened version): setup

Ten files go into your GitHub repository (upload them all at once):

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
app.js, index.html and sw.js. Never edit app.js,
logic.js or msal-browser.min.js by hand: the tamper check will stop the page from running if
they change.

## Removing access any time

Sign out with the ⎋ button, remove the app at myapps.microsoft.com (or account.live.com for a
personal account), or delete the app registration.
