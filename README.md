# MDY Team App

The MDY Aleh Zayis team app: your amudim, the Amud Tracker, submitting work, the Guide look-ups, deadlines, graphics, feedback and contacts, on phone and computer.

**Open it:** https://yaakov-bressler.github.io/mdy-team-app/

## Installing it

- **iPhone:** open the link in Safari, tap Share, then **Add to Home Screen**.
- **Android:** open it in Chrome, tap the ⋮ menu, then **Install app**.
- **Computer:** open it in Chrome or Edge and click the install icon at the right end of the address bar.

Sign in with your Aleh Zayis account. Updates arrive on their own.

## How it works

The app is a static page. It signs each person in with Microsoft and talks to the team SharePoint site through Microsoft Graph **as that person**, so it can only see or change what they already can on the site. No team data is stored in this repository or anywhere else.

- **Tracker:** reads and updates the Amud Tracker list (any team member can move a stage).
- **Submit:** uploads the file into the masechta's folder (Editors › your folder by default), named in house form, and moves the amud's stage.
- **Guide:** reads the Transliteration list, Definitions list, House Rulings and Style Guide live from Resources.
- **Deadlines, Feedback:** the site's Deadlines and Team Feedback lists. Feedback stays private to the person who sent it.
- **Graphics:** the masechta's Graphics stage folders and the pictures in them.
- **Contacts:** the site's own Contacts page.
- **Ask Claude:** the same assistant as the site's Ask Claude page, for style, process, edits and corpus reports. Questions go to the team's relay (`askRelay` in `config.js`), which holds the Claude API key and lets in only people signed in with an Aleh Zayis account. The app runs Claude's searches itself, over the rulebooks, the corpus and the transliteration and definitions lists; nothing about people, the tracker or deadlines is sent.

Settings (app registration, list ids) are in `config.js`. None of them are secrets.
