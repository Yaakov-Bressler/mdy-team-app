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

- **Tracker:** reads and updates the Amud Tracker list (any team member can move a stage). At the top, as on the site's Workflow page, one bar per volume in production (`volumes` in `config.js`), with Final and Closed shown together as Final. The Meilah volume (Meilah, Kinnim, Tamid, Middos) runs on eight stages only, set in `workflows` in `config.js`: Open, Generated, Edited, Mevukar, Kramer, Edited 2, Proofread, Final. The move sheet, Submit and the bars offer only those for its amudim. Kerisos (28 Oct 2026) has its own two lists, picked by the row's Workflow column: New is Open, Generated, Edited, Reviewed, Mevukar, Kramer, Edited 2, Proofread, Final; Old is Open, Written, Edited, Mevukar, Kramer, Edited 2, Proofread, Final. Run, Post Kramer, Mevukar 2, Edited 3 and Review 2 are gone from Kerisos; a row still on one of them stays there until someone moves it, and Next/Submit suggest the stage that replaced it (`retiredNext`).
- **Submit:** uploads the file into the masechta's folder (Editors › your folder by default), named in house form, and moves the amud's stage.
- **Guide:** reads the Transliteration list, Definitions list, House Rulings and Style Guide live from Resources. Phrases and Headers use the same index as the site's Guide page (Site Assets › MDYLookup), including the English lines with no Hebrew before them (`lines-<masechta>.json`).
- **Deadlines, Feedback:** the site's Deadlines and Team Feedback lists. Feedback stays private to the person who sent it.
  A deadline puts amudim on its owner's *My work* by its Masechta, Amudim and Done when columns, read the same way as the site's tally (`Claude/Deadlines/deadlines-core.js`): a volume ("Meilah volume" = Meilah, Kinnim, Tamid, Middos) or a list of masechtos, `all`, ranges, `all except …`, and terms naming a masechta. Only amudim not yet at the deadline's stage are listed. A deadline is yours when your surname is in its Who column or its title.
- **Graphics:** the masechta's Graphics stage folders and the pictures in them.
- **Contacts:** the site's own Contacts page.
- **Ask Chaim Levy** (formerly Ask Claude): the AI assistant, the same as the site's Ask Chaim Levy page (`SitePages/Ask-Chaim-Levy.aspx`), runs on Claude, for style, process, edits and corpus reports. Questions go to the team's relay (`askRelay` in `config.js`), which holds the Claude API key and lets in only people signed in with an Aleh Zayis account. The app runs Claude's searches itself, over the rulebooks, the corpus and the transliteration and definitions lists; nothing about people, the tracker or deadlines is sent. Its source texts (Claude/Ask Claude data) cover the Gemara or Mishnah, the commentary of record (Rashi; Tosafos on Meilah, the Rosh on Tamid, the Rambam on Kinnim and Middos) and the secondary commentaries, the Mesivta, and the delivered English of whole amudim, plus MDY's own Bechoros, which is labelled as not house style wherever it appears.
  Since 28 Oct 2026 (`ask-tools.js`, switched on by `askExtras` in `config.js` once the relay offers the tools in `relay-tools.json`) he can also find and read live files in the library (drafts, edited amudim, sheets) as the person asking, take attachments in the question box (Word, PDF, Excel, text, pictures), hand back a Word file with his edits as tracked changes and comments (downloaded on the device, never saved to SharePoint), and search the team chat log and the Proofreading and Graphics feedback sheets. To keep the cost down, files are read in parts of about 9,000 characters, an attachment goes with the question only as a short beginning, pictures are capped at two a question and dropped from the history after the answer, and the search results inside one question are held under 80,000 characters.

Settings (app registration, list ids) are in `config.js`. None of them are secrets.
