/* MDY Team App settings. None of these are secrets: the app can only read
   or change what the signed-in person is already allowed to on the site. */
window.MDY_CONFIG = {
  clientId: 'd52a2c56-5fe7-45c3-b64b-ca7894636b0a',
  tenantId: 'b1c142e7-4b0a-4f5f-a265-0144f240cfe0',
  siteHost: 'alehzayismdy.sharepoint.com',
  siteUrl: 'https://alehzayismdy.sharepoint.com',
  timeZone: 'Asia/Jerusalem',
  lists: {
    tracker: 'a3b84530-0e04-4c0c-9b37-c09e1c148337',
    deadlines: '97d2c140-5a78-45d6-aa0f-12751b213469',
    feedbackTitle: 'Team Feedback'
  },
  masechtos: ['Kerisos', 'Meilah', 'Tamid', 'Middos', 'Kinnim', 'Temurah'],
  prefixes: { Kerisos: 'KER', Meilah: 'MEI', Tamid: 'TAM', Middos: 'MID', Kinnim: 'KIN', Temurah: 'TEM' },
  resourcesFolder: 'Resources',
  siteAssetsDrive: 'b!-uVDfUf5LEeRULpN1gypU3ud3ezGcDxOi7ykzS1m9EU_rOsoceDARo88MgL_5cmC',
  contactsPage: 'Contacts.aspx',
  /* Volumes in production, shown at the top of the Tracker as on the site's Workflow page.
     A volume shows from its 'from' date (Israel time) until every amud in it is Final. */
  volumes: [
    { name: 'Kerisos', masechtos: ['Kerisos'], from: '' },
    { name: 'Meilah volume', sub: 'Meilah \u00b7 Kinnim \u00b7 Tamid \u00b7 Middos', masechtos: ['Meilah', 'Kinnim', 'Tamid', 'Middos'], from: '2026-10-25' }
  ],
  /* Ask Claude: the team's relay to the Claude API. It holds the key and checks the person's Aleh Zayis sign-in. */
  askRelay: 'https://mdy-claude.yaakovbressler1.workers.dev/',
  chatUrl: 'https://chat.google.com/room/AAQAnjn0HSU?cls=5',
  graphicsStages: ['Requests', 'Writer Requests', 'Prompts', 'Needs Approval', 'Ready for Review', 'Ready for Designer', 'Ready for Designing', 'Pending Charts', 'Ready to Go', 'Rejected'],
  /* The site's colour key, one colour per stage */
  stageColors: {
    'Open': '#e74c3c', 'Generated': '#ffa040', 'Written': '#ffa040', 'Edited': '#8b6348',
    'Reviewed': '#9b4f96', 'Run': '#9b4f96', 'Edited 2': '#1a54a8', 'Mevukar': '#7b1fa2',
    'Kramer': '#607d8b', 'Post Kramer': '#00bcd4', 'Edited 3': '#e91e63', 'Review 2': '#f9e000',
    'Mevukar 2': '#f9e000', 'Proofread': '#a9db5e', 'Check': '#00695c', 'Final': '#1db954', 'Closed': '#1db954'
  }
};
