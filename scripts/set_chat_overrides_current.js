// Paste this in the Teams page console (top frame) to pin the chat ID and host/region overrides.
(function () {
  localStorage.setItem(
    'teamsChatConversationIdOverride',
    '19:1511c8c5-cad7-47bb-98ae-74211b1cd404_46de1a27-3106-478b-bd49-6f675f88848d@unq.gbl.spaces'
  );
  localStorage.setItem('teamsChatHostOverride', 'teams.microsoft.com');
  localStorage.setItem('teamsChatRegionOverride', 'ca');
  console.log('[Overrides set] conversation/host/region pinned.');
})();
