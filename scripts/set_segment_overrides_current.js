// Paste in Teams console (top frame) to set conservative segment-walker overrides for your chat.
(function () {
  localStorage.setItem(
    'teamsChatConversationIdOverride',
    '19:22d90d37-8f68-497f-a6df-d604f3dca807_46de1a27-3106-478b-bd49-6f675f88848d@unq.gbl.spaces'
  );
  localStorage.setItem('teamsChatSegmentMaxSteps', '3');      // small to avoid crash
  localStorage.setItem('teamsChatSegmentPageSize', '20');     // small page
  localStorage.setItem('teamsChatSegmentDelayMs', '1200');    // at least 1200ms
  localStorage.setItem('teamsChatSegmentStartTime', '1748012400267'); // from metadata

  // Optional: resume from a known syncState (uncomment and paste full value)
  // localStorage.setItem('teamsChatSegmentInitialSyncState', '3e5b00000031393a323264...fc96010000');

  console.log('[Segment overrides set]');
})();
