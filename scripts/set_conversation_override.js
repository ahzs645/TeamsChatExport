// Paste this in the Teams page console, set the override, reload the tab, then run extraction.

// Option 1: conversation ID (two participants)
localStorage.setItem(
  'teamsChatConversationIdOverride',
  '19:22d90d37-8f68-497f-a6df-d604f3dca807_46de1a27-3106-478b-bd49-6f675f88848d@unq.gbl.spaces'
);

// Option 2: thread ID (try this if the conversation ID returns no messages)
// localStorage.setItem(
//   'teamsChatConversationIdOverride',
//   '19:46de1a27-3106-478b-bd49-6f675f4881639fbd62@unq.gbl.spaces'
// );

// After setting, reload the Teams tab so the override is picked up.
