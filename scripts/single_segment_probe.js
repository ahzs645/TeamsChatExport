// Single-segment probe: makes one chatsvc request using a provided syncState/startTime.
// Edit the values below before running in the Teams console (top frame).
(() => {
  const convId = '19:22d90d37-8f68-497f-a6df-d604f3dca807_46de1a27-3106-478b-bd49-6f675f88848d@unq.gbl.spaces';
  const host = 'teams.microsoft.com';
  const region = 'ca';
  const syncState = '3e5b00000031393a32326439306433372d386636382d343937662d613664662d6436303466336463613830375f34366465316132372d333130362d343738622d626434392d36663637356638383834386440756e712e67626c2e737061636573018b7ea8fd960100005311279e9a010000c21e6cfc96010000';
  const startTime = '1748012400267';
  const pageSize = '20';

  const log = (...args) => console.log('[SingleSegmentProbe]', ...args);
  const warn = (...args) => console.warn('[SingleSegmentProbe]', ...args);

  const token = localStorage.getItem('teamsChatAuthToken') || window.__teamsAuthToken;
  const headers = {
    'Accept': 'application/json',
    'x-ms-client-type': 'web',
    'x-ms-request-priority': '0',
    'x-ms-migration': 'True',
    'behavioroverride': 'redirectAs404',
    'x-ms-client-version': '1415/25110202315',
    'clientinfo': 'os=mac; osVer=10.15.7; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/25110202315; utcOffset=-08:00; timezone=America/Vancouver'
  };
  if (token) {
    const lower = token.toLowerCase();
    headers.Authorization = (lower.startsWith('bearer ') || lower.startsWith('skype_token')) ? token : `Bearer ${token}`;
  }

  const url = `https://${host}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(convId)}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${pageSize}&startTime=${startTime}&syncState=${encodeURIComponent(syncState)}`;
  log('fetching', url);
  fetch(url, { credentials: 'include', headers })
    .then(async (resp) => {
      const text = await resp.text();
      log('status', resp.status, resp.statusText);
      log('body', text.slice(0, 800));
    })
    .catch((err) => warn('fetch error', err));
})();
