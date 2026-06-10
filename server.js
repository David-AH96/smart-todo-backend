const express = require('express');
const cors = require('cors');
const { execFile } = require('child_process');

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

function runJXA(script, timeout) {
  timeout = timeout || 60000;
  return new Promise(function(resolve, reject) {
    execFile('osascript', ['-l', 'JavaScript', '-e', script], { maxBuffer: 1024 * 1024 * 10, timeout: timeout }, function(error, stdout, stderr) {
      if (error) { reject(new Error(stderr || error.message)); return; }
      try { resolve(JSON.parse(stdout.trim())); } catch (e) { resolve(stdout.trim()); }
    });
  });
}

function runAppleScript(script, timeout) {
  timeout = timeout || 60000;
  return new Promise(function(resolve, reject) {
    execFile('osascript', ['-e', script], { maxBuffer: 1024 * 1024 * 10, timeout: timeout }, function(error, stdout, stderr) {
      if (error) { reject(new Error(stderr || error.message)); return; }
      resolve(stdout.trim());
    });
  });
}

// ==========================================
//          OUTLOOK LOCAL INTEGRATION
// ==========================================

app.get('/api/health', async function(req, res) {
  try {
    var script = 'var o = Application("Microsoft Outlook"); var a = o.exchangeAccounts(); var email = "unknown"; var name = "unknown"; if (a.length > 0) { email = a[0].emailAddress(); name = a[0].name(); } JSON.stringify({ connected: true, name: name, email: email });';
    var result = await runJXA(script, 10000);
    res.json(result);
  } catch (err) { res.json({ connected: false, error: err.message, hint: 'Make sure Legacy Outlook for Mac is open.' }); }
});

app.get('/api/profile', async function(req, res) {
  try {
    var script = 'var o = Application("Microsoft Outlook"); var a = o.exchangeAccounts(); var result = []; for (var i = 0; i < a.length; i++) { result.push({ name: a[i].name(), email: a[i].emailAddress() }); } JSON.stringify({ name: result.length > 0 ? result[0].name : "Unknown", email: result.length > 0 ? result[0].email : "Unknown", accounts: result });';
    var result = await runJXA(script, 10000);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/folders', async function(req, res) {
  try {
    var script = 'var o = Application("Microsoft Outlook"); var f = o.mailFolders(); var r = []; for (var i = 0; i < Math.min(f.length, 30); i++) { try { var name = f[i].name(); if (name) r.push({ name: name, unreadCount: f[i].unreadCount() }); } catch(e) {} } JSON.stringify({ folders: r });';
    var result = await runJXA(script, 15000);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- BULK FLAG SYNC ENDPOINT (Safe Loop Method) ---
app.post('/api/emails/check-flags', async function(req, res) {
  var ids = req.body.ids || [];
  if (ids.length === 0) return res.json({ statuses: {} });
  var idsJson = JSON.stringify(ids);
  try {
    var script = 'var o = Application("Microsoft Outlook"); var inbox = o.inbox(); var msgs = inbox.messages; var targetIds = ' + idsJson + '; var statuses = {}; for (var j = 0; j < targetIds.length; j++) { statuses[targetIds[j]] = false; } for (var i = 0; i < 2000; i++) { try { var msg = msgs[i]; var msgId = String(msg.id()); if (targetIds.includes(msgId)) { var isFlagged = false; try { var tf = String(msg.todoFlag()).toLowerCase(); if (tf !== "not flagged" && tf !== "completed" && tf !== "null" && tf !== "undefined") { isFlagged = true; } } catch(e) {} statuses[msgId] = isFlagged; } } catch(e) { continue; } } JSON.stringify({ statuses: statuses });';
    var result = await runJXA(script);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- ALL OF TODAY'S EMAILS (Time filter removed for UI Checkboxes) ---
app.get('/api/emails', async function(req, res) {
  var count = parseInt(req.query.count) || 200; 
  var unreadOnly = req.query.unread === 'true';
  var filterLine = unreadOnly ? 'if (isRead) continue;' : '';
  
  var today = new Date();
  today.setHours(0,0,0,0);
  var startOfTodayTs = today.getTime();

  try {
    var script = 'var o = Application("Microsoft Outlook"); var inbox = o.inbox(); var msgs = inbox.messages; var stopTs = ' + startOfTodayTs + '; var emails = []; var collected = 0; for (var i = 0; i < 1500 && collected < ' + count + '; i++) { try { var msg = msgs[i]; var recvTime = msg.timeReceived(); var recvTs = recvTime.getTime(); if (recvTs < stopTs) break; var isRead = msg.isRead(); ' + filterLine + ' var isFlagged = false; try { var tf = String(msg.todoFlag()).toLowerCase(); if (tf !== "not flagged" && tf !== "completed" && tf !== "null" && tf !== "undefined") { isFlagged = true; } } catch(e) {} var importance = "normal"; try { var p = msg.priority(); if (p === "high priority" || p === "high") importance = "high"; else if (p === "low priority" || p === "low") importance = "low"; } catch(e) {} var sAddr = ""; var sName = ""; try { var s = msg.sender(); sAddr = s.address(); sName = s.name(); } catch(e) {} var body = ""; try { body = (msg.plainTextContent() || "").substring(0, 300); } catch(e) { try { body = (msg.content() || "").substring(0, 300); } catch(e2) {} } emails.push({ id: msg.id(), subject: msg.subject() || "(No subject)", body: body, from: sAddr, fromName: sName, receivedAt: recvTime.toISOString(), isRead: isRead, importance: importance, isFlagged: isFlagged }); collected++; } catch(e) { continue; } } JSON.stringify({ emails: emails, total: collected });';
    var result = await runJXA(script);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/emails/flagged', async function(req, res) {
  var count = parseInt(req.query.count) || 50;
  var today = new Date();
  today.setHours(0,0,0,0);
  var startOfTodayTs = today.getTime();

  try {
    var script = 'var o = Application("Microsoft Outlook"); var inbox = o.inbox(); var msgs = inbox.messages; var stopTs = ' + startOfTodayTs + '; var flagged = []; for (var i = 0; i < 1500 && flagged.length < ' + count + '; i++) { try { var msg = msgs[i]; var recvTime = msg.timeReceived(); if (recvTime.getTime() < stopTs) break; var isFlagged = false; try { var tf = String(msg.todoFlag()).toLowerCase(); if (tf !== "not flagged" && tf !== "completed" && tf !== "null" && tf !== "undefined") { isFlagged = true; } } catch(e) {} if (!isFlagged) continue; var sAddr = ""; var sName = ""; try { var s = msg.sender(); sAddr = s.address(); sName = s.name(); } catch(e) {} var body = ""; try { body = (msg.plainTextContent() || "").substring(0, 300); } catch(e) { try { body = (msg.content() || "").substring(0, 300); } catch(e2) {} } var importance = "normal"; try { var p = msg.priority(); if (p === "high priority" || p === "high") importance = "high"; } catch(e) {} flagged.push({ id: msg.id(), subject: msg.subject() || "(No subject)", body: body, from: sAddr, fromName: sName, receivedAt: recvTime.toISOString(), importance: importance, dueDate: null }); } catch(e) { continue; } } JSON.stringify({ flagged: flagged, total: flagged.length });';
    var result = await runJXA(script);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/emails/search', async function(req, res) {
  var query = (req.query.q || '').replace(/"/g, '\\"').toLowerCase();
  var count = parseInt(req.query.count) || 20;
  if (!query) return res.status(400).json({ error: 'Search query (q) is required' });
  try {
    var script = 'var o = Application("Microsoft Outlook"); var inbox = o.inbox(); var msgs = inbox.messages; var results = []; var query = "' + query + '"; for (var i = 0; i < 1000 && results.length < ' + count + '; i++) { try { var msg = msgs[i]; var subject = (msg.subject() || "").toLowerCase(); var body = ""; try { body = (msg.plainTextContent() || "").substring(0, 500).toLowerCase(); } catch(e) { try { body = (msg.content() || "").substring(0, 500).toLowerCase(); } catch(e2) {} } var sName = ""; var sAddr = ""; try { var s = msg.sender(); sName = s.name().toLowerCase(); sAddr = s.address().toLowerCase(); } catch(e) {} if (subject.includes(query) || body.includes(query) || sName.includes(query) || sAddr.includes(query)) { results.push({ id: msg.id(), subject: msg.subject() || "(No subject)", body: body.substring(0, 200), from: sAddr, fromName: sName, receivedAt: msg.timeReceived().toISOString(), isRead: msg.isRead() }); } } catch(e) { continue; } } JSON.stringify({ results: results, total: results.length, query: query });';
    var result = await runJXA(script);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/emails/:id', async function(req, res) {
  var emailId = req.params.id;
  try {
    var script = 'var o = Application("Microsoft Outlook"); var inbox = o.inbox(); var msgs = inbox.messages; var found = null; for (var i = 0; i < 2000; i++) { try { if (String(msgs[i].id()) === "' + emailId + '") { var msg = msgs[i]; var sAddr = ""; var sName = ""; try { var s = msg.sender(); sAddr = s.address(); sName = s.name(); } catch(e) {} var body = ""; try { body = msg.plainTextContent() || ""; } catch(e) { try { body = msg.content() || ""; } catch(e2) {} } found = { id: msg.id(), subject: msg.subject() || "(No subject)", body: body, from: sAddr, fromName: sName, receivedAt: msg.timeReceived().toISOString(), isRead: msg.isRead() }; break; } } catch(e) { continue; } } JSON.stringify(found || { error: "Email not found" });';
    var result = await runJXA(script);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/unread-count', async function(req, res) {
  try {
    var script = 'var o = Application("Microsoft Outlook"); var inbox = o.inbox(); JSON.stringify({ unreadCount: inbox.unreadCount() });';
    var result = await runJXA(script, 10000);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/calendar', async function(req, res) {
  var days = parseInt(req.query.days) || 7;
  try {
    var script = 'var o = Application("Microsoft Outlook"); var cal; try { cal = o.defaultCalendar(); } catch(e) { var cals = o.calendars(); if (cals.length > 0) cal = cals[0]; } if (!cal) { JSON.stringify({ events: [], total: 0 }); } else { var events = cal.calendarEvents(); var now = new Date(); var future = new Date(now.getTime() + ' + days + ' * 24 * 60 * 60 * 1000); var upcoming = []; for (var i = 0; i < Math.min(events.length, 500); i++) { try { var evt = events[i]; var start = evt.startTime(); if (start >= now && start <= future) { upcoming.push({ id: evt.id(), subject: evt.subject() || "(No subject)", startTime: start.toISOString(), endTime: evt.endTime().toISOString(), location: evt.location() || "", isAllDay: evt.isAllDayFlag() }); } } catch(e) { continue; } } upcoming.sort(function(a, b) { return new Date(a.startTime) - new Date(b.startTime); }); JSON.stringify({ events: upcoming, total: upcoming.length }); }';
    var result = await runJXA(script);
    res.json(result);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/api/sync', async function(req, res) {
  var syncCount = parseInt(req.query.count) || 50;
  try {
    var script = 'tell application "Microsoft Outlook"\nset cloudInbox to inbox\nset localInbox to mail folder "Inbox" of on my computer\nset recentMsgs to messages 1 thru ' + syncCount + ' of cloudInbox\nset copiedCount to 0\nrepeat with aMsg in recentMsgs\ntry\nduplicate aMsg to localInbox\nset copiedCount to copiedCount + 1\nend try\nend repeat\nreturn copiedCount\nend tell';
    var result = await runAppleScript(script);
    res.json({ success: true, message: 'Synced ' + result + ' emails to On My Computer', count: parseInt(result) || 0 });
  } catch (err) { res.json({ success: false, error: err.message }); }
});

// ==========================================
//          WEBEX CLOUD INTEGRATION
// ==========================================
const WEBEX_TOKEN = 'MzU0ZGNiMTMtOTgxNS00Mzc1LTkyNTMtZjBjZWFmOGVlNjVjMjQ5OGViYWUtMDg0_PF84_f305806c-42ae-4160-aabb-3ecc46b7c0a0';
const WEBEX_API_BASE = 'https://webexapis.com/v1';

const webexHeaders = {
  'Authorization': 'Bearer ' + WEBEX_TOKEN,
  'Content-Type': 'application/json'
};

// 1. Get user profile (Connection Check)
app.get('/api/webex/profile', async function(req, res) {
  try {
    var fetch = (await import('node-fetch')).default;
    const response = await fetch(`${WEBEX_API_BASE}/people/me`, { headers: webexHeaders });
    if (!response.ok) throw new Error('Webex API Error: ' + response.statusText);
    const data = await response.json();
    res.json({ connected: true, id: data.id, name: data.displayName, email: data.emails[0] });
  } catch (err) {
    res.status(500).json({ connected: false, error: err.message });
  }
});

// 2. Fetch Spaces (Rooms) AND Recent Messages with TIMEFRAME SUPPORT
app.get('/api/webex/sync', async function(req, res) {
  var days = parseInt(req.query.days) || 1; 
  
  var cutoffTs;
  if (days === 1) {
    var today = new Date();
    today.setHours(0,0,0,0);
    cutoffTs = today.getTime();
  } else {
    cutoffTs = Date.now() - (days * 24 * 60 * 60 * 1000);
  }
  
  var roomCount = days === 1 ? 15 : (days <= 7 ? 30 : 50);
  var msgCount = days === 1 ? 20 : (days <= 7 ? 50 : 100);

  try {
    var fetch = (await import('node-fetch')).default;
    const roomsRes = await fetch(`${WEBEX_API_BASE}/rooms?sortBy=lastactivity&max=${roomCount}`, { headers: webexHeaders });
    if (!roomsRes.ok) throw new Error('Failed to fetch rooms');
    const roomsData = await roomsRes.json();
    const rooms = roomsData.items || [];

    var allMessages = [];

    for (var i = 0; i < rooms.length; i++) {
      var roomId = rooms[i].id;
      var roomTitle = rooms[i].title;

      const msgRes = await fetch(`${WEBEX_API_BASE}/messages?roomId=${roomId}&max=${msgCount}`, { headers: webexHeaders });
      if (msgRes.ok) {
        const msgData = await msgRes.json();
        var msgs = msgData.items || [];
        
        msgs.forEach(m => {
          var createdTs = new Date(m.created).getTime();
          
          if (createdTs >= cutoffTs && m.personEmail !== 'RahulDavid.Dudde@pega.com') { 
            allMessages.push({
              id: m.id,
              roomId: roomId,
              roomTitle: roomTitle,
              text: m.text || '',
              html: m.html || '',
              from: m.personEmail,
              fromName: m.personDisplayName || 'Unknown',
              receivedAt: m.created,
              mentionedPeople: m.mentionedPeople || [],
              isFlagged: m.isFlagged === true || m.isStarred === true 
            });
          }
        });
      }
    }

    // Sort globally by oldest first for Sliding Window logic
    allMessages.sort((a, b) => new Date(a.receivedAt) - new Date(b.receivedAt));
    res.json({ messages: allMessages, total: allMessages.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, function() {
  console.log('');
  console.log('  ================================================');
  console.log('  Outlook Mac Bridge - Full Version Running!');
  console.log('  http://localhost:' + PORT);
  console.log('  ');
  console.log('  Includes Flag Syncing, UI Filters, & Webex Integrations');
  console.log('  ================================================');
  console.log('');
});
