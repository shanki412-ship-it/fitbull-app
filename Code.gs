/**
 * Strength Coach – Google Apps Script (the engine – the same for everyone)
 * ------------------------------------------------------------
 * Everything personal lives in Config.gs (USER_CONFIG). This file reads it.
 *
 * What it does
 *   • dailyBrief()    – every morning: today's session with target weights and today's targets, emailed to you.
 *   • weeklyReview()  – once a week: weight trend, maintenance estimate, next week's calorie target,
 *                       set changes, patterns, and a Gemini-written coaching note.
 *   • monthlyReview() – 1st of each month: month-on-month comparison.
 *   • doGet()         – the phone app (Log.html).
 *
 * One-time setup:
 *   1. Run start  – checks Config.gs, builds the tabs, switches on the AI (geminiKey in Config.gs),
 *                   schedules the emails and emails you the next step. Safe to run again any time.
 *   2. Deploy ▸ New deployment ▸ Web app (Execute as: Me, Access: Only myself), then run start once more –
 *      the email has a button to open the app on your phone.
 */

// =====================================================================
// ▶ START HERE – run "start" once after pasting the files (and again any time; it's safe)
// =====================================================================
/**
 * Does the whole setup in one go: checks Config.gs, builds the tabs, switches on the AI, schedules the emails,
 * connects Telegram if you added it, and emails you the last two steps (publish the app and add it to your phone).
 */
function start() {
  if (!hasConfig_()) { try { wizOpen(); return 'Opened the setup wizard'; } catch (e) { throw new Error('Not set up yet – open the Sheet and choose FitBull ▸ Set up FitBull.'); } }
  const problems = validateConfig_();
  if (problems.length) {
    const msg = 'Please fix these in Config.gs, save, then run start again:\n- ' + problems.join('\n- ');
    console.error(msg);
    throw new Error(msg);
  }
  const props = PropertiesService.getScriptProperties();
  const lines = [];
  setup();
  lines.push('✓ Your tabs are ready: Daily Log, Lifts, Runs, Food, Plan, Targets, Settings and the coach tabs.');
  // the script's own time zone should match yours, or dates can slip by a day
  let scriptTz = '';
  try { scriptTz = Session.getScriptTimeZone(); } catch (e) {}
  if (scriptTz && scriptTz !== tz_()) lines.push(`✗ One quick fix: Project Settings (gear icon on the left) → Time zone → choose ${tz_()} (it's ${scriptTz} now). Then run start again.`);

  // AI: take the key from Config.gs and test it
  const key = configKey_();
  if (key) props.setProperty('GEMINI_API_KEY', key);
  let aiOk = false;
  if (props.getProperty('GEMINI_API_KEY')) {
    props.deleteProperty('GEMINI_MODEL'); props.deleteProperty('GEMINI_SMART_MODEL');
    aiOk = !!callGemini_('Reply with JSON only: {"ok": true}', true);
    lines.push(aiOk ? `✓ AI is on (${props.getProperty('GEMINI_MODEL')}).` : `✗ AI key problem: ${lastGeminiError_}. Check the key in Config.gs, save, run start again.`);
  } else lines.push('• AI is off – paste your free Gemini key into geminiKey in Config.gs (aistudio.google.com → Get API key), save, run start again.');

  // emails + evening check-in
  createTriggers();
  const e = Object.assign({ dailyHour: 5, dailyMinute: 45, weeklyDay: 'Sun', weeklyHour: 19 }, cfg_().emails || {});
  lines.push(`✓ Emails scheduled: daily plan at ${e.dailyHour}:${String(e.dailyMinute || 0).padStart(2, '0')}, weekly review ${e.weeklyDay} ${e.weeklyHour}:00` +
    (agentCfg_().proactive ? `, coach check-in at ${agentCfg_().eveningHour}:00.` : '.'));
  props.setProperty('TRIGGERS_CHECKED', key_(new Date()));

  // Telegram (optional)
  let tgCode = '';
  const tg = cfg_().telegram || {};
  if (tg.botToken && tg.webhookUrl) {
    props.setProperty('TELEGRAM_BOT_TOKEN', String(tg.botToken).trim());
    props.setProperty('TELEGRAM_WEBHOOK_URL', String(tg.webhookUrl).trim());
    if (!telegramOn_()) { try { tgCode = connectTelegram(); lines.push(`✓ Telegram ready – send your bot:  /start ${tgCode}`); } catch (err) { lines.push('✗ Telegram: ' + err.message); } }
    else lines.push('✓ Telegram is connected.');
  }

  // where the app lives
  const url = appUrl_();
  lines.push(url ? `✓ Your app: ${url}` : '• Last step: publish the app (Deploy → New deployment → Web app) – the email explains.');
  props.setProperty('SETUP_DONE', new Date().toISOString());
  welcomeEmail_(url, aiOk, tgCode);
  const out = lines.join('\n');
  console.log(out);
  try { SpreadsheetApp.getUi().alert(appName_() + ' setup', out, SpreadsheetApp.getUi().ButtonSet.OK); } catch (err) {}
  return out;
}

function configKey_() {
  const k = String((typeof USER_CONFIG !== 'undefined' && USER_CONFIG && USER_CONFIG.geminiKey) || '').trim();
  return k && !/paste|your.?key|xxx/i.test(k) ? k : '';
}

/** The phone app's link once it's published (left out when Telegram uses a second, public deployment). */
function appUrl_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('APP_URL')) return p.getProperty('APP_URL');
  if (cfg_().appUrl) return String(cfg_().appUrl);
  if (telegramOn_() || (cfg_().telegram || {}).webhookUrl) return '';
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function welcomeEmail_(url, aiOk, tgCode) {
  const s = getSettings_(), name = profile_().name || '';
  const btn = u => `<p><a href="${u}" style="background:#1F5FBF;color:#fff;padding:12px 18px;border-radius:10px;text-decoration:none;font-weight:bold;display:inline-block">Open my coach app</a></p>`;
  const html = `<div style="font-family:Arial,sans-serif;font-size:15px;color:#1B2631;line-height:1.5">
    <h2 style="margin:0 0 6px">You're nearly there${name ? ', ' + name : ''} 💪</h2>
    <p>Your tabs, emails${aiOk ? ', AI coach' : ''} and settings are set up.</p>
    ${url ? `<h3 style="margin:14px 0 4px">Put the app on your phone</h3>${btn(url)}
      <ol><li>Open this email on your phone and tap the button.</li>
      <li>iPhone: Share → <b>Add to Home Screen</b>. Android: ⋮ → <b>Add to Home screen</b>.</li></ol>`
    : `<h3 style="margin:14px 0 4px">Last step: publish your phone app (1 minute)</h3>
      <ol><li>In Apps Script click <b>Deploy → New deployment</b>.</li>
      <li>Click the gear next to "Select type" → <b>Web app</b>.</li>
      <li>Execute as: <b>Me</b>. Who has access: <b>Only myself</b>. Click <b>Deploy</b> and allow access.</li>
      <li>Run <b>start</b> once more – you'll get a new email with a button to open the app on your phone.</li></ol>`}
    ${tgCode ? `<h3 style="margin:14px 0 4px">Telegram</h3><p>Open your bot in Telegram and send: <b>/start ${tgCode}</b></p>` : ''}
    <h3 style="margin:14px 0 4px">Day one</h3>
    <ul><li>Morning: weigh yourself and enter it on <b>Check-in</b>.</li>
      <li>Meals: tap your one-tap meals, snap a photo, or type what you ate on <b>Food</b>.</li>
      <li>Gym: enter the weight for each set and tap ✓ – the app sets your weights from then on. Runs go on <b>Gym → Run</b>.</li>
      <li>Evening: tap how you feel so tomorrow's plan adjusts.</li></ul>
    <p style="color:#5B6874">The daily plan email arrives each morning; the weekly review on ${(cfg_().emails || {}).weeklyDay || 'Sun'}.</p></div>`;
  MailApp.sendEmail({ name: appName_(), to: s['Email'] || Session.getActiveUser().getEmail(), subject: url ? 'Your coach app is ready – add it to your phone' : 'Coach setup: one last step', htmlBody: html });
}

/** Keeps the scheduled emails in place (runs at most once a day when the app opens). */
function ensureTriggers_() {
  const p = PropertiesService.getScriptProperties();
  const today = key_(new Date());
  if (p.getProperty('TRIGGERS_CHECKED') === today) return;
  p.setProperty('TRIGGERS_CHECKED', today);
  const have = new Set(ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()));
  const need = ['dailyBrief', 'weeklyReview', 'monthlyReview'].concat(agentCfg_().proactive ? ['agentEveningCheck'] : []);
  if (need.some(n => !have.has(n))) createTriggers();
}

// ---------- Engine settings ----------
const CONFIG = {
  // Tried in order; the first that works is remembered. Add newer models at the front when Google releases them.
  MODELS: ['gemini-3.8-flash', 'gemini-3.5-flash-lite', 'gemini-3.7-flash', 'gemini-2.5-flash'],
  // Photos, voice notes and the agent need a stronger model than quick text estimates. Tried in this order
  // (with one retry when a model is busy); "lite" models are only a last resort. Never remembered over MODELS.
  SMART_MODELS: ['gemini-3.8-flash', 'gemini-3.7-flash', 'gemini-2.5-flash', 'gemini-3.5-flash-lite'],
  KCAL_PER_KG: 7700,
  MAX_WEEKLY_CHANGE: 150,
};

const SHEETS = {
  LOG: 'Daily Log', LIFTS: 'Lifts', PLAN: 'Plan',
  TARGETS: 'Targets', SETTINGS: 'Settings', FOOD: 'Food',
};
const FOOD_HEADERS = ['Date', 'Time', 'Item', 'kcal', 'Protein g', 'Source', 'ID', 'Fat g', 'Sugar g', 'Meal'];
const MEAL_SLOTS = ['Morning', 'Afternoon', 'Evening', 'Night'];

/** Which part of the day a food belongs to. Uses what was picked; otherwise works it out from the clock time
 *  (Morning 4–12, Afternoon 12–17, Evening 17–21, Night 21–4). */
function mealSlot_(meal, time) {
  const m = String(meal || '').trim().toLowerCase();
  const hit = MEAL_SLOTS.find(s => s.toLowerCase() === m);
  if (hit) return hit;
  if (/break?fast|morning/.test(m)) return 'Morning';
  if (/lunch|afternoon|noon/.test(m)) return 'Afternoon';
  if (/dinner|evening|tea\s*time/.test(m)) return 'Evening';
  if (/night|supper|late/.test(m)) return 'Night';
  const hourOf = s => { const x = String(s || '').match(/^(\d{1,2}):/); return x ? Number(x[1]) : null; };
  let h = time instanceof Date ? hourOf(fmt_(time, 'HH:mm')) : hourOf(time);
  if (h === null) h = hourOf(fmt_(new Date(), 'HH:mm'));
  return h >= 4 && h < 12 ? 'Morning' : h >= 12 && h < 17 ? 'Afternoon' : h >= 17 && h < 21 ? 'Evening' : 'Night';
}

const LOG_HEADERS = ['Date', 'Weight kg', 'kcal', 'Protein g', 'Steps', 'Sleep h',
  'Session', 'Activity', 'Pump', 'Soreness', 'Joint pain', 'Waist cm', 'Notes', 'Feeling', 'Fat g', 'Sugar g', 'Skipped'];
const LIFT_HEADERS = ['Date', 'Session', 'Exercise', 'Set', 'Weight kg', 'Reps', 'RIR', 'ID'];
const PLAN_HEADERS = ['Session', 'Order', 'Exercise', 'Sets', 'Reps', 'Rep max',
  'Rest', 'Increment kg', 'Muscle'];
const TARGET_HEADERS = ['Week ending', 'Week #', '7-day avg weight', 'Weekly change kg',
  'Avg kcal', 'Avg protein', 'Avg steps', 'Avg sleep', 'Gym sessions', 'Sport/cardio',
  'Waist cm', 'Est. maintenance', 'Calorie target', 'Protein target', 'Step target',
  'Coach summary'];
const DAY_CODES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MUSCLES = ['Chest', 'Back', 'Shoulders', 'Rear delts', 'Quads', 'Hamstrings', 'Glutes',
  'Calves', 'Biceps', 'Triceps', 'Core', 'Forearms', 'Full body'];

// Swaps when equipment is busy: [exercise, weight increment kg]. Same sets and reps as the original.
// Each swap keeps its own history, so its targets progress separately. Config.gs can add more.
const BUILTIN_ALTERNATES = {
  'Back squat': [['Goblet squat', 2], ['Leg press', 10], ['Hack squat', 5]],
  'Front squat': [['Goblet squat', 2], ['Hack squat', 5], ['Leg press', 10]],
  'Goblet squat': [['Leg press', 10], ['Bulgarian split squat', 2], ['Hack squat', 5]],
  'Romanian deadlift': [['DB Romanian deadlift', 2], ['Seated leg curl', 2.5], ['45° back extension', 2.5]],
  'DB Romanian deadlift': [['Romanian deadlift', 5], ['45° back extension', 2.5], ['Seated leg curl', 2.5]],
  'Hanging knee raise': [["Captain's chair knee raise", 0], ['Reverse crunch', 0]],
  'Walking lunge (DB)': [['Bulgarian split squat', 2], ['Reverse lunge (DB)', 2], ['Step-up (DB)', 2]],
  'Bulgarian split squat': [['Reverse lunge (DB)', 2], ['Step-up (DB)', 2], ['Leg press', 10]],
  'Bench press': [['DB bench press', 2], ['Machine chest press', 5], ['Push-up', 0]],
  'DB bench press': [['Bench press', 2.5], ['Machine chest press', 5], ['Push-up', 0]],
  'Push-up': [['DB bench press', 2], ['Machine chest press', 5], ['Incline push-up', 0]],
  'Barbell row': [['Chest-supported DB row', 2], ['Seated cable row', 2.5], ['Machine row', 5]],
  'One-arm DB row': [['Chest-supported DB row', 2], ['Seated cable row', 2.5], ['Machine row', 5]],
  'Overhead press': [['Seated DB shoulder press', 2], ['Machine shoulder press', 2.5]],
  'Lat pulldown': [['Assisted pull-up', 0], ['Single-arm cable pulldown', 2.5]],
  'Lateral raise': [['Cable lateral raise', 1.25], ['Machine lateral raise', 2.5]],
  'Cable crunch': [['Ab wheel rollout', 0], ['Decline crunch', 0]],
  'Deadlift': [['Trap-bar deadlift', 5], ['Rack pull', 5], ['DB Romanian deadlift', 2]],
  'Trap-bar deadlift': [['Deadlift', 5], ['DB Romanian deadlift', 2], ['Rack pull', 5]],
  'Leg press': [['Hack squat', 5], ['Goblet squat', 2], ['Bulgarian split squat', 2]],
  'Leg extension': [['Sissy squat (assisted)', 0], ['Leg press', 10]],
  'Plank': [['Dead bug', 0], ['Side plank', 0]],
  'Hip thrust': [['Glute bridge (barbell)', 5], ['Hip thrust machine', 5], ['Cable pull-through', 2.5]],
  'Glute bridge (barbell)': [['Hip thrust', 5], ['Hip thrust machine', 5], ['Cable pull-through', 2.5]],
  'Standing calf raise': [['Seated calf raise', 5], ['Leg press calf raise', 10]],
  'Incline DB press': [['Incline barbell press', 2.5], ['Incline machine press', 5], ['Low-to-high cable fly', 2.5]],
  'Pull-up': [['Lat pulldown', 2.5], ['Assisted pull-up', 0]],
  'DB shoulder press': [['Machine shoulder press', 2.5], ['Arnold press', 2]],
  'Seated DB shoulder press': [['Machine shoulder press', 2.5], ['Arnold press', 2]],
  'Seated cable row': [['Chest-supported DB row', 2], ['Machine row', 5]],
  'DB curl': [['EZ-bar curl', 2.5], ['Cable curl', 2.5]],
  'Triceps pushdown': [['Overhead cable extension', 2.5], ['DB skull crusher', 2]],
  'Seated leg curl': [['Lying leg curl', 2.5], ['Nordic curl (assisted)', 0], ['Swiss ball leg curl', 0]],
  'Lying leg curl': [['Seated leg curl', 2.5], ['Swiss ball leg curl', 0]],
  'Face pull': [['Reverse pec deck', 2.5], ['Bent-over DB reverse fly', 1], ['Band pull-apart', 0]],
  'Chest fly (cable)': [['Pec deck', 2.5], ['DB fly', 1]],
};

// ---------- Reading Config.gs (always through these functions, never at file load) ----------
// Settings come from Config.gs (USER_CONFIG) – or, for apps set up with the FitBull ▸ Set up wizard,
// from the plan the wizard saved (Config.gs then just says USER_CONFIG = null).
let CFG_CACHE_ = null;
function cfg_() {
  if (typeof USER_CONFIG !== 'undefined' && USER_CONFIG) return USER_CONFIG;
  if (CFG_CACHE_) return CFG_CACHE_;
  const c = storedConfig_();
  if (!c) throw new Error('FitBull isn\'t set up yet – open the Sheet on a computer and choose FitBull ▸ Set up FitBull.');
  CFG_CACHE_ = c;
  return c;
}
function hasConfig_() { try { cfg_(); return true; } catch (e) { return false; } }
/** The wizard's plan is kept in Script Properties (split into pieces – each property holds up to 9 KB). */
function storedConfig_() {
  const p = PropertiesService.getScriptProperties(), n = Number(p.getProperty('CFG_PARTS')) || 0;
  if (!n) return null;
  let s = ''; for (let i = 0; i < n; i++) s += p.getProperty('CFG_' + i) || '';
  try { return JSON.parse(s); } catch (e) { console.error('Saved plan unreadable: ' + e.message); return null; }
}
function saveConfig_(c) {
  const p = PropertiesService.getScriptProperties(), s = JSON.stringify(c), size = 8000, old = Number(p.getProperty('CFG_PARTS')) || 0;
  const parts = []; for (let i = 0; i < s.length; i += size) parts.push(s.slice(i, i + size));
  const map = { CFG_PARTS: String(parts.length) }; parts.forEach((x, i) => { map['CFG_' + i] = x; });
  p.setProperties(map);
  for (let i = parts.length; i < old; i++) p.deleteProperty('CFG_' + i);
  CFG_CACHE_ = null;
}
function tz_() { return cfg_().timezone || Session.getScriptTimeZone(); }
function dayIndex_(code) { return DAY_CODES.findIndex(d => d.toLowerCase() === String(code || '').slice(0, 3).toLowerCase()); }
function planRows_() { return cfg_().plan || []; }
/** {dayIndex: session} – getDay() numbering, 0 = Sunday. */
function scheduleMap_() {
  const m = {}, sch = cfg_().schedule || {};
  Object.keys(sch).forEach(k => { const i = dayIndex_(k); if (i >= 0 && sch[k]) m[i] = String(sch[k]); });
  return m;
}
function sessionFor_(dow) { return scheduleMap_()[dow] || ''; }
function sessionsPerWeek_() { return Object.keys(scheduleMap_()).length || 1; }
/** Session names in the order they come up in the week (Mon → Sun), then any others in the plan. */
function sessionNames_() {
  const out = [], m = scheduleMap_();
  [1, 2, 3, 4, 5, 6, 0].forEach(i => { if (m[i] && !out.includes(m[i])) out.push(m[i]); });
  planRows_().forEach(r => { if (!out.includes(r[0])) out.push(r[0]); });
  return out;
}
function sports_() {
  return (cfg_().sports || []).map(s => Object.assign({ days: [], when: 'any', extraKcal: 0, restNextMorning: false, options: [] }, s));
}
function sportOn_(dow) { return sports_().find(s => s.days.some(d => dayIndex_(d) === dow)) || null; }
/** Which sport a logged activity text belongs to ("Badminton 90 min" → Badminton). */
function sportOfActivity_(text) {
  const a = String(text || '').toLowerCase();
  return a ? sports_().find(s => a.indexOf(String(s.name).toLowerCase()) >= 0) || null : null;
}
/** Activity buttons for the Check-in tab. */
function activityOptions_() {
  const out = [];
  sports_().forEach(s => (s.options.length ? s.options : ['']).forEach(o => {
    const v = (s.name + ' ' + o).trim();
    out.push({ label: v.replace(/ min( easy)?$/, 'm$1'), value: v });
  }));
  return out;
}
function trainingCfg_() {
  return Object.assign({ deloadEvery: 6, sessionMinutes: 45, maxSetsPerSession: 22, priorityMuscles: [], weeklySetCap: {}, mainLifts: [] },
    cfg_().training || {});
}
function stepsCfg_() { return Object.assign({ start: 5000, increase: 1500, everyWeeks: 2, max: 10000 }, cfg_().steps || {}); }
function mainLifts_() {
  const t = trainingCfg_();
  if (t.mainLifts && t.mainLifts.length) return t.mainLifts.slice();
  const out = [];
  planRows_().filter(r => Number(r[1]) === 1).forEach(r => { if (!out.includes(r[2])) out.push(r[2]); });
  return out;
}
function alternates_() { return Object.assign({}, BUILTIN_ALTERNATES, cfg_().alternates || {}); }
function appName_() { return String(cfg_().appName || 'FitBull').trim() || 'FitBull'; }
function appTagline_() { const t = cfg_().appTagline; return t === undefined ? 'Strong habits. Smarter training.' : String(t).trim(); }
function profile_() { return Object.assign({ name: 'you' }, cfg_().profile || {}); }
function meals_() {
  return (cfg_().meals || []).filter(m => m && m[0])
    .map(m => [m[0], m[1], m[2], m.length > 3 ? m[3] : null, m.length > 4 ? m[4] : null]);
}
/** The logged activity on a Daily Log row (older sheets call the column "Badminton/Run"). */
function activityOf_(r) { return String((r && (r['Activity'] || r['Badminton/Run'])) || '').trim(); }
function activityCol_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (!sh) return 'Activity';
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  return h.includes('Activity') || !h.includes('Badminton/Run') ? 'Activity' : 'Badminton/Run';
}
function settingsDefaults_() {
  const n = cfg_().nutrition || {};
  return {
    'Starting calorie target': n.startCalories, 'Weekly change goal kg': n.weeklyChangeKg,
    'Protein target g': n.proteinTarget, 'Protein minimum g': n.proteinMinimum,
    'Calorie floor': n.calorieFloor, 'Waist goal cm': n.waistGoalCm || '',
    'Goal weight kg': n.goalWeightKg || '',
  };
}

/** Run from the editor after editing Config.gs. Lists any problems; otherwise prints a summary. */
function checkConfig() {
  const problems = validateConfig_();
  if (problems.length) {
    const msg = 'Config.gs needs fixing:\n- ' + problems.join('\n- ');
    console.error(msg);
    throw new Error(msg);
  }
  const c = cfg_(), m = scheduleMap_();
  console.log(`Config OK for ${profile_().name}.\n` +
    `Gym days: ${[1, 2, 3, 4, 5, 6, 0].filter(i => m[i]).map(i => DAY_CODES[i] + ' ' + m[i]).join(', ')}\n` +
    `Sessions: ${sessionNames_().map(s => s + ' (' + planRows_().filter(r => r[0] === s).length + ' exercises)').join(', ')}\n` +
    `Sports: ${sports_().map(s => s.name + (s.days.length ? ' ' + s.days.join('/') : '')).join(', ') || 'none'}\n` +
    `Calories ${c.nutrition.startCalories}, protein ${c.nutrition.proteinTarget} g, weekly change ${c.nutrition.weeklyChangeKg} kg, timezone ${tz_()}`);
  return 'OK';
}

function validateConfig_() {
  const p = [];
  let c;
  try { c = cfg_(); } catch (e) { return [e.message]; }
  const num = (v, name, min, max) => {
    if (typeof v !== 'number' || isNaN(v)) p.push(`${name} must be a number (no quotes)`);
    else if ((min !== undefined && v < min) || (max !== undefined && v > max)) p.push(`${name} looks wrong (${v})`);
  };
  if (!c.profile || !c.profile.name) p.push('profile.name is missing');
  try { Utilities.formatDate(new Date(), c.timezone, 'yyyy'); } catch (e) { p.push(`timezone "${c.timezone}" isn't valid – use one like 'Europe/London'`); }
  if (!c.timezone) p.push('timezone is missing');
  const sessions = new Set((c.plan || []).map(r => r && r[0]));
  if (!Array.isArray(c.plan) || !c.plan.length) p.push('plan is empty');
  (c.plan || []).forEach((r, i) => {
    const where = `plan row ${i + 1} (${r && r[2] || '?'})`;
    if (!Array.isArray(r) || r.length !== 9) { p.push(`${where} must have 9 values`); return; }
    if (!r[0] || !r[2]) p.push(`${where}: session and exercise names are required`);
    if (typeof r[1] !== 'number') p.push(`${where}: order must be a number`);
    if (typeof r[3] !== 'number' || r[3] < 1 || r[3] > 6) p.push(`${where}: sets must be a number from 1 to 6`);
    if (r[5] !== '' && typeof r[5] !== 'number') p.push(`${where}: top of rep range must be a number, or '' for timed exercises`);
    if (typeof r[7] !== 'number') p.push(`${where}: kg to add must be a number (0 for bodyweight)`);
    if (!MUSCLES.includes(r[8])) p.push(`${where}: muscle "${r[8]}" – use one of: ${MUSCLES.join(', ')}`);
  });
  const sch = c.schedule || {};
  if (!Object.keys(sch).length) p.push('schedule has no gym days');
  Object.keys(sch).forEach(k => {
    if (dayIndex_(k) < 0) p.push(`schedule day "${k}" – use Mon, Tue, Wed, Thu, Fri, Sat or Sun`);
    if (!sessions.has(sch[k])) p.push(`schedule ${k}: session "${sch[k]}" isn't in the plan`);
  });
  (c.sports || []).forEach((s, i) => {
    if (!s.name) p.push(`sport ${i + 1} needs a name`);
    (s.days || []).forEach(d => { if (dayIndex_(d) < 0) p.push(`sport ${s.name}: day "${d}" isn't valid`); });
    if (s.restNextMorning) (s.days || []).forEach(d => {
      const next = (dayIndex_(d) + 1) % 7;
      if (sch[DAY_CODES[next]]) p.push(`${s.name} is on ${d} with restNextMorning, but ${DAY_CODES[next]} has a gym session – move one of them`);
    });
  });
  const n = c.nutrition || {};
  num(n.startCalories, 'nutrition.startCalories', 1200, 5000);
  num(n.weeklyChangeKg, 'nutrition.weeklyChangeKg', -1, 0.75);
  num(n.proteinTarget, 'nutrition.proteinTarget', 40, 300);
  num(n.proteinMinimum, 'nutrition.proteinMinimum', 30, 300);
  num(n.calorieFloor, 'nutrition.calorieFloor', 1200, 4000);
  if (n.goalWeightKg) num(n.goalWeightKg, 'nutrition.goalWeightKg', 30, 250);
  if (typeof n.calorieFloor === 'number' && typeof n.startCalories === 'number' && n.calorieFloor > n.startCalories) p.push('nutrition.calorieFloor is higher than startCalories');
  (c.meals || []).forEach((m, i) => {
    const ok = Array.isArray(m) && (m.length === 3 || m.length === 5) && m.slice(1).every(v => typeof v === 'number');
    if (!ok) p.push(`meal ${i + 1} must look like ['Name', kcal, protein, fat, sugar] (fat and sugar optional)`);
  });
  if (c.geminiKey !== undefined && typeof c.geminiKey !== 'string') p.push("geminiKey must be text in quotes, e.g. geminiKey: 'AIza...'");
  if (c.telegram && c.telegram.webhookUrl && !/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(String(c.telegram.webhookUrl))) p.push('telegram.webhookUrl must be the /exec link of your "Anyone" deployment');
  const rn = c.running;
  if (rn) {
    if (rn.goal && !RUN_CAPS[String(rn.goal).toLowerCase()]) p.push(`running.goal must be one of: ${Object.keys(RUN_CAPS).join(', ')}`);
    if (rn.runsPerWeek !== undefined) num(rn.runsPerWeek, 'running.runsPerWeek', 1, 6);
    if (rn.weeklyKm !== undefined) num(rn.weeklyKm, 'running.weeklyKm', 0, 150);
    (rn.days || []).forEach(d => { if (dayIndex_(d) < 0) p.push(`running.days: "${d}" isn't a day`); });
    if (rn.raceDate && !/^\d{4}-\d{2}-\d{2}$/.test(String(rn.raceDate))) p.push("running.raceDate must look like '2027-04-18' (or '')");
  }
  const hm = c.home || {};
  if (hm.equipment !== undefined && (!Array.isArray(hm.equipment) || hm.equipment.some(x => !['bodyweight', 'dumbbells', 'band'].includes(String(x).toLowerCase()))))
    p.push("home.equipment must be a list of 'bodyweight', 'dumbbells', 'band'");
  const ag = c.agent || {};
  if (ag.eveningHour !== undefined) num(ag.eveningHour, 'agent.eveningHour', 0, 23);
  const e = c.emails || {};
  if (e.dailyHour !== undefined) num(e.dailyHour, 'emails.dailyHour', 0, 23);
  if (e.weeklyDay !== undefined && dayIndex_(e.weeklyDay) < 0) p.push('emails.weeklyDay must be Mon–Sun');
  return p;
}

// ---------- Menu ----------
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  if (!hasConfig_()) {                                   // a fresh copy: one clear way in
    ui.createMenu('FitBull').addItem('▶ Set up FitBull', 'wizOpen').addToUi();
    try { welcomeSheet_(); } catch (e) {}
    return;
  }
  ui.createMenu('FitBull')
    .addItem('📱 Open my app', 'wizShowLink')
    .addItem('🔗 Set my app link', 'wizSetLink')
    .addItem('✏️ Change my goals or plan', 'wizOpen')
    .addSeparator()
    .addItem('▶ Start / repair setup', 'start')
    .addItem('Send today\'s brief now', 'dailyBrief')
    .addItem('Run weekly review now', 'weeklyReview')
    .addItem('Run monthly review now', 'monthlyReview')
    .addSeparator()
    .addItem('Setup tabs', 'setup')
    .addItem('Create schedules', 'createTriggers')
    .addSeparator()
    .addItem('Send evening check-in now', 'agentEveningCheck')
    .addItem('Connect Telegram', 'connectTelegram')
    .addToUi();
}

// ---------- Setup ----------
function setup() {
  checkConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(tz_());

  ensureSheet_(ss, SHEETS.LOG, LOG_HEADERS);
  ensureSheet_(ss, SHEETS.LIFTS, LIFT_HEADERS);
  ensureSheet_(ss, SHEETS.TARGETS, TARGET_HEADERS);
  ensureSheet_(ss, SHEETS.FOOD, FOOD_HEADERS);

  const plan = ensureSheet_(ss, SHEETS.PLAN, PLAN_HEADERS);
  if (plan.getLastRow() < 2) {
    const rows = planRows_();
    plan.getRange(2, 1, rows.length, PLAN_HEADERS.length).setValues(rows);
    PropertiesService.getScriptProperties().setProperty('PLAN_SIG', planSignature_());
  }

  const settings = ensureSheet_(ss, SHEETS.SETTINGS, ['Setting', 'Value']);
  if (settings.getLastRow() < 2) {
    const def = settingsDefaults_();
    const rows = [['Start date', new Date()], ['Email', Session.getActiveUser().getEmail()]]
      .concat(Object.keys(def).map(k => [k, def[k]]));
    settings.getRange(2, 1, rows.length, 2).setValues(rows);
  } else {
    // Existing Settings tab: add any settings newer versions use (e.g. Goal weight kg), leaving yours alone
    const have = settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
    const def = settingsDefaults_();
    Object.keys(def).filter(k => !have.includes(k)).forEach(k => settings.appendRow([k, def[k]]));
  }

  ensureAgentSheets_();
  runsSheet_();

  // Existing sheets: add any columns newer versions need (e.g. Fat g / Sugar g), keeping all your data
  ensureLogColumns_();
  foodSheet_();
  ensureLiftIds_();
  ensurePlanUpToDate_();

  ss.getSheetByName(SHEETS.LOG).getRange('A:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(SHEETS.LIFTS).getRange('A:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(SHEETS.TARGETS).getRange('A:A').setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(SHEETS.TARGETS).setColumnWidth(16, 500);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function createTriggers() {
  const managed = ['dailyBrief', 'weeklyReview', 'monthlyReview', 'agentEveningCheck'];
  ScriptApp.getProjectTriggers()
    .filter(t => managed.includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  const e = Object.assign({ dailyHour: 5, dailyMinute: 45, weeklyDay: 'Sun', weeklyHour: 19 }, cfg_().emails || {});
  const weekDays = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
  ScriptApp.newTrigger('dailyBrief').timeBased().everyDays(1)
    .atHour(e.dailyHour).nearMinute(e.dailyMinute || 0).inTimezone(tz_()).create();
  ScriptApp.newTrigger('weeklyReview').timeBased().onWeekDay(ScriptApp.WeekDay[weekDays[Math.max(0, dayIndex_(e.weeklyDay))]])
    .atHour(e.weeklyHour).inTimezone(tz_()).create();
  ScriptApp.newTrigger('monthlyReview').timeBased().onMonthDay(1)
    .atHour(8).inTimezone(tz_()).create();
  const ag = agentCfg_();
  if (ag.proactive) ScriptApp.newTrigger('agentEveningCheck').timeBased().everyDays(1)
    .atHour(ag.eveningHour).nearMinute(0).inTimezone(tz_()).create();
}

// ---------- Daily brief ----------
function dailyBrief() {
  ensurePlanUpToDate_();
  ensureLogColumns_();
  const s = getSettings_();
  const today = parseDay_(key_(new Date()));
  const dow = Number(fmt_(today, 'u')) % 7;     // ISO 1=Mon … 7=Sun → 0=Sun
  const session = sessionFor_(dow);
  const log = readTable_(SHEETS.LOG);
  const ready = readiness_(today, log);
  const t = dayTargets_(s, today, dow, ready);
  const sport = sportOn_(dow);

  let html = `<div style="font-family:Arial,sans-serif;font-size:15px;color:#1B2631">`;
  if (session && ready.rest) {
    html += `<h2 style="margin:0 0 4px">Rest today</h2><p style="background:#FFF4D6;padding:8px"><b>${ready.title}.</b> ${ready.message}</p>`;
    html += `<p style="color:#5B6874">${session} moves to your next free morning – the app keeps your targets as they are.</p>`;
  } else if (session) {
    const weekNum = weekNumber_(s['Start date'], today);
    const rows = sessionTargets_(session, today, readTable_(SHEETS.LIFTS), readTable_(SHEETS.PLAN), s, { ready, log });
    html += `<h2 style="margin:0 0 4px">${session}</h2>`;
    if (isDeloadWeek_(weekNum)) html += `<p style="background:#E6F4EA;padding:8px"><b>Deload week.</b> About 60% of your normal weights, 2 sets each. Easy on purpose – you'll come back stronger next week.</p>`;
    if (ready.title) html += `<p style="background:#FFF4D6;padding:8px"><b>${ready.title}.</b> ${ready.message}</p>`;
    html += `<p style="margin:0 0 12px;color:#5B6874">Warm up 5 min, then 2 ramp-up sets on the first lift. Woke up feeling different? Tap it on the Gym tab and the plan adjusts.</p>`;
    html += `<table cellpadding="6" style="border-collapse:collapse;width:100%">
      <tr style="background:#1B2631;color:#fff"><th align="left">Exercise</th><th>Sets × reps</th><th>Target</th><th align="left">Last time</th></tr>`;
    rows.forEach(r => {
      html += `<tr style="border-bottom:1px solid #ddd"><td>${r.exercise}</td>
        <td align="center">${r.sets} × ${r.reps}</td>
        <td align="center"><b>${r.target}</b></td>
        <td>${r.last}</td></tr>`;
    });
    html += `</table>`;
  } else {
    html += `<h2 style="margin:0 0 8px">${restLabel_(dow)}</h2>`;
    if (ready.title) html += `<p style="background:#FFF4D6;padding:8px"><b>${ready.title}.</b> ${ready.message}</p>`;
  }
  try { html += runEmailLine_(today); } catch (e) { console.error(e); }
  html += `<h3 style="margin:16px 0 4px">Today's targets</h3><ul>
    <li>Calories: <b>${t.kcal}</b>${t.note ? ' (' + t.note + ')' : ''}</li>
    <li>Protein: <b>${t.protein} g</b> (minimum ${s['Protein minimum g']} g)</li>
    <li>Steps: <b>${t.steps ? t.steps.toLocaleString() : 'paused – rest'}</b>${sport && t.steps ? ` (${sport.name.toLowerCase()} counts)` : ''}</li>
    <li>Sleep: 7+ hours${sport && sport.when === 'evening' ? ` – protect bedtime after ${sport.name.toLowerCase()}` : ''}</li></ul></div>`;

  try { const u = appUrl_(); if (u) html = html.replace(/<\/div>$/, `<p><a href="${u}" style="color:#1F5FBF;font-weight:bold">Open the app →</a></p></div>`); } catch (e) {}
  const subject = session && !ready.rest ? `Today: ${session} · ${t.kcal} kcal · ${t.protein} g protein`
    : `Today: rest · ${t.kcal} kcal${t.steps ? ' · ' + t.steps.toLocaleString() + ' steps' : ''}`;
  MailApp.sendEmail({ name: appName_(), to: s['Email'], subject: subject, htmlBody: html });
}

function restLabel_(dow) {
  const notes = cfg_().restDayNotes || {};
  const own = Object.keys(notes).find(k => dayIndex_(k) === dow);
  if (own) return notes[own];
  const prevSport = sportOn_((dow + 6) % 7);
  if (prevSport && prevSport.restNextMorning) return `Rest – recovery after ${prevSport.name.toLowerCase()} (an easy walk is fine)`;
  return 'Rest day (an easy walk is fine)';
}

// ---------- Readiness: how you feel decides today's plan ----------
// Feeling is set the evening before (for the next morning) or overridden in the morning on the Gym tab.
const FEELINGS = ['good', 'tired', 'cold', 'sick', 'sore'];

function readiness_(d, logRows) {
  const k = key_(d), y = key_(addDays_(d, -1));
  const row = logRows.find(r => r['Date'] instanceof Date && key_(r['Date']) === k) || {};
  const prev = logRows.find(r => r['Date'] instanceof Date && key_(r['Date']) === y) || {};
  const feeling = String(row['Feeling'] || '').toLowerCase();
  const sleep = num_(row['Sleep h']);
  const base = { mode: 'normal', rest: false, title: '', message: '', setDelta: 0, setCap: 0, loadFactor: 1, feeling: feeling || '' };
  if (feeling === 'sick') return Object.assign(base, { mode: 'sick', rest: true, title: "You're unwell – rest today",
    message: 'Fever, chest congestion or body aches mean no training. Rest, drink plenty, and eat around maintenance (shown below). See a doctor if a fever lasts more than 2–3 days or you feel worse.' });
  const prevSport = sportOfActivity_(activityOf_(prev));
  if (prevSport && prevSport.restNextMorning) return Object.assign(base, { mode: 'after-sport', rest: true,
    title: `You played ${prevSport.name.toLowerCase()} ${prevSport.when === 'evening' ? 'last night' : 'yesterday'}`, message: 'Recovery morning – no gym.' });
  if (feeling === 'cold') return Object.assign(base, { mode: 'light', title: 'Head cold – light session',
    message: 'Symptoms above the neck only: about 60% of your weights, 2 sets each – or swap it for a 20–30 min walk. Stop if you feel worse.', setCap: 2, loadFactor: 0.6 });
  if (feeling === 'sore') return Object.assign(base, { mode: 'sore', title: 'Sore or achy – easier session',
    message: 'Same weights, 1 set fewer. If an exercise hurts a joint, tap ⇄ Swap for a machine or dumbbell option, or skip it.', setDelta: -1 });
  if (feeling === 'tired' || (sleep !== null && sleep < 6)) return Object.assign(base, { mode: 'tired',
    title: sleep !== null && sleep < 6 ? 'Short sleep – trimmed session' : 'Tired – trimmed session',
    message: 'Same weights, 1 set fewer per exercise. Quality over quantity today.', setDelta: -1 });
  return base;
}

/** Daily guides for fat and total sugar, scaled from the UK reference intakes (70 g fat, 90 g sugars at 2,000 kcal). */
function fatSugarGuides_(kcal) {
  return { fat: Math.round(70 * kcal / 2000), sugar: Math.round(90 * kcal / 2000) };
}

/** Calories, protein and steps for a day, after sport-day and readiness adjustments. */
function dayTargets_(s, d, dow, ready) {
  const t = currentTargets_(s);
  let kcal = t.calories, steps = t.steps, note = '';
  const sport = sportOn_(dow);
  if (sport && sport.extraKcal) { kcal += Number(sport.extraKcal) || 0; note = `includes ${sport.name.toLowerCase()}`; }
  const runK = (() => { try { return runExtraKcal_(key_(d)); } catch (e) { return 0; } })();
  if (runK) { kcal += runK; note = note ? note + ' + your run' : `includes ${runK} kcal for your run`; }
  if (ready && ready.mode === 'sick') {
    const rows = readTable_(SHEETS.TARGETS);
    const m = rows.length ? num_(rows[rows.length - 1]['Est. maintenance']) : null;
    kcal = m || t.calories + 400; steps = 0; note = 'maintenance while unwell';
  }
  return { kcal, protein: t.protein, steps, note };
}

/** Days marked cold or sick: their lifts don't count for progression, their weights don't count for calories. */
function lightDays_(logRows) {
  const set = new Set();
  logRows.forEach(r => { if (r['Date'] instanceof Date && /^(cold|sick)$/i.test(String(r['Feeling'] || ''))) set.add(key_(r['Date'])); });
  return set;
}

function ensureLogColumns_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (!sh) return;
  const headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  LOG_HEADERS.forEach(h => {
    if (h === 'Activity' && headers.includes('Badminton/Run')) return;   // older sheets: keep the existing column
    if (!headers.includes(h)) { headers.push(h); sh.getRange(1, headers.length).setValue(h).setFontWeight('bold'); }
  });
}

/** Save how you feel for a date (the evening check-in saves it for tomorrow). */
function webSetFeeling(dateStr, feeling) {
  ensureLogColumns_();
  const f = FEELINGS.includes(String(feeling)) ? String(feeling) : 'good';
  const d = parseDay_(dateStr);
  upsertLog_(d, { 'Feeling': f }, []);
  const ready = readiness_(d, readTable_(SHEETS.LOG));
  const dow = Number(fmt_(d, 'u')) % 7;
  return { date: key_(d), feeling: f, readiness: ready, session: sessionFor_(dow) };
}

// ---------- Stall detection ----------
/** Same top weight for the last 3 sessions with no extra reps → suggest a 10% reset to build back up. */
function stallReset_(exercise, lifts) {
  const name = String(exercise).toLowerCase();
  const byDay = {};
  lifts.filter(l => String(l['Exercise']).toLowerCase() === name && l['Date'] instanceof Date).forEach(l => {
    const k = key_(l['Date']);
    const e = byDay[k] || (byDay[k] = { top: 0, reps: 0 });
    e.top = Math.max(e.top, Number(l['Weight kg']) || 0); e.reps += Number(l['Reps']) || 0;
  });
  const days = Object.keys(byDay).sort().slice(-3).map(k => byDay[k]);
  if (days.length < 3 || !days[0].top) return null;
  if (!days.every(x => x.top === days[0].top)) return null;
  if (days[2].reps > days[0].reps) return null;
  const x = days[0].top * 0.9;
  return { top: days[0].top, kg: x < 20 ? round_(x, 0.5) : round_(x, 2.5) };
}

// Double progression: all sets at top of rep range → add increment
function nextLoad_(planRow, lifts) {
  const name = String(planRow['Exercise']).toLowerCase();
  const rows = lifts.filter(l => String(l['Exercise']).toLowerCase() === name && l['Date']);
  if (!rows.length) return { target: '—', kg: '', last: 'first time: pick a weight you can do with 2 reps to spare' };
  const lastKey = rows.map(r => key_(r['Date'])).sort().pop();
  const last = rows.filter(r => key_(r['Date']) === lastKey);
  const weight = Math.max(...last.map(r => Number(r['Weight kg']) || 0));
  const reps = last.map(r => Number(r['Reps']) || 0);
  const repMax = Number(planRow['Rep max']);
  const repMin = parseInt(String(planRow['Reps']), 10) || 0;
  const inc = Number(planRow['Increment kg']) || 0;
  const lastTxt = `${weight ? weight + ' kg × ' : ''}${reps.join(', ')}`;

  if (repMax && reps.length >= Number(planRow['Sets']) && reps.every(r => r >= repMax) && inc) {
    const kg = round_(weight + inc, 0.5);
    return { target: `${kg} kg ↑`, kg, last: lastTxt };
  }
  if (repMin && reps.filter(r => r < repMin).length >= 2 && weight) {
    return { target: `${weight} kg (hold; drop 5% if it happens again)`, kg: weight, last: lastTxt };
  }
  return { target: weight ? `${weight} kg` : 'bodyweight', kg: weight || '', last: lastTxt };
}

function planSignature_() {
  return JSON.stringify(planRows_().map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]]));
}

/**
 * Keep the Plan tab in step with Config.gs. Runs only when the plan in Config.gs has changed:
 * adds new exercises, removes ones you deleted, updates order/reps/rest/increments/muscles.
 * Set counts the weekly review has adjusted are kept, unless you changed that exercise's sets in Config.gs.
 */
function ensurePlanUpToDate_() {
  const props = PropertiesService.getScriptProperties();
  const sig = planSignature_();
  const prevSig = props.getProperty('PLAN_SIG');
  if (prevSig === sig) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PLAN);
  if (!sh) return;
  const prevSets = {};
  try { JSON.parse(prevSig || '[]').forEach(r => { prevSets[r[0] + '|' + r[2]] = r[3]; }); } catch (e) {}
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const c = name => h.indexOf(name);
  const wanted = new Set(planRows_().map(r => r[0] + '|' + r[2]));
  // remove exercises no longer in Config.gs (bottom-up so row numbers stay valid)
  for (let i = values.length - 1; i >= 1; i--) {
    if (!wanted.has(values[i][c('Session')] + '|' + values[i][c('Exercise')])) sh.deleteRow(i + 1);
  }
  const now = sh.getDataRange().getValues();
  planRows_().forEach(row => {
    const k = row[0] + '|' + row[2];
    const i = now.findIndex((v, j) => j > 0 && v[c('Session')] + '|' + v[c('Exercise')] === k);
    if (i < 0) { sh.appendRow(row); return; }
    const keepSets = !prevSig || prevSets[k] === row[3];      // first sync, or sets unchanged in Config → keep the sheet's value
    const vals = row.slice();
    if (keepSets) vals[3] = now[i][c('Sets')];
    sh.getRange(i + 1, 1, 1, vals.length).setValues([vals]);
  });
  props.setProperty('PLAN_SIG', sig);
}

/** exercise name (lower case) → muscle, including every swap. */
function muscleMap_(plan) {
  const m = {};
  plan.forEach(p => {
    m[String(p['Exercise']).toLowerCase()] = p['Muscle'];
    (alternates_()[String(p['Exercise'])] || []).forEach(([alt]) => {
      if (!m[alt.toLowerCase()]) m[alt.toLowerCase()] = p['Muscle'];
    });
  });
  return m;
}

// ---------- Deload + session targets ----------

/** Every 6th week is a deload week (weeks 6, 12, 18 …). */
function isDeloadWeek_(weekNum) { const n = Number(trainingCfg_().deloadEvery) || 0; return n > 0 && weekNum > 0 && weekNum % n === 0; }

/** Lifts logged in deload weeks are ignored for progression, so light weeks don't reset your targets. */
function progressionLifts_(lifts, start) {
  return lifts.filter(l => l['Date'] instanceof Date && !isDeloadWeek_(weekNumber_(start, l['Date'])));
}

function deloadKg_(kg) {
  if (!kg) return '';
  const x = kg * 0.6;
  return x < 20 ? round_(x, 0.5) : round_(x, 2.5);
}

/**
 * Today's targets for one session: progression from history, then stall resets, deload and how you feel.
 * ctx (optional): { ready: readiness_(), log: Daily Log rows }.
 */
function sessionTargets_(sessionName, d, allLifts, plan, s, ctx) {
  ctx = ctx || {};
  const ready = ctx.ready || { mode: 'normal', setDelta: 0, setCap: 0, loadFactor: 1 };
  const light = lightDays_(ctx.log || []);
  const weekNum = weekNumber_(s['Start date'], d);
  const deload = isDeloadWeek_(weekNum);
  const dayKey = key_(d);
  const before = progressionLifts_(allLifts.filter(l => l['Date'] instanceof Date && key_(l['Date']) < dayKey), s['Start date'])
    .filter(l => !light.has(key_(l['Date'])));
  const scale = kg => { const x = kg * ready.loadFactor; return x < 20 ? round_(x, 0.5) : round_(x, 2.5); };

  // One function for the plan exercise and each swap, so all get the same adjustments
  const adjust = (n, name) => {
    let kg = n.kg, target = n.target;
    if (!deload && kg && !/↑/.test(target)) {
      const st = stallReset_(name, before);
      if (st) { kg = st.kg; target = `${kg} kg (reset −10%: stuck at ${st.top} kg for 3 sessions)`; }
    }
    if (deload) { kg = deloadKg_(Number(kg) || 0); target = kg ? `${kg} kg (deload)` : 'easy – deload week'; }
    else if (ready.loadFactor < 1 && kg) { kg = scale(Number(kg)); target = `${kg} kg (light day)`; }
    else if (ready.loadFactor < 1) target = 'easy – light day';
    return { kg, target };
  };

  return plan.filter(p => p['Session'] === sessionName)
    .sort((a, b) => a['Order'] - b['Order'])
    .map(p => {
      const n = adjust(nextLoad_(p, before), String(p['Exercise']));
      const last = nextLoad_(p, before).last;
      let sets = Number(p['Sets']) || 1;
      if (deload) sets = Math.min(2, sets);
      if (ready.setCap) sets = Math.min(ready.setCap, sets);
      if (ready.setDelta) sets = Math.max(1, sets + ready.setDelta);
      const alts = (alternates_()[String(p['Exercise'])] || []).map(([name, inc]) => {
        const raw = nextLoad_(Object.assign({}, p, { 'Exercise': name, 'Increment kg': inc }), before);
        const a = adjust(raw, name);
        return { exercise: name, target: a.target, kg: a.kg, last: raw.last };
      });
      return { exercise: String(p['Exercise']), order: Number(p['Order']), muscle: String(p['Muscle']),
        sets, reps: String(p['Reps']), rest: String(p['Rest']), target: n.target, kg: n.kg, last, alts };
    });
}

// ---------- Weekly review ----------
function weeklyReview() {
  ensurePlanUpToDate_();
  ensureLogColumns_();
  const s = getSettings_();
  const today = new Date();
  const log = readTable_(SHEETS.LOG);
  const lifts = readTable_(SHEETS.LIFTS);
  const plan = readTable_(SHEETS.PLAN);
  const targets = readTable_(SHEETS.TARGETS);

  const thisWeek = rowsBetween_(log, addDays_(today, -6), today);
  const lastWeek = rowsBetween_(log, addDays_(today, -13), addDays_(today, -7));
  const weekNum = weekNumber_(s['Start date'], today);
  const prevRow = targets.length ? targets[targets.length - 1] : null;
  const prevTarget = prevRow ? Number(prevRow['Calorie target']) : Number(s['Starting calorie target']);

  // Days you were unwell are left out of the calorie maths (appetite and water weight both go odd)
  const well = r => String(r['Feeling'] || '').toLowerCase() !== 'sick';
  const sickDays = thisWeek.filter(r => !well(r)).length;
  const avgW = avg_(thisWeek.filter(well).map(r => r['Weight kg']));
  const prevAvgW = avg_(lastWeek.filter(well).map(r => r['Weight kg'])) ||
    (prevRow ? Number(prevRow['7-day avg weight']) || null : null);
  const change = avgW && prevAvgW ? avgW - prevAvgW : null;
  const kcalDays = thisWeek.filter(well).filter(r => num_(r['kcal']) !== null).length;
  const avgKcal = avg_(thisWeek.filter(well).map(r => r['kcal']));
  const waist = lastValue_(thisWeek, 'Waist cm');
  const prevWaist = lastValue_(log.filter(r => r['Date'] && key_(r['Date']) < key_(addDays_(today, -6))), 'Waist cm');

  const cal = calorieDecision_({
    weekNum, kcalDays, avgKcal, change, prevTarget, waist, prevWaist,
    floor: Number(s['Calorie floor']), rate: Number(s['Weekly change goal kg']),
  });

  // Sets per muscle this week, from the Lifts log
  const muscleOf = muscleMap_(plan);
  const weekLifts = rowsBetween_(lifts, addDays_(today, -6), today);
  const setsPerMuscle = {};
  weekLifts.forEach(l => {
    const m = muscleOf[String(l['Exercise']).toLowerCase()] || 'Other';
    setsPerMuscle[m] = (setsPerMuscle[m] || 0) + 1;
  });

  // Main-lift progress: best this week vs best before
  const progress = mainLifts_().map(name => {
    const all = lifts.filter(l => String(l['Exercise']).toLowerCase() === name.toLowerCase() && l['Date']);
    const cutoff = key_(addDays_(today, -6));
    const now = Math.max(0, ...all.filter(l => key_(l['Date']) >= cutoff).map(l => Number(l['Weight kg']) || 0));
    const before = Math.max(0, ...all.filter(l => key_(l['Date']) < cutoff).map(l => Number(l['Weight kg']) || 0));
    return { lift: name, thisWeek: now || null, previousBest: before || null };
  });

  const feedback = thisWeek.filter(r => r['Session']).map(r => ({
    date: key_(r['Date']), session: r['Session'], pump: r['Pump'],
    soreness: r['Soreness'], jointPain: r['Joint pain'],
  }));

  // Volume: decide set changes from feedback and write them into the Plan tab
  const vol = volumeDecision_(plan, feedback, weekNum);
  applyPlanChanges_(vol.changes);
  const gymDays = new Set(weekLifts.filter(l => l['Date'] instanceof Date).map(l => key_(l['Date'])));

  // Protein: average over the days you logged food, and how many days hit the target / the minimum
  const pTarget = Number(s['Protein target g']) || 0, pMin = Number(s['Protein minimum g']) || 0;
  const pDays = thisWeek.map(r => num_(r['Protein g'])).filter(v => v !== null && v > 0);
  const protein = { avg: pDays.length ? Math.round(pDays.reduce((a, b) => a + b, 0) / pDays.length) : null, daysLogged: pDays.length,
    daysHitTarget: pDays.filter(v => pTarget && v >= pTarget).length, daysBelowMin: pDays.filter(v => pMin && v < pMin).length,
    target: pTarget, minimum: pMin,
    prevAvg: (() => { const v = lastWeek.map(r => num_(r['Protein g'])).filter(x => x !== null && x > 0); return v.length ? Math.round(v.reduce((a, b) => a + b, 0) / v.length) : null; })() };

  const summary = {
    weekNumber: weekNum,
    daysLogged: thisWeek.length,
    avgWeight7d: rnd_(avgW, 2), weeklyChangeKg: rnd_(change, 2),
    avgKcal: rnd_(avgKcal, 0), kcalDaysLogged: kcalDays,
    avgProtein: protein.avg, proteinDaysLogged: protein.daysLogged, proteinDaysHitTarget: protein.daysHitTarget,
    proteinDaysBelowMinimum: protein.daysBelowMin, proteinMinimum: pMin, previousWeekAvgProtein: protein.prevAvg,
    avgFat: rnd_(avg_(thisWeek.map(r => r['Fat g'])), 0), fatGuide: fatSugarGuides_(prevTarget).fat,
    avgSugar: rnd_(avg_(thisWeek.map(r => r['Sugar g'])), 0), sugarGuide: fatSugarGuides_(prevTarget).sugar,
    avgSteps: rnd_(avg_(thisWeek.map(r => r['Steps'])), 0),
    avgSleep: rnd_(avg_(thisWeek.map(r => r['Sleep h'])), 1),
    gymSessions: gymDays.size,
    sportOrCardio: thisWeek.filter(r => activityOf_(r)).length,
    waistCm: waist, previousWaistCm: prevWaist, waistGoal: s['Waist goal cm'],
    estMaintenance: cal.maintenance, previousCalorieTarget: prevTarget,
    newCalorieTarget: cal.target, calorieDecision: cal.reason,
    proteinTarget: s['Protein target g'], nextStepTarget: stepTarget_(weekNum + 1),
    setsPerMuscle, mainLiftProgress: progress, sessionFeedback: feedback,
    deloadDue: isDeloadWeek_(weekNum + 1) ? 'Next week is a deload week (about 60% weights, 2 sets)' : 'No',
    setChangesForNextWeek: vol.changes.map(c => `${c.session}: ${c.exercise} ${c.from} → ${c.to} sets`),
    volumeNotes: vol.notes,
    daysUnwell: sickDays,
    patternInsights: insights_(log, lifts, plan, s, today),
  };

  const coach = callGemini_(weeklyPrompt_(summary)) ||
    '(Coaching summary unavailable – check the GEMINI_API_KEY script property.)';

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TARGETS).appendRow([
    today, weekNum, summary.avgWeight7d, summary.weeklyChangeKg, summary.avgKcal,
    summary.avgProtein, summary.avgSteps, summary.avgSleep, summary.gymSessions,
    summary.sportOrCardio, waist, cal.maintenance || '', cal.target,
    s['Protein target g'], summary.nextStepTarget, stripHtml_(coach),
  ]);

  const html = `<div style="font-family:Arial,sans-serif;font-size:15px;color:#1B2631">
    <h2 style="margin:0">Week ${weekNum} review</h2>
    <table cellpadding="6" style="border-collapse:collapse;margin:12px 0">
      ${tr_('7-day avg weight', fmtNum_(summary.avgWeight7d, ' kg'))}
      ${tr_('Change vs last week', fmtNum_(summary.weeklyChangeKg, ' kg'))}
      ${tr_('Waist', waist ? waist + ' cm' + (prevWaist ? ` (was ${prevWaist})` : '') : 'not logged')}
      ${tr_('Avg kcal', fmtNum_(summary.avgKcal))}
      ${tr_('<b>Avg protein</b>', proteinLine_(protein))}
      ${tr_('Avg fat / sugar', `${fmtNum_(summary.avgFat, ' g')} / ${fmtNum_(summary.avgSugar, ' g')} <span style="color:#5B6874">(guides ~${summary.fatGuide} g / ~${summary.sugarGuide} g)</span>`)}
      ${tr_('Avg steps / sleep', `${fmtNum_(summary.avgSteps)} / ${fmtNum_(summary.avgSleep, ' h')}`)}
      ${tr_('Gym / sport + cardio', `${summary.gymSessions} of ${sessionsPerWeek_()} / ${summary.sportOrCardio}`)}
      ${tr_('Est. maintenance', cal.maintenance ? cal.maintenance + ' kcal' : 'from week 3')}
      ${tr_('<b>Next week calories</b>', `<b>${cal.target} kcal</b> – ${cal.reason}`)}
      ${tr_('<b>Next week steps</b>', `<b>${summary.nextStepTarget.toLocaleString()}</b>`)}
    </table>
    <h3 style="margin:12px 0 4px">Plan changes for next week</h3>
    <ul>${vol.changes.length ? vol.changes.map(c => `<li><b>${c.session}</b>: ${c.exercise} ${c.from} → ${c.to} sets</li>`).join('') : '<li>No set changes.</li>'}
    ${vol.notes.map(n => `<li style="color:#5B6874">${n}</li>`).join('')}
    ${isDeloadWeek_(weekNum + 1) ? '<li><b>Next week is a deload week</b>: about 60% weights, 2 sets each. The app sets the weights for you.</li>' : ''}</ul>
    ${summary.patternInsights.length ? `<h3 style="margin:12px 0 4px">What I noticed</h3><ul>${summary.patternInsights.map(x => `<li>${x}</li>`).join('')}</ul>` : ''}
    ${sickDays ? `<p style="color:#5B6874">${sickDays} day(s) marked unwell were left out of the calorie calculation.</p>` : ''}
    ${coach}</div>`;
  let runLine = ''; try { runLine = runWeekSummary_(today); } catch (e) { console.error(e); }
  MailApp.sendEmail({ name: appName_(), to: s['Email'], subject: `Week ${weekNum} review: ${cal.target} kcal next week` + (protein.avg ? ` · protein avg ${protein.avg} g` : ''),
    htmlBody: runLine ? html.replace('</table>', tr_('🏃 Running', runLine.replace(/^Running: /, '')) + '</table>') : html });
  try { agentWeekly_(summary, cal, vol); } catch (e) { console.error('Agent weekly: ' + e.message); }
}

/** "128 g a day (target 140 g) · hit target 4 of 6 days · ▲ 12 g vs last week" */
function proteinLine_(p) {
  if (!p.avg) return 'no food logged this week';
  const pct = p.target ? p.avg / p.target : 1;
  const colour = pct >= 0.95 ? '#1E8449' : pct >= 0.85 ? '#B9770E' : '#C0392B';
  let out = `<b style="color:${colour}">${p.avg} g a day</b>` + (p.target ? ` (target ${p.target} g)` : '');
  if (p.target) out += ` · hit target ${p.daysHitTarget} of ${p.daysLogged} day${p.daysLogged === 1 ? '' : 's'}`;
  if (p.daysBelowMin) out += ` · ${p.daysBelowMin} below the ${p.minimum} g minimum`;
  if (p.prevAvg) { const d = p.avg - p.prevAvg; out += d ? ` · ${d > 0 ? '▲' : '▼'} ${Math.abs(d)} g vs last week` : ' · same as last week'; }
  return out;
}

/**
 * Weekly set changes (pure function – no Sheets calls, easy to test).
 * Good session (pump ok/great, soreness none/some, no joint pain) → +1 set for one muscle in that session,
 * priority muscles first. Caps: weeklySetCap per muscle (default 14), 5 sets per exercise, maxSetsPerSession
 * per session (keeps it inside your time limit). After a deload week every exercise resets to the starting plan.
 */
function volumeDecision_(plan, feedback, weekNum) {
  const changes = [], notes = [];
  const rows = plan.map(p => ({ session: String(p['Session']), exercise: String(p['Exercise']), order: Number(p['Order']),
    muscle: String(p['Muscle']), sets: Number(p['Sets']) || 1 }));

  if (isDeloadWeek_(weekNum)) {
    rows.forEach(r => {
      const base = planRows_().find(b => b[0] === r.session && b[2] === r.exercise);
      if (base && base[3] !== r.sets) changes.push({ session: r.session, exercise: r.exercise, from: r.sets, to: base[3] });
    });
    notes.push('Deload week done: sets reset to the starting plan, and weights pick up where you left off.');
    return { changes, notes };
  }
  if (weekNum < 2) { notes.push('Week 1: sets held while you settle in.'); return { changes, notes }; }

  const weekly = {};
  rows.forEach(r => { weekly[r.muscle] = (weekly[r.muscle] || 0) + r.sets; });
  const tc = trainingCfg_();
  const capFor = m => Number((tc.weeklySetCap || {})[m]) || 14;
  const prio = tc.priorityMuscles || [];
  const isPrio = m => prio.includes(m) ? 0 : 1;
  const priority = prio.concat(['Chest', 'Shoulders', 'Back', 'Quads', 'Hamstrings', 'Glutes', 'Rear delts', 'Biceps', 'Triceps', 'Calves']
    .filter(m => !prio.includes(m)));
  const maxSets = Number(tc.maxSetsPerSession) || 22;

  sessionNames_().forEach(session => {
    const fb = feedback.filter(f => f.session === session).pop();
    if (!fb || (!fb.pump && !fb.soreness && !fb.jointPain)) { notes.push(`${session}: no feedback logged – sets held.`); return; }
    if (String(fb.jointPain).toUpperCase() === 'Y') { notes.push(`${session}: joint pain reported – sets held. Swap the exercise that hurt.`); return; }
    if (fb.soreness === 'still sore') { notes.push(`${session}: still sore – sets held.`); return; }
    if (fb.pump === 'low') { notes.push(`${session}: low pump – sets held.`); return; }

    const inSession = rows.filter(r => r.session === session);
    const total = inSession.reduce((a, r) => a + r.sets, 0);
    if (total >= maxSets) { notes.push(`${session}: already at the ${tc.sessionMinutes}-minute limit – sets held.`); return; }
    const muscles = priority.filter(m => inSession.some(r => r.muscle === m) && weekly[m] < capFor(m));
    muscles.sort((a, b) => {
      return isPrio(a) - isPrio(b) || weekly[a] - weekly[b];
    });
    // First pass: accessories only (extra sets on the heavy main lift cost the most time and recovery).
    // Second pass: allow the main lift if nothing else qualifies.
    for (const allowMain of [false, true]) {
      for (const m of muscles) {
        const cands = inSession.filter(r => r.muscle === m && r.sets < 5 && (allowMain || r.order !== 1))
          .sort((a, b) => a.sets - b.sets);
        if (!cands.length) continue;
        const r = cands[0];
        changes.push({ session, exercise: r.exercise, from: r.sets, to: r.sets + 1 });
        r.sets++; weekly[m]++;
        return;
      }
    }
    notes.push(`${session}: every muscle is at its weekly cap – sets held.`);
  });
  return { changes, notes };
}

/** Write set changes into the Plan tab. */
function applyPlanChanges_(changes) {
  if (!changes.length) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PLAN);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const cS = h.indexOf('Session'), cE = h.indexOf('Exercise'), cSets = h.indexOf('Sets');
  changes.forEach(c => {
    for (let i = 1; i < values.length; i++) {
      if (values[i][cS] === c.session && values[i][cE] === c.exercise) { sh.getRange(i + 1, cSets + 1).setValue(c.to); break; }
    }
  });
}

/** Patterns from the last 14 days, written as plain sentences for the weekly email. */
function insights_(log, lifts, plan, s, today) {
  const out = [];
  const t = currentTargets_(s);
  const rows = rowsBetween_(log, addDays_(today, -13), today);
  const last7 = rowsBetween_(log, addDays_(today, -6), today);
  const dowOf = r => Number(fmt_(r['Date'], 'u')) % 7;
  const names = ['Sundays', 'Mondays', 'Tuesdays', 'Wednesdays', 'Thursdays', 'Fridays', 'Saturdays'];

  // Weekday vs weekend protein
  const wd = rows.filter(r => num_(r['Protein g']) !== null && dowOf(r) >= 1 && dowOf(r) <= 5).map(r => Number(r['Protein g']));
  const we = rows.filter(r => num_(r['Protein g']) !== null && (dowOf(r) === 0 || dowOf(r) === 6)).map(r => Number(r['Protein g']));
  if (wd.length >= 2 && we.length >= 2) {
    const a = avg_(wd), b = avg_(we);
    if (a - b >= 20) out.push(`Protein drops at weekends: about ${Math.round(b)} g vs ${Math.round(a)} g on weekdays. Plan one high-protein weekend meal.`);
  }

  // Days that run over on calories
  const target = r => { const sp = sportOn_(dowOf(r)); return t.calories + (sp ? Number(sp.extraKcal) || 0 : 0); };
  const over = rows.filter(r => num_(r['kcal']) !== null && Number(r['kcal']) > target(r) + 100);
  const byDow = {};
  over.forEach(r => { byDow[dowOf(r)] = (byDow[dowOf(r)] || 0) + 1; });
  const repeat = Object.keys(byDow).filter(k => byDow[k] >= 2);
  if (repeat.length) out.push(`Calories ran over on ${repeat.map(k => names[k]).join(' and ')} in both of the last two weeks – worth planning those days ahead.`);
  else if (over.length >= 4) out.push(`Calories were over target on ${over.length} of the last 14 days.`);

  // Protein hit rate
  const pDays = last7.filter(r => num_(r['Protein g']) !== null);
  const hit = pDays.filter(r => Number(r['Protein g']) >= t.protein).length;
  if (pDays.length >= 3 && hit <= pDays.length / 2) out.push(`Protein target hit on ${hit} of ${pDays.length} logged days. A whey shake or protein yogurt closes most gaps.`);

  // Fat and sugar against the reference-intake guides
  const g = fatSugarGuides_(t.calories);
  const sug = avg_(last7.map(r => r['Sugar g'])), fatAvg = avg_(last7.map(r => r['Fat g']));
  const sugDays = last7.filter(r => num_(r['Sugar g']) !== null).length;
  if (sugDays >= 3 && sug > g.sugar * 1.15) out.push(`Sugar averaged about ${Math.round(sug)} g a day against a guide of about ${g.sugar} g. Sugar in chai, sweets, juice and fizzy drinks is the usual cause; fruit and milk are fine.`);
  if (sugDays >= 3 && fatAvg > g.fat * 1.2) out.push(`Fat averaged about ${Math.round(fatAvg)} g a day against a guide of about ${g.fat} g. Measuring cooking oil and ghee is the easiest fix.`);

  // Sleep
  const short = last7.filter(r => num_(r['Sleep h']) !== null && Number(r['Sleep h']) < 6).length;
  if (short >= 3) out.push(`Under 6 hours' sleep on ${short} nights this week. Short sleep slows both fat loss and strength gains.`);

  // Steps
  const stepDays = last7.filter(r => num_(r['Steps']) !== null);
  const met = stepDays.filter(r => Number(r['Steps']) >= t.steps || activityOf_(r)).length;
  if (stepDays.length >= 3 && met < stepDays.length / 2) out.push(`Step target reached on ${met} of ${stepDays.length} days – a 15-minute walk after lunch usually fixes it.`);

  // Logging consistency
  const foodDays = last7.filter(r => num_(r['kcal']) !== null).length;
  if (foodDays < 5) out.push(`Food logged on ${foodDays} of 7 days. The calorie adjustment needs at least 5.`);

  // Skipped sessions and why
  const skips = last7.filter(r => String(r['Skipped'] || '')).map(r => String(r['Skipped']));
  if (skips.length) {
    const why = {}; skips.forEach(x => { why[x] = (why[x] || 0) + 1; });
    const txt = Object.keys(why).map(k => `${why[k]} ${k}`).join(', ');
    const busy = (why.busy || 0) + (why.travel || 0);
    out.push(`Skipped ${skips.length} session${skips.length > 1 ? 's' : ''} this week (${txt}).` +
      (busy >= 2 ? ' Busy weeks happen – the 25-minute home workout (🏠 on the Gym tab) keeps progress moving.' :
       why.tired >= 2 ? ' Tiredness is behind more than one skip – sleep is worth a look before training harder.' : ''));
  }

  // Stalled lifts
  const seen = new Set();
  plan.forEach(p => {
    const n = String(p['Exercise']);
    if (seen.has(n)) return; seen.add(n);
    const st = stallReset_(n, lifts.filter(l => l['Date'] instanceof Date));
    if (st) out.push(`${n} has been stuck at ${st.top} kg for 3 sessions – next target resets to ${st.kg} kg to build back up.`);
  });
  return out;
}

/**
 * Adaptive calories (pure function – no Sheets calls, easy to test).
 * d.rate = goal weight change per week in kg (negative = lose, positive = gain, 0 = maintain).
 */
function calorieDecision_(d) {
  const roundTo10 = x => Math.round(x / 10) * 10;
  const rate = typeof d.rate === 'number' ? d.rate : -0.37;
  const band = Math.max(0.12, Math.abs(rate) / 3);
  if (d.weekNum < 3) return { target: d.prevTarget, maintenance: null, reason: 'weeks 1–2: holding while water weight settles' };
  if (d.kcalDays < 5 || d.avgKcal === null || d.change === null) {
    return { target: d.prevTarget, maintenance: null, reason: 'fewer than 5 days logged: target unchanged' };
  }
  const maintenance = roundTo10(d.avgKcal - (d.change * CONFIG.KCAL_PER_KG / 7));
  if (rate <= 0 && d.waist && d.prevWaist && d.waist < d.prevWaist && Math.abs(d.change) < 0.15) {
    return { target: d.prevTarget, maintenance, reason: 'waist down, weight flat: recomposition working, no change' };
  }
  const desired = maintenance + rate * CONFIG.KCAL_PER_KG / 7;
  let target = Math.min(d.prevTarget + CONFIG.MAX_WEEKLY_CHANGE, Math.max(d.prevTarget - CONFIG.MAX_WEEKLY_CHANGE, desired));
  target = Math.max(d.floor, roundTo10(target));
  const diff = d.change - rate, goal = `${rate > 0 ? '+' : ''}${rate} kg/week`;
  let reason;
  if (Math.abs(diff) <= band) reason = `on pace (goal ${goal}): fine-tuned to estimated maintenance`;
  else if (rate < 0) reason = diff < 0 ? `losing faster than the ${goal} goal: eating a little more to protect muscle`
    : `losing slower than the ${goal} goal: small trim`;
  else if (rate > 0) reason = diff > 0 ? `gaining faster than the ${goal} goal (extra would be mostly fat): small trim`
    : `gaining slower than the ${goal} goal: small increase`;
  else reason = diff < 0 ? 'weight drifting down: small increase' : 'weight drifting up: small trim';
  if (target === d.prevTarget) reason = 'on pace: no change';
  return { target, maintenance, reason };
}

/** One paragraph describing the person, for every Gemini prompt. */
function profileText_() {
  const p = profile_(), m = scheduleMap_(), sp = sports_();
  const days = [1, 2, 3, 4, 5, 6, 0].filter(i => m[i]).map(i => `${DAY_CODES[i]} ${m[i]}`).join(', ');
  return [
    `${p.name}: ${p.age || '?'}-year-old ${p.sex || ''}, ${p.heightCm || '?'} cm, about ${p.weightKg || '?'} kg.`,
    `Experience: ${p.experience || 'not stated'}. Goal: ${p.goal || 'not stated'}.`,
    `Gym: ${days} (${trainingCfg_().sessionMinutes}-minute sessions).`,
    sp.length ? `Sport/cardio: ${sp.map(s => s.name + (s.days.length ? ` (${s.days.join('/')} ${s.when})` : '') + (s.restNextMorning ? ', rest the morning after' : '')).join('; ')}.` : '',
    `Diet: ${p.diet || 'not stated'}.${p.avoid ? ' Avoids: ' + p.avoid + '.' : ''}`,
    p.injuries && !/^none$/i.test(p.injuries) ? `Injuries/limits: ${p.injuries}.` : '',
    p.lifestyle ? `Lifestyle: ${p.lifestyle}` : '',
    p.notes ? `Notes: ${p.notes}` : '',
  ].filter(Boolean).join('\n');
}

function weeklyPrompt_(summary) {
  const name = profile_().name;
  return `You are ${name}'s strength and nutrition coach.
${profileText_()}

Here is this week's data as JSON (calorie target already decided by the formula; do not change it):
${JSON.stringify(summary, null, 2)}

Write the weekly coaching note as simple HTML (<p>, <ul>, <li>, <b> only), under 220 words:
1. One-line verdict on the week.
2. Use patternInsights (already computed from ${name}'s logs) to make the advice specific. Do not repeat them word
   for word; pick the one or two that matter most.
3. Training volume: the set changes in setChangesForNextWeek and volumeNotes are already decided and applied
   to the plan. Explain them in one or two lines; do not invent other set changes. If feedback was missing,
   remind ${name} to tap pump/soreness/joint pain when finishing a session.
4. Protein in one line: avgProtein vs proteinTarget, proteinDaysHitTarget of proteinDaysLogged, and the change vs
   previousWeekAvgProtein. If it's short, name one or two easy protein foods from their own diet to close the gap.
5. Lift progress in one line.
6. Exactly 3 actions for next week (food, steps/sleep, training).
Be direct and honest; no empty praise. Never suggest calories below the floor. Mark missing data as missing.`;
}

// ---------- Monthly review ----------
function monthlyReview() {
  const s = getSettings_();
  const today = new Date();
  const log = readTable_(SHEETS.LOG);
  const lifts = readTable_(SHEETS.LIFTS);
  const thisM = rowsBetween_(log, addDays_(today, -30), addDays_(today, -1));
  const prevM = rowsBetween_(log, addDays_(today, -60), addDays_(today, -31));
  const best = (rows, from, to) => mainLifts_().map(n => ({
    lift: n,
    best: Math.max(0, ...rowsBetween_(rows, from, to)
      .filter(l => String(l['Exercise']).toLowerCase() === n.toLowerCase())
      .map(l => Number(l['Weight kg']) || 0)) || null,
  }));
  const data = {
    thisMonth: {
      avgWeight: rnd_(avg_(thisM.map(r => r['Weight kg'])), 1),
      latestWaist: lastValue_(thisM, 'Waist cm'),
      gymSessions: thisM.filter(r => String(r['Session'] || '').trim()).length,
      avgProtein: rnd_(avg_(thisM.map(r => r['Protein g'])), 0),
      avgSteps: rnd_(avg_(thisM.map(r => r['Steps'])), 0),
      avgSleep: rnd_(avg_(thisM.map(r => r['Sleep h'])), 1),
      bestLifts: best(lifts, addDays_(today, -30), addDays_(today, -1)),
    },
    previousMonth: {
      avgWeight: rnd_(avg_(prevM.map(r => r['Weight kg'])), 1),
      latestWaist: lastValue_(prevM, 'Waist cm'),
      gymSessions: prevM.filter(r => String(r['Session'] || '').trim()).length,
      bestLifts: best(lifts, addDays_(today, -60), addDays_(today, -31)),
    },
    gymAdherencePct: Math.round(100 * thisM.filter(r => String(r['Session'] || '').trim()).length / (sessionsPerWeek_() * 30 / 7)),
    waistGoal: s['Waist goal cm'],
    weekNumber: weekNumber_(s['Start date'], today),
  };
  const prompt = `You are ${profile_().name}'s strength coach.
${profileText_()}
Monthly data as JSON:
${JSON.stringify(data, null, 2)}
Write simple HTML (<p>, <ul>, <li>, <b>), under 200 words: verdict; waist and weight trend; lift progress;
adherence; then ONE recommendation: continue as is, add training volume, or adjust calories – with the reason.
If week 8–12 and lifts and waist are both stalled, recommend adding sets first. Mark missing data as missing.`;
  const coach = callGemini_(prompt) || '<p>(Summary unavailable – check GEMINI_API_KEY.)</p>';
  MailApp.sendEmail({ name: appName_(),
    to: s['Email'], subject: 'Monthly coach review',
    htmlBody: `<div style="font-family:Arial,sans-serif;font-size:15px;color:#1B2631"><h2>Monthly review</h2>${coach}</div>`,
  });
}

// ---------- Gemini ----------
let lastGeminiError_ = '';

/** Simple text prompt → text (json=true asks for a JSON reply). */
function callGemini_(prompt, json) {
  const body = { contents: [{ role: 'user', parts: [{ text: prompt }] }] };
  if (json) body.generationConfig = { responseMimeType: 'application/json' };
  return geminiGenerate_(body);
}

/** Sends a full generateContent body, trying each model in CONFIG.MODELS until one works. Returns text or null. */
function geminiGenerate_(body) {
  const r = geminiCall_(body, true);
  if (!r) return null;
  return r.text.replace(/^\x60{3}(?:html|json)?\s*/i, '').replace(/\x60{3}\s*$/, '').trim();   // strip code fences
}

/** Like geminiGenerate_ but returns all reply parts (text and function calls) – used by the agent. */
function geminiRaw_(body, smart) { return geminiCall_(body, false, smart); }

/** Strong-model call for photos (returns text). */
function geminiSmart_(body, budgetMs) {
  const r = geminiCall_(body, true, true, budgetMs);
  return r ? r.text.replace(/^\x60{3}(?:html|json)?\s*/i, '').replace(/\x60{3}\s*$/, '').trim() : null;
}

/** budgetMs (optional): stop trying more models / retries once this much time has passed, so the phone isn't left waiting. */
function geminiCall_(body, needText, smart, budgetMs) {
  lastGeminiError_ = '';
  const deadline = budgetMs ? Date.now() + budgetMs : 0;
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY') || configKey_();
  if (!key) { lastGeminiError_ = 'No Gemini key – paste it into geminiKey in Config.gs and run start'; return null; }
  const remembered = smart ? props.getProperty('GEMINI_SMART_MODEL') : props.getProperty('GEMINI_MODEL');
  const list = smart ? CONFIG.SMART_MODELS : CONFIG.MODELS;
  // a remembered "lite" model never jumps the queue for smart calls – a stronger model may be free again
  let models = remembered && !(smart && /lite/i.test(remembered)) ? [remembered].concat(list.filter(m => m !== remembered)) : list;
  // models whose free daily allowance ran out wait until it resets (each model has its own allowance)
  let cool = {}; try { cool = JSON.parse(props.getProperty('GEMINI_COOL') || '{}'); } catch (e) {}
  const nowMs = Date.now(), resting = models.filter(m => cool[m] && cool[m] > nowMs);
  if (resting.length && resting.length < models.length) models = models.filter(m => !resting.includes(m)).concat(resting);
  const tried = {}; let quotaHits = 0;

  for (let mi = 0; mi < models.length; mi++) {
    const model = models[mi];
    if (deadline && mi > 0 && Date.now() > deadline) { lastGeminiError_ = 'Gemini is slow right now (' + lastGeminiError_ + ')'; break; }
    try {
      const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body),
      });
      const code = res.getResponseCode();
      const raw = res.getContentText();
      if (code === 404 || code === 403 && /model|not found|not available|access/i.test(raw)) {
        lastGeminiError_ = `${model}: ${code}`;
        continue;                                   // model unavailable to this key – try the next one
      }
      if (code !== 200) {
        let msg = raw;
        try { msg = JSON.parse(raw).error.message; } catch (e) {}
        lastGeminiError_ = `${model}: ${code} – ${msg}`;
        console.error('Gemini error ' + lastGeminiError_);
        if (code === 429) {
          // a per-model limit (free tier, per minute or per day): park this model and try the next one
          quotaHits++;
          const secs = Number((String(msg).match(/retry in ([\d.]+)s/i) || [])[1]) || 0;
          const perDay = /per.?day|PerDay|free_tier_requests/i.test(msg) && !/per.?minute/i.test(msg);
          cool[model] = Date.now() + (perDay ? 6 * 3600e3 : Math.max(60, secs) * 1000);
          try { props.setProperty('GEMINI_COOL', JSON.stringify(cool)); } catch (e) {}
          lastGeminiError_ = 'Your free Gemini allowance is used up for now (' + model + ')';
          continue;
        }
        if (code === 400 || code === 401) return null;  // bad request / key – other models won't help
        if (smart && (code === 503 || code === 500) && !tried[model] && (!deadline || Date.now() + 1500 < deadline)) { tried[model] = true; Utilities.sleep(1500); mi--; }  // busy: one retry
        continue;
      }
      const data = JSON.parse(raw);
      const parts = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) || [];
      const text = parts.filter(p => !p.thought).map(p => p.text || '').join('').trim();
      const hasCall = parts.some(p => p.functionCall);
      if (needText ? !text : (!text && !hasCall)) { lastGeminiError_ = `${model}: empty reply`; continue; }
      if (remembered !== model) props.setProperty(smart ? 'GEMINI_SMART_MODEL' : 'GEMINI_MODEL', model);
      return { text, parts, model };
    } catch (e) {
      lastGeminiError_ = `${model}: ${e.message}`;
      console.error(e);
    }
  }
  if (quotaHits) lastGeminiError_ = 'Today\'s free Gemini allowance is used up on every model – it resets within a day. You can still type foods in or pick a one-tap meal.';
  console.error('All Gemini models failed. Last error: ' + lastGeminiError_);
  return null;
}

/** Run this from the editor to check your key and see which model works. */
function testGemini() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('GEMINI_MODEL'); props.deleteProperty('GEMINI_SMART_MODEL');
  const out = callGemini_('Reply with JSON only: {"kcal": 100, "protein": 5}', true);
  console.log(out ? `OK – quick estimates use ${props.getProperty('GEMINI_MODEL')}: ${out}` : `FAILED – ${lastGeminiError_}`);
  const smart = geminiSmart_({ contents: [{ role: 'user', parts: [{ text: 'Reply with JSON only: {"ok": true}' }] }], generationConfig: { responseMimeType: 'application/json' } });
  console.log(smart ? `OK – photos, voice and the coach use ${props.getProperty('GEMINI_SMART_MODEL')}` : `Photos: FAILED – ${lastGeminiError_}`);
}

// ---------- Phone logging page (web app) ----------
// Deploy ▸ New deployment ▸ Web app (Execute as: Me, Who has access: Only myself)

function doGet() {
  // The Telegram link is a second, public deployment. Only you (signed in) get the app – anyone else gets nothing.
  if (telegramOn_() && PropertiesService.getScriptProperties().getProperty('APP_GUARD') !== 'off' && !Session.getActiveUser().getEmail()) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">This link is private.</p>');
  }
  return HtmlService.createHtmlOutputFromFile('Log')
    .setTitle(appName_())
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Everything the page needs for one date. Only strings/numbers (no Date objects). */
function webGetDay(dateStr) {
  // housekeeping that only needs to happen once a day
  oncePerDay_('dayChecks', () => {
    ensureLiftIds_(); ensureLogColumns_();
    try { agentHousekeeping_(); } catch (e) { console.error(e); }
    try { updateDailyCheck_(); } catch (e) { console.error('Update check: ' + e.message); }
    try { ensureTriggers_(); } catch (e) { console.error(e); }
  });
  ensurePlanUpToDate_();                        // cheap: compares Config with the last sync
  return withMemo_(() => webGetDayData_(dateStr));
}

function webGetDayData_(dateStr) {
  const s = getSettings_();
  const d = dateStr ? parseDay_(dateStr) : parseDay_(key_(new Date()));
  const dow = Number(fmt_(d, 'u')) % 7;
  const sport = sportOn_(dow);
  const dayKey = key_(d);
  const weekNum = weekNumber_(s['Start date'], d);

  const allLifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date);
  const plan = readTable_(SHEETS.PLAN);
  const logRows = readTable_(SHEETS.LOG);
  const ready = readiness_(d, logRows);
  const dt = dayTargets_(s, d, dow, ready);
  const sessions = {};
  const sessionOrder = sessionNames_();
  sessionOrder.forEach(name => {
    sessions[name] = sessionTargets_(name, d, allLifts, plan, s, { ready, log: logRows });
  });
  const homeRows = homeRows_();
  [HOME_LOWER, HOME_UPPER].forEach(name => {
    sessions[name] = homeTargets_(sessionTargets_(name, d, allLifts, homeRows, s, { ready, log: logRows }), homeRows, allLifts, d);
  });
  const tomorrow = logRows.find(r => r['Date'] instanceof Date && key_(r['Date']) === key_(addDays_(d, 1)));

  const catchUp = catchUp_(d, logRows, allLifts, plan, s, ready);
  const found = findLogRow_(d);
  const log = {};
  if (found) Object.keys(found.values).forEach(k => {
    const v = found.values[k];
    if (!(v instanceof Date)) log[k] = v;
  });
  log['Activity'] = activityOf_(found ? found.values : {});

  return {
    date: dayKey,
    appName: appName_(), appTagline: appTagline_(),
    name: profile_().name,
    scheduled: sessionFor_(dow),
    restLabel: restLabel_(dow),
    sportToday: sport ? { name: sport.name, when: sport.when } : null,
    activities: activityOptions_(),
    meals: meals_(),
    sessionOrder,
    sessionsPerWeek: sessionsPerWeek_(),
    readiness: ready,
    tomorrowFeeling: tomorrow ? String(tomorrow['Feeling'] || '') : '',
    weekNum,
    deload: isDeloadWeek_(weekNum),
    sessions,
    targets: { kcal: dt.kcal, protein: dt.protein, proteinMin: Number(s['Protein minimum g']) || 120, steps: dt.steps, note: dt.note,
      fatGuide: fatSugarGuides_(dt.kcal).fat, sugarGuide: fatSugarGuides_(dt.kcal).sugar },
    log,
    foods: foodsFor_(d),
    recentFoods: recentFoods_(d),
    yesterdayFoods: foodsFor_(addDays_(d, -1)).filter(f => f.source !== 'pending').length,
    setsToday: allLifts.filter(l => key_(l['Date']) === dayKey).map(liftOut_),
    skipped: String(log['Skipped'] || ''),
    catchUp,
    home: { lower: HOME_LOWER, upper: HOME_UPPER, sessions: [HOME_LOWER, HOME_UPPER],
      forToday: homeFor_(sessionFor_(dow) || (catchUp || {}).session || '', plan) },
    nextWorkout: nextWorkout_(d, logRows, allLifts, plan, s),
    knownExercises: knownExercises_(plan, allLifts),
    recentExtras: recentExtras_(plan, allLifts, d),
    isFuture: dayKey > key_(new Date()),
  };
}

/** The next planned session after this date (within a week), with its target weights. */
function nextWorkout_(d, logRows, lifts, plan, s) {
  for (let i = 1; i <= 7; i++) {
    const x = addDays_(d, i), dow = Number(fmt_(x, 'u')) % 7, name = sessionFor_(dow);
    if (!name) continue;
    const ready = readiness_(x, logRows);
    if (ready.rest) continue;
    const ex = sessionTargets_(name, x, lifts, plan, s, { ready, log: logRows });
    return { date: key_(x), day: DAY_CODES[dow], inDays: i, session: name, mode: ready.mode, note: ready.title || '',
      deload: isDeloadWeek_(weekNumber_(s['Start date'], x)),
      exercises: ex.map(e => ({ exercise: e.exercise, sets: e.sets, reps: e.reps, target: e.target })) };
  }
  return null;
}

// ---------- Extra exercises (added on the day, "if time permits") ----------
/** Names for the extra-exercise box: plan, swaps, home moves and anything you've logged before. */
function knownExercises_(plan, lifts) {
  const out = [], seen = new Set();
  const add = n => { n = String(n || '').trim(); const k = n.toLowerCase(); if (n && !seen.has(k)) { seen.add(k); out.push(n); } };
  plan.forEach(p => add(p['Exercise']));
  Object.keys(alternates_()).forEach(k => { add(k); alternates_()[k].forEach(a => add(a[0])); });
  homeRows_().forEach(r => add(r['Exercise']));
  lifts.slice().reverse().forEach(l => add(l['Exercise']));
  return out.slice(0, 400);
}
/** Exercises you added as extras recently (logged but not in that session's plan), newest first. */
function recentExtras_(plan, lifts, d) {
  const inPlan = new Set(plan.map(p => String(p['Session']) + '|' + String(p['Exercise']).toLowerCase()));
  homeRows_().forEach(r => inPlan.add(r['Session'] + '|' + String(r['Exercise']).toLowerCase()));
  const alts = new Set(); Object.keys(alternates_()).forEach(k => alternates_()[k].forEach(a => alts.add(a[0].toLowerCase())));
  const from = key_(addDays_(d, -60)), out = [];
  lifts.filter(l => key_(l['Date']) >= from).sort((a, b) => b['Date'] - a['Date']).forEach(l => {
    const n = String(l['Exercise']);
    if (inPlan.has(String(l['Session']) + '|' + n.toLowerCase()) || alts.has(n.toLowerCase())) return;
    if (!out.some(x => x.toLowerCase() === n.toLowerCase())) out.push(n);
  });
  return out.slice(0, 8);
}
/** Target for an extra exercise from its own history: same top weight as last time, aim to beat the reps. */
function webExtraTarget(dateStr, name) {
  const n = String(name || '').trim().slice(0, 60);
  if (!n) throw new Error('Type an exercise name');
  const d = parseDay_(dateStr);
  const lifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date && key_(l['Date']) < key_(d));
  const plan = readTable_(SHEETS.PLAN);
  const canon = knownExercises_(plan, lifts).find(x => x.toLowerCase() === n.toLowerCase()) || (n.charAt(0).toUpperCase() + n.slice(1));
  const hist = lifts.filter(l => String(l['Exercise']).toLowerCase() === canon.toLowerCase());
  const pr = plan.find(p => String(p['Exercise']).toLowerCase() === canon.toLowerCase());
  const base = { exercise: canon, extra: true, sets: pr ? Number(pr['Sets']) || 3 : 3, reps: pr ? String(pr['Reps']) : '8–12', rest: pr ? String(pr['Rest']) : '75 s', alts: [] };
  if (!hist.length) return Object.assign(base, { target: 'first time – pick a weight you can do with 2 reps to spare', kg: '', last: 'first time' });
  const lastKey = hist.map(l => key_(l['Date'])).sort().pop();
  const last = hist.filter(l => key_(l['Date']) === lastKey);
  const kg = Math.max(...last.map(l => Number(l['Weight kg']) || 0));
  const reps = last.map(l => Number(l['Reps']) || 0);
  if (!pr) base.sets = Math.min(5, Math.max(2, last.length));
  return Object.assign(base, { kg: kg || '', target: kg ? `${kg} kg – beat last time` : `beat last time (${reps.join(', ')})`,
    last: `${fmt_(parseDay_(lastKey), 'd MMM')}: ${kg ? kg + ' kg × ' : ''}${reps.join(', ')}` });
}


// =====================================================================
// RUNNING – log runs, a weekly plan that builds distance safely, pace zones and stats
// =====================================================================
const RUNS_SHEET = 'Runs';
const RUN_HEADERS = ['Date', 'Type', 'Distance km', 'Duration min', 'Pace min/km', 'Avg HR', 'RPE', 'Notes', 'Source', 'ID'];
const RUN_TYPES = ['easy', 'long', 'tempo', 'intervals', 'race', 'other'];
const RUN_CAPS = { general: 20, '5k': 30, '10k': 40, half: 55, marathon: 75 };   // weekly km ceiling by goal

/** Running settings from Config.gs (null = no plan; you can still log runs). */
function runCfg_() {
  const r = cfg_().running;
  if (!r) return null;
  const goal = RUN_CAPS[String(r.goal || 'general').toLowerCase()] ? String(r.goal || 'general').toLowerCase() : 'general';
  return { goal, runsPerWeek: Math.max(1, Math.min(6, Number(r.runsPerWeek) || 2)),
    days: (r.days || []).filter(d => dayIndex_(d) >= 0).map(d => DAY_CODES[dayIndex_(d)]),
    weeklyKm: Math.max(2, Number(r.weeklyKm) || 8), longestRunKm: Math.max(2, Number(r.longestRunKm) || 5),
    recent5kMin: Number(r.recent5kMin) || 0, raceDate: /^\d{4}-\d{2}-\d{2}$/.test(String(r.raceDate || '')) ? String(r.raceDate) : '',
    eatBack: r.eatBack === undefined ? 0.5 : Math.max(0, Math.min(1, Number(r.eatBack) || 0)) };
}
function runsSheet_() { return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), RUNS_SHEET, RUN_HEADERS); }
function readRuns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(RUNS_SHEET)) return [];
  return readTable_(RUNS_SHEET).filter(r => r['Date'] instanceof Date && num_(r['Distance km']) !== null)
    .map(r => ({ date: key_(r['Date']), type: String(r['Type'] || 'easy'), km: Number(r['Distance km']), min: num_(r['Duration min']),
      hr: num_(r['Avg HR']), rpe: num_(r['RPE']), notes: String(r['Notes'] || ''), id: String(r['ID'] || ''), source: String(r['Source'] || '') }))
    .sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
}
function mondayOf_(d) { return addDays_(d, -((Number(fmt_(d, 'u')) + 6) % 7)); }   // keeps the time of day, so the date can't slip
function fmtPace_(secPerKm) { if (!secPerKm || !isFinite(secPerKm)) return ''; const s = Math.round(secPerKm); return Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0'); }
function round05_(x) { return Math.round(x * 2) / 2; }
function runKcal_(km) { return Math.round((Number(profile_().weightKg) || 70) * km); }   // ≈1 kcal per kg per km

/** 5k pace (sec/km) from your Config or your best recent effort (Riegel: T2 = T1 × (D2/D1)^1.06). */
function fiveKPace_(runs, rc, d) {
  if (rc && rc.recent5kMin) return rc.recent5kMin * 60 / 5;
  const from = key_(addDays_(d, -56));
  const best = runs.filter(r => r.date >= from && r.km >= 3 && r.min).map(r => (r.min * 60) * Math.pow(5 / r.km, 1.06) / 5);
  return best.length ? Math.min(...best) : null;
}
function paceZones_(p5) {
  if (!p5) return null;
  return { easy: fmtPace_(p5 * 1.22) + '–' + fmtPace_(p5 * 1.35), tempo: fmtPace_(p5 * 1.07), intervals: fmtPace_(p5 * 0.98), fiveK: fmtPace_(p5) };
}

/** Weekly distance target: +10% a week from what you actually ran, a cutback every 4th week, taper before a race. */
function runTargetKm_(ws, runs, rc, s) {
  const weeks = [1, 2, 3].map(i => { const a = key_(addDays_(ws, -7 * i)), b = key_(addDays_(ws, -7 * i + 6));
    return runs.filter(r => r.date >= a && r.date <= b).reduce((x, r) => x + r.km, 0); });
  const ran = weeks.filter(k => k > 0);
  let base = weeks[0] > 0 ? weeks[0] : ran.length ? ran.reduce((a, b) => a + b, 0) / ran.length : rc.weeklyKm;
  if (!runs.length) base = rc.weeklyKm / 1.1;
  let km = Math.min(RUN_CAPS[rc.goal], base * 1.1);
  const wn = weekNumber_(s['Start date'], ws);
  let note = '';
  if (wn > 0 && wn % 4 === 0) { km *= 0.8; note = 'cutback week – lighter on purpose'; }
  if (rc.raceDate) {
    const days = (parseDay_(rc.raceDate) - ws) / 864e5;
    if (days >= 0 && days < 7) { km *= 0.5; note = 'race week – taper'; }
    else if (days >= 7 && days < 14) { km *= 0.7; note = 'taper week before your race'; }
  }
  return { km: Math.max(2, round05_(km)), note };
}

/** This week's runs: types, distances and days, fitted around your gym and sport days. */
function runWeekPlan_(ws, runs, rc, s) {
  const tk = runTargetKm_(ws, runs, rc, s), n = rc.runsPerWeek, km = tk.km;
  const types = n === 1 ? ['easy'] : n === 2 ? ['easy', 'long'] : ['easy', rc.goal === 'general' ? 'easy' : (weekNumber_(s['Start date'], ws) % 2 ? 'tempo' : 'intervals'), 'long'];
  while (types.length < n) types.splice(1, 0, 'easy');
  const recentMax = Math.max(rc.longestRunKm, ...runs.filter(r => r.date >= key_(addDays_(ws, -28))).map(r => r.km), 0);
  const dist = {};
  // long run: about 40% of the week, but never less than your last long run (up to 60% of the week) and never more than +1.5 km on your longest
  const lastLong = Math.max(0, ...runs.filter(r => r.date >= key_(addDays_(ws, -14))).map(r => r.km));
  if (types.includes('long')) dist.long = Math.min(round05_(Math.max(km * 0.4, Math.min(lastLong, km * 0.6))), round05_(recentMax + 1.5));
  if (types.includes('tempo') || types.includes('intervals')) dist.q = round05_(km * 0.25);
  const left = km - (dist.long || 0) - (dist.q || 0), easyN = types.filter(x => x === 'easy').length;
  const easyKm = Math.max(2, round05_(left / Math.max(1, easyN)));
  // days: your preferred days first, then days with no gym and no sport, then any
  const order = [1, 2, 3, 4, 5, 6, 0];
  const pick = rc.days.map(c => dayIndex_(c));
  const free = order.filter(i => !pick.includes(i) && !sessionFor_(i) && !sportOn_(i));
  const rest = order.filter(i => !pick.includes(i) && !free.includes(i));
  const days = pick.concat(free, rest).slice(0, n).sort((a, b) => order.indexOf(a) - order.indexOf(b));
  // long run on the weekend day if there is one, otherwise the last day
  const longAt = types.includes('long') ? (days.findIndex(i => i === 6 || i === 0) >= 0 ? days.findIndex(i => i === 6 || i === 0) : days.length - 1) : -1;
  const queue = types.filter(x => x !== 'long');
  const plan = days.map((dow, i) => {
    const type = i === longAt ? 'long' : queue.shift();
    const dk = type === 'long' ? dist.long : (type === 'tempo' || type === 'intervals') ? dist.q : easyKm;
    const date = key_(addDays_(ws, order.indexOf(dow)));
    return { date, day: DAY_CODES[dow], type, km: Math.max(2, dk) };
  });
  return { targetKm: km, note: tk.note, plan };
}

function runDesc_(r, zones) {
  const z = zones || {};
  const at = p => p ? ` @ ${p} /km` : '';
  if (r.type === 'long') return `Long run ${r.km} km easy${at(z.easy)} – slower than you think; walk breaks are fine`;
  if (r.type === 'tempo') { const t = Math.max(1.5, round05_(r.km - 2.5));
    return `Tempo: 1.5 km easy, ${t} km comfortably hard${at(z.tempo)}, 1 km easy`; }
  if (r.type === 'intervals') { const reps = Math.max(3, Math.min(8, Math.round((r.km - 2.5) / 0.9)));
    return `Intervals: 1.5 km easy, then ${reps} × 3 min hard but controlled${at(z.intervals)} with 2 min easy jog, 1 km easy`; }
  return `Easy ${r.km} km${at(z.easy)} – conversational pace`;
}

/** Everything the Run screen needs for a date. */
function webGetRuns(dateStr) { return withMemo_(() => webGetRunsData_(dateStr)); }
function webGetRunsData_(dateStr) {
  const s = getSettings_();
  const d = dateStr ? parseDay_(dateStr) : parseDay_(key_(new Date()));
  const dayKey = key_(d), rc = runCfg_(), runs = readRuns_(), log = readTable_(SHEETS.LOG);
  const ws = mondayOf_(d), we = key_(addDays_(ws, 6));
  const weekRuns = runs.filter(r => r.date >= key_(ws) && r.date <= we);
  const p5 = fiveKPace_(runs, rc, d), zones = paceZones_(p5);
  let week = null, today = null;
  if (rc) {
    const wp = runWeekPlan_(ws, runs, rc, s);
    const used = new Set();
    const plan = wp.plan.map(p => {
      const done = weekRuns.find(r => r.date === p.date && !used.has(r.id));
      if (done) used.add(done.id);
      return Object.assign(p, { desc: runDesc_(p, zones), done: done ? { km: done.km, min: done.min, id: done.id } : null,
        missed: !done && p.date < key_(new Date()) });
    });
    // a run on a different day still counts towards the week
    const extra = weekRuns.filter(r => !used.has(r.id));
    extra.forEach(r => { const open = plan.find(p => !p.done && !p.missed); if (open) { open.done = { km: r.km, min: r.min, id: r.id, movedFrom: r.date }; used.add(r.id); } });
    week = { start: key_(ws), targetKm: wp.targetKm, note: wp.note, doneKm: round05_(weekRuns.reduce((a, r) => a + r.km, 0) * 2) / 2,
      runsDone: weekRuns.length, runsPlanned: plan.length, plan };
    const tp = plan.find(p => p.date === dayKey);
    if (tp) {
      const ready = readiness_(d, log);
      today = Object.assign({}, tp);
      if (ready.mode === 'sick') Object.assign(today, { type: 'rest', desc: 'Rest – no running while you\'re unwell' });
      else if (ready.mode === 'light') Object.assign(today, { type: 'easy', km: Math.min(tp.km, 3), desc: 'Head cold: optional easy 20 minutes at most – or a walk' });
      else if (['tired', 'sore'].includes(ready.mode) && tp.type !== 'easy') { today.type = 'easy'; today.km = round05_(tp.km * 0.8); today.desc = runDesc_(today, zones) + ' (eased off – ' + ready.mode + ')'; }
    }
  }
  // stats
  const weekly = [];
  for (let i = 7; i >= 0; i--) { const a = addDays_(ws, -7 * i), b = key_(addDays_(a, 6));
    weekly.push({ week: key_(a), km: Math.round(runs.filter(r => r.date >= key_(a) && r.date <= b).reduce((x, r) => x + r.km, 0) * 10) / 10 }); }
  const timed = runs.filter(r => r.min && r.km);
  const best = (minKm) => { const c = timed.filter(r => r.km >= minKm).sort((a, b) => a.min / a.km - b.min / b.km)[0];
    return c ? { date: c.date, km: c.km, pace: fmtPace_(c.min * 60 / c.km) } : null; };
  const longest = runs.slice().sort((a, b) => b.km - a.km)[0];
  const month = key_(addDays_(d, -29));
  return {
    date: dayKey, configured: !!rc, goal: rc ? rc.goal : '', raceDate: rc ? rc.raceDate : '',
    week, today, zones,
    recent: runs.slice(-10).reverse().map(r => Object.assign({}, r, { pace: r.min ? fmtPace_(r.min * 60 / r.km) : '' })),
    weekly,
    stats: { monthKm: Math.round(runs.filter(r => r.date >= month && r.date <= dayKey).reduce((a, r) => a + r.km, 0) * 10) / 10,
      longest: longest ? { km: longest.km, date: longest.date } : null, best5k: best(5), best10k: best(10),
      predicted5k: p5 ? (() => { const t = Math.round(p5 * 5); return Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0'); })() : '' },
    kcalPerKm: runKcal_(1), eatBack: rc ? rc.eatBack : 0.5,
  };
}

/** Save a run. run = {type, km, min, hr, rpe, notes, source, clientId}. Also notes it as today's activity. */
function webLogRun(dateStr, run) {
  run = run || {};
  const km = Number(run.km), min = num_(run.min);
  if (!(km > 0 && km <= 100)) throw new Error('Enter the distance in km');
  if (min !== null && (min <= 0 || min > 900)) throw new Error('Duration looks wrong');
  const d = parseDay_(dateStr);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = runsSheet_();
    const id = String(run.clientId || Utilities.getUuid());
    if (run.clientId && readRuns_().some(r => r.id === id)) return webGetRuns(dateStr);      // already saved (retry)
    const type = RUN_TYPES.includes(String(run.type)) ? String(run.type) : 'easy';
    const pace = min ? Math.round(min / km * 100) / 100 : '';
    sh.appendRow([d, type, Math.round(km * 100) / 100, min === null ? '' : min, pace, num_(run.hr) === null ? '' : Number(run.hr),
      num_(run.rpe) === null ? '' : Number(run.rpe), String(run.notes || '').slice(0, 200), String(run.source || 'app'), id]);
    const found = findLogRow_(d), act = found ? String(found.values[activityCol_()] || '') : '';
    if (!act || /^run\b/i.test(act)) {
      const total = readRuns_().filter(r => r.date === key_(d)).reduce((a, r) => a + r.km, 0);
      upsertLog_(d, { [activityCol_()]: `Run ${Math.round(total * 10) / 10} km` }, []);
    }
  } finally { lock.releaseLock(); }
  return webGetRuns(dateStr);
}

function webDeleteRun(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RUNS_SHEET);
  if (!sh) throw new Error('No runs yet');
  const values = sh.getDataRange().getValues(), ci = values[0].indexOf('ID');
  for (let i = values.length - 1; i >= 1; i--) if (String(values[i][ci]) === String(id)) {
    const date = values[i][0]; sh.deleteRow(i + 1);
    const left = readRuns_().filter(r => r.date === key_(date)).reduce((a, r) => a + r.km, 0);
    const found = findLogRow_(date);
    if (found && /^run\b/i.test(String(found.values[activityCol_()] || ''))) {
      const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
      const headers = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
      setCell_(h, headers, found.rowIndex, activityCol_(), left ? `Run ${Math.round(left * 10) / 10} km` : '');
    }
    return webGetRuns(key_(date));
  }
  throw new Error('Run not found – refresh and try again');
}

/** Read a run from a Garmin / Strava / watch screenshot. Nothing is saved – the form is filled for you to check. */
function webRunFromPhoto(data, mime) {
  if (!data) throw new Error('No screenshot received');
  const prompt = `This is a screenshot from a running app or watch (e.g. Garmin Connect, Strava, Apple Fitness, Nike Run Club).
Read the run's numbers exactly as shown. Convert miles to km (× 1.609) and pace per mile to per km if needed.
Reply with JSON only: {"date":"YYYY-MM-DD or empty","km":0,"minutes":0,"avgHr":0,"type":"easy|long|tempo|intervals|race|other","notes":"anything useful, e.g. elevation or splits, max 80 chars"}
Use 0 for anything not shown. minutes = moving time in minutes with decimals (e.g. 31:30 → 31.5).`;
  const text = geminiSmart_({ contents: [{ role: 'user', parts: [{ inline_data: { mime_type: mime || 'image/jpeg', data } }, { text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.1 } });
  if (!text) throw new Error('Could not read the screenshot (' + (lastGeminiError_ || 'no reply') + ') – type the numbers instead.');
  let j; try { const m = text.match(/\{[\s\S]*\}/); j = JSON.parse(m ? m[0] : text); } catch (e) { throw new Error('The screenshot reply was unclear – type the numbers instead.'); }
  const km = Math.round((Number(j.km) || 0) * 100) / 100;
  if (!km) throw new Error("Couldn't find a distance in that screenshot – try the run's summary screen.");
  return { date: /^\d{4}-\d{2}-\d{2}$/.test(String(j.date || '')) ? j.date : '', km, min: Math.round((Number(j.minutes) || 0) * 10) / 10 || '',
    hr: Math.round(Number(j.avgHr) || 0) || '', type: RUN_TYPES.includes(j.type) ? j.type : 'easy', notes: String(j.notes || '').slice(0, 80) };
}

/** Extra calories on days you run (default: eat back half of what the run burned). */
function runExtraKcal_(dateKey, runs) {
  const rc = runCfg_(); const share = rc ? rc.eatBack : 0.5;
  const km = (runs || readRuns_()).filter(r => r.date === dateKey).reduce((a, r) => a + r.km, 0);
  return km ? Math.round(runKcal_(km) * share / 10) * 10 : 0;
}

/** One line for emails. */
function runEmailLine_(d) {
  const r = webGetRuns(key_(d));
  if (!r.today || r.today.done) return '';
  return `<p style="background:#EAF2FF;padding:8px"><b>🏃 Run today:</b> ${r.today.desc}</p>`;
}
function runWeekSummary_(d) {
  const r = webGetRuns(key_(d));
  if (!r.configured || !r.week) return '';
  const next = runWeekPlan_(addDays_(mondayOf_(d), 7), readRuns_(), runCfg_(), getSettings_());
  return `Running: ${r.week.doneKm} of ${r.week.targetKm} km, ${r.week.runsDone} of ${r.week.runsPlanned} runs. Next week: ${next.targetKm} km` +
    (next.note ? ` (${next.note})` : '') + ` – ${next.plan.map(p => p.day + ' ' + p.type + ' ' + p.km + ' km').join(', ')}.`;
}

// ---------- 25-minute home workouts (for busy days) ----------
const HOME_LOWER = 'Home – Lower (25 min)', HOME_UPPER = 'Home – Upper (25 min)';
const LOWER_MUSCLES = ['Quads', 'Hamstrings', 'Glutes', 'Calves'];

/** Built from Config.gs home.equipment: 'bodyweight' (default), 'dumbbells', 'band'. Supersets keep it inside 25 minutes. */
function homeRows_() {
  const eq = ((cfg_().home || {}).equipment || ['bodyweight']).map(x => String(x).toLowerCase());
  const db = eq.includes('dumbbells'), band = eq.includes('band');
  const r = (session, order, ex, sets, reps, max, rest, inc, muscle) =>
    ({ 'Session': session, 'Order': order, 'Exercise': ex, 'Sets': sets, 'Reps': reps, 'Rep max': max, 'Rest': rest, 'Increment kg': inc, 'Muscle': muscle });
  return [
    r(HOME_LOWER, 1, db ? 'Bulgarian split squat (DB)' : 'Bulgarian split squat', 3, '8–12/leg', 12, 'superset', db ? 2 : 0, 'Quads'),
    r(HOME_LOWER, 2, 'Single-leg glute bridge', 3, '10–15/leg', 15, '60 s', 0, 'Glutes'),
    r(HOME_LOWER, 3, db ? 'Goblet squat' : 'Bodyweight squat (3-s lowering)', 3, '12–20', 20, 'superset', db ? 2 : 0, 'Quads'),
    r(HOME_LOWER, 4, 'Single-leg calf raise', 3, '12–20/leg', 20, '45 s', 0, 'Calves'),
    r(HOME_LOWER, 5, 'Dead bug', 2, '8/side', 8, '45 s', 0, 'Core'),
    r(HOME_UPPER, 1, 'Push-up', 4, '8–15', 15, 'superset', 0, 'Chest'),
    r(HOME_UPPER, 2, db ? 'One-arm DB row' : band ? 'Band row' : 'Backpack row', 4, '10–15', 15, '60 s', db ? 2 : 0, 'Back'),
    r(HOME_UPPER, 3, db ? 'Seated DB shoulder press' : 'Pike push-up', 3, db ? '8–12' : '6–10', db ? 12 : 10, 'superset', db ? 2 : 0, 'Shoulders'),
    r(HOME_UPPER, 4, band ? 'Band pull-apart' : 'Superman', 3, '12–15', 15, '45 s', 0, band ? 'Rear delts' : 'Back'),
    r(HOME_UPPER, 5, 'Side plank', 2, '30 s/side', '', '30 s', 0, 'Core'),
  ];
}

/** Which home version fits a gym session: the one that trains the same half of the body. */
function homeFor_(sessionName, plan) {
  const rows = plan.filter(p => p['Session'] === sessionName);
  if (!rows.length) return HOME_UPPER;
  const lower = rows.filter(p => LOWER_MUSCLES.includes(String(p['Muscle']))).reduce((a, p) => a + (Number(p['Sets']) || 0), 0);
  const all = rows.reduce((a, p) => a + (Number(p['Sets']) || 0), 0);
  return lower * 2 >= all ? HOME_LOWER : HOME_UPPER;
}

/** Bodyweight moves progress on reps: beat last time, then make it harder once every set hits the top of the range. */
function homeTargets_(list, rows, lifts, d) {
  const before = lifts.filter(l => l['Date'] instanceof Date && key_(l['Date']) < key_(d));
  return list.map(e => {
    const row = rows.find(r => r['Exercise'] === e.exercise) || {};
    if (Number(row['Increment kg'])) return e;                          // dumbbell moves use normal kg progression
    if (/deload|light day/.test(e.target)) return Object.assign({}, e, { target: 'easy – stop 4 reps before failure' });
    const hist = before.filter(l => String(l['Exercise']).toLowerCase() === e.exercise.toLowerCase());
    if (!hist.length) return Object.assign({}, e, { target: 'bodyweight – stop 2 reps before failure', last: 'first time – leave kg empty, just log reps' });
    const lastKey = hist.map(l => key_(l['Date'])).sort().pop();
    const reps = hist.filter(l => key_(l['Date']) === lastKey).map(l => Number(l['Reps']) || 0);
    const max = Number(row['Rep max']) || 0;
    const target = max && reps.length >= Number(row['Sets']) && reps.every(x => x >= max)
      ? 'harder ↑ – 3-s lowering + 1-s pause, same reps' : `beat last time (${reps.join(', ')})`;
    return Object.assign({}, e, { target });
  });
}

// ---------- Skipped sessions and catch-up ----------
const SKIP_REASONS = ['busy', 'tired', 'travel', 'sick', 'other'];

/** Record that a planned session was skipped (reason '' = undo). "sick" also uses the sick-day rules. */
function webSkipDay(dateStr, reason) {
  ensureLogColumns_();
  const r = SKIP_REASONS.includes(String(reason)) ? String(reason) : '';
  const d = parseDay_(dateStr);
  if (r) {
    upsertLog_(d, { 'Skipped': r }, []);
    if (r === 'sick') upsertLog_(d, { 'Feeling': 'sick' }, []);
  } else {
    const found = findLogRow_(d);
    if (found) {   // undo: clear the cell (and the sick feeling it set)
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const ci = headers.indexOf('Skipped'), fi = headers.indexOf('Feeling');
      if (fi >= 0 && String(found.values['Skipped']) === 'sick' && String(found.values['Feeling']) === 'sick') sh.getRange(found.rowIndex, fi + 1).setValue('');
      if (ci >= 0) sh.getRange(found.rowIndex, ci + 1).setValue('');
    }
  }
  return webGetDay(dateStr);
}

/**
 * Offer a missed session from earlier this week (Mon–Sun) on a free morning:
 * no session planned today, not a rest morning (sick, or the morning after a sport with restNextMorning,
 * logged or scheduled), and nothing trained yet today. Never two sessions in a day.
 * With several missed, pick the one whose muscles have had the fewest sets this week (keeps upper/lower balanced).
 */
function catchUp_(d, logRows, lifts, plan, s, ready) {
  const dow = Number(fmt_(d, 'u')) % 7, dayKey = key_(d);
  if (sessionFor_(dow) || (ready && ready.rest)) return null;
  const prevSport = sportOn_((dow + 6) % 7);
  if (prevSport && prevSport.restNextMorning) return null;
  if (lifts.some(l => key_(l['Date']) === dayKey)) return null;
  const weekStart = addDays_(d, -((Number(fmt_(d, 'u')) + 6) % 7));
  const startKey = s['Start date'] ? key_(s['Start date'] instanceof Date ? s['Start date'] : new Date(s['Start date'])) : '';
  const weekLifts = lifts.filter(l => key_(l['Date']) >= key_(weekStart) && key_(l['Date']) <= dayKey);
  const doneSessions = new Set(weekLifts.map(l => String(l['Session'])));
  const trainedDays = new Set(weekLifts.map(l => key_(l['Date'])));       // a home workout on a gym day counts for that day
  const logBy = {};
  logRows.forEach(r => { if (r['Date'] instanceof Date) logBy[key_(r['Date'])] = r; });
  const missed = [];
  for (let x = new Date(weekStart); key_(x) < dayKey; x = addDays_(x, 1)) {
    const k = key_(x), name = sessionFor_(Number(fmt_(x, 'u')) % 7);
    if (!name || (startKey && k < startKey) || doneSessions.has(name) || trainedDays.has(k)) continue;
    const r = logBy[k] || {};
    missed.push({ session: name, date: k, day: DAY_CODES[Number(fmt_(x, 'u')) % 7], reason: String(r['Skipped'] || (/^sick$/i.test(String(r['Feeling'] || '')) ? 'sick' : '')) });
  }
  if (!missed.length) return null;
  // sets done this week per muscle
  const muscleOf = {};
  plan.forEach(p => { muscleOf[String(p['Exercise']).toLowerCase()] = String(p['Muscle']); });
  const doneBy = {};
  weekLifts.forEach(l => { const m = muscleOf[String(l['Exercise']).toLowerCase()]; if (m) doneBy[m] = (doneBy[m] || 0) + 1; });
  const score = name => plan.filter(p => p['Session'] === name).reduce((a, p) => a + (doneBy[String(p['Muscle'])] || 0), 0);
  const pick = missed.slice().sort((a, b) => score(a.session) - score(b.session) || (a.date < b.date ? 1 : -1))[0];
  return Object.assign({}, pick, { others: missed.filter(m => m !== pick).map(m => m.session + ' (' + m.day + ')') });
}

function liftOut_(l) {
  return { id: String(l['ID'] || ''), session: String(l['Session']), exercise: String(l['Exercise']),
    set: Number(l['Set']) || 1, kg: l['Weight kg'] === '' ? '' : Number(l['Weight kg']), reps: Number(l['Reps']) || 0 };
}

// ----- Gym: one row per set, saved the moment it's ticked -----

/** Make sure the Lifts tab has an ID column and every row has an ID (older rows get one added). */
function ensureLiftIds_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIFTS);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let col = headers.indexOf('ID') + 1;
  if (!col) { col = LIFT_HEADERS.length; sh.getRange(1, col).setValue('ID').setFontWeight('bold'); }
  const n = sh.getLastRow() - 1;
  if (n < 1) return col;
  const ids = sh.getRange(2, col, n, 1).getValues();
  let changed = false;
  ids.forEach(r => { if (!r[0]) { r[0] = Utilities.getUuid(); changed = true; } });
  if (changed) sh.getRange(2, col, n, 1).setValues(ids);
  return col;
}

/** Save one set. clientId makes retries safe: sending the same set twice never creates two rows. */
function webLogSet(dateStr, session, exercise, set, kg, reps, clientId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const idCol = ensureLiftIds_();
    if (clientId && findLift_(clientId, idCol)) return { id: clientId, duplicate: true };
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIFTS);
    const id = clientId || Utilities.getUuid();
    const row = [parseDay_(dateStr), session, exercise, Number(set) || 1,
      kg === '' || kg === null || kg === undefined ? '' : Number(kg), Number(reps) || 0, '', id];
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return { id };
  } finally { lock.releaseLock(); }
}

function webUpdateSet(id, kg, reps) {
  const idCol = ensureLiftIds_();
  const f = findLift_(id, idCol);
  if (!f) throw new Error('Set not found – refresh and try again');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIFTS);
  sh.getRange(f.row, 5, 1, 2).setValues([[kg === '' ? '' : Number(kg), Number(reps) || 0]]);
  return { ok: true };
}

function webDeleteSet(id) {
  const idCol = ensureLiftIds_();
  const f = findLift_(id, idCol);
  if (f) SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIFTS).deleteRow(f.row);
  return { ok: true };
}

function findLift_(id, idCol) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIFTS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const ids = sh.getRange(2, idCol, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { row: i + 2 };
  return null;
}

// ----- Ask Coach: chat with Gemini about your own data -----

function coachSystem_() {
  const name = profile_().name;
  return `You are ${name}'s personal strength and nutrition coach, built into their tracking app.
${profileText_()}
Foods they eat: ${profile_().foods || 'see recent foods'}.

How to answer:
- Use the DATA snapshot from their logs. Quote their real numbers (weights, protein, calories, sleep) when relevant.
- Be direct, practical and brief: under 150 words unless asked for detail. Short bullets are fine.
- Suggest foods they actually eat (see foods, recent foods and meals) and portions in plain terms.
- The app sets the calorie target, protein target and lift weights. Explain them; don't override them.
  If they want a change, say what the app will do automatically or that they can edit the Settings tab.
- No medical diagnosis. For pain that persists, chest symptoms, dizziness, pregnancy questions or illness beyond a
  mild cold, tell them to see a doctor or physio. No crash diets. No supplements beyond whey, creatine and vitamin D.
- If the data needed is missing, say so and ask them to log it.
- Plain text only. You may use **bold** and lines starting with "- " for bullets.`;
}

/** A compact snapshot of the sheet for the coach (kept small so replies stay fast). */
function coachContext_() {
  ensureLogColumns_();
  const s = getSettings_();
  const today = parseDay_(key_(new Date()));
  const dow = Number(fmt_(today, 'u')) % 7;
  const log = readTable_(SHEETS.LOG);
  const lifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date);
  const plan = readTable_(SHEETS.PLAN);
  const ready = readiness_(today, log);
  const t = dayTargets_(s, today, dow, ready);
  const weekNum = weekNumber_(s['Start date'], today);

  const days = rowsBetween_(log, addDays_(today, -13), today).map(r => {
    const o = { date: key_(r['Date']) };
    [['Weight kg', 'weight'], ['kcal', 'kcal'], ['Protein g', 'protein'], ['Fat g', 'fat'], ['Sugar g', 'sugar'], ['Steps', 'steps'], ['Sleep h', 'sleep'],
      ['Waist cm', 'waist'], ['Session', 'gym'], ['Activity', 'activity'], ['Badminton/Run', 'activity'], ['Feeling', 'feeling'],
      ['Pump', 'pump'], ['Soreness', 'soreness'], ['Joint pain', 'jointPain']].forEach(([col, k]) => {
      if (r[col] !== '' && r[col] !== undefined && r[col] !== null) o[k] = r[col];
    });
    return o;
  });

  const byEx = {};
  lifts.forEach(l => {
    const ex = String(l['Exercise']), k = key_(l['Date']);
    byEx[ex] = byEx[ex] || {};
    (byEx[ex][k] = byEx[ex][k] || []).push(`${l['Weight kg'] === '' ? 'bw' : l['Weight kg']}×${l['Reps']}`);
  });
  const recentLifts = Object.keys(byEx).map(ex => {
    const ds = Object.keys(byEx[ex]).sort().slice(-2).reverse();
    return `${ex}: ` + ds.map(d => `${d} ${byEx[ex][d].join(', ')}`).join(' | ');
  });

  const sessions = {};
  sessionNames_().forEach(name => {
    sessions[name] = sessionTargets_(name, today, lifts, plan, s, { ready, log })
      .map(e => `${e.exercise} ${e.sets}×${e.reps} → ${e.target}`);
  });

  const targetsRows = readTable_(SHEETS.TARGETS);
  const lastReview = targetsRows.length ? targetsRows[targetsRows.length - 1] : null;
  return {
    today: key_(today), weekday: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][dow], weekNumber: weekNum,
    deloadWeek: isDeloadWeek_(weekNum), todaysSession: sessionFor_(dow) || 'rest day',
    readinessToday: ready.title ? `${ready.title}: ${ready.message}` : 'normal',
    targetsToday: { kcal: t.kcal, protein: t.protein, proteinMinimum: Number(s['Protein minimum g']) || 120, steps: t.steps,
      fatGuide: fatSugarGuides_(t.kcal).fat, sugarGuide: fatSugarGuides_(t.kcal).sugar },
    foodsToday: foodsFor_(today).map(f => `[${f.meal}] ${f.item} (${f.kcal} kcal, ${f.protein} g protein` +
      (f.fat !== null ? `, ${f.fat} g fat` : '') + (f.sugar !== null ? `, ${f.sugar} g sugar` : '') + ')'),
    recentFoods: recentFoods_(today).map(f => `${f.item} (${f.kcal} kcal, ${f.protein} g)`),
    last14Days: days,
    recentLifts,
    nextSessionTargets: sessions,
    lastWeeklyReview: lastReview ? { week: lastReview['Week #'], avgWeight: lastReview['7-day avg weight'],
      change: lastReview['Weekly change kg'], calorieTarget: lastReview['Calorie target'], maintenance: lastReview['Est. maintenance'] } : null,
    patterns: insights_(log, lifts, plan, s, today),
    waistGoalCm: s['Waist goal cm'],
  };
}

/** Ask the coach. history = [{role:'user'|'coach', text}] from this conversation (most recent last). */
function webAskCoach(question, history) {
  const q = String(question || '').trim();
  if (!q) throw new Error('Type a question first');
  const ctx = coachContext_();
  const turns = (history || []).slice(-8).map(h => ({ role: h.role === 'coach' ? 'model' : 'user', parts: [{ text: String(h.text).slice(0, 2000) }] }));
  const body = {
    systemInstruction: { parts: [{ text: coachSystem_() + '\n\nDATA (from their logs, as of today):\n' + JSON.stringify(ctx) }] },
    contents: turns.concat([{ role: 'user', parts: [{ text: q }] }]),
  };
  const answer = geminiGenerate_(body);
  if (!answer) throw new Error('The coach is unavailable right now (' + (lastGeminiError_ || 'no reply') + '). Try again in a minute.');
  return { answer };
}

// ----- Photo food logging -----

/** Estimate the foods in a meal photo. data = base64 JPEG (no prefix). Nothing is stored – only the estimates come back. */
/** Re-estimate one food after you correct its name (e.g. "soya chunk curry", not "meat curry"). Keeps your grams. */
function webRecalcItem(item, grams) {
  const name = String(item || '').trim().slice(0, 100);
  if (!name) throw new Error('Type what the food is first');
  const g = Math.max(0, Math.round(Number(grams) || 0));
  if (!g) {                                            // everything in the name is in the food table → instant, no AI
    const local = localEstimate_(name);
    if (local) return Object.assign({ item: name, grams: 0, source: 'table' }, local);
  }
  const out = recalcFood_(name, g, 20000);
  if (out) return out;
  const local = localEstimate_(name);
  if (local) return Object.assign({ item: name, grams: g, source: 'table' }, local);
  throw new Error(/slow right now/.test(lastGeminiError_) ? 'Gemini is busy right now – try again in a minute, or type the numbers.'
    : 'Could not recalculate (' + (lastGeminiError_ || 'no reply') + ') – try again in a minute, or type the numbers.');
}

/** The AI estimate behind Recheck: smart model, reference table, time-limited. Returns numbers or null. */
function recalcFood_(name, g, budgetMs) {
  const prompt = `Estimate the nutrition for this food as eaten: "${name}"${g ? `, portion ${g} g` : ' (assume one typical home portion and say how many grams that is)'}.
Assume home cooking with normal oil unless the name says otherwise. The name is the person's own correction – trust it exactly.
Use these reference values when the food (or an ingredient) is listed, scaled to the portion:
${foodReferenceText_()}
Reply with JSON only: {"item":"${name.replace(/"/g, "'")}","grams":0,"kcal":0,"protein":0,"fat":0,"sugar":0}`;
  const text = geminiSmart_({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0.1 } }, budgetMs);
  let j = null;
  if (text) { try { const m = text.match(/\{[\s\S]*\}/); j = JSON.parse(m ? m[0] : text); } catch (e) { j = null; } }
  if (!j) return null;
  const out = { item: name, grams: g || Math.max(0, Math.round(Number(j.grams) || 0)),
    kcal: Math.max(0, Math.round(Number(j.kcal) || 0)), protein: Math.max(0, Math.round(Number(j.protein) || 0)),
    fat: Math.max(0, Math.round(Number(j.fat) || 0)), sugar: Math.max(0, Math.round(Number(j.sugar) || 0)), source: 'ai' };
  const minKcal = out.protein * 4 + out.fat * 9;
  if (out.kcal && minKcal > out.kcal * 1.25) out.kcal = Math.round(minKcal * 1.1);
  return out;
}

/** Reference values for the foods this app knows, so photo, chat and typed estimates agree. */
function foodReferenceText_() {
  const nameOf = e => e[0].source.split('|')[0].replace(/\\s[*+]?/g, ' ').replace(/\\b/g, '').replace(/\[[^\]]*\]|[?()^$+*.\\]/g, '').replace(/\s+/g, ' ').trim();
  const fix = { 'idl': 'idly', 'protein yog': 'protein yogurt', 'greek yog': 'greek yogurt', 'blueberr': 'blueberries', 'blue berr': 'blueberries', 'straw berr': 'strawberries', 'cherr': 'cherries', 'dates': 'date', 'omelete': 'omelette', 'egg omelete': 'egg in an omelette', 'eggs': 'egg' };
  return FOOD_DB.map(e => `  ${fix[nameOf(e)] || nameOf(e)} (${e[6]}): ${e[1]} kcal, ${e[2]} g protein, ${e[3]} g fat, ${e[4]} g sugar` + (e[5] ? ` – per ${e[5]} g` : '') +
    (e[7] ? ` – so ONE piece is about ${rnd_(e[1] / e[7], 1)} kcal` : '')).join('\n');
}

/** Estimate the foods in a meal photo. data = base64 JPEG (no prefix). Nothing is stored – only the estimates come back. */
function webPhotoFood(data, mime, note) {
  if (!data) throw new Error('No photo received');
  const prompt = `You are a careful dietitian estimating a meal from a photo. The person's usual diet: ${profile_().diet || 'not stated'}.
Foods they often eat: ${String(profile_().foods || '').slice(0, 400) || 'not stated'}. It could be any cuisine.
Work in this order for EACH item:
1. Name it precisely (e.g. "chicken thigh curry", not "curry"; spot visible oil, ghee, cream, cheese, nuts, sauces).
2. Estimate the portion in grams (or ml) from visual cues: a dinner plate is about 26 cm, a katori/small bowl about 150 ml,
   a chapathi about 35 g, a dosa about 80 g, a tablespoon about 15 ml, a cooked chicken thigh about 90 g of meat.
3. Work out kcal, protein, fat and total sugars for THAT portion, assuming home cooking with normal oil unless it looks oily or fried.
Use these reference values for foods they eat often, scaled to the portion you see:
${foodReferenceText_()}
${note ? 'The person adds: "' + String(note).slice(0, 200) + '". This overrides what you see – use it to correct items or amounts.' : ''}
Reply with JSON only:
{"items":[{"item":"short name","grams":0,"kcal":0,"protein":0,"fat":0,"sugar":0}],"confidence":"low|medium|high","comment":"one short sentence on anything uncertain, e.g. hidden oil"}`;
  const body = {
    contents: [{ role: 'user', parts: [{ inline_data: { mime_type: mime || 'image/jpeg', data: data } }, { text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.2 },
  };
  const text = geminiSmart_(body);
  if (!text) throw new Error(/allowance/.test(lastGeminiError_) ? lastGeminiError_ : 'Could not analyse the photo (' + (lastGeminiError_ || 'no reply') + '). Try again, or add the food by name.');
  const m = text.match(/\{[\s\S]*\}/);
  let j;
  try { j = JSON.parse(m ? m[0] : text); } catch (e) { throw new Error('The photo reply was unclear – try again or add the food by name.'); }
  const items = (j.items || []).filter(x => x && String(x.item || '').trim()).map(x => {
    const out = {
      item: String(x.item).trim().slice(0, 80),
      grams: Math.max(0, Math.round(Number(x.grams) || 0)),
      kcal: Math.max(0, Math.round(Number(x.kcal) || 0)),
      protein: Math.max(0, Math.round(Number(x.protein) || 0)),
      fat: Math.max(0, Math.round(Number(x.fat) || 0)),
      sugar: Math.max(0, Math.round(Number(x.sugar) || 0)),
    };
    // sanity check: protein/fat/carb energy can't be far above the kcal figure
    const minKcal = out.protein * 4 + out.fat * 9;
    if (out.kcal && minKcal > out.kcal * 1.25) out.kcal = Math.round(minKcal * 1.1);
    return out;
  }).filter(x => x.item);
  if (!items.length) throw new Error("Couldn't spot any food in that photo – try a clearer shot from above.");
  return { items, confidence: String(j.confidence || 'medium'), comment: String(j.comment || ''), model: PropertiesService.getScriptProperties().getProperty('GEMINI_SMART_MODEL') || '' };
}

/** Add several foods in one go (used after reviewing a photo). */
function webAddFoods(dateStr, items, source, meal) {
  const d = parseDay_(dateStr);
  const list = (items || []).filter(x => x && String(x.item || '').trim());
  if (!list.length) throw new Error('Nothing selected');
  const time = fmt_(new Date(), 'HH:mm');
  const rows = list.map(x => [d, time, String(x.item).trim(), Number(x.kcal) || 0, Number(x.protein) || 0, source || 'photo', Utilities.getUuid(),
    num_(x.fat) === null ? '' : num_(x.fat), num_(x.sugar) === null ? '' : num_(x.sugar), mealSlot_(x.meal || meal, time)]);
  const sh = foodSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, FOOD_HEADERS.length).setValues(rows);
  syncFoodTotals_(d);
  const out = foodDay_(d);
  out.added = rows.length;
  return out;
}

// ----- Progress dashboard -----

/** Estimated 1-rep max (Epley). Only meaningful for 1–12 reps. */
function e1rm_(kg, reps) { return kg > 0 && reps > 0 && reps <= 12 ? kg * (1 + reps / 30) : 0; }

function webGetDashboard() { return withMemo_(webGetDashboardData_); }
function webGetDashboardData_() {
  const s = getSettings_();
  const today = parseDay_(key_(new Date()));
  const todayKey = key_(today);
  const weekNum = weekNumber_(s['Start date'], today);
  const t = currentTargets_(s);

  // Weight + 7-day average, last 12 weeks
  const log = readTable_(SHEETS.LOG).filter(r => r['Date'] instanceof Date);
  const recent = rowsBetween_(log, addDays_(today, -83), today);
  const wByDay = {};
  recent.forEach(r => { const w = num_(r['Weight kg']); if (w !== null) wByDay[key_(r['Date'])] = w; });
  const dayKeys = Object.keys(wByDay).sort();
  const weight = dayKeys.map(k => {
    const lo = key_(addDays_(parseDay_(k), -7));
    const vals = dayKeys.filter(x => x > lo && x <= k).map(x => wByDay[x]);
    return { date: k, w: wByDay[k], avg: rnd_(vals.reduce((a, b) => a + b, 0) / vals.length, 2) };
  });
  const waist = recent.filter(r => num_(r['Waist cm']) !== null).map(r => ({ date: key_(r['Date']), v: Number(r['Waist cm']) }));

  // This week (last 7 days)
  const week = rowsBetween_(log, addDays_(today, -6), today);
  const lifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date);
  const weekFrom = key_(addDays_(today, -6));
  const gymDays = new Set(lifts.filter(l => key_(l['Date']) >= weekFrom).map(l => key_(l['Date'])));
  const thisWeek = {
    avgProtein: rnd_(avg_(week.map(r => r['Protein g'])), 0),
    avgKcal: rnd_(avg_(week.map(r => r['kcal'])), 0),
    foodDays: week.filter(r => num_(r['kcal']) !== null).length,
    avgSteps: rnd_(avg_(week.map(r => r['Steps'])), 0),
    avgSleep: rnd_(avg_(week.map(r => r['Sleep h'])), 1),
    gymSessions: gymDays.size,
    activities: week.filter(r => activityOf_(r)).length,
    avgFat: rnd_(avg_(week.map(r => r['Fat g'])), 0),
    avgSugar: rnd_(avg_(week.map(r => r['Sugar g'])), 0),
  };

  // Personal records + strength trend per exercise
  const plan = readTable_(SHEETS.PLAN);
  const names = [];
  mainLifts_().concat(plan.map(p => String(p['Exercise'])), lifts.map(l => String(l['Exercise'])))
    .forEach(n => { if (n && !names.some(x => x.toLowerCase() === n.toLowerCase())) names.push(n); });
  const weekAgo = key_(addDays_(today, -7));
  const prs = names.map(name => {
    const rows = lifts.filter(l => String(l['Exercise']).toLowerCase() === name.toLowerCase());
    if (!rows.length) return null;
    const byDay = {};
    let bestKg = 0, bestE = 0, bestDate = '', bestSet = '';
    rows.forEach(l => {
      const kg = Number(l['Weight kg']) || 0, reps = Number(l['Reps']) || 0, k = key_(l['Date']);
      const e = e1rm_(kg, reps);
      if (kg > bestKg) bestKg = kg;
      if (e > bestE) { bestE = e; bestDate = k; bestSet = `${kg} kg × ${reps}`; }
      byDay[k] = Math.max(byDay[k] || 0, e);
    });
    const series = Object.keys(byDay).sort().filter(k => byDay[k] > 0).map(k => ({ date: k, v: rnd_(byDay[k], 1) }));
    return { exercise: name, bestKg, bestE1rm: rnd_(bestE, 1), bestDate, bestSet,
      isNew: bestDate && bestDate > weekAgo, sessions: Object.keys(byDay).length, series };
  }).filter(Boolean);

  // Logging streak: consecutive days with food or weight logged (today counts if logged)
  const logged = new Set();
  log.forEach(r => { if (num_(r['kcal']) !== null || num_(r['Weight kg']) !== null) logged.add(key_(r['Date'])); });
  let streak = 0;
  for (let i = logged.has(todayKey) ? 0 : 1; i < 400; i++) {
    if (logged.has(key_(addDays_(today, -i)))) streak++; else break;
  }

  const reviews = readTable_(SHEETS.TARGETS).slice(-6).map(r => ({
    date: r['Week ending'] instanceof Date ? key_(r['Week ending']) : String(r['Week ending']),
    week: r['Week #'], avgWeight: r['7-day avg weight'], change: r['Weekly change kg'],
    target: r['Calorie target'], maintenance: r['Est. maintenance'],
  }));

  return {
    today: todayKey, weekNum, deload: isDeloadWeek_(weekNum),
    nextDeloadWeek: (() => { const n = Number(trainingCfg_().deloadEvery) || 0; return !n ? null : isDeloadWeek_(weekNum) ? weekNum : Math.ceil(weekNum / n) * n; })(),
    sessionsPerWeek: sessionsPerWeek_(),
    targets: { kcal: t.calories, protein: t.protein, steps: t.steps, waistGoal: Number(s['Waist goal cm']) || null,
      fatGuide: fatSugarGuides_(t.calories).fat, sugarGuide: fatSugarGuides_(t.calories).sugar },
    weight, waist, thisWeek, prs, streak, reviews,
    ...goalsAndStreaks_(s, log, lifts, prs, today),
  };
}

// ---------- Share card: a summary of one day's workout ----------
/** Everything the phone needs to draw the share picture for a day: exercises, best sets, PRs, volume, streak. */
function webShareSummary(dateStr) {
  return withMemo_(() => {
    const d = parseDay_(dateStr), k = key_(d), s = getSettings_();
    const lifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date);
    const todays = lifts.filter(l => key_(l['Date']) === k && Number(l['Reps']) > 0);
    const before = lifts.filter(l => key_(l['Date']) < k);
    const order = [], by = {};
    todays.forEach(l => {
      const name = String(l['Exercise']), key = name.toLowerCase();
      if (!by[key]) { by[key] = { name, sets: 0, volume: 0, best: null }; order.push(key); }
      const kg = Number(l['Weight kg']) || 0, reps = Number(l['Reps']) || 0, e = e1rm_(kg, reps);
      const x = by[key]; x.sets++; x.volume += kg * reps;
      if (!x.best || e > x.best.e || (!e && !x.best.e && reps > x.best.reps)) x.best = { kg, reps, e };
    });
    const exercises = order.map(key => {
      const x = by[key], prev = before.filter(l => String(l['Exercise']).toLowerCase() === key);
      const prevBestE = Math.max(0, ...prev.map(l => e1rm_(Number(l['Weight kg']) || 0, Number(l['Reps']) || 0)));
      const prevBestReps = Math.max(0, ...prev.filter(l => !(Number(l['Weight kg']) > 0)).map(l => Number(l['Reps']) || 0));
      const pr = prev.length > 0 && (x.best.e ? x.best.e > prevBestE + 0.01 : x.best.reps > prevBestReps);
      return { name: x.name, sets: x.sets, best: x.best.kg ? `${rnd_(x.best.kg, 1)} kg × ${x.best.reps}` : `${x.best.reps} reps`,
        volume: Math.round(x.volume), pr };
    });
    const logRow = readTable_(SHEETS.LOG).find(r => r['Date'] instanceof Date && key_(r['Date']) === k) || {};
    const sessions = []; todays.forEach(l => { const n = String(l['Session'] || ''); if (n && !sessions.includes(n)) sessions.push(n); });
    let gymStreak = null;
    try {
      const g = goalsAndStreaks_(s, readTable_(SHEETS.LOG).filter(r => r['Date'] instanceof Date), lifts, [], parseDay_(key_(new Date())));
      const st = (g.streaks || []).find(x => x.id === 'gym'); if (st) gymStreak = { current: st.current, unit: st.unit };
    } catch (e) { console.error(e); }
    let run = null;
    try { const rc = runCfg_(); if (rc) { const r = readRuns_().find(x => x.date === k); if (r) run = { km: r.km, minutes: r.min, pace: r.min ? fmtPace_(r.min * 60 / r.km) : '' }; } } catch (e) {}
    return {
      appName: appName_(), appTagline: appTagline_(), date: k, weekday: fmt_(d, 'EEE'), dateLabel: fmt_(d, 'EEE d MMM'), name: profile_().name === 'you' ? '' : profile_().name,
      session: String(logRow['Session'] || sessions.join(' + ') || ''), weekNum: weekNumber_(s['Start date'], d),
      exercises, totalSets: exercises.reduce((a, x) => a + x.sets, 0), volumeKg: exercises.reduce((a, x) => a + x.volume, 0),
      prs: exercises.filter(x => x.pr).length, gymStreak, run,
    };
  });
}

// ---------- Goals and streaks (Progress tab) ----------
const STREAK_MILESTONES = [3, 7, 14, 21, 30, 50, 75, 100, 150, 200, 365];

/** For any date: the calorie target that applied then (incl. sport days), the step target and fat/sugar guides. */
function dayJudge_(s, reviews) {
  const allRuns = (() => { try { return readRuns_(); } catch (e) { return []; } })();
  const start = s['Start date'] instanceof Date ? s['Start date'] : new Date(s['Start date']);
  const revs = reviews.filter(r => r['Week ending'] instanceof Date && num_(r['Calorie target']) !== null)
    .map(r => ({ k: key_(r['Week ending']), kcal: Number(r['Calorie target']) }))
    .sort((a, b) => a.k < b.k ? -1 : a.k > b.k ? 1 : 0);
  const startKcal = Number(s['Starting calorie target']) || 2000;
  return k => {
    let kcal = startKcal;
    revs.forEach(r => { if (r.k < k) kcal = r.kcal; });
    const d = parseDay_(k), dow = Number(fmt_(d, 'u')) % 7;
    const sport = sportOn_(dow);
    if (sport && sport.extraKcal) kcal += Number(sport.extraKcal) || 0;
    kcal += runExtraKcal_(k, allRuns);
    return { kcal, steps: stepTarget_(weekNumber_(start, d)), guides: fatSugarGuides_(kcal), dow };
  };
}

/** Daily habits. Each test returns 'hit', 'miss' or 'skip' (a sick day, or no data that could count). */
function streakDefs_(s) {
  const pmin = Number(s['Protein minimum g']) || Number(s['Protein target g']) || 100;
  const kcalOf = r => r ? num_(r['kcal']) : null;
  return [
    { id: 'log', icon: '📝', label: 'Logging', rule: 'Food or weight logged',
      test: r => r && (num_(r['kcal']) !== null || num_(r['Weight kg']) !== null) ? 'hit' : 'miss' },
    { id: 'protein', icon: '🥚', label: 'Protein', rule: `${pmin} g+ protein`,
      test: (r, j, sick) => sick ? 'skip' : r && num_(r['Protein g']) !== null && Number(r['Protein g']) >= pmin ? 'hit' : 'miss' },
    { id: 'calories', icon: '🎯', label: 'Calories', rule: "Within 10% of the day's target",
      test: (r, j, sick) => { if (sick) return 'skip'; const k = kcalOf(r); return k !== null && k >= j.kcal * 0.9 && k <= j.kcal * 1.1 ? 'hit' : 'miss'; } },
    { id: 'sugar', icon: '🍬', label: 'Sugar', rule: 'Under the daily sugar guide',
      test: (r, j, sick) => { if (sick) return 'skip'; if (kcalOf(r) === null) return 'miss';
        const g = num_(r['Sugar g']); return g === null ? 'skip' : g <= j.guides.sugar ? 'hit' : 'miss'; } },
    { id: 'steps', icon: '👟', label: 'Steps', rule: 'Hit your step target',
      test: (r, j, sick) => sick ? 'skip' : r && num_(r['Steps']) !== null && Number(r['Steps']) >= j.steps ? 'hit' : 'miss' },
    { id: 'sleep', icon: '😴', label: 'Sleep', rule: '7 hours or more',
      test: r => r && num_(r['Sleep h']) !== null && Number(r['Sleep h']) >= 7 ? 'hit' : 'miss' },
    { id: 'weigh', icon: '⚖️', label: 'Weigh-in', rule: 'Morning weight logged',
      test: r => r && num_(r['Weight kg']) !== null ? 'hit' : 'miss' },
  ];
}

/** Current and best run. Skips neither break nor add; today counts only once it's a hit (until then it's "pending"). */
function runs_(statuses) {
  let best = 0, run = 0;
  statuses.forEach(st => { if (st === 'hit') { run++; best = Math.max(best, run); } else if (st === 'miss') run = 0; });
  let cur = 0;
  for (let i = statuses.length - 1; i >= 0; i--) {
    const st = statuses[i];
    if (i === statuses.length - 1 && st !== 'hit') continue;
    if (st === 'hit') cur++; else if (st === 'miss') break;
  }
  return { current: cur, best: Math.max(best, cur) };
}

function milestone_(n) {
  const next = STREAK_MILESTONES.find(m => m > n) || null;
  const reached = STREAK_MILESTONES.filter(m => m <= n).pop() || 0;
  return { next, toNext: next ? next - n : 0, reached };
}

function goalsAndStreaks_(s, log, lifts, prs, today) {
  const todayKey = key_(today);
  const judge = dayJudge_(s, readTable_(SHEETS.TARGETS));
  const byDay = {};
  log.forEach(r => { byDay[key_(r['Date'])] = r; });
  const sick = new Set(log.filter(r => /^sick$/i.test(String(r['Feeling'] || ''))).map(r => key_(r['Date'])));
  const gymDays = new Set(lifts.map(l => key_(l['Date'])));
  const start = s['Start date'] instanceof Date ? stripTime_(s['Start date']) : stripTime_(new Date(s['Start date']));
  const firstData = log.length ? log.map(r => key_(r['Date'])).sort()[0] : todayKey;
  const fromKey = [key_(start), firstData].sort()[0];
  const days = [];
  for (let i = 0; i < 400; i++) { const k = key_(addDays_(today, -i)); if (k < fromKey) break; days.unshift(k); }
  if (!days.length) days.push(todayKey);

  // ---- daily streaks ----
  const defs = streakDefs_(s);
  const status = {};
  defs.forEach(def => { status[def.id] = days.map(k => def.test(byDay[k], judge(k), sick.has(k))); });
  const streaks = defs.map(def => {
    const st = status[def.id], r = runs_(st), last = st[st.length - 1];
    const recent = st.slice(-14).map((v, i, a) => i === a.length - 1 && v !== 'hit' ? 'today' : v);
    while (recent.length < 14) recent.unshift('none');
    return { id: def.id, icon: def.icon, label: def.label, rule: def.rule, unit: 'day',
      current: r.current, best: r.best, todayDone: last === 'hit', recent, ...milestone_(r.current) };
  });

  // ---- gym: weeks (Mon–Sun) where you did every planned session ----
  const weekStart = d => addDays_(d, -((Number(fmt_(d, 'u')) + 6) % 7));
  const weeks = [];
  for (let w = weekStart(today), i = 0; i < 104 && key_(addDays_(w, 6)) >= fromKey; w = addDays_(w, -7), i++) {
    let planned = 0, done = 0;
    for (let j = 0; j < 7; j++) {
      const d = addDays_(w, j), k = key_(d);
      if (k > todayKey) { if (k >= key_(start) && sessionFor_(Number(fmt_(d, 'u')) % 7)) planned++; continue; }
      if (gymDays.has(k)) done++;
      if (k >= key_(start) && sessionFor_(Number(fmt_(d, 'u')) % 7) && !sick.has(k)) planned++;
    }
    weeks.unshift({ planned, done, current: i === 0 });
  }
  const wst = weeks.map(w => w.planned === 0 ? 'skip' : w.done >= w.planned ? 'hit' : 'miss');
  const wr = runs_(wst), thisWk = weeks[weeks.length - 1] || { planned: 0, done: 0 };
  const wRecent = wst.slice(-8).map((v, i, a) => i === a.length - 1 && v !== 'hit' ? 'today' : v);
  while (wRecent.length < 8) wRecent.unshift('none');
  streaks.splice(1, 0, { id: 'gym', icon: '🏋️', label: 'Gym', rule: 'Every planned session in the week', unit: 'week',
    current: wr.current, best: wr.best, todayDone: wst[wst.length - 1] === 'hit', recent: wRecent,
    week: { done: thisWk.done, planned: thisWk.planned }, ...milestone_(wr.current) });

  // ---- running: weeks where you did your planned number of runs ----
  const rc = runCfg_();
  if (rc) {
    const runs = readRuns_();
    const rweeks = [];
    for (let w = weekStart(today), i = 0; i < 52 && key_(addDays_(w, 6)) >= fromKey; w = addDays_(w, -7), i++) {
      const a = key_(w), b = key_(addDays_(w, 6));
      rweeks.unshift({ done: runs.filter(r => r.date >= a && r.date <= b).length, current: i === 0 });
    }
    const rst = rweeks.map(w => w.done >= rc.runsPerWeek ? 'hit' : 'miss');
    const rr = runs_(rst), thisR = rweeks[rweeks.length - 1] || { done: 0 };
    const rRecent = rst.slice(-8).map((v, i, a) => i === a.length - 1 && v !== 'hit' ? 'today' : v);
    while (rRecent.length < 8) rRecent.unshift('none');
    streaks.splice(2, 0, { id: 'run', icon: '🏃', label: 'Running', rule: `${rc.runsPerWeek} run${rc.runsPerWeek > 1 ? 's' : ''} in the week`, unit: 'week',
      current: rr.current, best: rr.best, todayDone: rst[rst.length - 1] === 'hit', recent: rRecent,
      week: { done: thisR.done, planned: rc.runsPerWeek }, ...milestone_(rr.current) });
  }

  // ---- scorecard: the last 7 full days vs the 7 before ----
  const windowStats = (fromI, toI) => {
    const ks = []; for (let i = fromI; i <= toI; i++) ks.push(key_(addDays_(today, -i)));
    const tally = id => {
      const def = defs.find(d => d.id === id);
      const st = ks.map(k => def.test(byDay[k], judge(k), sick.has(k)));
      return { hit: st.filter(v => v === 'hit').length, of: st.filter(v => v !== 'skip').length };
    };
    let planned = 0, done = 0;
    ks.forEach(k => { const d = parseDay_(k); if (gymDays.has(k)) done++;
      if (k >= key_(start) && sessionFor_(Number(fmt_(d, 'u')) % 7) && !sick.has(k)) planned++; });
    const over = ks.filter(k => !sick.has(k) && byDay[k] && num_(byDay[k]['kcal']) !== null && Number(byDay[k]['kcal']) > judge(k).kcal * 1.1).length;
    const fat = ks.reduce((a, k) => { const r = byDay[k], f = r ? num_(r['Fat g']) : null;
      if (f === null || sick.has(k)) return a; a.of++; if (f <= judge(k).guides.fat) a.hit++; return a; }, { hit: 0, of: 0 });
    const sleeps = ks.map(k => byDay[k] ? num_(byDay[k]['Sleep h']) : null).filter(v => v !== null);
    const parts = { protein: tally('protein'), calories: tally('calories'), sugar: tally('sugar'), fat,
      steps: tally('steps'), sleep: tally('sleep'), log: tally('log'), gym: { hit: Math.min(done, planned), of: planned, done } };
    if (rc) { const rk = new Set(ks); const n = readRuns_().filter(r => rk.has(r.date)).length; parts.run = { hit: Math.min(n, rc.runsPerWeek), of: rc.runsPerWeek, done: n }; }
    parts.calories.over = over;
    parts.sleep.avg = sleeps.length ? rnd_(sleeps.reduce((a, b) => a + b, 0) / sleeps.length, 1) : null;
    const W = { protein: 25, calories: 20, gym: 25, steps: 10, sleep: 10, log: 10 };
    let sw = 0, sc = 0;
    Object.keys(W).forEach(k => { if (parts[k].of > 0) { sw += W[k]; sc += W[k] * parts[k].hit / parts[k].of; } });
    const any = ks.some(k => byDay[k]) || done > 0;
    return { parts, score: sw && any ? Math.round(sc / sw * 100) : null };
  };
  const nowW = windowStats(1, 7), prevW = windowStats(8, 14);

  // ---- body goals ----
  const weights = log.filter(r => num_(r['Weight kg']) !== null).map(r => ({ k: key_(r['Date']), w: Number(r['Weight kg']) }))
    .sort((a, b) => a.k < b.k ? -1 : 1);
  const avgTo = k => { const lo = key_(addDays_(parseDay_(k), -7)); const v = weights.filter(x => x.k > lo && x.k <= k).map(x => x.w);
    return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
  const firstW = weights.slice(0, 7).filter(x => x.k <= key_(addDays_(parseDay_(weights.length ? weights[0].k : todayKey), 6))).map(x => x.w);
  const startW = firstW.length ? firstW.reduce((a, b) => a + b, 0) / firstW.length : null;
  const nowKey = weights.length ? weights[weights.length - 1].k : null;
  const nowW8 = nowKey ? avgTo(nowKey) : null;
  let rate = null;
  if (nowKey) {
    const ref = weights.filter(x => x.k <= key_(addDays_(parseDay_(nowKey), -28))).pop() ||
                weights.filter(x => x.k <= key_(addDays_(parseDay_(nowKey), -14))).pop();
    if (ref) { const wks = (parseDay_(nowKey) - parseDay_(ref.k)) / 864e5 / 7; rate = (nowW8 - avgTo(ref.k)) / wks; }
  }
  const goalRate = Number(s['Weekly change goal kg']) || 0;
  const goalW = num_(s['Goal weight kg']);
  let wStatus = rate === null ? 'early' : Math.abs(rate - goalRate) <= 0.2 ? 'on' :
    (goalRate < 0 ? (rate < goalRate ? 'fast' : 'slow') : goalRate > 0 ? (rate > goalRate ? 'fast' : 'slow') : 'off');
  let pct = null, etaWeeks = null;
  if (goalW !== null && startW !== null && nowW8 !== null && Math.abs(startW - goalW) > 0.2) {
    pct = Math.max(0, Math.min(100, Math.round((startW - nowW8) / (startW - goalW) * 100)));
    const left = goalW - nowW8;
    const useRate = rate !== null && Math.sign(rate) === Math.sign(left) && Math.abs(rate) >= 0.05 ? rate : (Math.sign(goalRate) === Math.sign(left) ? goalRate : null);
    if (useRate && Math.abs(left) > 0.2) etaWeeks = Math.ceil(left / useRate);
    if (Math.abs(left) <= 0.2) pct = 100;
  }
  const waists = log.filter(r => num_(r['Waist cm']) !== null).map(r => ({ k: key_(r['Date']), v: Number(r['Waist cm']) })).sort((a, b) => a.k < b.k ? -1 : 1);
  const waistGoal = num_(s['Waist goal cm']) || null;
  const wFirst = waists.length ? waists[0].v : null, wNow = waists.length ? waists[waists.length - 1].v : null;
  const waistPct = waistGoal && wFirst !== null && wFirst > waistGoal ? Math.max(0, Math.min(100, Math.round((wFirst - wNow) / (wFirst - waistGoal) * 100))) : null;

  // ---- strength ----
  const mains = mainLifts_().map(m => m.toLowerCase());
  const mainUp = prs.filter(p => mains.includes(p.exercise.toLowerCase()) && p.series.length >= 2)
    .map(p => ({ exercise: p.exercise, from: p.series[0].v, to: p.series[p.series.length - 1].v,
      pct: Math.round((p.series[p.series.length - 1].v / p.series[0].v - 1) * 100) }));
  const since = key_(addDays_(today, -7));
  const newPRs = prs.filter(p => p.series.length >= 2 && p.bestDate > since && p.bestDate !== p.series[0].date).map(p => p.exercise);
  const strength = { lifts: mainUp, avgPct: mainUp.length ? Math.round(mainUp.reduce((a, b) => a + b.pct, 0) / mainUp.length) : null, newPRs };

  const goals = {
    score: nowW.score, prevScore: prevW.score, week: nowW.parts,
    weight: { start: rnd_(startW, 1), now: rnd_(nowW8, 1), goal: goalW, rate: rnd_(rate, 2), goalRate, status: wStatus, pct, etaWeeks },
    waist: { start: wFirst, now: wNow, goal: waistGoal, pct: waistPct, change: wFirst !== null ? rnd_(wNow - wFirst, 1) : null },
    strength,
  };
  return { goals, streaks, cheer: cheer_(goals, streaks, s) };
}

/** Plain-English encouragement: a headline, up to 3 wins, one nudge and one focus for next week. */
function cheer_(g, streaks, s) {
  const name = profile_().name || '';
  const sc = g.score, wk = g.week;
  const headline = sc === null ? 'Log a few days and your scorecard fills in here.' :
    sc >= 85 ? `Outstanding week${name ? ', ' + name : ''}! This is what consistency looks like.` :
    sc >= 70 ? `Strong week${name ? ', ' + name : ''} – you're building momentum.` :
    sc >= 50 ? 'Good foundation – one small tweak lifts you higher.' :
    'Every day is a fresh start. Pick the one thing below and own it.';
  const wins = [];
  if (g.strength.newPRs.length) wins.push(`🏆 New PR${g.strength.newPRs.length > 1 ? 's' : ''} this week: ${g.strength.newPRs.slice(0, 3).join(', ')}`);
  streaks.slice().sort((a, b) => b.current - a.current).forEach(st => {
    if (wins.length >= 3 || st.current < 3) return;
    const u = st.unit === 'week' ? 'week' : 'day';
    wins.push(st.current >= st.best && st.best >= 3 ? `${st.icon} ${st.current}-${u} ${st.label.toLowerCase()} streak – your best ever`
      : `${st.icon} ${st.current}-${u} ${st.label.toLowerCase()} streak`);
  });
  if (wins.length < 3 && g.waist.change !== null && g.waist.change <= -0.5) wins.push(`📏 Waist down ${Math.abs(g.waist.change)} cm since you started`);
  if (wins.length < 3 && g.weight.status === 'on') wins.push('⚖️ Weight trend is right on plan');
  if (wins.length < 3 && g.strength.avgPct !== null && g.strength.avgPct >= 3) wins.push(`💪 Main lifts up ${g.strength.avgPct}% on average since you started`);
  if (wins.length < 3 && sc !== null && g.prevScore !== null && sc - g.prevScore >= 5) wins.push(`📈 Score up ${sc - g.prevScore} points on the week before`);

  // nudge: the streak closest to a best or a milestone
  let nudge = '';
  const alive = streaks.filter(st => st.current > 0);
  const tie = alive.find(st => st.best > st.current && st.best - st.current <= 2);
  const near = alive.filter(st => st.next && st.toNext <= 2).sort((a, b) => a.toNext - b.toNext)[0];
  const u = st => st.unit === 'week' ? (st.toNext === 1 ? 'week' : 'weeks') : (st.toNext === 1 ? 'day' : 'days');
  if (near) nudge = `${near.toNext} more ${u(near)} to a ${near.next}-${near.unit} ${near.label.toLowerCase()} streak ${near.icon}`;
  else if (tie) nudge = `${tie.best - tie.current} more ${tie.unit}${tie.best - tie.current > 1 ? 's' : ''} ties your best ${tie.label.toLowerCase()} streak (${tie.best})`;
  const pending = streaks.filter(st => st.unit === 'day' && st.current >= 3 && !st.todayDone).sort((a, b) => b.current - a.current)[0];
  if (!nudge && pending) nudge = `Keep your ${pending.current}-day ${pending.label.toLowerCase()} streak alive today ${pending.icon}`;

  // focus: the weakest area (under 70%)
  const meals = meals_().filter(m => m[2] >= 15).sort((a, b) => b[2] / Math.max(b[1], 1) - a[2] / Math.max(a[1], 1));
  const pTip = meals.length ? `Tap “${meals[0][0]}” (${meals[0][2]} g) on days you're short – the Food tab shows what's left.` : 'Add a protein-rich snack on days you are short.';
  const frac = k => wk[k] && wk[k].of ? wk[k].hit / wk[k].of : null;
  const tips = {
    protein: ['Protein', `You hit protein ${wk.protein.hit} of ${wk.protein.of} days. ${pTip}`],
    gym: ['Training', `${wk.gym.done} of ${wk.gym.of} planned sessions. On a rough morning, do the short version – showing up keeps the streak and the progress.`],
    calories: ['Calories', wk.calories.over > (wk.calories.of - wk.calories.hit) / 2
      ? `Over target on ${wk.calories.over} days. Log before you eat, not after – it's the simplest way to stay on plan.`
      : 'Mostly under target. Eating too little stalls strength – add a snack so you land within 10%.'],
    log: ['Logging', 'Log something every day, even a rough estimate – the coach can only adjust what it can see.'],
    steps: ['Steps', 'A 15-minute walk after lunch is about 2,000 steps.'],
    sleep: ['Sleep', `Average ${wk.sleep.avg !== null ? wk.sleep.avg + ' h' : 'sleep not logged'}. Lights out 30 minutes earlier is the cheapest recovery boost there is.`],
  };
  const order = ['protein', 'gym', 'calories', 'log', 'steps', 'sleep'];
  const weak = order.map(k => ({ k, f: frac(k) })).filter(x => x.f !== null && x.f < 0.7).sort((a, b) => a.f - b.f)[0];
  const focus = weak ? { area: tips[weak.k][0], tip: tips[weak.k][1] }
    : sc !== null ? { area: 'Keep going', tip: 'Every area is at 70% or better. Same again next week – boring consistency is what builds muscle.' }
    : { area: 'Getting started', tip: 'Log your weight and food today, and tick your sets at the gym.' };

  // body-goal message
  const w = g.weight;
  let body = '';
  if (w.status === 'early') body = 'Weigh in most mornings – after 2 weeks the app shows whether you\'re on pace.';
  else if (w.status === 'on') body = `On track: ${w.rate > 0 ? '+' : ''}${w.rate} kg a week against your ${w.goalRate > 0 ? '+' : ''}${w.goalRate} goal.`;
  else if (w.status === 'fast' && w.goalRate < 0) body = `Losing faster than planned (${w.rate} kg/week). Fine for a week or two – if it continues, the Sunday review adds calories to protect muscle.`;
  else if (w.status === 'slow' && w.goalRate < 0) body = (g.waist.change !== null && g.waist.change <= -0.5)
    ? `The scale is slower than planned, but your waist is down ${Math.abs(g.waist.change)} cm – that's recomposition working.`
    : 'Slower than planned. Keep logging honestly – the Sunday review adjusts your calories.';
  else if (w.status === 'fast') body = `Gaining faster than planned (+${w.rate} kg/week). The Sunday review trims calories to keep it lean.`;
  else if (w.status === 'slow') body = 'Gaining slower than planned. The Sunday review adds calories.';
  else body = `Weight is moving ${w.rate > 0 ? 'up' : 'down'} ${Math.abs(w.rate)} kg a week while you aim to maintain – the Sunday review adjusts.`;

  return { headline, wins, nudge, focus, body };
}

/** End of session: record which session and how it felt. */
function webFinishSession(dateStr, session, feedback) {
  upsertLog_(parseDay_(dateStr), {
    'Session': session,
    'Pump': feedback && feedback.pump,
    'Soreness': feedback && feedback.soreness,
    'Joint pain': feedback && feedback.jointPain,
  }, []);
  return { ok: true };
}

/** Recent sessions, newest first, with every set (for the History view). */
function webGetHistory(limit) {
  ensureLiftIds_();
  const lifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date);
  const groups = {};
  lifts.forEach(l => {
    const k = key_(l['Date']) + '|' + l['Session'];
    (groups[k] = groups[k] || []).push(l);
  });
  return Object.keys(groups).sort().reverse().slice(0, limit || 12).map(k => {
    const [date, session] = k.split('|');
    const sets = groups[k].map(liftOut_);
    const volume = sets.reduce((a, x) => a + (Number(x.kg) || 0) * x.reps, 0);
    return { date, session, sets, volume: Math.round(volume) };
  });
}

/** Morning/evening check-in: weight, sleep, steps, activity, waist, notes. */
function webSaveDaily(dateStr, fields) {
  ensureLogColumns_();
  fields = Object.assign({}, fields || {});
  const actKey = ['Activity', 'Badminton/Run'].find(k => fields[k] !== undefined);
  if (actKey) { const v = fields[actKey]; delete fields['Activity']; delete fields['Badminton/Run']; fields[activityCol_()] = v; }
  upsertLog_(parseDay_(dateStr), fields, []);
  return webGetDay(dateStr);
}

// ----- Food: one row per item in the Food tab; Daily Log kcal/protein are kept as the day's sums -----

/**
 * Add one food item and return immediately.
 * Numbers typed → saved as-is. Blank numbers → instant estimate from the built-in food table;
 * if the table doesn't recognise it, the item is saved as "pending" and the page asks Gemini in the background.
 */
function webAddFood(dateStr, kcal, protein, item, source, fat, sugar, meal) {
  const d = parseDay_(dateStr);
  const name = String(item || '').trim();
  let k = num_(kcal), p = num_(protein), f = num_(fat), su = num_(sugar), src = source || 'custom';
  if (k === null && p === null) {
    if (!name) throw new Error('Enter a food name or the numbers');
    const local = localEstimate_(name);
    if (local) { k = local.kcal; p = local.protein; f = f === null ? local.fat : f; su = su === null ? local.sugar : su; src = 'estimate'; }
    else { k = 0; p = 0; src = 'pending'; }
  } else if (name && (f === null || su === null)) {
    // numbers typed but not fat/sugar: fill those from the food table when the name is recognised
    const local = localEstimate_(name);
    if (local && local.kcal) {
      const scale = k !== null && local.kcal ? k / local.kcal : 1;
      if (f === null) f = Math.round(local.fat * scale);
      if (su === null) su = Math.round(local.sugar * scale);
    }
  }
  const id = Utilities.getUuid(), time = fmt_(new Date(), 'HH:mm');
  foodSheet_().appendRow([d, time, name || 'Food', k || 0, p || 0, src, id,
    f === null ? '' : f, su === null ? '' : su, mealSlot_(meal, time)]);
  syncFoodTotals_(d);
  const out = foodDay_(d);
  if (src === 'pending') out.pendingId = id;
  return out;
}

/** Background step for "pending" items: ask Gemini and fill in the numbers. */
function webEstimateFood(id) {
  const f = findFood_(id);
  if (!f) throw new Error('Item not found – refresh and try again');
  const name = String(f.values[2]);
  const est = localEstimate_(name) || recalcFood_(name, 0, 45000) || estimateFood_(name);
  const sh = foodSheet_();
  if (est) {
    sh.getRange(f.row, 4, 1, 3).setValues([[est.kcal, est.protein, 'estimate']]);
    sh.getRange(f.row, 8, 1, 2).setValues([[est.fat === null ? '' : est.fat, est.sugar === null ? '' : est.sugar]]);
  } else sh.getRange(f.row, 6).setValue('needs numbers');
  syncFoodTotals_(f.values[0]);
  const out = foodDay_(f.values[0]);
  if (!est) out.error = 'Could not estimate "' + f.values[2] + '" (' + (lastGeminiError_ || 'unclear reply') + '). Tap Edit to add the numbers.';
  return out;
}

function webUpdateFood(id, item, kcal, protein, fat, sugar, meal) {
  const f = findFood_(id);
  if (!f) throw new Error('Item not found – refresh and try again');
  const sh = foodSheet_();
  sh.getRange(f.row, 3, 1, 4).setValues([[String(item || f.values[2]).trim(), num_(kcal) || 0, num_(protein) || 0, 'edited']]);
  if (fat !== undefined || sugar !== undefined) {
    sh.getRange(f.row, 8, 1, 2).setValues([[num_(fat) === null ? '' : num_(fat), num_(sugar) === null ? '' : num_(sugar)]]);
  }
  if (meal) sh.getRange(f.row, 10).setValue(mealSlot_(meal, f.values[1]));
  syncFoodTotals_(f.values[0]);
  return foodDay_(f.values[0]);
}

/**
 * Rename a food you'd already saved (e.g. "Coffee" → "Coffee with 2 tsp sugar") and refresh its numbers.
 * Instant when the food table knows every part; otherwise it's marked "pending" and the phone asks the AI in the background.
 */
function webRenameFood(id, item, meal) {
  const f = findFood_(id);
  if (!f) throw new Error('Item not found – refresh and try again');
  const name = String(item || '').trim().slice(0, 100);
  if (!name) throw new Error('Type what the food is first');
  const sh = foodSheet_();
  const local = localEstimate_(name);
  if (local) {
    sh.getRange(f.row, 3, 1, 4).setValues([[name, local.kcal, local.protein, 'estimate']]);
    sh.getRange(f.row, 8, 1, 2).setValues([[local.fat, local.sugar]]);
  } else sh.getRange(f.row, 3, 1, 4).setValues([[name, f.values[3], f.values[4], 'pending']]);
  if (meal) sh.getRange(f.row, 10).setValue(mealSlot_(meal, f.values[1]));
  syncFoodTotals_(f.values[0]);
  const out = foodDay_(f.values[0]);
  if (!local) out.pendingId = String(id);
  return out;
}

/** Move a food to another part of the day (Morning / Afternoon / Evening / Night). */
function webSetFoodMeal(id, meal) {
  const f = findFood_(id);
  if (!f) throw new Error('Item not found – refresh and try again');
  foodSheet_().getRange(f.row, 10).setValue(mealSlot_(meal, f.values[1]));
  return foodDay_(f.values[0]);
}

function webDeleteFood(id) {
  const f = findFood_(id);
  if (!f) throw new Error('Item not found – refresh and try again');
  foodSheet_().deleteRow(f.row);
  syncFoodTotals_(f.values[0]);
  return foodDay_(f.values[0]);
}

/** Foods you've logged in the last 30 days (not counting today), most frequent first. */
function recentFoods_(d) {
  const sh = foodSheet_();
  if (sh.getLastRow() < 2) return [];
  const to = key_(addDays_(d, -1)), from = key_(addDays_(d, -30));
  const byName = {};
  foodValues_().forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const k = key_(r[0]);
    if (k < from || k > to) return;
    const src = String(r[5]);
    if (src === 'pending' || src === 'needs numbers') return;
    const name = String(r[2]).trim();
    if (!name || name === 'Food') return;
    const key = name.toLowerCase();
    const e = byName[key] || (byName[key] = { item: name, count: 0, last: '', kcal: 0, protein: 0, fat: null, sugar: null, slots: {} });
    e.count++;
    const slot = r[9] ? mealSlot_(r[9]) : mealSlot_('', r[1]);
    e.slots[slot] = (e.slots[slot] || 0) + 1;
    if (k >= e.last) { e.last = k; e.kcal = Number(r[3]) || 0; e.protein = Number(r[4]) || 0; e.fat = num_(r[7]); e.sugar = num_(r[8]); e.item = name; }
  });
  return Object.values(byName)
    .sort((a, b) => b.count - a.count || (a.last < b.last ? 1 : -1))
    .slice(0, 12)
    .map(e => ({ item: e.item, kcal: e.kcal, protein: e.protein, fat: e.fat, sugar: e.sugar, count: e.count, meals: e.slots }));
}

/** Copy every food from one day to another (e.g. "same as yesterday"). */
function webCopyFoods(fromStr, toStr) {
  const from = parseDay_(fromStr), to = parseDay_(toStr);
  const items = foodsFor_(from).filter(f => f.source !== 'pending' && f.source !== 'needs numbers');
  if (!items.length) throw new Error('Nothing logged on ' + fromStr);
  const sh = foodSheet_();
  const time = fmt_(new Date(), 'HH:mm');
  const rows = items.map(f => [to, time, f.item, f.kcal, f.protein, 'copied', Utilities.getUuid(),
    f.fat === null ? '' : f.fat, f.sugar === null ? '' : f.sugar, f.meal]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, FOOD_HEADERS.length).setValues(rows);
  syncFoodTotals_(to);
  const out = foodDay_(to);
  out.copied = rows.length;
  return out;
}

/** Small reply for food actions (much faster than rebuilding the whole day). */
function foodDay_(d) {
  return { date: key_(d), foods: foodsFor_(d) };
}

// ----- Built-in food table: [pattern, kcal, protein g, base grams (for "200g chicken"), label] -----
// Values per typical home portion. Order matters: specific items before general ones.
const FOOD_DB = [
  // [pattern, kcal, protein g, fat g, sugar g, base grams (for "200g chicken"), label] – per typical home portion
  [/egg\s*omel+e?t+e?|omel+e?t+e?.*\begg/, 90, 6.3, 7, 0.3, 0, 'per egg, cooked with a little oil'],
  [/omel+e?t+e?|omlet/, 180, 12.6, 14, 0.6, 0, '2-egg omelette'],
  [/egg\s*white/, 17, 3.6, 0, 0.2, 0, 'per white'],
  [/\beggs?\b/, 75, 6.3, 5, 0.6, 0, 'boiled egg'],
  [/soya\s*chunks?|soy\s*chunks?|meal\s*maker|nutrela/, 175, 26, 0.5, 0, 50, '50 g dry (about 150 g soaked)'],
  [/masala\s*dosa/, 300, 6, 12, 3, 0, 'masala dosa'],
  [/dosa|dosai/, 170, 4, 6, 1, 0, 'plain dosa'],
  [/idl[iy]|idlies/, 60, 2, 0.4, 0.3, 0, 'idly'],
  [/sambh?ar/, 130, 6, 4, 4, 0, 'bowl'],
  [/chutney/, 60, 1, 5, 1, 0, 'serving'],
  [/aloo\s*paratha|paratha|parotta/, 260, 6, 12, 2, 0, 'paratha'],
  [/chapath?i|chapati|roti|phulka/, 100, 3, 3, 0.4, 0, 'chapathi'],
  [/biryani/, 500, 20, 18, 3, 0, 'plate'],
  [/rice/, 205, 4.3, 0.4, 0.1, 0, 'cup cooked'],
  [/palak\s*dal|dal\b|daal|dhal/, 180, 9, 6, 2, 0, 'bowl'],
  [/rajma/, 230, 10, 7, 3, 0, 'bowl'],
  [/chole|chana\s*masala|chickpea/, 250, 10, 9, 5, 0, 'bowl'],
  [/paneer/, 300, 18, 24, 2, 100, '100 g'],
  [/chicken\s*(curry|gravy|masala)|butter\s*chicken/, 300, 25, 18, 4, 0, 'bowl'],
  [/chicken/, 300, 38, 16, 0, 150, '150 g skinless thigh'],
  [/fish\s*(curry|gravy)/, 250, 25, 14, 3, 0, 'bowl'],
  [/fish|salmon|basa|rohu|tuna/, 230, 33, 10, 0, 150, '150 g'],
  [/lamb|mutton/, 450, 25, 32, 4, 0, 'gravy bowl'],
  [/poha/, 300, 8, 11, 3, 0, 'plate with peanuts'],
  [/upma/, 250, 6, 9, 2, 0, 'plate'],
  [/pongal/, 300, 8, 12, 1, 0, 'bowl'],
  [/chilla|cheela/, 120, 6, 4, 1, 0, 'each'],
  [/vada|vadai/, 150, 4, 9, 0.5, 0, 'each'],
  [/samosa/, 260, 5, 17, 2, 0, 'each'],
  [/whey|protein\s*(shake|scoop|powder)/, 120, 24, 1.5, 2, 0, 'scoop'],
  [/protein\s*(yog|yoghurt|yogurt|pot)/, 150, 20, 1, 8, 0, 'pot'],
  [/greek\s*yog|hung\s*curd/, 130, 20, 0.5, 7, 200, '200 g'],
  [/curd|yog(h)?urt|raita/, 100, 6, 5, 7, 0, 'bowl'],
  [/oats|porridge/, 190, 6.5, 3.5, 0.5, 50, '50 g dry'],
  [/granola/, 230, 5, 9, 10, 50, '50 g'],
  [/pasta|spaghetti|noodles/, 220, 8, 1.3, 1, 0, 'cup cooked'],
  [/bread|toast/, 80, 3, 1, 1.5, 0, 'slice'],
  [/milk/, 150, 8, 8, 12, 0, 'glass'],
  [/chai|tea/, 80, 2, 2, 10, 0, 'cup with milk and sugar'],
  [/coffee/, 60, 2, 3, 5, 0, 'cup with milk'],
  [/banana/, 105, 1.3, 0.4, 14, 0, 'medium'],
  [/apple/, 95, 0.5, 0.3, 19, 0, 'medium'],
  [/grapes?/, 60, 0.6, 0.1, 13, 87, 'handful (about 15 grapes)', 15],
  [/blue\s*berr/, 40, 0.5, 0.2, 7, 70, 'handful (about 50 berries)', 50],
  [/straw\s*berr/, 32, 0.7, 0.3, 5, 100, 'handful (about 8 strawberries)', 8],
  [/cherr(y|ies)/, 50, 1, 0.2, 10, 80, 'handful (about 10 cherries)', 10],
  [/medjool/, 66, 0.4, 0, 16, 0, 'each (medjool date)'],
  [/\bdates?\b|khajur|khajoor/, 23, 0.2, 0, 5, 0, 'each (small date)'],
  [/raisins?|kishmish/, 90, 1, 0.1, 18, 30, 'small handful (30 g, about 90 raisins)', 90],
  [/avocado/, 240, 3, 22, 1, 0, 'whole (say "half avocado" for half)'],
  [/broccoli|salad|sabzi|vegetables|veg\b/, 60, 3, 3, 3, 0, 'serving'],
  [/peanuts?|nuts|almonds?|cashews?|pistachios?|walnuts?/, 170, 7, 14, 1.5, 30, 'handful (30 g, about 23 nuts)', 23],
  [/chia/, 60, 2, 4, 0, 0, 'tablespoon'],
  [/biscuit|cookie/, 50, 1, 2, 4, 0, 'each'],
  [/sweet|mithai|ladoo|laddu|gulab\s*jamun|jalebi|halwa/, 180, 3, 8, 22, 0, 'piece'],
  [/chocolate/, 110, 1.5, 6, 12, 20, '2 squares (20 g)', 2],
  [/juice/, 110, 1, 0.3, 22, 0, 'glass'],
  [/coke|cola|soft\s*drink|soda/, 140, 0, 0, 35, 0, 'can'],
  [/sugar|jaggery|honey/, 16, 0, 0, 4, 0, 'teaspoon'],
];
const WORD_NUMS = { a: 1, an: 1, one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, half: 0.5 };

/** Instant estimate when every part of the description is in FOOD_DB, e.g. "3 egg omelette, 2 chapathi and dal". */
function localEstimate_(desc) {
  const text = String(desc).toLowerCase()
    .replace(/(chai|tea|coffee)\s+(with\s+)?milk/g, '$1')     // milk is already in the chai/coffee numbers
    .replace(/(chai|tea)\s+(with\s+)?(\d+\s*(tsp|teaspoons?|spoons?)\s*)?sugar/g, '$1')   // chai numbers already include sugar
    .replace(/poha\s+(with\s+)?(peanuts?|groundnuts?)/g, 'poha');  // peanuts already in the poha numbers
  const parts = text.split(/,|\+|&|\band\b|\bwith\b|\bplus\b/).map(s => s.trim()).filter(Boolean);
  if (!parts.length) return null;
  let kcal = 0, protein = 0, fat = 0, sugar = 0;
  for (const part of parts) {
    const entry = FOOD_DB.find(e => e[0].test(part));
    if (!entry) return null;                                   // anything unknown → let Gemini handle it
    const [, k, p, f, su, baseG, , pieces] = entry;
    const g = part.match(/(\d+(?:\.\d+)?)\s*(g|gm|gms|grams?)\b/);
    let qty;
    if (g && baseG) qty = Number(g[1]) / baseG;
    else {
      const n = part.replace(/(\d+(?:\.\d+)?)\s*(g|gm|gms|grams?|ml)\b/g, '').match(/(\d+(?:\.\d+)?)/);
      const w = part.match(/\b(a|an|one|two|three|four|five|six|half)\b/);
      qty = n ? Number(n[1]) : w ? WORD_NUMS[w[1]] : 1;
      // small things counted one by one ("10 grapes", "12 almonds", "4 squares chocolate"): the number is pieces,
      // not portions – unless you say handful / bowl / cup / serving ("2 handfuls of grapes")
      if (pieces && (n || w) && !/handfuls?|bowls?|cups?|servings?|portions?|packs?|bars?|katori/.test(part)) qty = qty / pieces;
    }
    kcal += k * qty; protein += p * qty; fat += f * qty; sugar += su * qty;
  }
  return { kcal: Math.round(kcal), protein: Math.round(protein), fat: Math.round(fat), sugar: Math.round(sugar) };
}

function foodValues_() {
  return memo_('foodValues', () => { const sh = foodSheet_(); return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, FOOD_HEADERS.length).getValues(); });
}
function foodsFor_(d) {
  const k = key_(d);
  return foodValues_()
    .filter(r => r[0] instanceof Date && key_(r[0]) === k)
    .map(r => ({ time: r[1] instanceof Date ? fmt_(r[1], 'HH:mm') : String(r[1]), item: String(r[2]),
      kcal: Number(r[3]) || 0, protein: Number(r[4]) || 0, source: String(r[5]), id: String(r[6]),
      fat: num_(r[7]), sugar: num_(r[8]), meal: r[9] ? mealSlot_(r[9]) : mealSlot_('', r[1]) }));
}

function findFood_(id) {
  const sh = foodSheet_();
  if (sh.getLastRow() < 2) return null;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, FOOD_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) if (String(values[i][6]) === String(id)) return { row: i + 2, values: values[i] };
  return null;
}

/** Daily Log kcal and protein = sum of that day's Food rows. */
function syncFoodTotals_(d) {
  const items = foodsFor_(d);
  const kcal = items.reduce((a, x) => a + x.kcal, 0);
  const protein = items.reduce((a, x) => a + x.protein, 0);
  const anyFat = items.some(x => x.fat !== null), anySugar = items.some(x => x.sugar !== null);
  const fat = items.reduce((a, x) => a + (x.fat || 0), 0);
  const sugar = items.reduce((a, x) => a + (x.sugar || 0), 0);
  const found = findLogRow_(d);
  if (!found && !items.length) return;
  ensureLogColumns_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const vals = { 'kcal': items.length ? kcal : '', 'Protein g': items.length ? protein : '',
    'Fat g': anyFat ? Math.round(fat) : '', 'Sugar g': anySugar ? Math.round(sugar) : '' };
  if (!found) upsertLog_(d, vals, []);
  else Object.keys(vals).forEach(h => {
    const c = headers.indexOf(h);
    if (c >= 0) sh.getRange(found.rowIndex, c + 1).setValue(vals[h]);
  });
}

function foodSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureSheet_(ss, SHEETS.FOOD, FOOD_HEADERS);
  // older Food tabs: add the Fat g / Sugar g headers once (checked once per page load)
  memo_('foodHeaders', () => {
    const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    if (h[7] !== 'Fat g') sh.getRange(1, 8, 1, 2).setValues([['Fat g', 'Sugar g']]).setFontWeight('bold');
    if (h[9] !== 'Meal') sh.getRange(1, 10).setValue('Meal').setFontWeight('bold');
    return true;
  });
  return sh;
}

/** Ask Gemini for kcal and protein of a described portion. Returns {kcal, protein} or null. */
function estimateFood_(desc) {
  const text = callGemini_(`Estimate total calories and protein for this food, as a typical home portion for someone whose
diet is "${profile_().diet || 'mixed'}", unless a quantity is given. The text may contain typos (e.g. "omleter bekasft" = omelette breakfast); interpret sensibly.
Food: "${desc}"
Sugar = total sugars in grams (as on a nutrition label), including sugar added to drinks like chai.
Reply with JSON only, no other text: {"kcal": <number>, "protein": <grams>, "fat": <grams>, "sugar": <grams>}`, true);
  if (!text) return null;
  const m = text.match(/\{[\s\S]*\}/);
  if (!m) return null;
  try {
    const j = JSON.parse(m[0]);
    const kcal = Math.round(Number(j.kcal)), protein = Math.round(Number(j.protein));
    const fat = num_(j.fat), sugar = num_(j.sugar);
    return isNaN(kcal) || isNaN(protein) ? null
      : { kcal, protein, fat: fat === null ? null : Math.round(fat), sugar: sugar === null ? null : Math.round(sugar) };
  } catch (e) { return null; }
}

/** Saves a gym session: one Lifts row per set, plus session + feedback on the Daily Log row. */
function webSaveWorkout(dateStr, session, sets, feedback) {
  const d = parseDay_(dateStr);
  const rows = (sets || [])
    .filter(x => x && x.exercise && Number(x.reps) > 0)
    .map(x => [d, session, x.exercise, Number(x.set) || 1, x.kg === '' ? '' : Number(x.kg) || 0, Number(x.reps), x.rir === '' || x.rir === undefined ? '' : Number(x.rir), Utilities.getUuid()]);
  if (rows.length) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIFTS);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  upsertLog_(d, {
    'Session': session,
    'Pump': feedback && feedback.pump,
    'Soreness': feedback && feedback.soreness,
    'Joint pain': feedback && feedback.jointPain,
  }, []);
  return { saved: rows.length, day: webGetDay(dateStr) };
}

/** Insert or update the Daily Log row for a date. Fields in addTo are summed, Notes are appended. */
function upsertLog_(d, fields, addTo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const found = findLogRow_(d);
  let rowIdx, current;
  if (found) {
    rowIdx = found.rowIndex;
    current = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
  } else {
    rowIdx = sh.getLastRow() + 1;
    current = headers.map(() => '');
    current[0] = d;
  }
  Object.keys(fields || {}).forEach(k => {
    const i = headers.indexOf(k);
    const v = fields[k];
    if (i < 0 || v === '' || v === null || v === undefined) return;
    if (addTo.includes(k)) current[i] = (Number(current[i]) || 0) + Number(v);
    else if (k === 'Notes' && current[i]) current[i] = `${current[i]}; ${v}`;
    else current[i] = v;
  });
  sh.getRange(rowIdx, 1, 1, headers.length).setValues([current]);
}

function findLogRow_(d) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (sh.getLastRow() < 2) return null;
  const values = sheetValues_(SHEETS.LOG);
  const headers = values[0];
  const k = key_(d);
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] instanceof Date && key_(values[i][0]) === k) {
      const o = {};
      headers.forEach((h, j) => { o[h] = values[i][j]; });
      return { rowIndex: i + 1, values: o };
    }
  }
  return null;
}

/** 'yyyy-MM-dd' → midnight Dublin time, so dates never slip a day. */
function parseDay_(str) { return Utilities.parseDate(String(str), tz_(), 'yyyy-MM-dd'); }

// ---------- Helpers ----------
/** Settings tab values (you can edit them in the sheet), falling back to Config.gs for anything missing. */
function getSettings_() { return memo_('settings', getSettingsRaw_); }
function getSettingsRaw_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const out = Object.assign({}, settingsDefaults_());
  if (sh) sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues()
    .forEach(([k, v]) => { if (k && v !== '' && v !== null) out[k] = v; });
  return out;
}

function currentTargets_(s) {
  const targets = readTable_(SHEETS.TARGETS);
  const last = targets.length ? targets[targets.length - 1] : null;
  const weekNum = weekNumber_(s['Start date'], new Date());
  return {
    calories: last ? Number(last['Calorie target']) : Number(s['Starting calorie target']),
    protein: Number(s['Protein target g']),
    steps: stepTarget_(weekNum),
  };
}

// ---------- Speed: read each tab once per page load, and run housekeeping once a day ----------
const ENGINE_VERSION = '2026-10-14.14';      // bumped on every release – the Update button compares against this
let MEMO_ = null;                       // only switched on while a read-only screen is being built
function withMemo_(fn) {
  const outer = MEMO_ === null;
  if (outer) MEMO_ = {};
  try { return fn(); } finally { if (outer) MEMO_ = null; }
}
function memo_(key, fn) {
  if (!MEMO_) return fn();
  if (!(key in MEMO_)) MEMO_[key] = fn();
  return MEMO_[key];
}
/** Runs fn at most once a day (and again after an app update). setup()/start() run everything regardless. */
function oncePerDay_(name, fn) {
  const p = PropertiesService.getScriptProperties();
  const stamp = key_(new Date()) + '|' + ENGINE_VERSION;
  if (p.getProperty('DAILY_' + name) === stamp) return;
  fn();
  p.setProperty('DAILY_' + name, stamp);
}

function readTable_(name) { return memo_('table:' + name, () => readTableRaw_(name)); }
function sheetValues_(name) {
  return memo_('values:' + name, () => { const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); return sh ? sh.getDataRange().getValues() : []; });
}
function readTableRaw_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sheetValues_(name).slice();
  const headers = values.shift();
  return values.filter(r => r.some(c => c !== '')).map(r => {
    const o = {};
    headers.forEach((h, i) => { o[h] = r[i]; });
    return o;
  });
}

function rowsBetween_(rows, from, to) {
  const a = key_(from), b = key_(to);
  return rows.filter(r => r['Date'] instanceof Date && key_(r['Date']) >= a && key_(r['Date']) <= b)
    .sort((x, y) => x['Date'] - y['Date']);
}

function stepTarget_(weekNum) {
  const c = stepsCfg_();
  return Math.min(c.max, c.start + c.increase * Math.floor((Math.max(weekNum, 1) - 1) / Math.max(1, c.everyWeeks)));
}

function weekNumber_(start, today) {
  const s = start instanceof Date ? start : new Date(start);
  return Math.floor((stripTime_(today) - stripTime_(s)) / 864e5 / 7) + 1;
}

function lastValue_(rows, col) {
  const vals = rows.filter(r => num_(r[col]) !== null);
  return vals.length ? Number(vals[vals.length - 1][col]) : null;
}

function num_(v) { return v === '' || v === null || v === undefined || isNaN(Number(v)) ? null : Number(v); }
function avg_(arr) { const n = arr.map(num_).filter(v => v !== null); return n.length ? n.reduce((a, b) => a + b, 0) / n.length : null; }
function rnd_(v, dp) { return v === null || v === undefined ? null : Number(v.toFixed(dp)); }
function round_(v, step) { return Math.round(v / step) * step; }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function stripTime_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function fmt_(d, pattern) { return Utilities.formatDate(d, tz_(), pattern); }
function key_(d) { return fmt_(d instanceof Date ? d : new Date(d), 'yyyy-MM-dd'); }
function fmtNum_(v, unit) { return v === null || v === undefined ? 'missing' : `${Number(v).toLocaleString()}${unit || ''}`; }
function tr_(label, value) { return `<tr><td style="color:#5B6874;padding-right:16px">${label}</td><td>${value}</td></tr>`; }
function stripHtml_(h) { return String(h).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); }


// =====================================================================
// AI AGENT – a coach that can act, with approvals, memory, undo,
// proactive check-ins and an optional Telegram chat.
// =====================================================================
const AGENT_SHEETS = { PROPOSALS: 'Agent proposals', LOG: 'Agent log', NOTES: 'Coach notes' };
const PROPOSAL_HEADERS = ['ID', 'Created', 'Kind', 'Summary', 'Details', 'Reason', 'Status', 'Decided', 'Until', 'Undo'];
const AGENTLOG_HEADERS = ['Time', 'Source', 'Action', 'Summary', 'Details', 'Undo', 'ID', 'Seen'];
const NOTES_HEADERS = ['Date', 'Note', 'Source', 'ID'];
const AGENT_MAX_STEPS = 6;

function agentCfg_() {
  const a = cfg_().agent || {};
  return { proactive: a.proactive !== false, eveningHour: a.eveningHour === undefined ? 21 : Number(a.eveningHour) };
}

function ensureAgentSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, AGENT_SHEETS.PROPOSALS, PROPOSAL_HEADERS);
  ensureSheet_(ss, AGENT_SHEETS.LOG, AGENTLOG_HEADERS);
  ensureSheet_(ss, AGENT_SHEETS.NOTES, NOTES_HEADERS);
}

function sheetRows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { sh, headers: sh ? sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] : [], rows: [] };
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return { sh, headers, rows: values.map((r, i) => { const o = { _row: i + 2 }; headers.forEach((h, j) => { o[h] = r[j]; }); return o; }) };
}
function setCell_(sh, headers, row, col, value) { const c = headers.indexOf(col); if (c >= 0) sh.getRange(row, c + 1).setValue(value); }

/** Write Daily Log cells, including blanks (used by undo). */
function setLogCells_(d, fields) {
  ensureLogColumns_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let found = findLogRow_(d);
  if (!found) { upsertLog_(d, {}, []); found = findLogRow_(d); }
  Object.keys(fields).forEach(k => setCell_(sh, headers, found.rowIndex, k, fields[k] === undefined || fields[k] === null ? '' : fields[k]));
}

function setSetting_(name, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const i = rows.findIndex(r => String(r[0]) === name);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(value); else sh.appendRow([name, value]);
}

// ---------- Agent log (every action, with undo) ----------
function logAction_(source, action, summary, details, undo) {
  ensureAgentSheets_();
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AGENT_SHEETS.LOG).appendRow([new Date(), source || 'app', action, summary,
    JSON.stringify(details || {}), undo ? JSON.stringify(undo) : '', id, action === 'message' ? '' : 'yes']);
  return id;
}

/** Undo an agent action (from the app's Undo button or "undo" in Telegram). */
function webUndoAction(id) {
  ensureAgentSheets_();
  const t = sheetRows_(AGENT_SHEETS.LOG);
  const r = t.rows.find(x => String(x['ID']) === String(id));
  if (!r) throw new Error('Action not found');
  if (!r['Undo']) throw new Error('That one can\'t be undone (or already was)');
  const u = JSON.parse(r['Undo']);
  if (u.type === 'foods') u.ids.forEach(fid => { try { webDeleteFood(fid); } catch (e) {} });
  else if (u.type === 'runs') u.ids.forEach(rid => { try { webDeleteRun(rid); } catch (e) {} });
  else if (u.type === 'food_update') webUpdateFood(u.id, u.prev.item, u.prev.kcal, u.prev.protein, u.prev.fat === null ? '' : u.prev.fat, u.prev.sugar === null ? '' : u.prev.sugar);
  else if (u.type === 'log') setLogCells_(parseDay_(u.date), u.prev);
  else if (u.type === 'note') { const n = sheetRows_(AGENT_SHEETS.NOTES); const row = n.rows.find(x => String(x['ID']) === u.id); if (row) n.sh.deleteRow(row._row); }
  else if (u.type === 'proposal') revertProposal_(u.id, 'undone');
  setCell_(t.sh, t.headers, r._row, 'Undo', '');
  setCell_(t.sh, t.headers, r._row, 'Summary', 'Undone: ' + r['Summary']);
  return { ok: true, summary: 'Undone: ' + r['Summary'] };
}

// ---------- Coach notes (the agent's memory) ----------
function coachNotes_() {
  return sheetRows_(AGENT_SHEETS.NOTES).rows.slice(-40).map(r => ({ date: r['Date'] instanceof Date ? key_(r['Date']) : String(r['Date']), note: String(r['Note']) }));
}

// ---------- Proposals (bigger changes need your OK) ----------
const PROPOSAL_KINDS = ['calories', 'protein', 'sets', 'replace_exercise', 'add_exercise', 'remove_exercise'];

function currentCalorieTarget_() {
  const s = getSettings_();
  return currentTargets_(s).calories;
}
function latestWeight_() {
  const log = readTable_(SHEETS.LOG).filter(r => r['Date'] instanceof Date && num_(r['Weight kg']) !== null).sort((a, b) => a['Date'] - b['Date']);
  return log.length ? Number(log[log.length - 1]['Weight kg']) : Number(profile_().weightKg) || 75;
}

/** Hard safety rails. Returns an error string, or '' if the change is allowed. Gemini can't get round these. */
function checkProposal_(p) {
  const s = getSettings_();
  const today = key_(new Date());
  if (!PROPOSAL_KINDS.includes(p.kind)) return `Unknown change "${p.kind}"`;
  if (p.until && (!/^\d{4}-\d{2}-\d{2}$/.test(p.until) || p.until <= today || p.until > key_(addDays_(new Date(), 60)))) return 'until must be a date in the next 60 days (YYYY-MM-DD)';
  if (p.kind === 'calories') {
    const cur = currentCalorieTarget_(), k = Math.round(Number(p.kcal));
    const floor = Number(s['Calorie floor']) || 1500;
    if (!k) return 'kcal is required';
    if (k < floor) return `Calories can't go below your floor of ${floor} kcal`;
    if (Math.abs(k - cur) > CONFIG.MAX_WEEKLY_CHANGE) return `Calories change by at most ${CONFIG.MAX_WEEKLY_CHANGE} kcal at a time (now ${cur})`;
    const recent = sheetRows_(AGENT_SHEETS.PROPOSALS).rows.find(r => r['Kind'] === 'calories' && r['Status'] === 'applied' &&
      r['Decided'] instanceof Date && (new Date() - r['Decided']) < 6 * 864e5);
    if (recent) return 'Calories were already changed in the last 6 days – give it a week';
    return '';
  }
  if (p.kind === 'protein') {
    const g = Math.round(Number(p.protein_g)), w = latestWeight_(), cur = Number(s['Protein target g']) || 0;
    if (!g) return 'protein_g is required';
    if (g < w * 1.4 || g > w * 2.4) return `Protein should stay between ${Math.round(w * 1.4)} and ${Math.round(w * 2.4)} g (1.4–2.4 g per kg)`;
    if (Math.abs(g - cur) > 30) return 'Protein changes by at most 30 g at a time';
    return '';
  }
  const plan = readTable_(SHEETS.PLAN);
  const rows = plan.filter(r => r['Session'] === p.session);
  if (!rows.length) return `There's no session called "${p.session}"`;
  const cap = Number(trainingCfg_().maxSetsPerSession) || 24;
  const total = rows.reduce((a, r) => a + (Number(r['Sets']) || 0), 0);
  const find = n => rows.find(r => String(r['Exercise']).toLowerCase() === String(n || '').toLowerCase());
  const setsOk = n => Number.isInteger(Number(n)) && Number(n) >= 1 && Number(n) <= 6;
  if (p.kind === 'sets') {
    const r = find(p.exercise); if (!r) return `${p.exercise} isn't in ${p.session}`;
    if (!setsOk(p.sets)) return 'Sets must be 1–6';
    if (Math.abs(Number(p.sets) - Number(r['Sets'])) > 2) return 'Sets change by at most 2 at a time';
    if (total - Number(r['Sets']) + Number(p.sets) > cap) return `That would take ${p.session} over ${cap} sets`;
    return '';
  }
  if (p.kind === 'replace_exercise') {
    const r = find(p.exercise); if (!r) return `${p.exercise} isn't in ${p.session}`;
    if (!p.new_exercise || String(p.new_exercise).length > 60) return 'new_exercise is required';
    if (p.sets !== undefined && !setsOk(p.sets)) return 'Sets must be 1–6';
    const ns = p.sets !== undefined ? Number(p.sets) : Number(r['Sets']);
    if (total - Number(r['Sets']) + ns > cap) return `That would take ${p.session} over ${cap} sets`;
    return '';
  }
  if (p.kind === 'add_exercise') {
    if (!p.new_exercise || String(p.new_exercise).length > 60) return 'new_exercise is required';
    if (find(p.new_exercise)) return `${p.new_exercise} is already in ${p.session}`;
    if (!setsOk(p.sets)) return 'Sets must be 1–6';
    if (!MUSCLES.includes(p.muscle)) return `muscle must be one of: ${MUSCLES.join(', ')}`;
    if (rows.length >= 10) return 'A session can have at most 10 exercises';
    if (total + Number(p.sets) > cap) return `That would take ${p.session} over ${cap} sets`;
    return '';
  }
  if (p.kind === 'remove_exercise') {
    if (!find(p.exercise)) return `${p.exercise} isn't in ${p.session}`;
    if (rows.length <= 2) return 'A session needs at least 2 exercises';
    return '';
  }
  return 'Not allowed';
}

function proposalSummary_(p) {
  const until = p.until ? ` until ${p.until}` : '';
  switch (p.kind) {
    case 'calories': return `Daily calories ${currentCalorieTarget_()} → ${Math.round(p.kcal)} kcal`;
    case 'protein': return `Protein target ${getSettings_()['Protein target g']} → ${Math.round(p.protein_g)} g`;
    case 'sets': return `${p.session}: ${p.exercise} → ${p.sets} sets${until}`;
    case 'replace_exercise': return `${p.session}: ${p.exercise} → ${p.new_exercise}${p.sets ? ` (${p.sets} × ${p.reps || 'same reps'})` : ''}${until}`;
    case 'add_exercise': return `${p.session}: add ${p.new_exercise} ${p.sets} × ${p.reps || '10–12'}${until}`;
    case 'remove_exercise': return `${p.session}: remove ${p.exercise}${until}`;
  }
  return p.kind;
}

function createProposal_(p, source) {
  ensureAgentSheets_();
  const err = checkProposal_(p);
  if (err) return { error: err };
  const id = 'P' + Utilities.getUuid().slice(0, 8);
  const summary = proposalSummary_(p);
  const details = Object.assign({}, p); delete details.reason;
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AGENT_SHEETS.PROPOSALS)
    .appendRow([id, new Date(), p.kind, summary, JSON.stringify(details), String(p.reason || ''), 'pending', '', p.until || '', '']);
  logAction_(source, 'proposal', 'Suggested: ' + summary, details, null);
  return { id, summary, reason: String(p.reason || ''), status: 'pending' };
}

function pendingProposals_() {
  return sheetRows_(AGENT_SHEETS.PROPOSALS).rows.filter(r => r['Status'] === 'pending')
    .map(r => ({ id: String(r['ID']), summary: String(r['Summary']), reason: String(r['Reason']), kind: String(r['Kind']), until: r['Until'] ? String(r['Until']) : '' }));
}

/** Approve or decline a suggestion (app button or Telegram button). */
function webDecideProposal(id, approve) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureAgentSheets_();
    const t = sheetRows_(AGENT_SHEETS.PROPOSALS);
    const r = t.rows.find(x => String(x['ID']) === String(id));
    if (!r) throw new Error('Suggestion not found');
    if (r['Status'] !== 'pending') return { id, status: r['Status'], summary: String(r['Summary']) };
    if (!approve) {
      setCell_(t.sh, t.headers, r._row, 'Status', 'declined'); setCell_(t.sh, t.headers, r._row, 'Decided', new Date());
      return { id, status: 'declined', summary: String(r['Summary']) };
    }
    const p = JSON.parse(r['Details']);
    const err = checkProposal_(p);                  // re-check: things may have changed since it was suggested
    if (err) {
      setCell_(t.sh, t.headers, r._row, 'Status', 'blocked'); setCell_(t.sh, t.headers, r._row, 'Decided', new Date());
      return { id, status: 'blocked', summary: String(r['Summary']), error: err };
    }
    const undo = applyProposal_(p);
    setCell_(t.sh, t.headers, r._row, 'Status', 'applied');
    setCell_(t.sh, t.headers, r._row, 'Decided', new Date());
    setCell_(t.sh, t.headers, r._row, 'Undo', JSON.stringify(undo));
    const logId = logAction_('you', 'approved', String(r['Summary']), p, { type: 'proposal', id: String(id) });
    return { id, status: 'applied', summary: String(r['Summary']), actionId: logId };
  } finally { lock.releaseLock(); }
}

function applyProposal_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.kind === 'calories') {
    const targets = ss.getSheetByName(SHEETS.TARGETS);
    if (targets && targets.getLastRow() >= 2) {
      const col = TARGET_HEADERS.indexOf('Calorie target') + 1, row = targets.getLastRow();
      const prev = targets.getRange(row, col).getValue();
      targets.getRange(row, col).setValue(Math.round(p.kcal));
      return { type: 'targets', row, col, prev };
    }
    const prev = getSettings_()['Starting calorie target'];
    setSetting_('Starting calorie target', Math.round(p.kcal));
    return { type: 'setting', name: 'Starting calorie target', prev };
  }
  if (p.kind === 'protein') {
    const prev = getSettings_()['Protein target g'];
    setSetting_('Protein target g', Math.round(p.protein_g));
    return { type: 'setting', name: 'Protein target g', prev };
  }
  // plan changes: snapshot the session's rows, then rewrite them
  const sh = ss.getSheetByName(SHEETS.PLAN);
  const values = sh.getDataRange().getValues(), headers = values[0];
  const ci = n => headers.indexOf(n);
  const mine = [], others = [];
  values.slice(1).forEach(r => (r[ci('Session')] === p.session ? mine : others).push(r));
  const before = mine.map(r => r.slice());
  const findI = n => mine.findIndex(r => String(r[ci('Exercise')]).toLowerCase() === String(n).toLowerCase());
  if (p.kind === 'sets') mine[findI(p.exercise)][ci('Sets')] = Number(p.sets);
  if (p.kind === 'replace_exercise') {
    const r = mine[findI(p.exercise)];
    r[ci('Exercise')] = p.new_exercise;
    if (p.sets !== undefined) r[ci('Sets')] = Number(p.sets);
    if (p.reps) r[ci('Reps')] = String(p.reps);
    if (p.rep_max !== undefined) r[ci('Rep max')] = Number(p.rep_max);
    if (p.increment_kg !== undefined) r[ci('Increment kg')] = Number(p.increment_kg);
  }
  if (p.kind === 'add_exercise') {
    const row = headers.map(() => '');
    row[ci('Session')] = p.session; row[ci('Order')] = mine.length + 1; row[ci('Exercise')] = p.new_exercise;
    row[ci('Sets')] = Number(p.sets); row[ci('Reps')] = String(p.reps || '10–12');
    row[ci('Rep max')] = p.rep_max !== undefined ? Number(p.rep_max) : (parseInt(String(p.reps || '12').split(/[–-]/).pop(), 10) || 12);
    row[ci('Rest')] = p.rest || '60 s'; row[ci('Increment kg')] = p.increment_kg !== undefined ? Number(p.increment_kg) : 2.5;
    row[ci('Muscle')] = p.muscle;
    mine.push(row);
  }
  if (p.kind === 'remove_exercise') mine.splice(findI(p.exercise), 1);
  mine.forEach((r, i) => { r[ci('Order')] = i + 1; });
  writePlan_(sh, headers, others.concat(mine));
  return { type: 'plan', session: p.session, rows: before };
}

function writePlan_(sh, headers, rows) {
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows.map(r => { const x = r.slice(0, headers.length); while (x.length < headers.length) x.push(''); return x; }));
}

/** Put back what a proposal changed. */
function revertProposal_(id, status) {
  const t = sheetRows_(AGENT_SHEETS.PROPOSALS);
  const r = t.rows.find(x => String(x['ID']) === String(id));
  if (!r || r['Status'] !== 'applied' || !r['Undo']) return false;
  const u = JSON.parse(r['Undo']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (u.type === 'targets') ss.getSheetByName(SHEETS.TARGETS).getRange(u.row, u.col).setValue(u.prev);
  if (u.type === 'setting') setSetting_(u.name, u.prev);
  if (u.type === 'plan') {
    const sh = ss.getSheetByName(SHEETS.PLAN);
    const values = sh.getDataRange().getValues(), headers = values[0];
    const others = values.slice(1).filter(x => x[headers.indexOf('Session')] !== u.session);
    writePlan_(sh, headers, others.concat(u.rows));
  }
  setCell_(t.sh, t.headers, r._row, 'Status', status || 'reverted');
  return true;
}

/** Temporary changes (e.g. a travel week) switch back by themselves after their "until" date. */
function agentHousekeeping_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(AGENT_SHEETS.PROPOSALS)) return;
  const today = key_(new Date());
  sheetRows_(AGENT_SHEETS.PROPOSALS).rows.forEach(r => {
    const until = r['Until'] instanceof Date ? key_(r['Until']) : String(r['Until'] || '');
    if (r['Status'] === 'applied' && until && until < today) {
      if (revertProposal_(r['ID'], 'ended')) logAction_('coach', 'message', `Temporary change ended – back to normal: ${r['Summary']}`, {}, null);
    }
    if (r['Status'] === 'pending' && r['Created'] instanceof Date && (new Date() - r['Created']) > 14 * 864e5) {
      const t = sheetRows_(AGENT_SHEETS.PROPOSALS);
      setCell_(t.sh, t.headers, r._row, 'Status', 'expired');
    }
  });
}

// ---------- Tools the coach can use ----------
const AGENT_TOOLS = [
  { name: 'get_day', description: 'Everything about one day: targets, logged weight/sleep/steps/feeling, foods eaten, the planned session with target weights, readiness, missed-session catch-up and the next workout.',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING', description: 'YYYY-MM-DD, default today' } } } },
  { name: 'get_progress', description: 'Week score, food/health hit rates, weight trend vs goal, waist, strength gains and current streaks.' },
  { name: 'get_lift_history', description: 'Recent sessions of one exercise (date, kg × reps per set).',
    parameters: { type: 'OBJECT', properties: { exercise: { type: 'STRING' }, sessions: { type: 'INTEGER', description: 'default 5' } }, required: ['exercise'] } },
  { name: 'get_plan', description: 'The training plan rows (session, exercise, sets, reps, rest, kg step, muscle).',
    parameters: { type: 'OBJECT', properties: { session: { type: 'STRING' } } } },
  { name: 'log_food', description: 'Log foods the user says they ATE (not ones they are asking about). One entry per distinct food, with your best estimates using typical Indian/Irish portions.',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING', description: 'YYYY-MM-DD, default today' },
      meal: { type: 'STRING', enum: MEAL_SLOTS, description: 'when they ate it (breakfast = Morning, lunch = Afternoon, dinner = Evening, late snack = Night); leave out to use the current time of day' },
      items: { type: 'ARRAY', items: { type: 'OBJECT', properties: { item: { type: 'STRING' }, kcal: { type: 'NUMBER' }, protein: { type: 'NUMBER' }, fat: { type: 'NUMBER' }, sugar: { type: 'NUMBER' } }, required: ['item', 'kcal', 'protein'] } } },
      required: ['items'] } },
  { name: 'correct_food', description: 'Fix a food already logged when the user corrects it (e.g. "that was soya chunks, not chicken" or "it was 3 chapathis"). Finds the latest matching item on that day and replaces its name and numbers.',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING' }, find: { type: 'STRING', description: 'words from the logged item name, e.g. "chicken"' },
      item: { type: 'STRING', description: 'corrected name with amount' }, kcal: { type: 'NUMBER' }, protein: { type: 'NUMBER' }, fat: { type: 'NUMBER' }, sugar: { type: 'NUMBER' } },
      required: ['find', 'item', 'kcal', 'protein'] } },
  { name: 'log_run', description: 'Log a run the user did (from text, a voice note or a watch/Garmin/Strava screenshot). Distance in km (convert miles).',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING' }, km: { type: 'NUMBER' }, minutes: { type: 'NUMBER', description: 'moving time, e.g. 31.5' },
      type: { type: 'STRING', enum: RUN_TYPES }, avg_hr: { type: 'INTEGER' }, rpe: { type: 'INTEGER', description: 'effort 1–10' }, notes: { type: 'STRING' } }, required: ['km'] } },
  { name: 'get_runs', description: "This week's run plan (types, distances, paces, done/missed), recent runs, weekly km and running bests.",
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING' } } } },
  { name: 'log_daily', description: 'Save morning/evening numbers the user tells you: weight, sleep, steps, waist, sport/cardio done.',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING' }, weight_kg: { type: 'NUMBER' }, sleep_h: { type: 'NUMBER' }, steps: { type: 'INTEGER' }, waist_cm: { type: 'NUMBER' },
      activity: { type: 'STRING', description: 'e.g. "Badminton 60 min", "Run 30 min easy"' } } } },
  { name: 'set_feeling', description: 'Record how the user feels for a day (usually tomorrow in the evening). This changes that day\'s session automatically.',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING' }, feeling: { type: 'STRING', enum: ['good', 'tired', 'cold', 'sick', 'sore'] } }, required: ['feeling'] } },
  { name: 'skip_workout', description: 'Mark a planned gym session as skipped with a reason. A catch-up is then offered on the next free morning.',
    parameters: { type: 'OBJECT', properties: { date: { type: 'STRING' }, reason: { type: 'STRING', enum: ['busy', 'tired', 'travel', 'sick', 'other'] } }, required: ['reason'] } },
  { name: 'add_coach_note', description: 'Remember something lasting about the user: injuries, preferences, equipment, schedule changes, foods they dislike.',
    parameters: { type: 'OBJECT', properties: { note: { type: 'STRING' } }, required: ['note'] } },
  { name: 'propose_change', description: 'Suggest a change to calories, protein or the training plan. The user must approve it in the app or Telegram – it is NOT applied yet. Hard limits are enforced; if it returns an error, explain it and adjust.',
    parameters: { type: 'OBJECT', properties: {
      kind: { type: 'STRING', enum: PROPOSAL_KINDS },
      session: { type: 'STRING' }, exercise: { type: 'STRING', description: 'existing exercise' }, new_exercise: { type: 'STRING' },
      sets: { type: 'INTEGER' }, reps: { type: 'STRING', description: 'e.g. "8–10"' }, rep_max: { type: 'INTEGER' }, increment_kg: { type: 'NUMBER' },
      muscle: { type: 'STRING', enum: MUSCLES }, kcal: { type: 'INTEGER' }, protein_g: { type: 'INTEGER' },
      until: { type: 'STRING', description: 'YYYY-MM-DD – for temporary changes (e.g. a travel week); switches back automatically' },
      reason: { type: 'STRING', description: 'one short sentence the user will see' } }, required: ['kind', 'reason'] } },
];

/** Run one tool. Returns a small JSON-able result for the model; records actions with undo. */
function execTool_(name, args, ctx) {
  args = args || {};
  const today = key_(new Date());
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(args.date || '')) ? String(args.date) : today;
  if (date > key_(addDays_(new Date(), 7))) return { error: 'Date is too far ahead' };
  const act = (action, summary, details, undo) => { const id = logAction_(ctx.source, action, summary, details, undo); ctx.actions.push({ id, summary, undoable: !!undo }); return id; };
  switch (name) {
    case 'get_day': {
      const d = webGetDay(date);
      return { date: d.date, scheduled: d.scheduled || 'rest day', readiness: d.readiness.title || 'normal', targets: d.targets,
        logged: d.log, skipped: d.skipped, foods: d.foods.map(f => `${f.item} ${f.kcal} kcal ${f.protein} g P`),
        session: (d.sessions[d.log['Session'] || d.scheduled || (d.catchUp && d.catchUp.session)] || []).map(e => `${e.exercise} ${e.sets}×${e.reps} → ${e.target}`),
        setsDone: d.setsToday.length, catchUp: d.catchUp, nextWorkout: d.nextWorkout && { date: d.nextWorkout.date, session: d.nextWorkout.session,
          exercises: d.nextWorkout.exercises.map(e => `${e.exercise} ${e.sets}×${e.reps} → ${e.target}`) } };
    }
    case 'get_progress': {
      const d = webGetDashboard();
      return { score: d.goals.score, previousScore: d.goals.prevScore, lastWeek: d.goals.week, weight: d.goals.weight, waist: d.goals.waist,
        strength: d.goals.strength, streaks: d.streaks.map(s => `${s.label}: ${s.current} ${s.unit}s (best ${s.best})`), cheer: d.cheer };
    }
    case 'get_lift_history': {
      const n = Math.min(10, Number(args.sessions) || 5), ex = String(args.exercise || '').toLowerCase();
      const by = {};
      readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date && String(l['Exercise']).toLowerCase() === ex)
        .forEach(l => { const k = key_(l['Date']); (by[k] = by[k] || []).push(`${l['Weight kg'] === '' ? 'bw' : l['Weight kg']}×${l['Reps']}`); });
      return { exercise: args.exercise, sessions: Object.keys(by).sort().slice(-n).map(k => `${k}: ${by[k].join(', ')}`) };
    }
    case 'get_plan':
      return { plan: readTable_(SHEETS.PLAN).filter(r => !args.session || r['Session'] === args.session)
        .map(r => `${r['Session']} #${r['Order']}: ${r['Exercise']} ${r['Sets']}×${r['Reps']} rest ${r['Rest']} +${r['Increment kg']} kg (${r['Muscle']})`) };
    case 'log_food': {
      const items = (args.items || []).filter(x => x && x.item).slice(0, 12).map(x => {
        const k = Math.max(0, Math.min(3000, Math.round(Number(x.kcal) || 0)));
        return { item: String(x.item).slice(0, 80), kcal: k, protein: Math.max(0, Math.round(Number(x.protein) || 0)),
          fat: num_(x.fat) === null ? '' : Math.round(Number(x.fat)), sugar: num_(x.sugar) === null ? '' : Math.round(Number(x.sugar)) };
      });
      if (!items.length) return { error: 'No foods given' };
      const before = new Set(foodsFor_(parseDay_(date)).map(f => f.id));
      const out = webAddFoods(date, items, 'coach', args.meal);
      const ids = out.foods.filter(f => !before.has(f.id)).map(f => f.id);
      const kcal = items.reduce((a, b) => a + b.kcal, 0), p = items.reduce((a, b) => a + b.protein, 0);
      const slot = (out.foods.find(f => ids.indexOf(f.id) >= 0) || {}).meal || '';
      act('log_food', `Logged ${items.map(i => i.item).join(', ')} (${kcal} kcal, ${p} g protein)${slot ? ' – ' + slot : ''}${date !== today ? ' for ' + date : ''}`, { date, items }, { type: 'foods', ids });
      const fs = out.foods.filter(f => f.source !== 'pending');
      return { ok: true, logged: items, dayTotals: { kcal: fs.reduce((a, f) => a + (Number(f.kcal) || 0), 0), protein: fs.reduce((a, f) => a + (Number(f.protein) || 0), 0) } };
    }
    case 'correct_food': {
      const want = String(args.find || '').toLowerCase().trim();
      const list = foodsFor_(parseDay_(date)).filter(f => want && String(f.item).toLowerCase().indexOf(want) >= 0);
      if (!list.length) return { error: `No food matching "${args.find}" logged on ${date}` };
      const f = list[list.length - 1];
      const prev = { item: f.item, kcal: f.kcal, protein: f.protein, fat: f.fat, sugar: f.sugar };
      const n = { item: String(args.item).slice(0, 80), kcal: Math.max(0, Math.round(Number(args.kcal) || 0)), protein: Math.max(0, Math.round(Number(args.protein) || 0)),
        fat: num_(args.fat) === null ? '' : Math.round(Number(args.fat)), sugar: num_(args.sugar) === null ? '' : Math.round(Number(args.sugar)) };
      webUpdateFood(f.id, n.item, n.kcal, n.protein, n.fat, n.sugar);
      act('correct_food', `Corrected ${prev.item} → ${n.item} (${n.kcal} kcal, ${n.protein} g protein)`, { date, from: prev, to: n }, { type: 'food_update', id: f.id, prev });
      return { ok: true, was: prev, now: n };
    }
    case 'log_run': {
      const before = new Set(readRuns_().map(r => r.id));
      const r = webLogRun(date, { km: args.km, min: args.minutes, type: args.type, hr: args.avg_hr, rpe: args.rpe, notes: args.notes, source: ctx.source });
      const added = readRuns_().filter(x => !before.has(x.id)).map(x => x.id);
      act('log_run', `Logged a ${args.type || 'easy'} run: ${args.km} km${args.minutes ? ' in ' + args.minutes + ' min' : ''}${date !== today ? ' (' + date + ')' : ''}`,
        { date, args }, { type: 'runs', ids: added });
      return { ok: true, week: r.week ? { doneKm: r.week.doneKm, targetKm: r.week.targetKm, runsDone: r.week.runsDone, runsPlanned: r.week.runsPlanned } : null };
    }
    case 'get_runs': {
      const r = webGetRuns(date);
      return { configured: r.configured, goal: r.goal, week: r.week, today: r.today, zones: r.zones, recent: r.recent.slice(0, 5), weekly: r.weekly, stats: r.stats };
    }
    case 'log_daily': {
      const map = { weight_kg: 'Weight kg', sleep_h: 'Sleep h', steps: 'Steps', waist_cm: 'Waist cm', activity: 'Activity' };
      const fields = {};
      Object.keys(map).forEach(k => { if (args[k] !== undefined && args[k] !== null && args[k] !== '') fields[map[k]] = args[k]; });
      if (fields['Weight kg'] !== undefined && (fields['Weight kg'] < 30 || fields['Weight kg'] > 250)) return { error: 'Weight looks wrong' };
      if (fields['Sleep h'] !== undefined && (fields['Sleep h'] < 0 || fields['Sleep h'] > 16)) return { error: 'Sleep looks wrong' };
      if (!Object.keys(fields).length) return { error: 'Nothing to save' };
      const found = findLogRow_(parseDay_(date)) || { values: {} };
      const col = activityCol_();
      const prev = {}; Object.keys(fields).forEach(f => { const c = f === 'Activity' ? col : f; prev[c] = found.values[c] === undefined ? '' : found.values[c]; });
      webSaveDaily(date, fields);
      act('log_daily', 'Saved ' + Object.keys(fields).map(f => `${f.replace(' kg', '').replace(' h', '').replace(' cm', '').toLowerCase()} ${fields[f]}`).join(', ') + (date !== today ? ' for ' + date : ''), { date, fields }, { type: 'log', date, prev });
      return { ok: true, saved: fields };
    }
    case 'set_feeling': {
      const prevRow = findLogRow_(parseDay_(date));
      const prev = { 'Feeling': prevRow ? prevRow.values['Feeling'] || '' : '' };
      const r = webSetFeeling(date, args.feeling);
      act('set_feeling', `Feeling "${r.feeling}" for ${date === today ? 'today' : date}` + (r.readiness.title ? ` – ${r.readiness.title.toLowerCase()}` : ''), { date, feeling: r.feeling }, { type: 'log', date, prev });
      return { ok: true, feeling: r.feeling, sessionThatDay: r.session, effect: r.readiness.title ? r.readiness.title + ': ' + r.readiness.message : 'normal session' };
    }
    case 'skip_workout': {
      const prevRow = findLogRow_(parseDay_(date));
      const prev = { 'Skipped': prevRow ? prevRow.values['Skipped'] || '' : '', 'Feeling': prevRow ? prevRow.values['Feeling'] || '' : '' };
      const d = webSkipDay(date, args.reason);
      act('skip_workout', `Skipped ${d.scheduled || 'workout'} on ${date === today ? 'today' : date} (${args.reason})`, { date, reason: args.reason }, { type: 'log', date, prev });
      return { ok: true, catchUp: d.catchUp };
    }
    case 'add_coach_note': {
      ensureAgentSheets_();
      const note = String(args.note || '').trim().slice(0, 300);
      if (!note) return { error: 'Empty note' };
      const id = Utilities.getUuid();
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AGENT_SHEETS.NOTES).appendRow([new Date(), note, ctx.source, id]);
      act('note', 'Noted: ' + note, { note }, { type: 'note', id });
      return { ok: true };
    }
    case 'propose_change': {
      if (ctx.proposals.length >= 3) return { error: 'At most 3 suggestions at a time' };
      const r = createProposal_(args, ctx.source);
      if (r.error) return { error: r.error };
      ctx.proposals.push(r);
      return { ok: true, id: r.id, summary: r.summary, note: 'Waiting for the user to approve' };
    }
  }
  return { error: 'Unknown tool ' + name };
}

function agentSystem_(source) {
  const today = parseDay_(key_(new Date()));
  const notes = coachNotes_();
  return coachSystem_().replace(/- The app sets the calorie target[\s\S]*?Settings tab\.\n/,
    '') + `

You are also an AGENT: you can act on the user's tracker with tools.
- Today is ${fmt_(today, 'EEEE')} ${key_(today)}. Work out dates for "yesterday", "tomorrow", "Saturday" etc. yourself.
- Do it, don't just say it: when the user TELLS you facts ("had 2 dosa and sambar", "weighed 74.6", "slept 6h", "played badminton 90 min", "feeling tired for tomorrow", "can't train today – work"), call the matching tool straight away. Don't log things they only ask about.
- If the user corrects something already logged ("that was soya chunks, not chicken", "only 1 chapathi"), use correct_food – don't log it twice.
- Food: one log_food item per distinct food, always with kcal, protein, fat and sugar estimates (typical Indian home portions unless stated).
- Meal time: set log_food "meal" from what they say (breakfast → Morning, lunch → Afternoon, dinner → Evening, late snack → Night). If they don't say, leave it out and the current time of day is used.
  For a meal PHOTO: first judge each portion in grams from visual cues (dinner plate ~26 cm, katori ~150 ml, chapathi ~35 g, tablespoon ~15 ml),
  then work out the macros; put the amount in the item name (e.g. "Chicken curry ~200 g"). Use these reference values for their usual foods:
${foodReferenceText_()}
- Bigger changes – calories, protein, sets, swapping/adding/removing exercises – go through propose_change. Say clearly it's a SUGGESTION waiting for their approval; never claim it's done. For travel or temporary limits use "until".
- The code enforces safety limits (calorie floor, max ±150 kcal at a time, sets ±2 and 1–6, protein 1.4–2.4 g/kg). If a tool returns an error, explain briefly and offer what IS possible.
- Runs: use log_run for runs they tell you about or show in a watch/Strava screenshot; get_runs for their run plan and paces. Never add more than about 10% to weekly running distance.
- Short on time or can't get to the gym? Suggest the 25-minute home workout ("${HOME_LOWER}" / "${HOME_UPPER}", the 🏠 button on the Gym tab). It counts as that day's session.
- Never plan training on a day marked sick. Lift weights come from the app's progression – don't invent target kg.
- Use add_coach_note when you learn something lasting (injury, equipment, preference, schedule). Respect these notes:
${notes.length ? notes.map(n => '  • ' + n.date + ': ' + n.note).join('\n') : '  (no notes yet)'}
- After acting, reply in 1–4 short lines confirming what changed. ${source === 'telegram' ? 'This is a Telegram chat: plain text, no markdown tables, keep it short. Tell them they can reply "undo".' : 'They can tap Undo in the app.'}`;
}

/**
 * The agent loop: Gemini decides which tools to call, we run them (with the safety rails), and feed results back.
 * input = { text, source: 'app'|'telegram'|'coach', history: [{role, text}], media: [{mime, data}], toolFilter }
 */
function runAgent_(input) {
  ensureAgentSheets_();
  const ctx = { source: input.source || 'app', actions: [], proposals: [] };
  const tools = input.toolFilter ? AGENT_TOOLS.filter(t => input.toolFilter.includes(t.name)) : AGENT_TOOLS;
  const userParts = [{ text: String(input.text || '').slice(0, 4000) || 'Please look at this.' }]
    .concat((input.media || []).map(m => ({ inline_data: { mime_type: m.mime, data: m.data } })));
  const contents = (input.history || []).slice(-8).map(h => ({ role: h.role === 'coach' ? 'model' : 'user', parts: [{ text: String(h.text).slice(0, 1500) }] }))
    .concat([{ role: 'user', parts: userParts }]);
  const system = agentSystem_(ctx.source) + '\n\nDATA SNAPSHOT (from their logs, as of now):\n' + JSON.stringify(coachContext_());
  let reply = '';
  for (let step = 0; step < AGENT_MAX_STEPS; step++) {
    const r = geminiRaw_({ systemInstruction: { parts: [{ text: system }] }, contents,
      tools: [{ functionDeclarations: tools }], toolConfig: { functionCallingConfig: { mode: 'AUTO' } },
      generationConfig: { temperature: 0.3 } }, true);
    if (!r) throw new Error('The coach is unavailable right now (' + (lastGeminiError_ || 'no reply') + '). Try again in a minute.');
    const calls = r.parts.filter(p => p.functionCall);
    contents.push({ role: 'model', parts: r.parts });          // keep parts exactly as sent back (thought signatures)
    if (!calls.length) { reply = r.text; break; }
    contents.push({ role: 'user', parts: calls.map(c => {
      let out;
      try { out = execTool_(c.functionCall.name, c.functionCall.args, ctx); } catch (e) { out = { error: e.message }; }
      return { functionResponse: { name: c.functionCall.name, response: { result: out } } };
    }) });
  }
  if (!reply) reply = ctx.actions.length || ctx.proposals.length ? 'Done.' : 'Sorry – I couldn\'t finish that. Try rephrasing.';
  return { reply: reply.replace(/^\x60{3}\w*\s*/, '').replace(/\x60{3}\s*$/, '').trim(), actions: ctx.actions, proposals: ctx.proposals };
}

// ---------- Phone app: Coach tab ----------
/** Chat with the agent from the app. history = [{role:'user'|'coach', text}] */
function webAgentChat(message, history) {
  const q = String(message || '').trim();
  if (!q) throw new Error('Type a question or message first');
  const out = runAgent_({ text: q, source: 'app', history });
  out.answer = out.reply;      // older pages read "answer"
  return out;
}
/** Kept for older versions of the page. */
function webAskCoach(question, history) { return webAgentChat(question, history); }

/** Pending suggestions, recent actions (with undo) and unread check-in messages for the Coach tab. */
function webGetAgentState() {
  ensureAgentSheets_();
  try { agentHousekeeping_(); } catch (e) {}
  const t = sheetRows_(AGENT_SHEETS.LOG);
  const unread = t.rows.filter(r => r['Action'] === 'message' && r['Seen'] !== 'yes');
  unread.forEach(r => setCell_(t.sh, t.headers, r._row, 'Seen', 'yes'));
  const recent = t.rows.filter(r => r['Action'] !== 'message' && r['Action'] !== 'proposal').slice(-12).reverse()
    .map(r => ({ id: String(r['ID']), when: r['Time'] instanceof Date ? fmt_(r['Time'], 'EEE HH:mm') : '', source: String(r['Source']), summary: String(r['Summary']), undoable: !!r['Undo'] }));
  return { proposals: pendingProposals_(), recent, messages: unread.map(r => ({ text: String(r['Summary']), at: r['Time'] instanceof Date ? r['Time'].getTime() : Date.now() })),
    notes: coachNotes_().slice(-10).reverse(), telegram: telegramOn_() };
}

// ---------- Proactive check-ins ----------
/** Deliver a coach message: Telegram if connected, and always to the app's Coach tab. */
function agentDeliver_(text, proposals) {
  logAction_('coach', 'message', text, {}, null);
  if (telegramOn_()) {
    tgSend_(text);
    (proposals || []).forEach(p => tgSendProposal_(p));
    tgRemember_('coach', text);
  }
}

/** Runs every evening (createTriggers). Looks at today and tomorrow and sends a short, specific nudge. */
function agentEveningCheck() {
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) return;
  const today = key_(new Date());
  const d = webGetDay(today);
  const missed = (() => {
    const weekStart = addDays_(parseDay_(today), -((Number(fmt_(parseDay_(today), 'u')) + 6) % 7));
    const lifts = readTable_(SHEETS.LIFTS).filter(l => l['Date'] instanceof Date && key_(l['Date']) >= key_(weekStart));
    const done = new Set(lifts.map(l => String(l['Session'])));
    const log = readTable_(SHEETS.LOG);
    let n = 0;
    for (let x = new Date(weekStart); key_(x) <= today; x = addDays_(x, 1)) {
      const name = sessionFor_(Number(fmt_(x, 'u')) % 7);
      const r = log.find(l => l['Date'] instanceof Date && key_(l['Date']) === key_(x)) || {};
      if (name && !done.has(name) && !/^sick$/i.test(String(r['Feeling'] || ''))) n++;
    }
    return n;
  })();
  const prompt = `Write tonight's check-in for ${profile_().name} (it's ${fmt_(new Date(), 'HH:mm')}). Use get_day for today and tomorrow if you need detail.
Cover, in at most 6 short lines:
1. Food: protein and calories so far vs target today; if short, name 1–2 of THEIR foods that close the gap (with grams of protein).
2. Training: if they trained today, one line on it (a PR, or a lift that's stuck). If they missed sessions, mention the catch-up.${missed >= 2 ? ` They have missed ${missed} planned sessions this week – suggest a lighter, shorter week and you may propose_change once if it clearly helps.` : ''}
3. Tomorrow: what's planned (or rest/recovery), in one line.
4. End by asking how they feel for tomorrow: they can reply good, tired, sore, cold or sick.
Do NOT log anything. Friendly, specific, no fluff.`;
  const out = runAgent_({ text: prompt, source: 'coach', toolFilter: ['get_day', 'get_progress', 'get_lift_history', 'get_plan', 'propose_change'] });
  agentDeliver_(out.reply, out.proposals);
  return out;
}

/** Called at the end of the Sunday weekly review: the agent reads the week and may suggest up to 2 changes. */
function agentWeekly_(summary, cal, vol) {
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) return;
  const prompt = `The weekly review just ran. Summary: ${JSON.stringify({ summary, calories: cal, setChanges: vol && vol.changes, notes: vol && vol.notes })}.
Calories and sets were already adjusted by the app's rules – don't repeat those. Look for ONE or TWO other improvements that are clearly worth it
(a stuck lift that needs a variation swap, a joint niggle in the coach notes, protein consistently short, too many skipped sessions)
and use propose_change for them. If nothing is clearly worth changing, propose nothing.
Then write a 3-line note: the week's headline win, the one focus for next week, and what you suggested (if anything).`;
  const out = runAgent_({ text: prompt, source: 'coach', toolFilter: ['get_progress', 'get_lift_history', 'get_plan', 'get_day', 'propose_change'] });
  agentDeliver_(out.reply, out.proposals);
  return out;
}

// ---------- Telegram (optional) ----------
// Setup: see connectTelegram(). The bot only ever answers the one chat you pair with a code.
function tgProps_() {
  const p = PropertiesService.getScriptProperties(), tg = cfg_().telegram || {};
  // Telegram settings can live in Config.gs – copy them in so the rest of the code finds them
  if (tg.botToken && !p.getProperty('TELEGRAM_BOT_TOKEN')) p.setProperty('TELEGRAM_BOT_TOKEN', String(tg.botToken).trim());
  if (tg.webhookUrl && !p.getProperty('TELEGRAM_WEBHOOK_URL')) p.setProperty('TELEGRAM_WEBHOOK_URL', String(tg.webhookUrl).trim());
  return p;
}
function telegramOn_() { const p = PropertiesService.getScriptProperties(); return !!(p.getProperty('TELEGRAM_BOT_TOKEN') && p.getProperty('TELEGRAM_CHAT_ID')); }

function tgApi_(method, payload) {
  const token = tgProps_().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) throw new Error('TELEGRAM_BOT_TOKEN is not set');
  const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/${method}`, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload || {}), muteHttpExceptions: true });
  let j = {};
  try { j = JSON.parse(res.getContentText()); } catch (e) {}
  if (!j.ok) console.error(`Telegram ${method}: ${res.getContentText().slice(0, 300)}`);
  return j;
}
function tgPlain_(text) { return String(text || '').replace(/\*\*(.+?)\*\*/g, '$1').slice(0, 4000); }
function tgSend_(text, extra) {
  const chat = tgProps_().getProperty('TELEGRAM_CHAT_ID');
  if (!chat) return null;
  return tgApi_('sendMessage', Object.assign({ chat_id: chat, text: tgPlain_(text) }, extra || {}));
}
function tgSendProposal_(p) {
  return tgSend_(`Suggestion: ${p.summary}${p.reason ? '\nWhy: ' + p.reason : ''}`, { reply_markup: { inline_keyboard: [[
    { text: '✅ Approve', callback_data: 'ap:' + p.id }, { text: '✖ Decline', callback_data: 'de:' + p.id }]] } });
}
/** Short rolling chat history for Telegram, so replies like "tired" make sense after the evening check-in. */
function tgHistory_() { try { return JSON.parse(tgProps_().getProperty('TG_HISTORY') || '[]'); } catch (e) { return []; } }
function tgRemember_(role, text) {
  let h = tgHistory_(); h.push({ role, text: String(text).slice(0, 700) }); h = h.slice(-10);
  while (JSON.stringify(h).length > 8000) h.shift();
  tgProps_().setProperty('TG_HISTORY', JSON.stringify(h));
}
function tgDownload_(fileId) {
  const token = tgProps_().getProperty('TELEGRAM_BOT_TOKEN');
  const f = tgApi_('getFile', { file_id: fileId });
  if (!f.ok || !f.result || !f.result.file_path) throw new Error('Could not download that file from Telegram');
  if (f.result.file_size && f.result.file_size > 15 * 1024 * 1024) throw new Error('That file is too big (15 MB max)');
  const blob = UrlFetchApp.fetch(`https://api.telegram.org/file/bot${token}/${f.result.file_path}`).getBlob();
  return Utilities.base64Encode(blob.getBytes());
}

/**
 * Run from the editor after: (1) creating a bot with @BotFather and saving TELEGRAM_BOT_TOKEN in Script properties,
 * (2) a second deployment (Web app, Execute as Me, Who has access: Anyone) saved as TELEGRAM_WEBHOOK_URL.
 */
function connectTelegram() {
  const p = tgProps_();
  const url = p.getProperty('TELEGRAM_WEBHOOK_URL');
  if (!p.getProperty('TELEGRAM_BOT_TOKEN')) throw new Error('Add TELEGRAM_BOT_TOKEN in Project Settings → Script properties first');
  if (!url || !/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(url)) throw new Error('Add TELEGRAM_WEBHOOK_URL (the /exec URL of your "Anyone" deployment) in Script properties first');
  const secret = Utilities.getUuid().replace(/-/g, '');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  p.setProperty('TELEGRAM_SECRET', secret);
  p.setProperty('TELEGRAM_PAIR_CODE', code);
  p.deleteProperty('TELEGRAM_CHAT_ID'); p.deleteProperty('TG_LAST_UPDATE'); p.deleteProperty('TG_HISTORY');
  const r = tgApi_('setWebhook', { url: url + '?tg=' + secret, drop_pending_updates: true, allowed_updates: ['message', 'callback_query'] });
  const me = tgApi_('getMe', {});
  const msg = r.ok ? `OK – now open Telegram, find @${me.result ? me.result.username : 'your bot'} and send:  /start ${code}` : 'FAILED: ' + JSON.stringify(r);
  console.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return code;
}
function disconnectTelegram() {
  tgApi_('deleteWebhook', { drop_pending_updates: true });
  ['TELEGRAM_CHAT_ID', 'TELEGRAM_SECRET', 'TELEGRAM_PAIR_CODE', 'TG_LAST_UPDATE', 'TG_HISTORY'].forEach(k => tgProps_().deleteProperty(k));
  console.log('Telegram disconnected');
}
/** Check the webhook from the editor. */
function telegramStatus() { console.log(JSON.stringify(tgApi_('getWebhookInfo', {}))); }

/** Telegram calls this. Always answers with HtmlService (a ContentService reply makes Telegram see a redirect and retry). */
function doPost(e) {
  try { handleTelegram_(e); } catch (err) { console.error(err); try { tgSend_('Sorry – something went wrong: ' + err.message); } catch (x) {} }
  return HtmlService.createHtmlOutput('ok');
}

function handleTelegram_(e) {
  const p = tgProps_();
  if (!e || !e.parameter || !p.getProperty('TELEGRAM_SECRET') || e.parameter.tg !== p.getProperty('TELEGRAM_SECRET')) return;
  const u = JSON.parse((e.postData && e.postData.contents) || '{}');
  // ignore repeats (Telegram re-sends an update if it thinks we were slow)
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const last = Number(p.getProperty('TG_LAST_UPDATE') || 0);
    if (u.update_id && u.update_id <= last) return;
    if (u.update_id) p.setProperty('TG_LAST_UPDATE', String(u.update_id));
  } finally { lock.releaseLock(); }

  const paired = p.getProperty('TELEGRAM_CHAT_ID');
  if (u.callback_query) {
    const cq = u.callback_query;
    if (String(cq.message && cq.message.chat && cq.message.chat.id) !== paired) return;
    const m = String(cq.data || '').match(/^(ap|de):(.+)$/);
    if (!m) return;
    const r = webDecideProposal(m[2], m[1] === 'ap');
    const msg = r.status === 'applied' ? `✅ Done: ${r.summary}` : r.status === 'declined' ? `✖ Declined: ${r.summary}` :
      r.status === 'blocked' ? `Couldn't apply – ${r.error}` : `Already ${r.status}: ${r.summary}`;
    tgApi_('answerCallbackQuery', { callback_query_id: cq.id, text: msg.slice(0, 190) });
    tgApi_('editMessageText', { chat_id: paired, message_id: cq.message.message_id, text: msg });
    return;
  }
  const msg = u.message;
  if (!msg || !msg.chat) return;
  const chatId = String(msg.chat.id), text = String(msg.text || msg.caption || '').trim();

  if (/^\/start\b/.test(text)) {
    const code = text.split(/\s+/)[1];
    if (!paired && code && code === p.getProperty('TELEGRAM_PAIR_CODE')) {
      p.setProperty('TELEGRAM_CHAT_ID', chatId); p.deleteProperty('TELEGRAM_PAIR_CODE');
      tgSend_(`Connected 💪 I'm your coach. Tell me what you ate, your weight or sleep, send a meal photo or a voice note, or ask anything.\nReply "undo" to reverse my last change.`);
    }
    return;                                        // wrong code or someone else: stay silent
  }
  if (chatId !== paired) return;                   // only your chat

  if (/^\/?undo$/i.test(text)) {
    const last = sheetRows_(AGENT_SHEETS.LOG).rows.filter(r => r['Source'] === 'telegram' && r['Undo']).pop();
    tgSend_(last ? webUndoAction(last['ID']).summary : 'Nothing to undo.');
    return;
  }
  if (/^\/help$/i.test(text)) {
    tgSend_('Try: "had 2 dosa and sambar", "weighed 74.6, slept 6h", "travelling Mon–Thu, hotel gym has dumbbells to 25 kg", "how am I doing this week?", a meal photo, or a voice note. "undo" reverses my last change.');
    return;
  }
  tgApi_('sendChatAction', { chat_id: paired, action: 'typing' });
  const media = [];
  let prompt = text;
  if (msg.photo && msg.photo.length) {
    const best = msg.photo[msg.photo.length - 1];
    media.push({ mime: 'image/jpeg', data: tgDownload_(best.file_id) });
    if (!prompt) prompt = 'Here is a photo of what I ate – log it.';
  }
  const voice = msg.voice || msg.audio;
  if (voice) {
    media.push({ mime: voice.mime_type || 'audio/ogg', data: tgDownload_(voice.file_id) });
    prompt = (prompt ? prompt + '\n' : '') + 'This is a voice note from me – listen and do what I say.';
  }
  if (!prompt && !media.length) return;
  const out = runAgent_({ text: prompt, source: 'telegram', history: tgHistory_(), media });
  tgRemember_('user', prompt);
  tgRemember_('coach', out.reply);
  tgSend_(out.reply);
  out.proposals.forEach(pr => tgSendProposal_(pr));
}

// =====================================================================================
// ---------- Updates: the "Update now" button ----------
// The app code (Code + Log) is published at a public address. Once a day the app checks for a newer
// version; you get an email and a banner in the app, and one tap installs it into YOUR project.
// Your Config, Sheet, data, Gemini key, triggers and app link are not touched.
// Needs: the Apps Script API switched on (script.google.com/home/usersettings) and the
// script.projects / script.deployments permissions (in appsscript.json).
// Config (optional): updates: { mode: 'ask' | 'auto' | 'off', source: 'https://…/' }
// =====================================================================================
const UPDATE_SOURCE_DEFAULT = 'https://raw.githubusercontent.com/shanki412-ship-it/fitbull-app/main/';
function updatesCfg_() { return Object.assign({ mode: 'ask', source: UPDATE_SOURCE_DEFAULT }, cfg_().updates || {}); }

/** '2026-10-13.1' > '2026-10-13' > '2026-10-12.9' */
function versionNewer_(a, b) {
  const pa = String(a || '').split(/[^0-9]+/).filter(Boolean).map(Number), pb = String(b || '').split(/[^0-9]+/).filter(Boolean).map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) { const x = pa[i] || 0, y = pb[i] || 0; if (x !== y) return x > y; }
  return false;
}

/** What's published right now: {version, date, notes:[...]} or null. */
function latestRelease_() {
  const src = updatesCfg_().source.replace(/\/?$/, '/');
  const res = UrlFetchApp.fetch(src + 'version.json?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  try { const j = JSON.parse(res.getContentText()); return j && j.version ? j : null; } catch (e) { return null; }
}

/** For the app: is there a newer version? (checked at most every 6 hours unless force) */
function webCheckUpdate(force) {
  const u = updatesCfg_(), p = PropertiesService.getScriptProperties();
  if (u.mode === 'off' || PropertiesService.getScriptProperties().getProperty('UPDATE_BLOCKED')) return { mode: 'off', current: ENGINE_VERSION, available: false };
  let rel = null;
  const cached = p.getProperty('UPDATE_LATEST'), at = Number(p.getProperty('UPDATE_CHECKED_AT')) || 0;
  if (!force && cached && Date.now() - at < 6 * 3600e3) { try { rel = JSON.parse(cached); } catch (e) {} }
  if (!rel) {
    rel = latestRelease_();
    if (rel) { p.setProperty('UPDATE_LATEST', JSON.stringify(rel)); p.setProperty('UPDATE_CHECKED_AT', String(Date.now())); }
  }
  const available = !!(rel && versionNewer_(rel.version, ENGINE_VERSION));
  return { mode: u.mode, current: ENGINE_VERSION, available, version: rel ? rel.version : '', date: rel ? rel.date || '' : '',
    notes: rel && Array.isArray(rel.notes) ? rel.notes.slice(0, 8) : [], dismissed: p.getProperty('UPDATE_DISMISSED') === (rel && rel.version) };
}
function webDismissUpdate(version) { PropertiesService.getScriptProperties().setProperty('UPDATE_DISMISSED', String(version || '')); return true; }

/** Once a day: email about a new version (mode 'ask'), or install it (mode 'auto'). */
function updateDailyCheck_() {
  const st = webCheckUpdate(true);
  if (!st.available) return;
  const p = PropertiesService.getScriptProperties();
  if (st.mode === 'auto') {
    try { webInstallUpdate(st.version); return; } catch (e) { console.error('Auto-update failed: ' + e.message); }
  }
  if (p.getProperty('UPDATE_EMAILED') === st.version) return;
  const s = getSettings_(), url = appUrl_();
  const notes = st.notes.length ? '<ul>' + st.notes.map(n => '<li>' + String(n).replace(/</g, '&lt;') + '</li>').join('') + '</ul>' : '';
  MailApp.sendEmail({ name: appName_(), to: s['Email'] || Session.getActiveUser().getEmail(), subject: `New in ${appName_()}: an update is ready`,
    htmlBody: `<div style="font-family:Arial,sans-serif;font-size:15px;color:#1B2631"><h2 style="margin:0 0 8px">${appName_()} update (${st.date || st.version})</h2>
      <p>What's new:</p>${notes}
      ${url ? `<p><a href="${url}?update=1" style="background:#2E9E62;color:#fff;padding:12px 18px;border-radius:10px;text-decoration:none;font-weight:bold">Update now</a></p>` : ''}
      <p style="color:#5B6874">Takes about 30 seconds. Your goals, logs and settings stay exactly as they are.</p></div>` });
  p.setProperty('UPDATE_EMAILED', st.version);
}

function scriptApi_(method, path, body) {
  const res = UrlFetchApp.fetch('https://script.googleapis.com/v1/' + path, {
    method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, payload: body ? JSON.stringify(body) : undefined });
  const code = res.getResponseCode(), text = res.getContentText();
  if (code >= 200 && code < 300) return text ? JSON.parse(text) : {};
  let msg = text; try { msg = JSON.parse(text).error.message; } catch (e) {}
  if (/has not been used in project|it is disabled/i.test(msg)) {
    PropertiesService.getScriptProperties().setProperty('UPDATE_BLOCKED', '1');   // stop offering a button that can't work here
    throw new Error("Google doesn't allow one-tap updates for this app without extra Google Cloud setup. Updates will come to you another way – ask the person who shared FitBull with you. (Google said: " + String(msg).slice(0, 120) + ')');
  }
  if (/User has not enabled the Apps Script API/i.test(msg))
    throw new Error('One-time switch needed: open script.google.com/home/usersettings, turn "Google Apps Script API" On, then tap Update again.');
  if (code === 403 && /scope|insufficient|permission/i.test(msg))
    throw new Error('This app needs one extra permission. In Apps Script, make sure appsscript.json has the update permissions, choose "start" and press Run once to approve, then Deploy → Manage deployments → New version, and tap Update again. (Google said: ' + String(msg).slice(0, 160) + ')');
  throw new Error('Update failed (' + code + '): ' + msg);
}

/** Installs the latest version into this project and updates the app (same link). Returns what changed. */
function webInstallUpdate(expectVersion) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('An update is already running – try again in a minute.');
  try {
    const rel = latestRelease_();
    if (!rel) throw new Error("Couldn't reach the update server – check the connection and try again.");
    if (!versionNewer_(rel.version, ENGINE_VERSION)) return { ok: true, upToDate: true, version: ENGINE_VERSION };
    const src = updatesCfg_().source.replace(/\/?$/, '/');
    const get = f => { const r = UrlFetchApp.fetch(src + f + '?v=' + encodeURIComponent(rel.version), { muteHttpExceptions: true }); if (r.getResponseCode() !== 200) throw new Error('Download failed for ' + f + ' (' + r.getResponseCode() + ')'); return r.getContentText(); };
    const code = get('Code.gs'), page = get('Log.html');
    // safety checks before anything is replaced
    const vm = code.match(/const ENGINE_VERSION = '([^']+)'/);
    if (!vm || vm[1] !== rel.version || code.length < 50000 || !/function doGet\(/.test(code)) throw new Error('The downloaded code looks incomplete – nothing was changed. Try again later.');
    if (page.length < 20000 || !/<\/html>/i.test(page)) throw new Error('The downloaded page looks incomplete – nothing was changed. Try again later.');
    const id = ScriptApp.getScriptId();
    const content = scriptApi_('get', 'projects/' + id + '/content');
    const files = (content.files || []).map(f => ({ name: f.name, type: f.type, source: f.source }));
    // the engine is the server file that declares ENGINE_VERSION (whatever it's called); the page is "Log"
    const eng = files.find(f => f.type === 'SERVER_JS' && /const ENGINE_VERSION\s*=/.test(f.source)) || files.find(f => f.type === 'SERVER_JS' && f.name === 'Code');
    const log = files.find(f => f.type === 'HTML' && f.name === 'Log');
    if (!eng || !log) throw new Error("Couldn't find the Code and Log files in this project – nothing was changed.");
    eng.source = code; log.source = page;
    scriptApi_('put', 'projects/' + id + '/content', { files });
    const ver = scriptApi_('post', 'projects/' + id + '/versions', { description: appName_() + ' ' + rel.version });
    const deps = (scriptApi_('get', 'projects/' + id + '/deployments').deployments || [])
      .filter(d => d.deploymentConfig && d.deploymentConfig.versionNumber && (d.entryPoints || []).some(e => e.entryPointType === 'WEB_APP'));
    const p = PropertiesService.getScriptProperties();
    let url = ''; try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
    const known = p.getProperty('UPDATE_DEPLOYMENT_ID');
    const dep = deps.find(d => d.deploymentId === known) || deps.find(d => url && url.indexOf(d.deploymentId) >= 0) || deps[0];
    if (!dep) throw new Error('The new code is saved, but no web-app deployment was found. In Apps Script: Deploy → Manage deployments → New version.');
    scriptApi_('put', 'projects/' + id + '/deployments/' + dep.deploymentId, { deploymentConfig: {
      scriptId: id, versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: appName_() + ' ' + rel.version } });
    p.setProperty('UPDATE_DEPLOYMENT_ID', dep.deploymentId);
    p.setProperty('UPDATE_INSTALLED', rel.version + ' ' + new Date().toISOString());
    p.deleteProperty('UPDATE_LATEST');
    return { ok: true, version: rel.version, notes: rel.notes || [], appUrl: appUrl_() };
  } finally { lock.releaseLock(); }
}
