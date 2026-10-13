/**
 * FitBull – setup wizard (FitBull ▸ Set up FitBull) and the plan builder it uses.
 * Part of the app code – the same for everyone.
 */
// =====================================================================================
// ---------- Setup wizard: FitBull ▸ Set up FitBull (for friends who copy the template Sheet) ----------
// The wizard asks the questions in a form, builds the plan with the same rules as the setup interview,
// saves it (no code editing), builds the tabs, schedules the emails and publishes the phone app.
// =====================================================================================
const WIZ_DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const r50_ = x => Math.round(x / 50) * 50, r5_ = x => Math.round(x / 5) * 5;

/** Exercise library: [name, muscle, kg to add, type] – type main = compound, acc = accessory, core, small. */
const WIZ_LIB = {
  gym: {
    'Full Body A': [['Back squat', 'Quads', 5, 'main'], ['Bench press', 'Chest', 2.5, 'main'], ['Seated cable row', 'Back', 2.5, 'acc'], ['Romanian deadlift', 'Hamstrings', 5, 'acc'], ['Lateral raise', 'Shoulders', 1, 'small'], ['Plank', 'Core', 0, 'core']],
    'Full Body B': [['Trap-bar deadlift', 'Hamstrings', 5, 'main'], ['Seated DB shoulder press', 'Shoulders', 1, 'main'], ['Lat pulldown', 'Back', 2.5, 'acc'], ['Leg press', 'Quads', 5, 'acc'], ['Incline DB press', 'Chest', 1, 'acc'], ['Hanging knee raise', 'Core', 0, 'core']],
    'Full Body C': [['Goblet squat', 'Quads', 2, 'main'], ['Incline DB press', 'Chest', 1, 'main'], ['One-arm DB row', 'Back', 1, 'acc'], ['Hip thrust', 'Glutes', 5, 'acc'], ['Face pull', 'Rear delts', 2.5, 'small'], ['Dead bug', 'Core', 0, 'core']],
    'Upper A': [['Bench press', 'Chest', 2.5, 'main'], ['Barbell row', 'Back', 2.5, 'main'], ['Seated DB shoulder press', 'Shoulders', 1, 'acc'], ['Lat pulldown', 'Back', 2.5, 'acc'], ['DB curl', 'Biceps', 1, 'small'], ['Triceps pushdown', 'Triceps', 2.5, 'small']],
    'Lower A': [['Back squat', 'Quads', 5, 'main'], ['Romanian deadlift', 'Hamstrings', 5, 'main'], ['Leg press', 'Quads', 5, 'acc'], ['Seated leg curl', 'Hamstrings', 2.5, 'acc'], ['Standing calf raise', 'Calves', 2.5, 'small'], ['Plank', 'Core', 0, 'core']],
    'Upper B': [['Incline DB press', 'Chest', 1, 'main'], ['Lat pulldown', 'Back', 2.5, 'main'], ['One-arm DB row', 'Back', 1, 'acc'], ['Lateral raise', 'Shoulders', 1, 'small'], ['Face pull', 'Rear delts', 2.5, 'small'], ['Hanging knee raise', 'Core', 0, 'core']],
    'Lower B': [['Trap-bar deadlift', 'Hamstrings', 5, 'main'], ['Bulgarian split squat', 'Quads', 1, 'acc'], ['Hip thrust', 'Glutes', 5, 'acc'], ['Lying leg curl', 'Hamstrings', 2.5, 'acc'], ['Standing calf raise', 'Calves', 2.5, 'small'], ['Cable crunch', 'Core', 2.5, 'core']],
  },
  dumbbells: {
    'Full Body A': [['Goblet squat', 'Quads', 2, 'main'], ['DB bench press', 'Chest', 1, 'main'], ['One-arm DB row', 'Back', 1, 'acc'], ['DB Romanian deadlift', 'Hamstrings', 2, 'acc'], ['Lateral raise', 'Shoulders', 1, 'small'], ['Plank', 'Core', 0, 'core']],
    'Full Body B': [['Reverse lunge (DB)', 'Quads', 1, 'main'], ['DB shoulder press', 'Shoulders', 1, 'main'], ['Chest-supported DB row', 'Back', 1, 'acc'], ['Single-leg glute bridge', 'Glutes', 0, 'acc'], ['Push-up', 'Chest', 0, 'acc'], ['Dead bug', 'Core', 0, 'core']],
    'Full Body C': [['Bulgarian split squat (DB)', 'Quads', 1, 'main'], ['Incline DB press', 'Chest', 1, 'main'], ['One-arm DB row', 'Back', 1, 'acc'], ['DB Romanian deadlift', 'Hamstrings', 2, 'acc'], ['DB curl', 'Biceps', 1, 'small'], ['Side plank', 'Core', 0, 'core']],
    'Upper A': [['DB bench press', 'Chest', 1, 'main'], ['One-arm DB row', 'Back', 1, 'main'], ['DB shoulder press', 'Shoulders', 1, 'acc'], ['Push-up', 'Chest', 0, 'acc'], ['DB curl', 'Biceps', 1, 'small'], ['DB skull crusher', 'Triceps', 1, 'small']],
    'Lower A': [['Goblet squat', 'Quads', 2, 'main'], ['DB Romanian deadlift', 'Hamstrings', 2, 'main'], ['Reverse lunge (DB)', 'Quads', 1, 'acc'], ['Single-leg glute bridge', 'Glutes', 0, 'acc'], ['Single-leg calf raise', 'Calves', 0, 'small'], ['Plank', 'Core', 0, 'core']],
    'Upper B': [['Incline DB press', 'Chest', 1, 'main'], ['Chest-supported DB row', 'Back', 1, 'main'], ['Lateral raise', 'Shoulders', 1, 'small'], ['Bent-over DB reverse fly', 'Rear delts', 1, 'small'], ['Push-up', 'Chest', 0, 'acc'], ['Dead bug', 'Core', 0, 'core']],
    'Lower B': [['Bulgarian split squat (DB)', 'Quads', 1, 'main'], ['DB Romanian deadlift', 'Hamstrings', 2, 'main'], ['Step-up (DB)', 'Quads', 1, 'acc'], ['Single-leg glute bridge', 'Glutes', 0, 'acc'], ['Single-leg calf raise', 'Calves', 0, 'small'], ['Side plank', 'Core', 0, 'core']],
  },
  bodyweight: {
    'Full Body A': [['Bodyweight squat (3-s lowering)', 'Quads', 0, 'main'], ['Push-up', 'Chest', 0, 'main'], ['Superman', 'Back', 0, 'acc'], ['Single-leg glute bridge', 'Glutes', 0, 'acc'], ['Plank', 'Core', 0, 'core']],
    'Full Body B': [['Bulgarian split squat', 'Quads', 0, 'main'], ['Incline push-up', 'Chest', 0, 'main'], ['Superman', 'Back', 0, 'acc'], ['Single-leg glute bridge', 'Glutes', 0, 'acc'], ['Side plank', 'Core', 0, 'core']],
    'Full Body C': [['Bodyweight squat (3-s lowering)', 'Quads', 0, 'main'], ['Push-up', 'Chest', 0, 'main'], ['Single-leg calf raise', 'Calves', 0, 'small'], ['Single-leg glute bridge', 'Glutes', 0, 'acc'], ['Dead bug', 'Core', 0, 'core']],
  },
};
const WIZ_SWAPS = {
  knee: { 'Back squat': 'Leg press', 'Bulgarian split squat': 'Hip thrust', 'Bulgarian split squat (DB)': 'Single-leg glute bridge', 'Reverse lunge (DB)': 'DB Romanian deadlift', 'Step-up (DB)': 'Single-leg glute bridge', 'Goblet squat': 'Glute bridge (barbell)' },
  back: { 'Romanian deadlift': 'Seated leg curl', 'Trap-bar deadlift': 'Hip thrust', 'Back squat': 'Leg press', 'Barbell row': 'Seated cable row', 'DB Romanian deadlift': 'Single-leg glute bridge' },
  shoulder: { 'Bench press': 'DB bench press', 'Seated DB shoulder press': 'Lateral raise', 'DB shoulder press': 'Lateral raise', 'Overhead press': 'Lateral raise' },
};

/** Builds a complete settings object from the wizard's answers. Returns {config, summary, warnings}. */
function wizBuildConfig_(a) {
  const n = x => Number(x) || 0, warnings = [];
  const sex = a.sex === 'female' ? 'female' : 'male', age = n(a.age), h = n(a.heightCm), w = n(a.weightKg);
  if (age < 18) throw new Error('FitBull is built for adults (18+). For under-18s, train with a qualified coach and a parent or guardian.');
  if (h < 120 || h > 230 || w < 35 || w > 250) throw new Error('Please check height (cm) and weight (kg).');
  const gymDays = WIZ_DAYS.filter(d => (a.gymDays || []).includes(d));
  const caution = !!(a.health && (a.health.heart || a.health.faint || a.health.meds));
  const pregnant = a.women === 'pregnant', postpartum = a.women === 'postpartum' || a.women === 'breastfeeding';
  const noDeficit = !!(a.health && a.health.ed) || pregnant || postpartum;
  // ---- calories (Mifflin–St Jeor) ----
  const bmr = 10 * w + 6.25 * h - 5 * age + (sex === 'male' ? 5 : -161);
  const acts = (a.activities || []).filter(x => x && x.name);
  const sessions = gymDays.length + acts.reduce((s, x) => s + (x.days || []).length, 0) + (a.run && a.run.goal !== 'none' ? n(a.run.perWeek) : 0);
  let factor = a.job === 'physical' ? 1.65 : sessions >= 6 ? 1.55 : sessions >= 4 ? 1.45 : 1.375;
  if (a.job === 'feet') factor = Math.max(factor, 1.5);
  const maint = bmr * factor;
  const goal = noDeficit && /fat|recomp/.test(a.goal) ? 'health' : a.goal;
  const wc = { fat: -0.006 * w, recomp: -0.0025 * w, muscle: a.experience === 'beginner' ? 0.3 : 0.2, strength: 0.1, health: 0, run: 0 }[goal] || 0;
  const weeklyChangeKg = Math.max(-1, Math.round(wc * 20) / 20);
  const floor = Math.max(r50_(bmr), sex === 'female' ? 1200 : 1500);
  const startCalories = Math.max(floor, r50_(maint + weeklyChangeKg * 7700 / 7));
  const bmi = w / Math.pow(h / 100, 2), protW = bmi > 30 ? 25 * Math.pow(h / 100, 2) : w;
  const gkg = { fat: 2.0, recomp: 1.9, muscle: 1.8, strength: 1.8, health: 1.6, run: 1.6 }[goal] || 1.6;
  const proteinTarget = r5_(protW * gkg), proteinMinimum = Math.min(proteinTarget, r5_(protW * 1.5));
  if (noDeficit && /fat|recomp/.test(a.goal)) warnings.push('Calories are set to maintenance (no deficit) because of your health answers.');
  // ---- training plan ----
  const where = ['gym', 'dumbbells', 'bodyweight'].includes(a.where) ? a.where : 'gym';
  const lib = WIZ_LIB[where];
  const nd = Math.min(5, gymDays.length);
  let names = nd <= 1 ? ['Full Body A'] : nd === 2 ? ['Full Body A', 'Full Body B'] : nd === 3 ? ['Full Body A', 'Full Body B', 'Full Body C']
    : nd === 4 ? ['Upper A', 'Lower A', 'Upper B', 'Lower B'] : ['Upper A', 'Lower A', 'Upper B', 'Lower B', 'Full Body A'];
  if (where === 'bodyweight') names = names.map((x, i) => ['Full Body A', 'Full Body B', 'Full Body C'][i % 3]);
  if (gymDays.length > 5) warnings.push('More than 5 gym days – the plan uses 5 and leaves the rest as rest/active days.');
  const minutes = [30, 45, 60, 75].includes(n(a.minutes)) ? n(a.minutes) : 45;
  const maxEx = minutes <= 30 ? 4 : minutes <= 45 ? 5 : 6;
  const inj = String(a.injuries || '').toLowerCase();
  const swaps = Object.assign({}, /knee|acl|menisc/.test(inj) ? WIZ_SWAPS.knee : {}, /back|spine|disc|sciatic/.test(inj) ? WIZ_SWAPS.back : {}, /shoulder|rotator/.test(inj) ? WIZ_SWAPS.shoulder : {});
  const known = {}; Object.values(WIZ_LIB).forEach(l => Object.values(l).forEach(rows => rows.forEach(r => { known[r[0]] = r; })));
  known['Glute bridge (barbell)'] = ['Glute bridge (barbell)', 'Glutes', 5, 'acc'];
  const lvl = a.experience === 'advanced' ? 2 : a.experience === 'intermediate' ? 1 : 0;
  const reps = type => type === 'main' ? (lvl === 0 || caution || pregnant ? ['8–10', 10] : ['6–8', 8]) : type === 'core' ? ['10–15', 15] : type === 'small' ? ['12–15', 15] : ['10–12', 12];
  const sets = type => Math.max(2, Math.min(4, (type === 'main' ? 3 : type === 'core' ? 2 : lvl === 0 ? 2 : 3) + (lvl === 2 && type === 'main' ? 1 : 0) - (caution || pregnant || postpartum ? 1 : 0)));
  const plan = [], uniq = [...new Set(names)];
  uniq.forEach(sess => {
    const seen = new Set(); let order = 0;
    (lib[sess] || lib['Full Body A']).slice(0, maxEx).forEach(r => {
      let ex = swaps[r[0]] ? known[swaps[r[0]]] || [swaps[r[0]], r[1], r[2], r[3]] : r;
      if (seen.has(ex[0])) return; seen.add(ex[0]);
      const type = r[3], [rs, top] = ex[0] === 'Plank' || ex[0] === 'Side plank' ? ['30–45 s', ''] : reps(type);
      plan.push([sess, ++order, ex[0], sets(type), rs, top, type === 'main' ? (lvl ? '2 min' : '90 s') : type === 'core' ? '45 s' : '75 s', ex[2], ex[1]]);
    });
  });
  const schedule = {}; gymDays.slice(0, 5).forEach((d, i) => { schedule[d] = names[i % names.length]; });
  if (!Object.keys(schedule).length) { schedule.Mon = 'Full Body A'; schedule.Thu = 'Full Body B' in lib ? 'Full Body B' : 'Full Body A'; warnings.push('No gym days picked – suggested Mon and Thu (change any time with ✏️ Change my plan).'); if (!plan.some(r => r[0] === 'Full Body B') && lib['Full Body B']) lib['Full Body B'].slice(0, maxEx).forEach((r, i) => plan.push(['Full Body B', i + 1, r[0], sets(r[3]), reps(r[3])[0], reps(r[3])[1], '75 s', r[2], r[1]])); }
  // ---- other activities → sports ----
  const sports = acts.slice(0, 4).map(x => {
    const mins = n(x.minutes) || 45, nm = String(x.name).trim().slice(0, 30);
    const per = /yoga|pilates|stretch/i.test(nm) ? 2.5 : /walk/i.test(nm) ? 2 : 5;
    return { name: nm, days: WIZ_DAYS.filter(d => (x.days || []).includes(d)), when: x.when === 'morning' ? 'morning' : 'evening', extraKcal: Math.round(mins * per / 10) * 10, restNextMorning: false, options: [mins + ' min'] };
  }).filter(s => s.days.length);
  // ---- running ----
  let running;
  if (a.run && a.run.goal && a.run.goal !== 'none') {
    running = { goal: ['general', '5k', '10k', 'half', 'marathon'].includes(a.run.goal) ? a.run.goal : 'general', runsPerWeek: Math.max(1, Math.min(6, n(a.run.perWeek) || 2)),
      days: WIZ_DAYS.filter(d => (a.run.days || []).includes(d)), weeklyKm: Math.max(n(a.run.weeklyKm) || 6, 2), longestRunKm: Math.max(2, n(a.run.longestKm) || Math.round((n(a.run.weeklyKm) || 6) / 2)),
      recent5kMin: n(a.run.fiveK) || 0, raceDate: /^\d{4}-\d{2}-\d{2}$/.test(String(a.run.raceDate || '')) ? a.run.raceDate : '', eatBack: 0.5 };
  }
  const goalText = { fat: 'Lose fat while keeping muscle', recomp: 'Recomposition: build muscle and lose fat', muscle: 'Build muscle (lean gain)', strength: 'Get stronger', health: 'Stay fit and healthy', run: 'Get fitter for running' }[a.goal] || 'Get fitter';
  const flags = [caution ? 'Has a heart, fainting or medication flag – keep it conservative and get a doctor\'s OK.' : '',
    pregnant ? 'Pregnant – no deficit, no new heavy lifting or running, avoid lying flat on the back after the first trimester; midwife/doctor OK first.' : '',
    postpartum ? 'Had a baby in the last 12 months / breastfeeding – no deficit, rebuild gradually, pelvic-floor friendly; doctor/physio OK first.' : '',
    a.health && a.health.ed ? 'History of an eating disorder – no calorie deficit; focus on training and regular meals.' : ''].filter(Boolean).join(' ');
  const wake = Math.max(4, Math.min(11, n(a.wakeHour) || 7)), bed = Math.max(19, Math.min(24, n(a.bedHour) || 23));
  const config = {
    appName: 'FitBull', appTagline: 'Strong habits. Smarter training.', updates: { mode: 'off' }, geminiKey: '',
    profile: { name: String(a.name || '').trim().slice(0, 40) || 'Friend', age, sex, heightCm: h, weightKg: w,
      experience: { beginner: 'Beginner – new to lifting or under 6 months', intermediate: 'Intermediate – lifting consistently 6 months to 2 years', advanced: 'Advanced – lifting consistently 2+ years' }[a.experience] || 'Beginner',
      goal: goalText + (a.target ? ' – ' + String(a.target).slice(0, 80) : ''),
      diet: [{ nonveg: 'Non-vegetarian', veg: 'Vegetarian (dairy ok)', egg: 'Vegetarian + eggs', vegan: 'Vegan', pesc: 'Pescatarian', halal: 'Halal' }[a.diet] || 'Mixed', a.cuisine ? a.cuisine + ' food' : '', a.powder ? 'uses protein powder' : ''].filter(Boolean).join('; '),
      foods: String(a.foods || '').slice(0, 400), avoid: String(a.avoid || '').slice(0, 200), injuries: String(a.injuries || 'None').slice(0, 200) || 'None',
      lifestyle: [({ desk: 'Desk job', feet: 'On my feet at work', physical: 'Physical job', shifts: 'Shift work' })[a.job] || '', gymDays.length ? 'Gym ' + gymDays.join('/') + ', ' + minutes + ' min' : '',
        acts.map(x => x.name + ' ' + (x.days || []).join('/')).join(', '), 'wakes ~' + wake + ':00, bed ~' + (bed === 24 ? '0' : bed) + ':00'].filter(Boolean).join('. '),
      notes: flags },
    timezone: String(a.timezone || '') || Session.getScriptTimeZone(),
    emails: { dailyHour: wake, dailyMinute: 0, weeklyDay: 'Sun', weeklyHour: 19 },
    schedule, restDayNotes: {}, sports,
    nutrition: { startCalories, weeklyChangeKg, proteinTarget, proteinMinimum, calorieFloor: floor, waistGoalCm: 0, goalWeightKg: n(a.goalWeightKg) || 0 },
    steps: { start: Math.max(4000, n(a.steps) || 5000), increase: 1000, everyWeeks: 2, max: Math.max(8000, Math.min(12000, (n(a.steps) || 5000) + 4000)) },
    training: { deloadEvery: lvl === 0 ? 8 : 6, sessionMinutes: minutes, maxSetsPerSession: Math.min(30, Math.round(minutes / 2)), priorityMuscles: goal === 'muscle' ? ['Chest', 'Back'] : goal === 'fat' || goal === 'recomp' ? ['Glutes', 'Back'] : [], weeklySetCap: {}, mainLifts: plan.filter(r => r[1] === 1).map(r => r[2]) },
    plan, alternates: {},
    agent: { proactive: true, eveningHour: Math.max(18, Math.min(23, bed - 2)) },
    home: { equipment: (a.homeEquip && a.homeEquip.length ? a.homeEquip : ['bodyweight']).filter(x => ['bodyweight', 'dumbbells', 'band'].includes(x)) },
    telegram: { botToken: '', webhookUrl: '' },
    meals: Array.isArray(a.meals) && a.meals.length ? a.meals : wizDefaultMeals_(a.diet),
  };
  if (running) config.running = running;
  const days = WIZ_DAYS.map(d => schedule[d] ? d + ': ' + schedule[d] : sports.find(s => s.days.includes(d)) ? d + ': ' + sports.find(s => s.days.includes(d)).name : running && running.days.includes(d) ? d + ': run' : d + ': rest');
  const summary = { name: config.profile.name, calories: startCalories, protein: proteinTarget, proteinMin: proteinMinimum, weeklyChangeKg, floor, steps: config.steps.start,
    week: days, sessions: uniq.map(s => ({ name: s, exercises: plan.filter(r => r[0] === s).map(r => r[2] + ' ' + r[3] + '×' + r[4]) })),
    running: running ? `${running.runsPerWeek} runs a week (${running.goal}), starting ~${running.weeklyKm} km/week` : '', warnings,
    doctor: caution || pregnant || postpartum };
  return { config, summary };
}

/** Starter one-tap meals by diet (the AI personalises them from your foods when it can). */
function wizDefaultMeals_(diet) {
  const base = [['Greek yogurt + berries', 220, 20, 4, 14], ['Oats with milk + banana', 380, 14, 8, 22], ['Protein shake', 130, 24, 2, 3]];
  const byDiet = {
    vegan: [['Tofu scramble + toast', 380, 24, 16, 4], ['Lentil dal + rice', 520, 20, 10, 5], ['Chickpea salad', 420, 17, 15, 7], ['Soya chunk curry', 330, 30, 9, 6], ['Soy yogurt + nuts', 260, 12, 14, 9], ['Pea protein shake', 140, 25, 2, 1]],
    veg: [['Paneer bhurji + roti', 520, 28, 30, 5], ['Dal + rice + curd', 600, 22, 14, 9], ['Greek yogurt + berries', 220, 20, 4, 14], ['Chickpea curry + rice', 560, 19, 13, 8], ['Cheese toastie', 420, 20, 20, 4], ['Protein shake', 130, 24, 2, 3]],
    egg: [['3-egg omelette + toast', 380, 24, 22, 3], ['Dal + rice + curd', 600, 22, 14, 9], ['Greek yogurt + berries', 220, 20, 4, 14], ['Paneer wrap', 480, 26, 24, 4], ['2 boiled eggs', 150, 12, 10, 1], ['Protein shake', 130, 24, 2, 3]],
  };
  const nonveg = [['Chicken + rice + veg', 550, 40, 14, 3], ['3-egg omelette + toast', 380, 24, 22, 3], ['Tuna sandwich', 400, 30, 12, 4], ['Salmon + potatoes', 560, 34, 24, 3], ['2 boiled eggs', 150, 12, 10, 1]];
  return (byDiet[diet === 'vegan' ? 'vegan' : diet === 'veg' ? 'veg' : diet === 'egg' ? 'egg' : ''] || nonveg.concat(base)).slice(0, 8);
}

/** Ask the AI for one-tap meals from what they actually eat (falls back to the starter list). */
function wizAiMeals_(a) {
  try {
    const text = callGemini_(`Create 8 one-tap meals/snacks for a food tracker for someone whose diet is "${a.diet || 'mixed'}", cuisine "${a.cuisine || 'any'}".
Foods they often eat: ${String(a.foods || 'not given').slice(0, 400)}. Never include: ${String(a.avoid || 'nothing').slice(0, 200)}.
Realistic home portions. At least 3 with 20 g+ protein. Short names (max 28 characters).
Reply with JSON only: [["Name", kcal, protein_g, fat_g, sugar_g], ...]`, true);
    const arr = JSON.parse((text || '').match(/\[[\s\S]*\]/)[0]);
    const ok = arr.filter(m => Array.isArray(m) && m.length >= 3 && m[0] && [1, 2].every(i => typeof m[i] === 'number' && m[i] >= 0))
      .map(m => [String(m[0]).slice(0, 28), Math.round(m[1]), Math.round(m[2]), Math.round(Number(m[3]) || 0), Math.round(Number(m[4]) || 0)])
      .filter(m => m[1] > 0 && m[1] < 2000).slice(0, 10);
    return ok.length >= 4 ? ok : null;
  } catch (e) { return null; }
}

// ---- menu actions ----
function wizOpen() {
  const html = HtmlService.createHtmlOutputFromFile('Wizard').setWidth(620).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, hasConfig_() ? 'Change your FitBull plan' : 'Set up FitBull');
}
function wizShowLink() {
  const url = appUrl_();
  const html = HtmlService.createHtmlOutput(`<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.5">${url
    ? `<p>Your FitBull app:</p><p><a href="${url}" target="_blank" style="background:#2E9E62;color:#fff;padding:10px 16px;border-radius:10px;text-decoration:none;font-weight:bold">Open my app</a></p><p style="color:#5B6874;font-size:13px">On your phone: open the welcome email and tap the button, then Add to Home screen.<br>${url}</p>`
    : '<p>The app isn\'t published yet – run FitBull ▸ Set up FitBull and finish the last step.</p>'}</div>`).setWidth(460).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'FitBull');
}
/** A friendly first screen in a fresh copy (runs when the Sheet opens). */
function welcomeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheets()[0];
  if (sh.getLastRow() > 0 || ss.getSheets().length > 1) return;
  sh.setName('Start here');
  sh.getRange('B2').setValue('👋 Welcome to FitBull – Strong habits. Smarter training.').setFontSize(18).setFontWeight('bold');
  sh.getRange('B4:B10').setValues([['0. Use an up-to-date Chrome, Edge or Safari on a computer (not a phone, not an old work laptop).'], ['1. On a computer, click the FitBull menu at the top (next to Help).'], ['2. Choose ▶ Set up FitBull.'],
    ['3. Google asks for permission → Continue → choose your account → Advanced → Go to FitBull → Allow.'], ['4. Answer the questions (about 10 minutes).'],
    ['5. Open the email on your phone, tap Open my app, then Add to Home screen.'], ['Your data stays in this Sheet, in your own Google account.']]).setFontSize(13);
  sh.setColumnWidth(2, 900);
}

// ---- called from the wizard page ----
function webWizState() {
  const p = PropertiesService.getScriptProperties();
  let saved = null; try { saved = JSON.parse(p.getProperty('WIZ_ANSWERS') || 'null'); } catch (e) {}
  return { configured: hasConfig_(), hasKey: !!p.getProperty('GEMINI_API_KEY'), answers: saved, appUrl: hasConfig_() ? appUrl_() : '',
    email: Session.getActiveUser().getEmail() || '', scriptTz: Session.getScriptTimeZone() };
}
function webWizSaveKey(key) {
  // tidy whatever was pasted: spaces, line breaks, quotes, invisible characters, "API key:" labels
  key = String(key || '').replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/^\s*(api\s*key|key)\s*[:=]\s*/i, '').replace(/[\s"'`]/g, '');
  if (key.length < 30) throw new Error('That looks too short – copy the whole key from aistudio.google.com/apikey (the Copy button next to it).');
  if (!/^[A-Za-z0-9._-]+$/.test(key)) throw new Error('That has characters a key never has – copy it again with the Copy button on aistudio.google.com/apikey.');
  const p = PropertiesService.getScriptProperties(), old = p.getProperty('GEMINI_API_KEY');
  p.setProperty('GEMINI_API_KEY', key); p.deleteProperty('GEMINI_MODEL'); p.deleteProperty('GEMINI_SMART_MODEL');
  // quick check: one "list models" call (well under a second) instead of a full AI reply
  let ok = false;
  try {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200',
      { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const have = ((JSON.parse(res.getContentText()).models) || []).map(m => String(m.name).replace(/^models\//, ''));
      const pick = l => l.find(m => have.indexOf(m) >= 0);
      if (pick(CONFIG.MODELS)) p.setProperty('GEMINI_MODEL', pick(CONFIG.MODELS));
      if (pick(CONFIG.SMART_MODELS)) p.setProperty('GEMINI_SMART_MODEL', pick(CONFIG.SMART_MODELS));
      ok = true;
    } else {
      let msg = res.getContentText(); try { msg = JSON.parse(msg).error.message; } catch (e) {}
      lastGeminiError_ = res.getResponseCode() + ' – ' + msg;
    }
  } catch (e) { lastGeminiError_ = e.message; }
  if (!ok && !/API key|PERMISSION|400|401|403/i.test(lastGeminiError_)) ok = !!callGemini_('Reply with JSON only: {"ok": true}', true);  // network hiccup: fall back to a real call
  if (!ok) { if (old) p.setProperty('GEMINI_API_KEY', old); else p.deleteProperty('GEMINI_API_KEY'); throw new Error('The key didn\'t work: ' + (lastGeminiError_ || 'no reply') + '. Copy it again from aistudio.google.com.'); }
  return { ok: true, model: p.getProperty('GEMINI_MODEL') || '' };
}
function webWizBuild(answers) {
  const a = Object.assign({}, answers);
  if (!a.meals && PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) a.meals = wizAiMeals_(a) || undefined;
  const out = wizBuildConfig_(a);
  out.meals = out.config.meals;
  return out;
}
/** Saves the plan, builds everything and publishes the app. Returns {appUrl} or {manual:true} (then the person clicks Deploy). */
function webWizFinish(answers, meals) {
  const a = Object.assign({}, answers, meals ? { meals } : {});
  const { config } = wizBuildConfig_(a);
  const p = PropertiesService.getScriptProperties();
  const before = storedConfig_();
  saveConfig_(config);
  const problems = validateConfig_();
  if (problems.length) { if (before) saveConfig_(before); else p.deleteProperty('CFG_PARTS'); throw new Error('Plan check failed: ' + problems.join('; ')); }
  p.setProperty('WIZ_ANSWERS', JSON.stringify(Object.assign({}, answers, { meals: config.meals })));
  setup();
  // plan changed later: the Settings tab's targets follow the new answers (your logs are kept)
  if (before) { const def = settingsDefaults_(); Object.keys(def).forEach(k => setSetting_(k, def[k])); }
  createTriggers();
  p.setProperty('TRIGGERS_CHECKED', key_(new Date()));
  // publish (or keep) the phone app
  let url = appUrl_();
  if (!url || !p.getProperty('WIZ_DEPLOYED')) {
    const r = wizDeploy_();
    if (r.manual) return { manual: true, scriptUrl: 'https://script.google.com/d/' + ScriptApp.getScriptId() + '/edit' };
    url = r.url;
  }
  const aiOk = !!p.getProperty('GEMINI_API_KEY');
  if (!p.getProperty('SETUP_DONE')) { p.setProperty('SETUP_DONE', new Date().toISOString()); welcomeEmail_(url, aiOk, ''); }
  return { ok: true, appUrl: url, email: getSettings_()['Email'] || '' };
}
function webWizRetryDeploy() { const r = wizDeploy_(); if (r.manual) return webWizManualDone(); const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SETUP_DONE')) { p.setProperty('SETUP_DONE', new Date().toISOString()); welcomeEmail_(r.url, !!p.getProperty('GEMINI_API_KEY'), ''); }
  return { ok: true, appUrl: r.url, email: getSettings_()['Email'] || '' }; }

/** Publishes the phone app through the Apps Script API (needs the API switched on in the user's settings). */
function wizDeploy_() {
  const id = ScriptApp.getScriptId(), p = PropertiesService.getScriptProperties();
  try {
    const existing = (scriptApi_('get', 'projects/' + id + '/deployments').deployments || [])
      .filter(d => d.deploymentConfig && d.deploymentConfig.versionNumber && (d.entryPoints || []).some(e => e.entryPointType === 'WEB_APP'));
    const ver = scriptApi_('post', 'projects/' + id + '/versions', { description: appName_() + ' ' + ENGINE_VERSION + ' (setup)' });
    let dep;
    if (existing.length) dep = scriptApi_('put', 'projects/' + id + '/deployments/' + existing[0].deploymentId, { deploymentConfig: { scriptId: id, versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: appName_() } });
    else dep = scriptApi_('post', 'projects/' + id + '/deployments', { versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: appName_() });
    const web = (dep.entryPoints || []).find(e => e.entryPointType === 'WEB_APP');
    const url = (web && web.webApp && web.webApp.url) || ('https://script.google.com/macros/s/' + dep.deploymentId + '/exec');
    p.setProperty('APP_URL', url); p.setProperty('UPDATE_DEPLOYMENT_ID', dep.deploymentId); p.setProperty('WIZ_DEPLOYED', '1');
    return { url };
  } catch (e) {
    // Google only allows this when the person has their own Cloud project – almost nobody does,
    // so the wizard shows the 4-click Deploy steps instead.
    console.log('Publishing via the API not available: ' + e.message);
    return { manual: true };
  }
}

/** After the person clicked Deploy themselves: pick up the app link and send the welcome email. */
function webWizManualDone(pasted) {
  const p = PropertiesService.getScriptProperties();
  let url = '';
  if (pasted) { url = cleanAppLink_(pasted); if (!url) throw new Error('That isn\'t an app link – copy the "Web app URL" (it starts with https://script.google.com/macros/s/).'); }
  else { try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {} }
  if (!url) return { notYet: true };
  p.setProperty('APP_URL', url); p.setProperty('WIZ_DEPLOYED', '1');
  if (!p.getProperty('SETUP_DONE')) { p.setProperty('SETUP_DONE', new Date().toISOString()); welcomeEmail_(url, !!p.getProperty('GEMINI_API_KEY'), ''); }
  return { ok: true, appUrl: url, email: getSettings_()['Email'] || '' };
}
/** Accepts a pasted web-app link (…/macros/s/<id>/dev or /exec); returns it tidied, or '' if it isn't one. */
function cleanAppLink_(u) {
  u = String(u || '').replace(/[\s"'<>]/g, '');
  const m = u.match(/^https:\/\/script\.google\.com\/(?:a\/[^/]+\/)?macros\/s\/[A-Za-z0-9_-]{20,}\/(?:dev|exec)/);
  return m ? m[0] : '';
}
/** Menu: paste your own app link (Deploy ▸ Test deployments ▸ Web app URL) so emails and "Open my app" use it. */
function wizSetLink() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Your FitBull app link',
    'In Extensions ▸ Apps Script click Deploy ▸ Test deployments, copy the Web app URL (ends in /dev) and paste it here.\n' +
    'That link always runs the newest version, so updates reach you without redeploying.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = cleanAppLink_(r.getResponseText());
  if (!url) { ui.alert('That isn\'t an app link – it should start with https://script.google.com/macros/s/'); return; }
  PropertiesService.getScriptProperties().setProperty('APP_URL', url);
  wizShowLink();
}
function setSetting_(name, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS); if (!sh) return;
  const v = sh.getDataRange().getValues(); for (let i = 1; i < v.length; i++) if (v[i][0] === name) { if (value !== '' && value !== undefined) sh.getRange(i + 1, 2).setValue(value); return; }
}
