const WORKBOOK_PATH = './Race_Shed_Fantasy_V7.xlsx';

const DRIVER_NAMES = {
  "1":"Ross Chastain","2":"Austin Cindric","3":"Austin Dillon","4":"Noah Gragson","5":"Kyle Larson",
  "6":"Brad Keselowski","7":"Justin Haley","8":"Kyle Busch","9":"Chase Elliott","10":"Ty Dillon",
  "11":"Denny Hamlin","12":"Ryan Blaney","16":"AJ Allmendinger","17":"Chris Buescher","19":"Chase Briscoe",
  "20":"Christopher Bell","21":"Josh Berry","22":"Joey Logano","23":"Bubba Wallace","24":"William Byron",
  "34":"Todd Gilliland","35":"Riley Herbst","38":"Zane Smith","41":"Cole Custer","42":"John H. Nemechek",
  "43":"Erik Jones","45":"Tyler Reddick","47":"Ricky Stenhouse Jr.","48":"Alex Bowman","51":"Cody Ware",
  "54":"Ty Gibbs","60":"Ryan Preece","71":"Michael McDowell","77":"Carson Hocevar","88":"Shane van Gisbergen",
  "99":"Daniel Suárez"
};

const TRACK_TYPES = {
  1:"Superspeedway",2:"Intermediate",3:"Road Course",4:"Short Track",5:"Intermediate",6:"Short Track",
  7:"Short Track",8:"Road Course",9:"Short Track",10:"Superspeedway",11:"Intermediate",12:"Intermediate",
  13:"All-Star / Intermediate",14:"Intermediate",15:"Intermediate",16:"Road Course",17:"Road Course",
  18:"Short Track",19:"Street Course",20:"Road Course",21:"Intermediate",22:"Superspeedway",
  23:"Short Track",24:"Intermediate",25:"Short Track",26:"Road Course",27:"Intermediate",
  28:"Short Track",29:"Road Course",30:"Intermediate",31:"Short Track",32:"Superspeedway",
  33:"Road Course",34:"Short Track",35:"Intermediate",36:"Championship / Short Track"
};

const COLOR_MAP = {
  "5":["#f6c544","#111"],"20":["#ffe16a","#111"],"23":["#58a8ff","#08111a"],"11":["#e10600","#fff"],
  "12":["#d9dde4","#08111a"],"9":["#25d75f","#08111a"],"24":["#ffd94a","#08111a"],"45":["#66e0d8","#08111a"],
  "17":["#2f8bff","#fff"],"19":["#f35b4f","#fff"],"22":["#ffec8a","#111"],"1":["#ff6b6b","#fff"]
};

function el(id){ return document.getElementById(id); }
function asNum(v){ const n = Number(v); return Number.isFinite(n) ? n : null; }
function normCar(v){ return String(v ?? '').replace(/\.0$/,'').trim(); }

async function loadWorkbook(){
  const res = await fetch(WORKBOOK_PATH, { cache: 'no-store' });
  if(!res.ok) throw new Error(`Could not load workbook: ${res.status}`);
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, { type: 'array' });
}

function sheetToRows(wb, name){
  const ws = wb.Sheets[name];
  if(!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}

function latestRaceWithResults(results){
  const nums = [...new Set(results.filter(r => asNum(r.finish)).map(r => asNum(r.raceNumber)))].filter(Boolean);
  return nums.length ? Math.max(...nums) : 0;
}

function raceScore(player, raceNo, picks, pointsByRaceCar){
  const row = picks.find(p => p.player === player && p.raceNumber === raceNo);
  if(!row) return 0;
  return [row.pick1,row.pick2,row.pick3].reduce((sum, car) => sum + (pointsByRaceCar[`${raceNo}-${normCar(car)}`] || 0), 0);
}

function badge(car){
  if(!car) return '';
  const carStr = normCar(car);
  const name = DRIVER_NAMES[carStr] || 'Driver TBD';
  const colors = COLOR_MAP[carStr] || ['#dfe6f1','#08111a'];
  return `<span class="badge"><span class="num" style="background:${colors[0]};color:${colors[1]}">#${carStr}</span><span>${name}</span></span>`;
}

function renderError(err){
  document.body.innerHTML += `<div class="shell"><div class="error-box"><strong>Workbook load failed.</strong><div style="margin-top:8px">${String(err.message || err)}</div></div></div>`;
}

function getFeaturedRaceNumber(schedule, results){
  const latest = latestRaceWithResults(results);
  if(latest > 0) {
    const next = latest + 1;
    if(schedule.some(r => r.raceNumber === next)) return next;
  }
  return 5;
}

function renderDashboard(data){
  const { players, schedule, picks, results, featuredRaceNo, featuredPicks, featuredPot, pointsByRaceCar, standings, latestCompleted } = data;
  const featuredRace = schedule.find(r => r.raceNumber === featuredRaceNo) || schedule[0] || {};
  const completedRaceNos = [...new Set(results.map(r => r.raceNumber))].sort((a,b)=>a-b);

  el('workbookStatus').textContent = 'Connected';
  el('featuredRaceName').textContent = featuredRace.raceName || 'Race Shed Fantasy';
  el('featuredMeta').textContent = [featuredRace.date, featuredRace.track].filter(Boolean).join(' • ');

  el('heroPills').innerHTML = [
    `<span class="pill">$5 / race</span>`,
    `<span class="pill">${players.length} players in league</span>`,
    `<span class="pill">$${featuredPot} featured race pot</span>`,
    `<span class="pill">${completedRaceNos.length} completed league races</span>`
  ].join('');

  el('picksSubmittedCount').textContent = `${featuredPicks.length} / ${players.length}`;

  el('trackTitle').textContent = featuredRace.raceName || 'Track details';
  el('trackBadge').textContent = (featuredRace.track || 'AUTO').toString().split(' ')[0].slice(0,4).toUpperCase();
  const trackDetails = [
    ['Track', featuredRace.track || '-'],
    ['Date', featuredRace.date || '-'],
    ['Track type', featuredRace.trackType || '-'],
    ['Featured pot', `$${featuredPot}`],
    ['Workbook', 'Race_Shed_Fantasy_V7.xlsx'],
    ['Sync mode', 'Auto from Excel']
  ];
  el('trackGrid').innerHTML = trackDetails.map(([label, value]) => `<div class="track-item"><div class="label">${label}</div><div class="value">${value}</div></div>`).join('');
  el('trackNote').textContent = 'Update the workbook in GitHub, commit the file, and Vercel refreshes this board. No league-data.js file needed in V14.';

  el('statsGrid').innerHTML = [
    {label:'League leader', value: standings[0]?.player || '-'},
    {label:'Leader points', value: String(standings[0]?.seasonPoints || 0)},
    {label:'Featured race pot', value: `$${featuredPot}`},
    {label:'Latest winning score', value: latestCompleted ? String(Math.max(...players.map(p => raceScore(p, latestCompleted, picks, pointsByRaceCar)), 0)) : '-'}
  ].map(s => `<div class="panel stat"><div class="label">${s.label}</div><div class="value">${s.value}</div></div>`).join('');

  const picker = el('racePicker');
  picker.innerHTML = '';
  schedule.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r.raceNumber;
    opt.textContent = `${r.raceNumber}. ${r.raceName}`;
    if(r.raceNumber === featuredRaceNo) opt.selected = true;
    picker.appendChild(opt);
  });

  function renderRace(raceNo){
    const n = Number(raceNo);
    const raceRows = picks.filter(p => p.raceNumber === n && [p.pick1,p.pick2,p.pick3].some(Boolean));
    el('picksSummary').textContent = raceRows.length ? `${raceRows.length} players currently on the board for this race.` : 'No picks entered yet for this race.';
    el('picksBoard').innerHTML = raceRows.length ? `<table>
      <thead><tr><th>Player</th><th>Picks</th><th>Total</th></tr></thead>
      <tbody>
      ${raceRows.map(r => {
        const total = raceScore(r.player, n, picks, pointsByRaceCar);
        return `<tr>
          <td class="player-cell">${r.player}</td>
          <td><div class="pick-badges">${badge(r.pick1)}${badge(r.pick2)}${badge(r.pick3)}</div></td>
          <td>${total || '-'}</td>
        </tr>`;
      }).join('')}
      </tbody>
    </table>` : `<div class="empty">Awaiting picks.</div>`;

    const raceResults = results.filter(r => r.raceNumber === n).sort((a,b)=>a.finish-b.finish).slice(0,12);
    el('resultsTitle').textContent = raceResults.length ? `${raceResults[0].raceName} results` : 'Race results';
    el('raceResults').innerHTML = raceResults.length ? `<table>
      <thead><tr><th>Fin</th><th>Car</th><th>Driver</th><th>Pts</th></tr></thead>
      <tbody>${raceResults.map(r => `<tr><td>${r.finish}</td><td>#${r.carNumber}</td><td>${DRIVER_NAMES[r.carNumber] || '-'}</td><td>${r.finishPts}</td></tr>`).join('')}</tbody>
    </table>` : `<div class="empty">No race results loaded yet.</div>`;
  }

  picker.addEventListener('change', e => renderRace(e.target.value));
  renderRace(featuredRaceNo);

  el('standingsBoard').innerHTML = `<table>
    <thead><tr><th>Rank</th><th>Player</th><th>Season Pts</th><th>Wins</th><th>Best Week</th></tr></thead>
    <tbody>${standings.map((s,i)=> `<tr><td class="rank">${i+1}</td><td class="player-cell">${s.player}</td><td>${s.seasonPoints}</td><td>${s.wins}</td><td>${s.bestWeek}</td></tr>`).join('')}</tbody>
  </table>`;

  const latestDriverPts = [];
  if(latestCompleted){
    const seen = new Set();
    results.filter(r => r.raceNumber === latestCompleted).sort((a,b)=> a.finish - b.finish).forEach(r => {
      if(seen.has(r.carNumber)) return;
      seen.add(r.carNumber);
      latestDriverPts.push({ car:r.carNumber, driver:DRIVER_NAMES[r.carNumber] || 'Driver TBD', pts:r.finishPts });
    });
  }
  el('latestScoring').innerHTML = latestDriverPts.length ? `<div class="latest-list">${
    latestDriverPts.slice(0,10).map(r => `<div class="latest-item"><div class="driver-title">#${r.car} ${r.driver}</div><div class="pts">${r.pts} pts</div></div>`).join('')
  }</div>` : `<div class="empty">Latest scoring will appear once race points are loaded.</div>`;

  if(latestCompleted){
    const podium = players.map(player => ({ player, score: raceScore(player, latestCompleted, picks, pointsByRaceCar) }))
      .sort((a,b)=> b.score - a.score || a.player.localeCompare(b.player))
      .slice(0,3);
    el('podiumBoard').innerHTML = `<div class="podium">${
      podium.map((p,i) => `<div class="podium-card ${i===0?'first':''}"><div class="spot">${['1st','2nd','3rd'][i]}</div><div class="name">${p.player}</div><div class="score">${p.score} pts</div></div>`).join('')
    }</div>`;
  } else {
    el('podiumBoard').innerHTML = '';
  }

  el('scheduleBoard').innerHTML = `<div class="schedule-list">${
    schedule.map(r => {
      const cls = completedRaceNos.includes(r.raceNumber) ? 'completed' : 'pending';
      return `<div class="schedule-item ${cls}">
        <div class="schedule-num">${r.raceNumber}</div>
        <div><div class="schedule-name">${r.raceName}</div><div class="schedule-meta">${[r.date, r.track].filter(Boolean).join(' • ')}</div></div>
        <div class="schedule-type">${r.trackType || '-'}</div>
      </div>`;
    }).join('')
  }</div>`;
}

async function init(){
  try{
    const wb = await loadWorkbook();
    const scheduleRows = sheetToRows(wb, 'Schedule').map(r => ({
      raceNumber: asNum(r['Race #']),
      raceName: r['Race'],
      date: r['Date'],
      track: r['Track'],
      tv: r['TV'],
      trackType: r['Track Type'] || TRACK_TYPES[asNum(r['Race #'])] || ''
    })).filter(r => r.raceNumber);

    const picksRows = sheetToRows(wb, 'WeeklyPicks').map(r => ({
      raceNumber: asNum(r['Race #']),
      raceName: r['Race'],
      player: r['Player'],
      pick1: r['Pick 1'],
      pick2: r['Pick 2'],
      pick3: r['Pick 3'],
      weeklyTotal: asNum(r['Weekly Total']) || 0
    })).filter(r => r.raceNumber && r.player);

    const resultsRows = sheetToRows(wb, 'RaceResults').map(r => ({
      raceNumber: asNum(r['Race #']),
      raceName: r['Race'],
      finish: asNum(r['Finish']),
      carNumber: normCar(r['Car #']),
      finishPts: asNum(r['Finish Pts']) || 0
    })).filter(r => r.raceNumber && r.finish && r.carNumber);

    const playerRows = sheetToRows(wb, 'Players').map(r => String(r['Player'] || '').trim()).filter(Boolean);

    const featuredRaceNo = getFeaturedRaceNumber(scheduleRows, resultsRows);
    const featuredPicks = picksRows.filter(p => p.raceNumber === featuredRaceNo && [p.pick1,p.pick2,p.pick3].some(Boolean));
    const featuredPot = featuredPicks.length * 5;

    const pointsByRaceCar = {};
    resultsRows.forEach(r => { pointsByRaceCar[`${r.raceNumber}-${r.carNumber}`] = r.finishPts; });

    const latestCompleted = latestRaceWithResults(resultsRows);

    const standings = playerRows.map(player => {
      const playerRowsFiltered = picksRows.filter(p => p.player === player);
      let seasonPoints = 0, bestWeek = 0, wins = 0;
      playerRowsFiltered.forEach(row => {
        const total = raceScore(player, row.raceNumber, picksRows, pointsByRaceCar);
        seasonPoints += total;
        if(total > bestWeek) bestWeek = total;
      });
      return { player, seasonPoints, bestWeek, wins };
    });

    const completedRaceNos = [...new Set(resultsRows.map(r => r.raceNumber))];
    completedRaceNos.forEach(raceNo => {
      const scores = playerRows.map(player => ({ player, score: raceScore(player, raceNo, picksRows, pointsByRaceCar) }));
      const top = Math.max(0, ...scores.map(s => s.score));
      scores.filter(s => s.score === top && top > 0).forEach(s => {
        const found = standings.find(x => x.player === s.player);
        if(found) found.wins += 1;
      });
    });

    standings.sort((a,b)=> b.seasonPoints - a.seasonPoints || b.wins - a.wins || a.player.localeCompare(b.player));

    renderDashboard({
      players: playerRows,
      schedule: scheduleRows,
      picks: picksRows,
      results: resultsRows,
      featuredRaceNo,
      featuredPicks,
      featuredPot,
      pointsByRaceCar,
      standings,
      latestCompleted
    });
  } catch(err){
    el('workbookStatus').textContent = 'Load failed';
    renderError(err);
  }
}

init();
