/* ============================================================================
   LIVE POWER BI CONNECTION MODULE
   ----------------------------------------------------------------------------
   Replaces the embedded-SEED / parseWorkbook() data layer in
   South_Region_Sales_Command_Center_PBI.html with real-time DAX queries
   against the PUBLISHED NGD_Full_Dataset, via MSAL.js + the Power BI REST
   "Execute Queries" endpoint.

   HOW TO WIRE THIS INTO THE EXISTING HTML FILE:
   1. Fill in the four CONFIG values just below.
   2. Add this script tag in <head>, BEFORE this file:
        <script src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js"></script>
   3. Add this file itself as a <script src="live-connect.js"></script>,
      placed AFTER the main dashboard <script> block (it needs MONTHS,
      normStage, computeSeasonal, boot, DATA, C etc. to already exist).
   4. In the main dashboard script, find the block at the very bottom that
      reads:
            try{const bin=atob(SEED); ... parseWorkbook(wb); boot();}
      and replace it with:
            initApp();
      NOT loadLive() directly -- initApp() checks for a cached MSAL session
      first. Calling loadLive() (and its loginPopup() call inside getToken())
      straight from page load with no cached session gets the popup silently
      blocked by the browser (empty_window_error / popup_window_error), since
      it isn't a direct user gesture. initApp() shows a "Sign In" button
      instead when no session is cached, so the popup fires from a real click
      and succeeds on the first try. Once a session exists, initApp() skips
      the button and goes straight into loadLive() as before.
      You can also delete the `const SEED = "...";` line and the
      <script src=".../xlsx@0.18.5/..."> tag entirely -- neither is used
      anymore in live mode. parseWorkbook() itself can stay in the file
      unused, or be deleted; nothing else calls it once initApp() is wired in.
   5. Remove/skip this if you still want a "Refresh Now" button: just call
      loadLive() again on click (not initApp() -- a session already exists
      by then); it re-runs everything and re-renders.
   ============================================================================ */

const PBI_CONFIG = {
  TENANT_ID:    "11bfe7ed-a96d-47b8-93b9-f1d5ced7091b",
  CLIENT_ID:    "026eb3d4-356a-4d3b-8f44-7adcd56f75c4",
  WORKSPACE_ID: "003c3aa9-61d8-4530-b8f7-e64e531babac",
  DATASET_ID:   "9feaf6d8-5d92-4212-a381-444dbe7ba277",
};

// ---------------------------------------------------------------------------
// SharePoint / Microsoft Graph config -- MMR Leader Input Log
// Resolved via: GET https://graph.microsoft.com/v1.0/shares/{encoded-url}/driveItem
// File: MMR_Leader_Input_Log.xlsx, site: PowerBI-CollaborationTeam
// ---------------------------------------------------------------------------
const SHAREPOINT_CONFIG = {
  DRIVE_ID: "b!UpDL1A9FgEGL6TSnvGTjSB5qvG0M8kBHgVt2i-xLcGj3ikKTWNKkRoPWUUZqeQy2",
  ITEM_ID:  "01JMH5MWOMMVXQBHCXPVHL3GHDYE4GFBUV",
};

const D365_BASE = "https://onedigital.crm.dynamics.com/main.aspx?appid=c8495106-ec1a-e911-a952-000d3a1d55a5&pagetype=entityrecord&etn=opportunity&id=";
const SOUTH_MARKETS = ['Carolinas','Tennessee','North/Central Florida','South Florida','Georgia','Texas'];
const BUSINESS_TO_PRACTICE = {'Employee Benefits':'EB','Human Resources':'HRC','P&C':'P&C'};
const DISPLAY_PRACTICE = {'Employee Benefits':'Employee Benefits','Human Resources':'HRC','P&C':'P&C'};

// ---------------------------------------------------------------------------
// Auth (MSAL.js -- Public Client / SPA, no secret, PKCE flow)
// ---------------------------------------------------------------------------
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: PBI_CONFIG.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${PBI_CONFIG.TENANT_ID}`,
    redirectUri: window.location.href
  },
  cache: { cacheLocation: "sessionStorage" }
});

let _msalInitialized = false;
async function ensureMsalInit(){
  if(_msalInitialized) return;
  await msalInstance.initialize();
  _msalInitialized = true;
}

// ---------------------------------------------------------------------------
// Token acquisition -- guarded against concurrent calls.
//
// WHY THE GUARD: loadLive() fires 6 DAX queries at once via Promise.all(),
// and each one calls getToken() independently. On a fresh sign-in with no
// cached session, all 6 would otherwise try to open an MSAL popup at nearly
// the same instant. MSAL only permits ONE interactive login in flight at a
// time -- the first call wins and shows the real popup, but the other 5
// immediately fail with "interaction_in_progress", and since Promise.all
// rejects on the first failure, the whole load appears to fail even though
// the actual login succeeded underneath (which is why clicking Retry right
// after always worked -- a session was already cached by then).
//
// FIX: concurrent callers share the SAME in-flight request instead of each
// starting their own competing popup attempt.
// ---------------------------------------------------------------------------
let _tokenPromise = null;
async function getToken(){
  if(_tokenPromise) return _tokenPromise;
  _tokenPromise = (async () => {
    await ensureMsalInit();
    const request = { scopes: ["https://analysis.windows.net/powerbi/api/Dataset.Read.All"] };
    const accounts = msalInstance.getAllAccounts();
    if(accounts.length > 0){
      try{
        const res = await msalInstance.acquireTokenSilent({ ...request, account: accounts[0] });
        return res.accessToken;
      }catch(e){ /* fall through to interactive */ }
    }
    const res = await msalInstance.loginPopup(request);
    return res.accessToken;
  })();
  try{
    return await _tokenPromise;
  }finally{
    _tokenPromise = null; // clear once settled so a later (non-concurrent) call can re-trigger if the session expires
  }
}

// ---------------------------------------------------------------------------
// Graph token -- for writing MMR leader input to the SharePoint Excel log.
// Reuses the same msalInstance/ensureMsalInit as getToken() above; just a
// different scope. First call may show a one-time consent popup for
// Files.ReadWrite.All even if the user already consented to the Power BI
// scope. NOTE: must exactly match the permission granted admin consent in
// Entra ID -- "Files.ReadWrite" and "Files.ReadWrite.All" are DIFFERENT,
// distinct Graph permissions; requesting one when only the other was
// consented fails silently from the user's perspective (mmrSave() shows
// "Saved locally only"). Same concurrent-call guard as getToken() above --
// not currently needed here since saveMMRToSharePoint() awaits its 4
// graphAddRows() calls sequentially, but hardened the same way in case that
// ever changes.
// ---------------------------------------------------------------------------
let _graphTokenPromise = null;
async function getGraphToken(){
  if(_graphTokenPromise) return _graphTokenPromise;
  _graphTokenPromise = (async () => {
    await ensureMsalInit();
    const request = { scopes: ["Files.ReadWrite.All"] };
    const accounts = msalInstance.getAllAccounts();
    if(accounts.length > 0){
      try{
        const res = await msalInstance.acquireTokenSilent({ ...request, account: accounts[0] });
        return res.accessToken;
      }catch(e){ /* fall through to interactive */ }
    }
    const res = await msalInstance.loginPopup(request);
    return res.accessToken;
  })();
  try{
    return await _graphTokenPromise;
  }finally{
    _graphTokenPromise = null;
  }
}

// ---------------------------------------------------------------------------
// Write MMR leader input to the SharePoint Excel log (MMR_Leader_Input_Log.xlsx)
// via Microsoft Graph. Table names below are the underlying Excel Table
// names (Table Design tab), NOT the sheet tab names -- confirmed unchanged
// even after the sheet tabs were renamed to MMR_MonthlyCommit /
// MMR_DealLevelForecast / MMR_ProducerFocus / MMR_RecentWins.
// ---------------------------------------------------------------------------
async function graphAddRows(tableName, rowsArray){
  if(!rowsArray || !rowsArray.length) return;
  const token = await getGraphToken();
  const url = `https://graph.microsoft.com/v1.0/drives/${SHAREPOINT_CONFIG.DRIVE_ID}` +
              `/items/${SHAREPOINT_CONFIG.ITEM_ID}/workbook/tables('${tableName}')/rows/add`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values: rowsArray })  // rowsArray = [[col1,col2,...], [col1,col2,...]]
  });
  if(!res.ok){
    const err = await res.text();
    throw new Error(`Graph write to ${tableName} failed (${res.status}): ${err}`);
  }
  return res.json();
}

// Returns the signed-in user's email (verified via MSAL/Entra ID), NOT the
// free-text "Market Leader" field a user types into the form -- this is the
// actual authenticated identity of whoever clicked Save.
async function currentUserEmail(){
  await ensureMsalInit();
  const acc = msalInstance.getAllAccounts()[0];
  return acc ? acc.username : "unknown";
}

async function saveMMRToSharePoint(state){
  const saveId = `${state.market||'ALL'}::${MONTHS[CUR_MONTH_IDX]}${CUR_YEAR}::${Date.now()}`;
  const savedAtUTC = new Date().toISOString();
  const savedByUser = await currentUserEmail();
  // Must match the on-screen section headers exactly: "This Month" / "Quarter End" / "Next Month"
  const HZ_LABEL = { dThis: "This Month", dQtr: "Quarter End", dNext: "Next Month" };

  // AdditionalConfidentToClose_Calculated is NOT a user input -- it's mmrComputeAddl(state),
  // already computed inside mmrGather(). Stored here as a point-in-time snapshot.
  await graphAddRows("MMR_Summary", [[
    saveId, savedAtUTC, savedByUser, state.market, MONTHS[CUR_MONTH_IDX], CUR_YEAR,
    state.leader, state.commitHz, state.addl, state.nextOpps
  ]]);

  const dealRows = [];
  ["dThis","dQtr","dNext"].forEach(hz => {
    (state[hz]||[]).forEach(d => {
      if(!d.deal && !d.amt) return; // skip empty rows
      dealRows.push([saveId, savedAtUTC, savedByUser, HZ_LABEL[hz], d.deal, d.producer, d.amt, d.close, d.conf, d.notes]);
    });
  });
  if(dealRows.length) await graphAddRows("MMR_Deals", dealRows);

  const prodRows = (state.prod||[])
    .filter(p => p.producer)
    .map(p => [saveId, savedAtUTC, savedByUser, p.producer, p.gap, p.plan, p.outcome]);
  if(prodRows.length) await graphAddRows("MMR_ProducersBehindPace", prodRows);

  const winRows = (state.wins||[])
    .filter(w => w.producer)
    .map(w => [saveId, savedAtUTC, savedByUser, w.producer, w.deal, w.teeup, w.advance, w.client, w.team]);
  if(winRows.length) await graphAddRows("MMR_RecentWins", winRows);
}

// ---------------------------------------------------------------------------
// DAX execution helper. Returns an array of plain objects with bracket-free
// keys, e.g. a row column "[Producer]" becomes row.Producer.
// ---------------------------------------------------------------------------
async function runDax(query){
  const token = await getToken();
  const url = `https://api.powerbi.com/v1.0/myorg/groups/${PBI_CONFIG.WORKSPACE_ID}/datasets/${PBI_CONFIG.DATASET_ID}/executeQueries`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      queries: [{ query }],
      serializerSettings: { includeNulls: true }
    })
  });
  if(!res.ok){
    const text = await res.text();
    throw new Error(`Power BI API error ${res.status}: ${text}`);
  }
  const json = await res.json();
  const rawRows = json.results[0].tables[0].rows;
  return rawRows.map(row => {
    const out = {};
    for(const key in row){
      const cleanKey = key.replace(/^\[/, '').replace(/\]$/, '');
      out[cleanKey] = row[key];
    }
    return out;
  });
}

function n(v){ const x = parseFloat(v); return isNaN(x) ? 0 : x; }
function normMarket(m){ return m === 'North / Central Florida' ? 'North/Central Florida' : m; }

// ---------------------------------------------------------------------------
// The DAX queries themselves. Same filters validated this session:
// South region, EB/HRC/P&C practice, Compass excluded, current year via
// YEAR(TODAY()) so this never needs updating for next year.
// ---------------------------------------------------------------------------
const REGION_FILTER  = `FILTER(VALUES('Dim Broker'[PL Region]), 'Dim Broker'[PL Region]="South")`;
const BROKER_FILTER  = `FILTER(VALUES('Dim Broker'[PL Broker]), 'Dim Broker'[PL Broker]<>"Compass")`;
const PRACTICE_FILTER = `FILTER(VALUES(FactNetGrowthKPIs[Business Practice]), FactNetGrowthKPIs[Business Practice] IN {"Employee Benefits","Human Resources","P&C"})`;
const CONV_SPLIT_EXPR = `CALCULATE(SUMX(FactNetGrowthKPIs, FactNetGrowthKPIs[Referral Consultant Split%] + FactNetGrowthKPIs[Secondary Referral Consultant Split%] + FactNetGrowthKPIs[Third Referral Consultant Split%]))`;

const DAX_QUERIES = {

  // 6-year rolling window (current year and 5 prior) so the seasonal pacing
  // curve on Goal Trending has real history, not the linear fallback.
  clientLevel: `
    EVALUATE
    SELECTCOLUMNS(
      FILTER(
        SUMMARIZECOLUMNS(
          DimProducer[Producer], DimClientProspect[Company Name], DimClientProspect[Company Type],
          DimClientProspect[Size], DimClientProspect[Industry Group], 'Dim Broker'[Geographic Market],
          FactNetGrowthKPIs[Business Practice], 'Date'[Month Name Short], 'Date'[Year], FactNetGrowthKPIs[Opportunity ID],
          ${REGION_FILTER}, ${BROKER_FILTER}, ${PRACTICE_FILTER},
          FILTER(ALL('Date'[Year]), 'Date'[Year] >= YEAR(TODAY())-5 && 'Date'[Year] <= YEAR(TODAY())),
          "Won", [Won_Revenue],
          "NonRec", [NonRecurring_Won_Revenue],
          "Conv", ${CONV_SPLIT_EXPR}
        ),
        [Won] <> 0
      ),
      "Producer",[Producer], "Company",[Company Name], "CoType",[Company Type], "Size",[Size],
      "Industry",[Industry Group], "Market",[Geographic Market], "Practice",[Business Practice],
      "Month",[Month Name Short], "Year",[Year], "OppID",[Opportunity ID],
      "Won",[Won], "NonRec",[NonRec], "Conv",[Conv]
    )
  `,

  // Producer goals from FactOwner. AVERAGE (not SUM) per Owner/Year avoids
  // the known x12 fan-out (FactOwner stores one row per producer per month).
  // Do NOT add 'Dim Broker'[Geographic Market] to this grouping -- there is
  // no real relationship between DimProducer and Dim Broker, and doing so
  // fans out to all 6 markets per producer. Market is derived later, in JS,
  // via majority vote across each producer's own client-level rows.
  producerGoals: `
    EVALUATE
    FILTER(
      SELECTCOLUMNS(
        SUMMARIZECOLUMNS(
          DimProducer[Producer], DimProducer[Type], DimProducer[Market], DimProducer[BR Sales Leader],
          FILTER(VALUES(FactOwner[Region]), FactOwner[Region]="South"),
          FILTER(VALUES(FactOwner[Year]), FactOwner[Year] = FORMAT(YEAR(TODAY()), "0")),
          FILTER(VALUES(FactOwner[Sub-Vertical]), FactOwner[Sub-Vertical] IN {"Employee Benefits","HRC","P&C"}),
          "Goal", AVERAGE(FactOwner[Producer Goal])
        ),
        "Producer",[Producer], "Type",[Type], "Market",[Market], "Leader",[BR Sales Leader], "Goal",[Goal]
      ),
      [Goal] > 0
    )
  `,

  // Current-year producer performance: Won/NonRec/Conv (row-level fix from
  // this session), 12-month pipeline, delinquent pipeline, the REAL
  // Total Contribution Revenue measure, and Meetings from the Activity table.
  producerPerf: `
    EVALUATE
    SELECTCOLUMNS(
      SUMMARIZECOLUMNS(
        DimProducer[Producer],
        ${REGION_FILTER}, ${BROKER_FILTER},
        FILTER(VALUES('Date'[Year]), 'Date'[Year] = YEAR(TODAY())),
        ${PRACTICE_FILTER},
        "Won", [Won_Revenue],
        "NonRec", [NonRecurring_Won_Revenue],
        "Conv", ${CONV_SPLIT_EXPR},
        "Pipe12", [Pipeline Rolling 12 Months],
        "Delinq", [Pipeline Delinquent by Producer],
        "TotalContribution", [Total Contribution Revenue],
        "Meetings", CALCULATE(COUNTROWS(Activity), Activity[Activity Type]="Meeting")
      ),
      "Producer",[Producer], "Won",[Won], "NonRec",[NonRec], "Conv",[Conv],
      "Pipe12",[Pipe12], "Delinq",[Delinq], "TotalContribution",[TotalContribution], "Meetings",[Meetings]
    )
  `,

  pipeline: `
    EVALUATE
    SELECTCOLUMNS(
      FILTER(
        SUMMARIZECOLUMNS(
          DimProducer[Producer], DimClientProspect[Company Name], DimClientProspect[Company Type],
          DimClientProspect[Industry Group], DimClientProspect[Size], 'Dim Broker'[Geographic Market],
          DimOpportunity[Pipeline], 'Date'[Month Name Short], 'Date'[Year], DimOpportunity[Opportunity ID],
          ${REGION_FILTER}, ${BROKER_FILTER},
          FILTER(VALUES(DimOpportunity[Status]), DimOpportunity[Status]="Open"),
          ${PRACTICE_FILTER},
          "Rev", [Total Pipeline]
        ),
        [Rev] <> 0
      ),
      "Producer",[Producer], "Company",[Company Name], "CoType",[Company Type], "Industry",[Industry Group],
      "Size",[Size], "Market",[Geographic Market], "Pipeline",[Pipeline], "Month",[Month Name Short],
      "Year",[Year], "OppID",[Opportunity ID], "Rev",[Rev]
    )
  `,

  convergence: `
    EVALUATE
    SELECTCOLUMNS(
      FILTER(
        SUMMARIZECOLUMNS(
          DimClientProspect[Company Name], 'Date'[Month Name Short], 'Date'[Year], 'Dim Broker'[Geographic Market],
          DimOpportunity[Referral Consultant], DimOpportunity[Primary Producer],
          ${REGION_FILTER}, ${BROKER_FILTER},
          FILTER(VALUES('Date'[Year]), 'Date'[Year] = YEAR(TODAY())),
          ${PRACTICE_FILTER},
          "Rev", ${CONV_SPLIT_EXPR}
        ),
        [Rev] <> 0
      ),
      "Company",[Company Name], "Month",[Month Name Short], "Year",[Year], "Market",[Geographic Market],
      "Referrer",[Referral Consultant], "Producer",[Primary Producer], "Rev",[Rev]
    )
  `,

  // The model's own Convergence Goal Tracker table. It only stores combined
  // "Tennessee + Carolinas" and "Florida" totals -- the submarket split is
  // computed client-side in JS from real won-revenue share (see splitGoal()).
  convergenceGoalTracker: `
    EVALUATE
    SELECTCOLUMNS(
      'Convergence Goal Tracker',
      "Market", 'Convergence Goal Tracker'[Market],
      "Goal", 'Convergence Goal Tracker'[EB/HRC/P&C]
    )
  `
};

// ---------------------------------------------------------------------------
// Majority-vote helpers: DimProducer <-> Dim Broker has no real relationship,
// so a producer's market/practice is derived from their own real transactions
// rather than trusted from a direct (and fan-out-prone) join.
// ---------------------------------------------------------------------------
function majorityVote(rows, key, producerName){
  const counts = {};
  for(const r of rows){
    if(r.Producer === producerName && r[key]){
      counts[r[key]] = (counts[r[key]] || 0) + 1;
    }
  }
  let best = null, bestCount = 0;
  for(const k in counts){ if(counts[k] > bestCount){ best = k; bestCount = counts[k]; } }
  return best;
}

function splitGoal(combinedGoal, subA, subB, wonBySubmarket){
  const wa = wonBySubmarket[subA] || 0;
  const wb = wonBySubmarket[subB] || 0;
  const total = wa + wb;
  if(total <= 0) return [combinedGoal/2, combinedGoal/2];
  return [combinedGoal * wa / total, combinedGoal * wb / total];
}

// ---------------------------------------------------------------------------
// Main entry point. Runs all 6 queries in parallel, assembles the exact same
// DATA shape parseWorkbook() used to produce, then calls the dashboard's own
// computeSeasonal() + boot() so every render function works unmodified.
// ---------------------------------------------------------------------------
async function loadLive(){
  const el = document.getElementById('loadingMsg');
  if(el) el.innerHTML = 'Signing in and querying the live model&hellip;';

  try{
    const [clientRowsRaw, goalRows, perfRows, pipeRowsRaw, convRowsRaw, goalTrackerRows] = await Promise.all([
      runDax(DAX_QUERIES.clientLevel),
      runDax(DAX_QUERIES.producerGoals),
      runDax(DAX_QUERIES.producerPerf),
      runDax(DAX_QUERIES.pipeline),
      runDax(DAX_QUERIES.convergence),
      runDax(DAX_QUERIES.convergenceGoalTracker)
    ]);

    if(el) el.innerHTML = 'Assembling data&hellip;';

    // --- normalize market spelling everywhere ---
    clientRowsRaw.forEach(r => r.Market = normMarket(r.Market));
    pipeRowsRaw.forEach(r => r.Market = normMarket(r.Market));
    convRowsRaw.forEach(r => r.Market = normMarket(r.Market));

    // --- Client level Data -> DATA.cli ---
    const cli = clientRowsRaw.map(r => ({
      producer: r.Producer, market: r.Market,
      practice: BUSINESS_TO_PRACTICE[r.Practice] || r.Practice,
      region: 'South', company: r.Company, type: r.CoType, size: r.Size,
      industry: r.Industry, oppLink: r.OppID ? (D365_BASE + r.OppID) : '',
      month: r.Month, year: n(r.Year),
      won: n(r.Won), nonrec: n(r.NonRec), conv: n(r.Conv)
    })).filter(r => r.year > 0 && MONTHS.includes(r.month));

    const years = [...new Set(cli.map(r => r.year))].sort();
    const curYear = Math.max(...years);

    // --- dominant market/practice per producer (majority vote) ---
    const allProducers = [...new Set(goalRows.map(g => g.Producer).filter(Boolean))];
    const dominantMarket = {}, dominantPractice = {};
    allProducers.forEach(name => {
      dominantMarket[name] = majorityVote(clientRowsRaw, 'Market', name);
      dominantPractice[name] = majorityVote(clientRowsRaw, 'Practice', name);
    });

    // --- Producer Data -> DATA.producers ---
    const perfByName = {};
    perfRows.forEach(r => { perfByName[r.Producer] = r; });
    const producers = goalRows
      .filter(g => n(g.Goal) > 0)
      .map(g => {
        const p = perfByName[g.Producer] || {};
        return {
          producer: g.Producer, goal: n(g.Goal),
          won: n(p.Won), nonrec: n(p.NonRec), conv: n(p.Conv),
          total: p.TotalContribution != null ? n(p.TotalContribution) : null,
          pipeline12: p.Pipe12 != null ? n(p.Pipe12) : null,
          delinq: p.Delinq != null ? n(p.Delinq) : null,
          market: dominantMarket[g.Producer] || g.Market,
          region: 'South',
          practice: DISPLAY_PRACTICE[dominantPractice[g.Producer]] || dominantPractice[g.Producer],
          leader: g.Leader, type: g.Type,
          meetings: p.Meetings != null ? n(p.Meetings) : 0
        };
      });

    // --- Pipeline Data -> DATA.pipe ---
    const pipe = pipeRowsRaw.map(r => ({
      company: r.Company, producer: r.Producer, rev: n(r.Rev),
      stage: normStage(r.Pipeline), month: r.Month, year: n(r.Year),
      size: r.Size, market: r.Market, region: 'South',
      type: r.CoType, industry: r.Industry,
      oppLink: r.OppID ? (D365_BASE + r.OppID) : ''
    })).filter(r => r.rev > 0 && r.stage);

    // --- Revenue.Plan equivalent: planByMarket / planByPractice ---
    const planByMarket = {}, planByPractice = {};
    SOUTH_MARKETS.forEach(m => { planByMarket[m] = 0; planByPractice[m] = { EB:0, HRC:0, 'P&C':0 }; });
    goalRows.filter(g => n(g.Goal) > 0).forEach(g => {
      const mkt = dominantMarket[g.Producer] || normMarket(g.Market);
      const prac = DISPLAY_PRACTICE[dominantPractice[g.Producer]] === 'HRC' ? 'HRC'
                 : (BUSINESS_TO_PRACTICE[dominantPractice[g.Producer]] || dominantPractice[g.Producer]);
      if(planByPractice[mkt] && prac && planByPractice[mkt][prac] !== undefined){
        planByPractice[mkt][prac] += n(g.Goal);
        planByMarket[mkt] += n(g.Goal);
      }
    });

    // --- Convergence Goal Tracker -> DATA.convGoalByMarket ---
    const liveGoalByName = {};
    goalTrackerRows.forEach(r => { liveGoalByName[r.Market] = n(r.Goal); });
    const wonBySubmarket = {};
    clientRowsRaw.filter(r => n(r.Year) === curYear).forEach(r => {
      wonBySubmarket[r.Market] = (wonBySubmarket[r.Market] || 0) + n(r.Won);
    });
    const [carolinasGoal, tennesseeGoal] = splitGoal(liveGoalByName['Tennessee + Carolinas'] || 0, 'Carolinas', 'Tennessee', wonBySubmarket);
    const [ncFlGoal, sFlGoal] = splitGoal(liveGoalByName['Florida'] || 0, 'North/Central Florida', 'South Florida', wonBySubmarket);
    const convGoalByMarket = {
      'Carolinas': carolinasGoal, 'Tennessee': tennesseeGoal,
      'North/Central Florida': ncFlGoal, 'South Florida': sFlGoal,
      'Georgia': liveGoalByName['Georgia'] || 0, 'Texas': liveGoalByName['Texas'] || 0
    };

    // --- Convergence transactions -> convByMarket / convByProducer / convByMonth ---
    const convByMarket = {}, convByProducer = {}, convByMonth = Array(12).fill(0);
    convRowsRaw.forEach(r => {
      const rev = n(r.Rev);
      convByMarket[r.Market] = (convByMarket[r.Market] || 0) + rev;
      const mi = MONTHS.indexOf(r.Month);
      if(mi >= 0) convByMonth[mi] += rev;
      if(r.Producer) convByProducer[r.Producer] = (convByProducer[r.Producer] || 0) + rev;
    });

    // --- assemble DATA exactly as parseWorkbook() used to, then reuse the
    //     dashboard's own seasonal-curve + render pipeline unmodified ---
    DATA = { cli, producers, pipe, years, planByMarket, planByPractice, convGoalByMarket, convByMarket, convByProducer, convByMonth };
    CUR_YEAR = curYear;
    computeSeasonal();
    boot();

  }catch(err){
    console.error(err);
    if(el) el.innerHTML = `Live query failed: ${err.message}. <button onclick="loadLive()">Retry</button>`;
  }
}

// ---------------------------------------------------------------------------
// Entry point used by the page (replaces the old direct "loadLive();" call).
//
// WHY THIS EXISTS: browsers block window.open() popups that aren't triggered
// by a direct user gesture (click/tap). loadLive() -> getToken() -> loginPopup()
// used to fire automatically on page load, so on any visit without an existing
// cached session, the browser silently killed the popup (empty_window_error /
// popup_window_error) and the user had to click "Retry" -- which worked only
// because THAT click counted as a real user gesture.
//
// FIX: check for a cached MSAL account first.
//   - Returning visit, session still cached -> straight into loadLive(), silent
//     token refresh, no popup, no button, unchanged from before.
//   - First visit / expired session -> show a "Sign In" button immediately
//     instead of attempting a doomed automatic popup. The click on THAT button
//     is what makes the resulting loginPopup() succeed on the first try.
// ---------------------------------------------------------------------------
async function initApp(){
  const el = document.getElementById('loadingMsg');
  await ensureMsalInit();
  const hasSession = msalInstance.getAllAccounts().length > 0;

  if(hasSession){
    loadLive();
    return;
  }

  if(el){
    el.innerHTML = 'Sign in to load live data from Power BI.<br><br>' +
      '<button onclick="loadLive()" style="font-size:16px;padding:10px 24px;cursor:pointer">Sign In</button>';
  }
}