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
            loadLive();
      You can also delete the `const SEED = "...";` line and the
      <script src=".../xlsx@0.18.5/..."> tag entirely -- neither is used
      anymore in live mode. parseWorkbook() itself can stay in the file
      unused, or be deleted; nothing else calls it once loadLive() is wired in.
   5. Remove/skip this if you still want a "Refresh Now" button: just call
      loadLive() again on click; it re-runs everything and re-renders.
   ============================================================================ */

const PBI_CONFIG = {
  TENANT_ID:    "YOUR_TENANT_ID_HERE",
  CLIENT_ID:    "YOUR_APP_CLIENT_ID_HERE",
  WORKSPACE_ID: "YOUR_WORKSPACE_ID_HERE",
  DATASET_ID:   "YOUR_DATASET_ID_HERE",
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

async function getToken(){
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
