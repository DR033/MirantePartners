/**
 * Mirante Partners Toolkit - v1.4 - 07.19.2025
 */

// ---------------------------------------------------------------------------
// Configuration handling using a dedicated sheet "// Config"
// ---------------------------------------------------------------------------

function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('// Config');
}

function ensureConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('// Config');
  if (!sheet) {
    sheet = ss.insertSheet('// Config');
    sheet.getRange(1, 1, 1, 2).setValues([['Description', 'Value']]);
    sheet.hideColumns(3); // store keys internally
  }
  // Ensure columns A and B use "Clip" wrapping so long text does not expand rows
  sheet.getRange(1, 1, sheet.getMaxRows(), 2)
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  return sheet;
}

function readConfig(key) {
  const sheet = getConfigSheet();
  if (!sheet) return '';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === key) return data[i][1];
  }
  return '';
}

function writeConfig(key, value) {
  const sheet = ensureConfigSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow(['', value, key]);
}

function getColumnConfig(propName) {
  const value = readConfig(propName);
  if (!value) {
    SpreadsheetApp.getUi().alert(`Configuration Error! The column for "${propName}" has not been set for this sheet. Please ask an administrator to configure it.`);
    throw new Error(`Missing configuration for ${propName}`);
  }
  return value;
}

function getLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('logs');
}

function ensureLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('logs');
  if (!sheet) {
    sheet = ss.insertSheet('logs');
    sheet.getRange(1, 1, 1, 4)
         .setValues([[
           'Timestamp',
           'User Email',
           'Function',
           'Status'
         ]]);
  }
  return sheet;
}

function logAction(funcName, status) {
  const sheet = ensureLogsSheet();
  sheet.appendRow([
    new Date(),
    Session.getActiveUser().getEmail(),
    funcName,
    status
  ]);
}

// Default Configuration Values (overridden via setup)
const MODEL = 'claude-sonnet-4-20250514';
const TEMPERATURE = 0.4;
const MAX_TOKENS = 350;
const EMAIL_PROMPT_ROW = 13;       // current prompt used by Claude
const PREVIOUS_PROMPT_ROW = 14;    // stores the previous customization style
const LAST_EMAIL_ROW = 15;
const INFO_PROMPT_ROW = 16;        // active custom info guidelines
const PREVIOUS_INFO_ROW = 17;      // previous custom info prompt
const LAST_INFO_ROW = 18;

// Default text for the email writing guidelines section of the system prompt
const DEFAULT_EMAIL_GUIDELINES = `Context & Purpose
  You are Diego Rosário from Mirante Partners, an M&A investor writing a personalized cold outreach email to the owner of a private company.
  Your objective is to demonstrate genuine research and understanding of their business while professionally expressing acquisition interest.

  CRITICAL PERSONALIZATION REQUIREMENTS:
  1. Find ONE genuinely interesting, specific detail about their business (avoid obvious/generic facts)
  2. Use conversational, natural language - avoid business jargon and buzzwords
  3. Reference something that shows you actually read their content (not just skimmed)
  4. Keep the owner's background mention subtle and relevant
  5. Sound like a human, not a marketing email

  2. Greeting: "Hi [Owner's first name],"

  3. Opening (1-2 sentences):
      - Start with a simple, genuine observation about their business
      - No superlatives or excessive praise
      - Sound like you're having a conversation, not pitching

  4. Purpose & Ask (2-3 sentences):
      - State that you're Diego from Mirante Partners, an investor who works with business owners on exits
      - Express interest in having a conversation if they're ever considering an exit
      - Simple request: "Would you be open to a brief conversation if this might be an appropriate time?"

  5. Professional Signature Block:
      Best regards,
      Diego Rosário
      Mirante Partners

  Tone Guidelines:
  - Write like you're sending a message to a colleague, not a sales prospect
  - Be genuinely curious rather than impressed
  - Avoid business buzzwords: "strategic," "compelling," "leverage," "ecosystem," "synergies"
  - No excessive praise or superlatives
  - Keep it under 100 words
  - Just write the body of the email, without subject line
  - Be respectful of their timing and decision-making autonomy
  - Frame as "if you're ever considering an exit" not "I want to acquire your company"
  - Use phrases like "appropriate time" or "if this might be of interest"
  - Don't assume they want to sell - respect that it's their choice

  AVOID AT ALL COSTS:
  - "Caught my attention" / "I came across"
  - "Impressive" / "exciting" / "compelling"
  - "Strategic alignment" / "partnership opportunity" (use "acquisition" instead)
  - "Proven track record" / "demonstrated success"
  - Industry jargon and consultant-speak
  - Multiple compliments in one email
  - Overly detailed business analysis
  - Vague language about "exploring opportunities" (be direct about acquisition)

  SAMPLE PHRASES TO USE:
  - "I'd love to have a conversation if you think this might be an appropriate time"
  - "If you're ever considering an exit, I'd be interested in discussing"
  - "Would you be open to a brief conversation if this is something you might consider"
  - "I work with business owners exploring exit opportunities"

  Quality Checklist:
  - Does it respect their autonomy and timing?
  - Would you be comfortable receiving this email?
  - Does it sound like genuine interest, not aggressive pursuit?
  - Is the specific detail meaningful, not just factual?`;

const DEFAULT_CUSTOM_INFO_GUIDELINES = `Context & Purpose
You are generating personalized observations for business outreach emails. These observations demonstrate research and create connection points while maintaining professional authenticity. The goal is to sound like a knowledgeable professional who has genuinely researched the recipient, not someone trying to impress with industry expertise.

Length & Structure
Aim for ~250 characters (1–2 brief sentences). Always start with "By the way" or "As an aside".

Content Guidelines
- Make a brief, intelligent observation about: Their background or career moves; Company situation or recent developments; Industry context (surface-level only)
- Show you've done research without being overly detailed
- Include moderate empathy and occasional opinion qualifiers: "I think that..."; "I believe that...", "Sounds like..."

Tone
- Conversational and direct – not overly pleasant or flattering
- Sound like you're stating an interesting fact you noticed
- Avoid trying to impress with deep industry knowledge
- No technical rambling or extensive analysis
- Professional but authentic

Style Examples
Example 1:
"By the way, saw that you bought Precision Safe Sidewalks shortly after launching Blue Zone, and now you're on your second acquisition in the ADA space. Sounds like both ventures are thriving!"

Example 2:
"By the way, saw that you guys were very early adopters of Apple Newton while writing your own software for ems charting. I think it's a testament to you and Shawn's care for solving problems proactively and making life just a bit easier for our frontline responders!"

Key Success Criteria
- Demonstrates genuine research without being invasive
- Creates natural conversation starter
- Maintains professional boundaries
- Shows authentic interest rather than sales manipulation
- Balances knowledge with humility
- ALWAYS addresses the recipient directly using "you/your" rather than their name in third person`;

/**
 * Convert a column letter (e.g. "A", "AA") to its 1-based index.
 */
function letterToColumn(letter) {
  let col = 0;
  letter = letter.toUpperCase();
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

/**
 * Retrieve a script property value. Throws if not set.
 */
function getConfig(name) {
  const val = readConfig(name);
  if (!val) {
    SpreadsheetApp.getUi().alert(
      `Configuration Error!\n` +
      `Property "${name}" not found.\n\n` +
      `Please run 👑 Admin Setup → Setup Columns & Sheet Name.`
    );
    throw new Error(`Missing configuration for ${name}`);
  }
  return val;
}


/**
 * Normalize a URL to its domain (remove protocol, www, and paths).
 */
function normalizeDomain(url) {
  try {
    const u = new URL(url);
    let host = u.hostname;
    if (host.startsWith('www.')) host = host.slice(4);
    return host;
  } catch (e) {
    return url.replace(/https?:\/\//, '').split('/')[0].replace(/^www\./, '');
  }
}

function buildSystemPrompt() {
  // Use getColumnConfig to dynamically insert the owner column letter.
  const ownerCol = getColumnConfig('OWNER_COL_LETTER');

  const dataSection = `Input Data Structure:
  - Company Name
  - Homepage content (scraped)
  - About Us section (scraped)
  - Column ${ownerCol}: Recipient's name (use for context - address them directly as "you" in the observation)
  - Owner+Company combined text (scraped)
  - Scraped Owner LinkedIn profile text
  - Additional scraped page contents (up to 5 pages)`;

  const guidelines = readConfig('CUSTOM_INFO_GUIDELINES') || DEFAULT_CUSTOM_INFO_GUIDELINES;

  return `${dataSection}\n\n${guidelines}`;
}

function buildSystemPrompt2() {
  // We still need the owner column for the prompt's instructions.
  const ownerCol = getColumnConfig('OWNER_COL_LETTER');

  const dataSection = `Input Data Structure:
  - Company name
  - Homepage content
  - About Us section
  - Column ${ownerCol}: Owner's first name (for "you/your" references)
  - Combined Owner+Company text
  - Owner's LinkedIn bio and recent activity
  - Additional page contents (up to 5 pages)`;

  const guidelines = readConfig('EMAIL_GUIDELINES') || DEFAULT_EMAIL_GUIDELINES;

  return `${dataSection}\n\n${guidelines}`;
}

function onInstall(e) {
  onOpen(e);
}

/**
 * Creates the add-on menu in the spreadsheet UI.
 * This function runs automatically when the spreadsheet is opened.
 * @param {Object} e The event object.
 */
const ADMIN_USERS = [
  'fred@mirantepartners.com',
  'diego@mirantepartners.com'
];

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail();

  // Warn if not configured by checking for essential properties.
  const sheetName = readConfig('SHEET_NAME');
  const companyCol = readConfig('COMPANY_COL_LETTER');
  if (!sheetName || !companyCol) {
    ui.alert(
      'Toolkit not configured yet!',
      'Please run *⚙️ 0 - Setup Columns & Sheet Name* before using any of the menu commands.',
      ui.ButtonSet.OK
    );
  }

  // Build the main add‑on menu
  const menu = ui.createAddonMenu()
    .addItem('⚙️ 0 - Setup Columns & Sheet Name', 'setupColumnsForThisSheet')
    // 🔍 Find…
    .addItem('🔍 1 - Enrich Data',       'enrichData')
    .addSubMenu(ui.createMenu('✨ 2 - Create Customization')
      .addItem('✨ Create Customization', 'runCombinedScrapesOptimized')
      .addItem('🪄 Change Custom Info Style', 'changeCustomInfoStyle')
      .addItem('↩️ Revert to previous custom info', 'revertToPreviousInfoStyle')
      .addItem('🔄 Revert to default custom info', 'revertToDefaultInfoStyle'))

    // 🚀 Upload to Apollo
    .addSubMenu(ui.createMenu('🚀 3 - Upload to Apollo')
      .addItem('🚀 Upload Contacts to Apollo',  'uploadContacts')
      .addItem('🔄 Refresh Senders & Sequences','refreshLookups')
    );
  // Only add "Admin Functions" if the user is in your ADMIN_USERS list
  if (ADMIN_USERS.indexOf(userEmail) !== -1) {
      menu
        .addSubMenu(ui.createMenu('👑 Admin Functions')
          .addItem('🔑 Setup API Keys (Global)',    'setupApiKey')
          .addSubMenu(ui.createMenu('✉️ Create Full Email - beta')
            .addItem('✉️ Create Full Email (beta)',    'createFullEmail')
            .addItem('🪄 Change Email Optimization Style', 'changeEmailOptimizationStyle')
            .addItem('↩️ Revert to previous style',      'revertToPreviousCustomization')
            .addItem('🔄 Revert to default style',       'revertToDefaultCustomization'))
          .addSeparator()
        );
  }

  menu.addToUi();
}

/**
 * Prompt the administrator to configure column letters for this sheet.
 * Saved in the // Config sheet so each spreadsheet retains its own setup.
 */
function setupColumnsForThisSheet() {
  const fn = 'setupColumnsForThisSheet';
  try {
    requireAdmin();

    const ui    = SpreadsheetApp.getUi();
    const sheet = ensureConfigSheet();

  // Clear existing config and write headers
  sheet.clear();
  sheet.getRange(1, 1, 1, 2)
       .setValues([['Description', 'Value']]);

  const rows = [
    ['Sheet name (e.g., Sheet1)', '', 'SHEET_NAME'],
    ['Column letter for Company name', '', 'COMPANY_COL_LETTER'],
    ['Column letter for Website URL', '', 'WEBSITE_COL_LETTER'],
    ['Column letter for Industry', '', 'CUSTOM_INDUSTRY_COL_LETTER'],
    ['Column letter for Owner/Founder name', '', 'OWNER_COL_LETTER'],
    ['Column letter for Owner email address', '', 'OWNER_EMAIL_COL_LETTER'],
    ['Column letter for Owner LinkedIn URL', '', 'OWNER_LINKEDIN_COL_LETTER'],
    ['Column letter for output customization', '', 'OUTPUT_COL_LETTER'],
    ['Column letter to flag which rows to process', '', 'FIND_OWNER_INFO_COL_LETTER'],
    ['Column letter for Sequence ID', '', 'SEQUENCE_ID_COL_LETTER'],
    ['Column letter for Sender ID', '', 'SENDER_ID_COL_LETTER'],
    ['Email writing guidelines (edit if desired)', DEFAULT_EMAIL_GUIDELINES, 'EMAIL_GUIDELINES']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(2, 1, rows.length, 1).setFontStyle('italic');
  sheet.hideColumns(3);

  // Reserve rows for previous prompt and last generated email
  const prevRow = rows.length + 2; // starting row 2 plus length
  sheet.getRange(prevRow, 1, 1, 2)
       .setValues([['Previous Email Prompt', '']]);
  sheet.getRange(prevRow + 1, 1, 1, 2)
       .setValues([['Last Email Customization', '']]);
  sheet.getRange(prevRow, 1, 2, 1).setFontStyle('italic');

  const infoRow = prevRow + 2;
  sheet.getRange(infoRow, 1, 1, 3)
       .setValues([['Email (custom_info paragraph) writing guidelines', DEFAULT_CUSTOM_INFO_GUIDELINES, 'CUSTOM_INFO_GUIDELINES']]);
  sheet.getRange(infoRow + 1, 1, 1, 2)
       .setValues([['Previous custom_info Prompt', '']]);
  sheet.getRange(infoRow + 2, 1, 1, 2)
       .setValues([['Last Email Customization (custom_info paragraph)', '']]);
    sheet.getRange(infoRow, 1, 3, 1).setFontStyle('italic');

    ensureLogsSheet();
    ui.alert('✔ // Config sheet created. Please fill in the values before running any function.');
    logAction(fn, 'Completed');
  } catch (e) {
    logAction(fn, 'Failed: ' + e.message);
    throw e;
  }
}

/**
 * Prompt for API keys (Anthropic, Apollo, OpenAI and BrightData).
 */
function setupApiKey() {
  requireAdmin();
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKeys = [
    'ANTHROPIC_API_KEY',
    'APOLLO_API_KEY',
    'OPENAI_API_KEY',
    'BRIGHTDATA_API_KEY',
    'BRIGHTDATA_DATASET_ID',
    'BRIGHTDATA_UNLOCKER_ZONE',
    'BRIGHTDATA_SERP_ZONE'
  ];

  apiKeys.forEach(name => {
    const resp = ui.prompt(
      'API Key Setup',
      `Enter your new ${name}:`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() === ui.Button.OK) {
      scriptProps.setProperty(name, resp.getResponseText().trim());
    }
  });

  ui.alert('API keys saved to Script Properties');
}

/**
 * Run a search via Bright Data’s SERP API and return the top result URLs.
 */
function searchWithBrightData(query, numResults) {
  const apiKey   = getApiKey('BRIGHTDATA_API_KEY');
  const zoneName = getApiKey('BRIGHTDATA_SERP_ZONE');
  const endpoint = 'https://api.brightdata.com/request';

  // Build a Google‐style SERP URL with JSON parsing enabled
  const searchUrl = 'https://www.google.com/search?q=' +
                    encodeURIComponent(query) +
                    '&gl=us&hl=en&brd_json=1';

  const payload = {
    zone:   zoneName,
    url:    searchUrl,
    format: 'raw'
  };

  const options = {
    method:            'post',
    contentType:       'application/json',
    headers:           { Authorization: `Bearer ${apiKey}` },
    payload:           JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(endpoint, options);
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  console.log('🔍 BrightData SERP status:', code);
  console.log('🔍 BrightData SERP body:', body);

  if (code !== 200) {
    throw new Error(`SERP API Error: HTTP ${code} — ${body}`);
  }

  const data = JSON.parse(body);

  // pull from the “organic” array in your zone’s response
  const hits =
       data.organic
    || data.organic_results
    || data.results
    || data.data?.organic_results
    || data.items
    || [];

  return hits
    .slice(0, numResults)
    .map(item => item.link || item.url)
    .filter(u => !!u);
}

/**
 * Scrape via Bright Data’s Web Unlocker API (renders JS over HTTP).
 */
function callBrightDataUnlockerAPI(url, apiKey, zone, returnRawHtml = false) {
  const endpoint = 'https://api.brightdata.com/request';
  const payload  = {
    zone:   zone,
    url:    url,
    format: returnRawHtml ? 'raw' : 'json'
  };
  const opts = {
    method:            'post',
    contentType:       'application/json',
    headers:           { 'Authorization': `Bearer ${apiKey}` },
    payload:           JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const resp = UrlFetchApp.fetch(endpoint, opts);
  const code = resp.getResponseCode(), body = resp.getContentText();
  if (code !== 200) {
    console.error(`Unlocker API Error (${code}): ${body}`);
    return `Error: HTTP ${code}`;
  }
  
  // If raw HTML was requested, just return it
  if (returnRawHtml) {
    return body;
  }
  
  // Otherwise parse the JSON and extract the HTML field
  let html = '';
  try {
    const obj = JSON.parse(body);
    // Try a few common places where Unlocker might put the HTML
    if (obj.html) {
      html = obj.html;
    } else if (obj.response && obj.response.html) {
      html = obj.response.html;
    } else if (obj.data && obj.data.html) {
      html = obj.data.html;
    } else if (obj.body) {
      html = obj.body;
    } else {
      // Fallback: stringify whole obj in case it *is* the HTML
      html = typeof obj === 'string' ? obj : JSON.stringify(obj);
    }
  } catch (e) {
    console.error(`Failed to parse Unlocker JSON, returning raw body: ${e}`);
    html = body;
  }
  
  // Strip tags and collapse whitespace
  return stripHtml(html);
}

/**
 * Strip HTML tags and collapse whitespace.
 */
function stripHtml(html) {
  return html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
             .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
             .replace(/<[^>]+>/g, ' ')
             .replace(/&nbsp;/gi, ' ')
             .replace(/&amp;/gi, '&')
             .replace(/\s+/g, ' ')
             .trim();
}

function triggerBrightDataScrape(url, apiKey, datasetId) {
  const endpoint = "https://api.brightdata.com/datasets/v3/trigger";
  
  const headers = {
    "Authorization": `Bearer ${apiKey}`,
    "Content-Type": "application/json"
  };
  
  const params = {
    "dataset_id": datasetId,
    "include_errors": "true"
  };
  
  const data = [{"url": url}];
  
  // Build URL with parameters
  const urlWithParams = endpoint + "?" + Object.keys(params)
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');
  
  const options = {
    method: 'POST',
    headers: headers,
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(urlWithParams, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
    }
    
    const result = JSON.parse(response.getContentText());
    return result.id || result.snapshot_id;
    
  } catch (error) {
    console.error(`Error triggering Bright Data scrape: ${error.message}`);
    throw error;
  }
}

/**
 * Customize a single row by calling Anthropic and writing the result (with robust error handling).
 *
 * @param {number} rowNum   1‑based sheet row to process
 * @param {Sheet=} sheet    Optional Sheet object; if omitted, we look it up by name
 * @return {number}         1 if customization succeeded, 0 on error
 */
function runCustomization(rowNum, sheet = null) {
  console.log(`runCustomization called with rowNum: ${rowNum}, sheet: ${sheet}`);

  // If sheet isn't passed in, grab it by config
  if (!sheet) {
    console.log('Sheet is null, looking it up...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = getConfig('SHEET_NAME');
    console.log(`Looking for sheet "${name}" among: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
    sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`Sheet not found: "${name}"`);
  }

  // Resolve all column indexes
  const companyCol                = letterToColumn(getColumnConfig('COMPANY_COL_LETTER'));
  const homepageCol               = letterToColumn(getColumnConfig('HOMEPAGE_COL_LETTER'));
  const aboutUsCol                = letterToColumn(getColumnConfig('ABOUTUS_COL_LETTER'));
  const ownerCol                  = letterToColumn(getColumnConfig('OWNER_COL_LETTER'));
  const pageCols                  = [
    'PAGE1_COL_LETTER','PAGE2_COL_LETTER','PAGE3_COL_LETTER',
    'PAGE4_COL_LETTER','PAGE5_COL_LETTER'
  ].map(key => letterToColumn(getColumnConfig(key)));
  const linkedinScrapeCol         = letterToColumn(getColumnConfig('OWNER_LINKEDIN_SCRAPE_COL_LETTER'));
  const ownerNameCompanyCol       = letterToColumn(getColumnConfig('OWNER_NAME_COMPANY_NAME_SCRAPE_LETTER'));
  const outputCol                 = letterToColumn(getColumnConfig('OUTPUT_COL_LETTER'));

  // Read that single row
  const lastCol = sheet.getLastColumn();
  const rowData = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];

  const company             = rowData[companyCol - 1]             || '';
  const homepage            = rowData[homepageCol - 1]            || '';
  const aboutUs             = rowData[aboutUsCol - 1]             || '';
  const owner               = rowData[ownerCol - 1]               || '';
  const pages               = pageCols.map(c => rowData[c - 1]   || '');
  const linkedinProfileText = rowData[linkedinScrapeCol - 1]     || '';
  const ownerNameCompany    = rowData[ownerNameCompanyCol - 1]   || '';

  // Build prompts
  const systemPrompt = buildSystemPrompt();
  const userMsg = [
    `COMPANY: ${company}`,
    `HOMEPAGE: ${homepage}`,
    `ABOUT_US: ${aboutUs}`,
    `OWNER: ${owner}`,
    `OWNER_NAME_COMPANY_TEXT: ${ownerNameCompany}`,
    `LINKEDIN_PROFILE_TEXT: ${linkedinProfileText}`,
    `SCRAPED_PAGES:\n${pages.join('\n\n')}`
  ].join('\n');

  // Call Anthropic with our rate‑limiter
  const response = rateLimitedAnthropicFetch(
    'https://api.anthropic.com/v1/messages',
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': getApiKey('ANTHROPIC_API_KEY'),
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: MODEL,
        max_tokens: MAX_TOKENS,
        temperature: TEMPERATURE,
        system: systemPrompt,
        messages: [{ role: 'user', content: userMsg }]
      }),
      muteHttpExceptions: true
    }
  );

  // Check HTTP status
  const code = response.getResponseCode();
  const raw  = response.getContentText();
  if (code !== 200) {
    console.error(`Anthropic API error ${code}: ${raw}`);
    sheet.getRange(rowNum, outputCol).setValue(`ERROR: HTTP ${code}`);
    return 0;
  }

  // Parse JSON defensively
  let text = '';
  try {
    const obj = JSON.parse(raw);

    if (Array.isArray(obj.content)) {
      // Standard Claude response
      text = obj.content[0]?.text;
    } else if (typeof obj.completion === 'string') {
      // Fallback if it returned a single string
      text = obj.completion;
    } else if (obj.choices?.length) {
      // OpenAI‐style
      text = obj.choices[0].message?.content;
    } else {
      throw new Error('Unexpected response format');
    }

    text = (text || '').trim();
  } catch (e) {
    console.error('Response parsing error:', e, raw);
    sheet.getRange(rowNum, outputCol).setValue('ERROR: Invalid response');
    return 0;
  }

  // Write back to sheet
  sheet.getRange(rowNum, outputCol)
       .setValue(text)
       .setFontColor('red')
       .setFontStyle('italic');

  // Pause to respect rate limits
  Utilities.sleep(2000);
  return 1;
}

function applyBatchUpdates(sheet, updates) {
  if (updates.length === 0) return;

  console.log(`Applying ${updates.length} batch updates...`);
  
  // Group updates by row for efficiency
  const rowUpdates = {};
  
  for (const update of updates) {
    if (!rowUpdates[update.row]) {
      rowUpdates[update.row] = [];
    }
    rowUpdates[update.row].push(update);
  }

  // Apply updates row by row (still more efficient than individual cells)
  for (const [row, rowUpdateList] of Object.entries(rowUpdates)) {
    try {
      for (const update of rowUpdateList) {
        sheet.getRange(parseInt(row), update.col).setValue(update.value);
      }
    } catch (error) {
      console.error(`Error updating row ${row}: ${error.message}`);
    }
  }
  
  console.log('Batch updates completed');
}

function applyBatchUpdates(sheet, updates) {
  if (updates.length === 0) return;

  console.log(`Applying ${updates.length} batch updates...`);
  
  // Group updates by row for efficiency
  const rowUpdates = {};
  
  for (const update of updates) {
    if (!rowUpdates[update.row]) {
      rowUpdates[update.row] = [];
    }
    rowUpdates[update.row].push(update);
  }

  // Apply updates row by row (still more efficient than individual cells)
  for (const [row, rowUpdateList] of Object.entries(rowUpdates)) {
    try {
      for (const update of rowUpdateList) {
        sheet.getRange(parseInt(row), update.col).setValue(update.value);
      }
    } catch (error) {
      console.error(`Error updating row ${row}: ${error.message}`);
    }
  }
  
  console.log('Batch updates completed');
}




function getBrightDataSnapshot(snapshotId, apiKey) {
  const endpoint = `https://api.brightdata.com/datasets/v3/snapshot/${snapshotId}`;
  
  const headers = {
    "Authorization": `Bearer ${apiKey}`
  };
  
  const options = {
    method: 'GET',
    headers: headers,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // Handle both 200 (complete) and 202 (running) as valid responses
    if (responseCode === 200 || responseCode === 202) {
      const result = JSON.parse(responseText);
      return result;
    } else {
      throw new Error(`HTTP ${responseCode}: ${responseText}`);
    }
    
  } catch (error) {
    console.error(`Error getting Bright Data snapshot: ${error.message}`);
    throw error;
  }
}

function formatProfileForMCP(profileJson) {
  const profile = profileJson.profile || profileJson;
  const lines = [];

  // Basic info
  if (profile.name) {
    lines.push(`Name: ${profile.name}`);
  }
  if (profile.city) {
    lines.push(`City: ${profile.city}`);
  }
  if (profile.position) {
    lines.push(`Position: ${profile.position}`);
  }
  if (profile.about) {
    lines.push(`\nAbout: ${profile.about}`);
  }
  if (profile.location) {
    lines.push(`Location: ${profile.location}`);
  }

  // Posts
  if (Array.isArray(profile.posts) && profile.posts.length > 0) {
    lines.push(`\nPosts:`);
    profile.posts.forEach(post => {
      const title = post.title || '(no title)';
      lines.push(`• ${title}`);
    });
  }

  // Current company
  if (profile.current_company) {
    const cc = profile.current_company;
    lines.push(`\nCurrent Company:`);
    if (cc.name)     lines.push(`• Name: ${cc.name}`);
    if (cc.title)    lines.push(`• Title: ${cc.title}`);
    if (cc.location) lines.push(`• Location: ${cc.location}`);
  }

  // Experience (up to first 15)
  if (Array.isArray(profile.experience) && profile.experience.length > 0) {
    lines.push(`\nExperience:`);
    profile.experience.slice(0, 15).forEach(exp => {
      const title     = exp.title      || '(no title)';
      const company   = exp.company    || '(no company)';
      const location  = exp.location   || '';
      const startDate = exp.start_date || '';
      const endDate   = exp.end_date   || '';
      const desc      = exp.description || '';
      let entry = `• ${title} at ${company}`;
      if (location)  entry += `, ${location}`;
      if (startDate || endDate) {
        entry += ` (${startDate || '?'} – ${endDate || '?'})`;
      }
      lines.push(entry);
      if (desc) {
        lines.push(`   ${desc}`);
      }
    });
  }

  // Education (up to first 10)
  if (Array.isArray(profile.education) && profile.education.length > 0) {
    lines.push(`\nEducation:`);
    profile.education.slice(0, 10).forEach(edu => {
      const school    = edu.title      || '(no school)';
      const degree    = edu.degree     || '';
      const field     = edu.field      || '';
      const startYear = edu.start_year || '';
      const endYear   = edu.end_year   || '';
      const desc      = edu.description || '';
      let entry = `• ${school}`;
      if (degree) entry += ` – ${degree}`;
      if (field)  entry += ` in ${field}`;
      if (startYear || endYear) {
        entry += ` (${startYear || '?'} – ${endYear || '?'})`;
      }
      lines.push(entry);
      if (desc) {
        lines.push(`   ${desc}`);
      }
    });
  }

  // Certifications
  if (Array.isArray(profile.certifications) && profile.certifications.length > 0) {
    lines.push(`\nCertifications:`);
    profile.certifications.forEach(cert => {
      const title = cert.title || '(no title)';
      const meta  = cert.meta  || cert.subtitle || '';
      lines.push(meta
        ? `• ${title} (${meta})`
        : `• ${title}`);
    });
  }

  // Organizations
  if (Array.isArray(profile.organizations) && profile.organizations.length > 0) {
    lines.push(`\nOrganizations:`);
    profile.organizations.forEach(org => {
      const title = org.title           || '(no title)';
      const type  = org.membership_type || '';
      const start = org.start_date      || '';
      const end   = org.end_date        || '';
      let entry = `• ${title}`;
      if (type)  entry += ` — ${type}`;
      if (start || end) {
        entry += ` (${start || '?'} – ${end || '?'})`;
      }
      lines.push(entry);
    });
  }

  return lines.join('\n');
}

/**
 * Fetch all email‐accounts (senders) from Apollo and write them
 * into a hidden lookup sheet in the *active* spreadsheet.
 */
function refreshSenders() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  let lookupSheet = ss.getSheetByName('Senders_Lookup');

  // create & hide if missing, otherwise clear
  if (!lookupSheet) {
    lookupSheet = ss.insertSheet('Senders_Lookup');
    lookupSheet.hideSheet();
  } else {
    lookupSheet.clearContents();
  }

  // header row
  lookupSheet.getRange(1, 1, 1, 3)
             .setValues([['ID','Email','Provider']]);

  // fetch from Apollo
  const apolloKey = getApiKey('APOLLO_API_KEY');
  const resp = UrlFetchApp.fetch('https://api.apollo.io/api/v1/email_accounts', {
    method: 'get',
    contentType: 'application/json',
    headers: { 'X-Api-Key': apolloKey },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch senders: HTTP ${resp.getResponseCode()}`);
  }

  const accounts = JSON.parse(resp.getContentText()).email_accounts || [];
  if (accounts.length) {
    const rows = accounts.map(a => [ a.id, a.email, a.provider_display_name ]);
    lookupSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

/**
 * Fetch all email sequences via Apollo’s Search endpoint
 * and write them into a hidden lookup sheet in the *active* spreadsheet.
 */
function refreshSequences() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  let lookupSheet = ss.getSheetByName('Sequences_Lookup');

  // create & hide if missing, otherwise clear
  if (!lookupSheet) {
    lookupSheet = ss.insertSheet('Sequences_Lookup');
    lookupSheet.hideSheet();
  } else {
    lookupSheet.clearContents();
  }

  // header row
  lookupSheet.getRange(1, 1, 1, 2)
             .setValues([['ID','Name']]);

  const apolloKey = getApiKey('APOLLO_API_KEY');
  const perPage   = 100;
  let   page      = 1;
  let   allSeqs   = [];
  let   batch;

  // page through until fewer than perPage results
  do {
    const resp = UrlFetchApp.fetch('https://api.apollo.io/api/v1/emailer_campaigns/search', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Api-Key': apolloKey },
      payload: JSON.stringify({
        q_name:   '',
        page:     String(page),
        per_page: String(perPage)
      }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch sequences (page ${page}): HTTP ${resp.getResponseCode()}`);
    }

    batch    = JSON.parse(resp.getContentText()).emailer_campaigns || [];
    allSeqs  = allSeqs.concat(batch);
    page++;
  } while (batch.length === perPage);

  // write them out
  if (allSeqs.length) {
    const rows = allSeqs.map(s => [ s.id, s.name ]);
    lookupSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

/**
 * Upload contacts flagged in FIND_OWNER_INFO_COL to Apollo sequences.
 * CORRECTED VERSION based on Clay community discussion about typed_custom_fields
 *
 * For each row flagged "1":
 *   ① Verify the email sequence exists
 *   ② Ensure an Account (company) exists
 *   ③ Create or match the Contact
 *   ④ Set custom field using typed_custom_fields
 *   ⑤ Enrol the contact in the chosen sequence
 */
function uploadContacts() {
  // --- As configurações globais e da planilha continuam iguais ---
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = activeSS.getSheetByName(getConfig('SHEET_NAME'));
  const apolloKey = getApiKey('APOLLO_API_KEY');

  // --- CORREÇÃO: Buscando TODAS as colunas do jeito novo ---
  const flagCol = letterToColumn(getColumnConfig('FIND_OWNER_INFO_COL_LETTER'));
  const emailCol = letterToColumn(getColumnConfig('OWNER_EMAIL_COL_LETTER'));
  const ownerCol = letterToColumn(getColumnConfig('OWNER_COL_LETTER'));
  const companyCol = letterToColumn(getColumnConfig('COMPANY_COL_LETTER'));
  const websiteCol = letterToColumn(getColumnConfig('WEBSITE_COL_LETTER'));
  const seqFriendlyCol = letterToColumn(getColumnConfig('SEQUENCE_ID_COL_LETTER'));
  const senderFriendlyCol = letterToColumn(getColumnConfig('SENDER_ID_COL_LETTER'));
  const customInfoCol = letterToColumn(getColumnConfig('OUTPUT_COL_LETTER'));
  const customIndustryCol = letterToColumn(getColumnConfig('CUSTOM_INDUSTRY_COL_LETTER'));

  // --- O resto da função continua exatamente igual ---
  // your custom field IDs
  const customFieldId = '606b3624b9802d00fe84ac4a';
  const customIndustryFieldId = '681b6bfa10baf6000de50185';

const seqMap    = buildLookupMap(activeSS, 'Sequences_Lookup');
const senderMap = buildLookupMap(activeSS, 'Senders_Lookup');

  const rows = sheet.getDataRange().getValues();
  let processed = 0;

  rows.slice(1).forEach((row, idx) => {
    const sheetRow = idx + 2;
    if (row[flagCol - 1] !== 1) return;

    const email = (row[emailCol - 1] || '').trim();
    const owner = (row[ownerCol - 1] || '').trim();
    const companyName = (row[companyCol - 1] || '').trim();
    const rawSite = (row[websiteCol - 1] || '').trim();
    const seqName = (row[seqFriendlyCol - 1] || '').trim();
    const senderName = (row[senderFriendlyCol - 1] || '').trim();
    const customInfo = (row[customInfoCol - 1] || '').trim();
    const customIndustry = (row[customIndustryCol - 1] || '').trim();

    const sequenceId = seqMap[seqName];
    const senderId = senderMap[senderName];
    if (!email || !sequenceId || !senderId) {
      sheet.getRange(sheetRow, flagCol).setValue('');
      return;
    }

    // 1️⃣ Sequence exists?
    const seqResp = UrlFetchApp.fetch(
      `https://api.apollo.io/api/v1/emailer_campaigns/${sequenceId}`,
      { method: 'get', headers: { 'X-Api-Key': apolloKey }, muteHttpExceptions: true }
    );
    if (seqResp.getResponseCode() !== 200) {
      sheet.getRange(sheetRow, flagCol).setValue('');
      return;
    }

    // 2️⃣ Ensure Account exists
    const domain = rawSite ? extractDomain(rawSite) : '';
    const websiteUrl = rawSite ? normaliseWebsite(rawSite) : '';
    const accountId = companyName
      ? findOrCreateAccount(companyName, domain, apolloKey)
      : '';

    // 3️⃣ Find or create the Contact
    const contactId = findOrCreateContact(
      email, owner, companyName, websiteUrl, accountId, apolloKey,
      sheet, sheetRow, flagCol
    );
    if (!contactId) return;

    // 4️⃣ Set custom fields using the CORRECT Apollo API approach
    if (customInfo || customIndustry) {
      console.log(`Setting custom fields for contact ${contactId}`);
      if (customInfo) console.log(`  Custom info: ${customInfo}`);
      if (customIndustry) console.log(`  Custom industry: ${customIndustry}`);
      
      Utilities.sleep(1000);
      
      const customFieldsSet = setCustomFieldsCorrectly(contactId, customFieldId, customInfo, customIndustryFieldId, customIndustry, apolloKey);
      if (!customFieldsSet) {
        console.log(`FAILED to set custom fields for contact ${contactId}`);
        sheet.getRange(sheetRow, flagCol).setValue('');
        return;
      }
      
      Utilities.sleep(2000);
    }

    // 5️⃣ Enrol in sequence
    const enrol = UrlFetchApp.fetch(
      `https://api.apollo.io/api/v1/emailer_campaigns/${sequenceId}/add_contact_ids`,
      {
        method: 'post',
        contentType: 'application/json',
        headers: { 'X-Api-Key': apolloKey },
        payload: JSON.stringify({
          emailer_campaign_id: sequenceId,
          contact_ids: [contactId],
          send_email_from_email_account_id: senderId
        }),
        muteHttpExceptions: true
      }
    );

    if ([200, 201].includes(enrol.getResponseCode())) {
      sheet.getRange(sheetRow, flagCol).setValue('');
      processed++;
    } else {
      console.log(`Enrollment failed: ${enrol.getContentText()}`);
      sheet.getRange(sheetRow, flagCol).setValue('');
    }

    Utilities.sleep(1000);
  });

  SpreadsheetApp.getUi().alert(`Uploaded ${processed} contacts to Apollo.`);
}

/**
 * Set custom fields using the CORRECT Apollo API approach
 * Now handles both custom info and custom industry fields
 */
function setCustomFieldsCorrectly(contactId, customFieldId, customValue, customIndustryFieldId, customIndustryValue, apolloKey) {
  console.log(`Setting custom fields for contact ${contactId}`);
  
  // Build the typed_custom_fields object with only the fields that have values
  const customFields = {};
  
  if (customValue) {
    customFields[customFieldId] = customValue;
    console.log(`  Adding custom info field: ${customValue}`);
  }
  
  if (customIndustryValue) {
    customFields[customIndustryFieldId] = customIndustryValue;
    console.log(`  Adding custom industry field: ${customIndustryValue}`);
  }
  
  // Only make the API call if we have fields to set
  if (Object.keys(customFields).length === 0) {
    console.log('No custom fields to set');
    return true;
  }

  // Use the approach from the Clay discussion
  const updateResp = UrlFetchApp.fetch(
    `https://api.apollo.io/api/v1/contacts/${contactId}`, {
      method: 'put',
      contentType: 'application/json',
      headers: {
        'Cache-Control': 'no-cache',
        'Content-Type': 'application/json',
        'accept': 'application/json',
        'X-Api-Key': apolloKey
      },
      payload: JSON.stringify({
        typed_custom_fields: customFields
      }),
      muteHttpExceptions: true
    }
  );

  const responseCode = updateResp.getResponseCode();
  const responseText = updateResp.getContentText();
  
  console.log(`Custom fields update response: ${responseCode}`);
  console.log(`Response body: ${responseText}`);

  if ([200, 201].includes(responseCode)) {
    console.log(`✅ Successfully set custom fields for contact ${contactId}`);
    return true;
  } else {
    console.log(`❌ Failed to set custom fields for contact ${contactId}: ${responseText}`);
    return false;
  }
}

/**
 * Find an existing contact by email or create a new one.
 */
function findOrCreateContact(
  email, owner, companyName, websiteUrl, accountId, apolloKey,
  sheet, rowIdx, flagCol
) {
  // A) look for existing
  const search = UrlFetchApp.fetch(
    'https://api.apollo.io/api/v1/contacts/search', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Api-Key': apolloKey },
      payload: JSON.stringify({ q_email: email, per_page: 1 }),
      muteHttpExceptions: true
    }
  );
  
  if (search.getResponseCode() === 200) {
    const found = JSON.parse(search.getContentText()).contacts || [];
    if (found.length) {
      console.log(`Found existing contact: ${found[0].id}`);
      return found[0].id;
    }
  }

  // B) create new contact
  const parts = owner.split(/\s+/);
  const payload = {
    email,
    first_name: parts[0] || '',
    last_name: parts.length > 1 ? parts.pop() : '',
    organization_name: companyName || undefined,
    website_url: websiteUrl || undefined,
    account_id: accountId || undefined
  };

  const create = UrlFetchApp.fetch(
    'https://api.apollo.io/api/v1/contacts', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Api-Key': apolloKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  if ([200, 201].includes(create.getResponseCode())) {
    const createdContact = JSON.parse(create.getContentText()).contact;
    console.log(`Created new contact: ${createdContact.id}`);
    return createdContact.id;
  } else {
    console.log(`Contact creation failed: ${create.getContentText()}`);
    sheet.getRange(rowIdx, flagCol).setValue('');
    return null;
  }
}

// ── helper functions ──
function buildLookupMap(ss, sheetName) {
  const data = ss.getSheetByName(sheetName).getDataRange().getValues().slice(1);
  const m = {};
  data.forEach(r => { if (r[1]) m[r[1].toString().trim()] = r[0].toString().trim(); });
  return m;
}

function extractDomain(site) {
  return site.replace(/^https?:\/\//i, '').replace(/^www\./i, '').split(/[\/?#]/)[0].trim();
}

function normaliseWebsite(site) {
  let u = site.trim();
  if (!/^https?:\/\//i.test(u)) u = 'https://' + u;
  return u.replace(/\/.*$/, '');
}

function findOrCreateAccount(name, domain, apolloKey) {
  const resp = UrlFetchApp.fetch(
    'https://api.apollo.io/api/v1/accounts/search', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Api-Key': apolloKey },
      payload: JSON.stringify({ q_organization_name: name, per_page: 1 }),
      muteHttpExceptions: true
    }
  );
  if (resp.getResponseCode() === 200) {
    const list = JSON.parse(resp.getContentText()).accounts || [];
    if (list.length) return list[0].id;
  }
  const pay = { name };
  if (domain) pay.domain = domain;
  const cr = UrlFetchApp.fetch(
    'https://api.apollo.io/api/v1/accounts', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Api-Key': apolloKey },
      payload: JSON.stringify(pay),
      muteHttpExceptions: true
    }
  );
  return [200,201].includes(cr.getResponseCode())
    ? JSON.parse(cr.getContentText()).account.id
    : '';
}

/**
 * Returns the final domain part of the MX records for the domain of the given email address.
 *
 * For example:
 *   "1 sobreirocapital-pt.mail.protection.outlook.com." returns "outlook"
 *   "50 alt4.aspmx.l.google.com., 20 alt1.aspmx.l.google.com., ..." returns "google"
 *
 * @param {string} email The email address to look up.
 * @return {string} A comma-separated list of unique extracted domains from the MX records, or an error message.
 * @customfunction
 */


function getMXDomain(email) {
  if (!email || email.indexOf("@") === -1) {
    return "Invalid email address";
  }
  
  // Extract the domain from the email address.
  var domain = email.split("@")[1];
  // Query Google's DNS-over-HTTPS API for MX records.
  var url = "https://dns.google/resolve?name=" + domain + "&type=MX";
  
  try {
    var response = UrlFetchApp.fetch(url);
    var data = JSON.parse(response.getContentText());
    
    if (data.Status === 0 && data.Answer) {
      // Use a Set to collect unique extracted domains.
      var domainSet = new Set();
      // Regex to capture the last two segments of a domain that end with a period.
      var regex = /([^.]+\.[^.]+)\.$/;
      
      for (var i = 0; i < data.Answer.length; i++) {
        var answer = data.Answer[i];
        if (answer.type === 15) { // MX record type is 15.
          var match = answer.data.match(regex);
          if (match && match[1]) {
            // Example: "outlook.com" -> split on "." and take the first part.
            var fullDomain = match[1];
            var parts = fullDomain.split(".");
            if (parts.length > 0) {
              domainSet.add(parts[0]);
            }
          }
        }
      }
      
      if (domainSet.size > 0) {
        return Array.from(domainSet).join(", ");
      } else {
        return "No MX domain part found";
      }
    } else {
      return "No MX records found or an error occurred";
    }
  } catch (e) {
    return "Error fetching MX records: " + e;
  }
}

function refreshLookups() {
   refreshSenders();
   refreshSequences();
   applyLookupDropdowns();
}

function getApiKey(keyName) {
  const val = PropertiesService
                .getScriptProperties()
                .getProperty(keyName);
  if (!val) {
    throw new Error(`${keyName} not found in Script Properties`);
  }
  return val;
}


function requireAdmin() {
  const user = Session.getActiveUser().getEmail();
  if (!ADMIN_USERS.includes(user)) {
    throw new Error('Only administrators may run this.');
  }
}

/**
 * After column‑setup, apply data‑validation dropdowns to any blank
 * Sequence & Sender cells, using the *external* lookup spreadsheet.
 */
/**
 * After setup, apply dynamic dropdowns to any blank Sequence‑ID
 * and Sender‑ID cells by pointing to the active sheet’s lookup ranges.
 */
function applyLookupDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Identify your data sheet & columns
  const dataSheetName = getConfig('SHEET_NAME');
  if (!dataSheetName) throw new Error('SHEET_NAME not configured');
  const dataSheet = ss.getSheetByName(dataSheetName);
  if (!dataSheet) throw new Error(`Sheet "${dataSheetName}" not found`);

  const seqCol = letterToColumn(getColumnConfig('SEQUENCE_ID_COL_LETTER'));
  const senderCol = letterToColumn(getColumnConfig('SENDER_ID_COL_LETTER'));

  // 2. Get the lookup sheets in this spreadsheet
  const seqLookup    = ss.getSheetByName('Sequences_Lookup');
  const senderLookup = ss.getSheetByName('Senders_Lookup');
  if (!seqLookup || !senderLookup) {
    throw new Error('Lookup sheets Sequences_Lookup and/or Senders_Lookup are missing');
  }

  // 3. Build dynamic ranges starting at B2 down to the last row
  const seqLastRow    = seqLookup.getLastRow();
  const senderLastRow = senderLookup.getLastRow();
  const seqRange      = seqLookup.getRange(2, 2, seqLastRow - 1, 1);
  const senderRange   = senderLookup.getRange(2, 2, senderLastRow - 1, 1);

  // 4. Create data‑validation rules that reference those ranges
  const seqRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(seqRange, true)
    .setAllowInvalid(false)
    .build();
  const senderRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(senderRange, true)
    .setAllowInvalid(false)
    .build();

  // 5. Loop rows 2–10 and apply the rule if the cell is blank
  for (let row = 2; row <= 10; row++) {
    const seqCell    = dataSheet.getRange(row, seqCol);
    const senderCell = dataSheet.getRange(row, senderCol);

    if (!seqCell.getValue())    seqCell.setDataValidation(seqRule);
    if (!senderCell.getValue()) senderCell.setDataValidation(senderRule);
  }
}


/**
 * REVISED - Searches for a LinkedIn profile using Bright Data's SERP API.
 * @param {string} ownerName The name of the person to search for.
 * @param {string} companyName The company name to refine the search.
 * @returns {string} The URL of the first LinkedIn profile found, or a status message.
 */
function findFirstLinkedInResult(ownerName, companyName) {
  // 1. Validate inputs
  if (!ownerName || !companyName) {
    console.log('Skipping search: Missing ownerName or companyName.');
    return 'Missing name or company';
  }
  console.log(`Searching for: "${ownerName}", "${companyName}"`);

  // 2. Construct the query
  const query = `site:linkedin.com/in/ "${ownerName}" "${companyName}"`;
  
  try {
    // 3. Call the SERP API (your searchWithBrightData function has its own logging)
    console.log(`Executing search with query: ${query}`);
    const results = searchWithBrightData(query, 1); 

    // 4. Process the results
    if (results && results.length > 0) {
      const firstResult = results[0];
      console.log(`SUCCESS: Found LinkedIn URL -> ${firstResult}`);
      // Basic validation that it's a LinkedIn URL
      if (firstResult.includes('linkedin.com/in/')) {
        return firstResult;
      } else {
        console.log(`WARNING: Result found, but may not be a valid Profile URL: ${firstResult}`);
        return 'Found, but not a profile URL';
      }
    } else {
      console.log(`INFO: No results returned from BrightData for this query.`);
      return 'Not found';
    }
  } catch (e) {
    // 5. Handle any errors during the process
    console.error(`ERROR: The 'searchWithBrightData' call failed. Details: ${e.message}`);
    return `API Error`;
  }
}


/**
 * =================================================================
 * Main Standalone Orchestration Function (Optimized for Speed)
 * =================================================================
 * This function controls the entire data enrichment process,
 * running API calls in concurrent batches for maximum efficiency.
 *
 * NOTE: This function depends on helper functions (e.g., getConfig,
 * getApiKey, letterToColumn, normalizeDomain, etc.) which must be
 * defined elsewhere in your Apps Script project.
 */
function enrichData() {
  const fn = 'enrichData';
  try {
    console.log('🚀 Starting the main enrichment process (Optimized)...');
    const startTime = new Date();

  // --- General Setup ---
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig('SHEET_NAME'));
    if (!sheet) {
      console.error(`Sheet with name "${getConfig('SHEET_NAME')}" not found.`);
      SpreadsheetApp.getUi().alert(`Error: Sheet "${getConfig('SHEET_NAME')}" not found.`);
      logAction(fn, 'Failed: sheet not found');
      return;
    }
  const allData = sheet.getDataRange().getValues();

  // --- API Keys & Column Configs ---
  const openaiKey = getApiKey('OPENAI_API_KEY');
  const apolloKey = getApiKey('APOLLO_API_KEY');
  const flagCol = letterToColumn(getColumnConfig('FIND_OWNER_INFO_COL_LETTER'));
  const websiteCol = letterToColumn(getColumnConfig('WEBSITE_COL_LETTER'));
  const industryCol = letterToColumn(getColumnConfig('CUSTOM_INDUSTRY_COL_LETTER'));
  const ownerCol = letterToColumn(getColumnConfig('OWNER_COL_LETTER'));
  const emailCol = letterToColumn(getColumnConfig('OWNER_EMAIL_COL_LETTER'));
  const linkedinCol = letterToColumn(getColumnConfig('OWNER_LINKEDIN_COL_LETTER'));
  const companyCol = letterToColumn(getColumnConfig('COMPANY_COL_LETTER'));

  // --- Identify rows to process ---
  const flaggedRows = [];
  for (let i = 1; i < allData.length; i++) {
    if (Number(allData[i][flagCol - 1]) === 1) {
      flaggedRows.push({
        index: i,
        rowNum: i + 1,
        data: allData[i]
      });
    }
  }

  if (flaggedRows.length === 0) {
    SpreadsheetApp.getUi().alert('No rows were flagged with "1" to process.');
    return;
  }
  console.log(`Found ${flaggedRows.length} rows to process.`);
  let allSheetUpdates = [];

  try {
    // =================================================================
    // Step 1 & 2: Find Industry & Owner Name (ChatGPT) in Batch
    // =================================================================
    console.log('--- Steps 1 & 2: Building ChatGPT requests...');
    const chatGptRequests = [];
    const chatGptMeta = []; // To map responses back to rows/columns

    flaggedRows.forEach(({ index, rowNum, data }) => {
      const website = data[websiteCol - 1] || '';
      if (!website) return;

      // --- Industry Descriptor Request ---
      const existingIndustry = (data[industryCol - 1] || '').trim();
      if (!existingIndustry) {
        const promptText = `You must respond with ONLY the industry descriptor - no explanations, no additional text, no citations, no sentences.\n\nFor the website "${website}", provide a 1-3 word noun phrase describing their primary offering.\n\nRequirements:\n- Must be 1-3 words maximum\n- Must be lowercase (except acronyms like HVAC, IT)\n- Must NOT include "Solutions" or "Services"\n- Must fit naturally in: "I found your company while searching for [YOUR ANSWER] business"\n- Must be based on how they describe themselves\n\nRespond with ONLY the descriptor phrase, nothing else:`;
        const payload = { model: "gpt-4.1", tools: [{ type: "web_search_preview", search_context_size: "medium", user_location: { type: "approximate", country: "US" } }], tool_choice: { type: "web_search_preview" }, input: promptText };
        chatGptRequests.push({ url: 'https://api.openai.com/v1/responses', method: 'post', contentType: 'application/json', headers: { 'Authorization': `Bearer ${openaiKey}` }, payload: JSON.stringify(payload), muteHttpExceptions: true });
        chatGptMeta.push({ rowNum, col: industryCol, type: 'industry' });
      }

      // --- Owner Name Request ---
      const existingOwner = (data[ownerCol - 1] || '').trim();
      if (!existingOwner) {
        const promptText = `Your task is to identify the primary owner, founder, or CEO for the website: ${website}.\n\nFollow these rules exactly:\n1. You MUST respond with ONLY the person's first and last name.\n2. If you cannot find the name, you MUST respond with the exact text "Not found".`;
        const payload = { model: "gpt-4.1", tools: [{ type: "web_search_preview", search_context_size: "medium", user_location: { type: "approximate", country: "US" } }], tool_choice: { type: "web_search_preview" }, input: promptText };
        chatGptRequests.push({ url: 'https://api.openai.com/v1/responses', method: 'post', contentType: 'application/json', headers: { 'Authorization': `Bearer ${openaiKey}` }, payload: JSON.stringify(payload), muteHttpExceptions: true });
        chatGptMeta.push({ rowNum, index, col: ownerCol, type: 'owner' });
      }
    });

    if (chatGptRequests.length > 0) {
      console.log(`Executing ${chatGptRequests.length} ChatGPT requests in parallel...`);
      const responses = UrlFetchApp.fetchAll(chatGptRequests);
      responses.forEach((resp, i) => {
        const meta = chatGptMeta[i];
        let value = 'Error';
        if (resp.getResponseCode() === 200) {
          const rd = JSON.parse(resp.getContentText());
          if (rd.output_text) { value = rd.output_text.trim(); }
          else if (Array.isArray(rd.output)) { const msg = rd.output.find(o => o.type === 'message'); const txt = msg?.content?.find(c => c.type === 'output_text'); if (txt) value = txt.text.trim(); }
        }

        if (meta.type === 'owner' && !/^[A-Za-zÀ-ÖØ-öø-ÿ]+ [A-Za-zÀ-ÖØ-öø-ÿ]+$/.test(value)) {
          value = 'Not found';
        }
        if (meta.type === 'industry') {
          value = value.replace(/,/g, '').trim();
        }

        allSheetUpdates.push({ row: meta.rowNum, col: meta.col, value });
        // Update in-memory data for subsequent steps
        if (meta.type === 'owner') {
          allData[meta.index][ownerCol - 1] = value;
        }
      });
    }
    console.log('✅ ChatGPT lookups complete.');

    // =================================================================
    // Step 3: Find Owner Details from Name (Apollo) in Batch
    // =================================================================
    console.log('--- Step 3: Building Apollo-by-Name requests...');
    const apolloNameRequests = [];
    const apolloNameMeta = [];

    flaggedRows.forEach(({ index, rowNum, data }) => {
      const owner = allData[index][ownerCol - 1] || ''; // Use potentially updated data
      const website = data[websiteCol - 1] || '';
      if (!owner || owner === 'Not found' || !website) return;

      const existingEmail = data[emailCol - 1];
      const existingLinkedIn = data[linkedinCol - 1];
      if (existingEmail && existingLinkedIn) return;

      const parts = owner.trim().split(' ');
      if (parts.length < 2) return;

      const payload = { first_name: parts[0], last_name: parts[parts.length - 1], domain: normalizeDomain(website), reveal_linkedin_url: true, reveal_personal_emails: true, reveal_company_emails: true };
      apolloNameRequests.push({ url: 'https://api.apollo.io/api/v1/people/match', method: 'post', contentType: 'application/json', headers: { 'X-Api-Key': apolloKey }, payload: JSON.stringify(payload), muteHttpExceptions: true });
      apolloNameMeta.push({ rowNum, index, existingEmail, existingLinkedIn });
    });

    if (apolloNameRequests.length > 0) {
      console.log(`Executing ${apolloNameRequests.length} Apollo-by-Name requests in parallel...`);
      const responses = UrlFetchApp.fetchAll(apolloNameRequests);
      responses.forEach((resp, i) => {
        const meta = apolloNameMeta[i];
        if (resp.getResponseCode() === 200) {
          const person = (JSON.parse(resp.getContentText()).person) || {};
          const email = person.company_email || person.email || '';
          const linkedin = person.linkedin_url || '';
          if (!meta.existingEmail && email) {
            allSheetUpdates.push({ row: meta.rowNum, col: emailCol, value: email });
            allData[meta.index][emailCol - 1] = email;
          }
          if (!meta.existingLinkedIn && linkedin) {
            allSheetUpdates.push({ row: meta.rowNum, col: linkedinCol, value: linkedin });
            allData[meta.index][linkedinCol - 1] = linkedin;
          }
        }
      });
    }
    console.log('✅ Apollo name lookup complete.');

    // =================================================================
    // Step 4: Find Missing LinkedIn URLs (BrightData) - Sequential
    // NOTE: This step remains sequential as it involves its own search logic.
    // =================================================================
    console.log('--- Step 4: Finding Missing LinkedIn URLs (BrightData)...');
    flaggedRows.forEach(({ index, rowNum, data }) => {
      const ownerName = allData[index][ownerCol - 1];
      const companyName = data[companyCol - 1];
      const existingUrl = allData[index][linkedinCol - 1];

      if (existingUrl || !ownerName || ownerName === 'Not found' || !companyName) return;

      console.log(`▶️ Row ${rowNum}: BrightData lookup for ${ownerName} at ${companyName}`);
      const linkedinUrl = findFirstLinkedInResult(ownerName, companyName);
      if (linkedinUrl && linkedinUrl !== 'Not found' && !linkedinUrl.startsWith('Missing') && !linkedinUrl.startsWith('API Error')) {
        allSheetUpdates.push({ row: rowNum, col: linkedinCol, value: linkedinUrl });
        allData[index][linkedinCol - 1] = linkedinUrl;
      }
      Utilities.sleep(1500); // Rate-limit this sequential step
    });
    console.log('✅ BrightData lookup complete.');

    // =================================================================
    // Step 5: Find Owner Email from LinkedIn (Apollo) in Batch
    // =================================================================
    console.log('--- Step 5: Building Apollo-by-LinkedIn requests...');
    const apolloLinkedInRequests = [];
    const apolloLinkedInMeta = [];

    flaggedRows.forEach(({ index, rowNum, data }) => {
      if (allData[index][emailCol - 1]) return; // Skip if email already found
      const linkedinUrl = allData[index][linkedinCol - 1] || '';
      if (!linkedinUrl) return;

      const payload = { linkedin_url: linkedinUrl, reveal_personal_emails: true, reveal_company_emails: true };
      apolloLinkedInRequests.push({ url: 'https://api.apollo.io/api/v1/people/match', method: 'post', contentType: 'application/json', headers: { 'X-Api-Key': apolloKey }, payload: JSON.stringify(payload), muteHttpExceptions: true });
      apolloLinkedInMeta.push({ rowNum, col: emailCol });
    });

    if (apolloLinkedInRequests.length > 0) {
      console.log(`Executing ${apolloLinkedInRequests.length} Apollo-by-LinkedIn requests in parallel...`);
      const responses = UrlFetchApp.fetchAll(apolloLinkedInRequests);
      responses.forEach((resp, i) => {
        const meta = apolloLinkedInMeta[i];
        if (resp.getResponseCode() === 200) {
          const person = JSON.parse(resp.getContentText()).person || {};
          const emailFound = person.company_email || person.email || '';
          if (emailFound) {
            allSheetUpdates.push({ row: meta.rowNum, col: meta.col, value: emailFound });
          }
        }
      });
    }
    console.log('✅ Apollo LinkedIn lookup complete.');

    // =================================================================
    // Step 6: Finalize and update flags
    // =================================================================
    console.log('--- Step 6: Applying all updates to the sheet...');
    flaggedRows.forEach(({ rowNum }) => {
      allSheetUpdates.push({ row: rowNum, col: flagCol, value: '' });
    });

    applyBatchUpdates(sheet, allSheetUpdates);
    console.log('✅ All updates applied.');

    const duration = Math.round((new Date() - startTime) / 1000);
    console.log(`✅🎉 Enrichment process completed successfully in ${duration} seconds!`);
    SpreadsheetApp.getUi().alert(`Success!`, `Enriched ${flaggedRows.length} rows.`, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    console.error(`A critical error occurred during the enrichment process: ${e.toString()}`, e.stack);
    SpreadsheetApp.getUi().alert('An unexpected error occurred. Please check the logs for details.');
  }
  
  // Catch any unexpected errors from the outer try block
} catch (e) {
  console.error('Unexpected failure in enrichData:', e);
  SpreadsheetApp.getUi().alert('An unexpected error occurred. Please check the logs for details.');
}

}

/**
 * Fetch from Anthropic with a 30‑calls/minute shared rate limit.
 */
function rateLimitedAnthropicFetch(url, options) {
  const CACHE_KEY = 'anthropic_calls';
  const WINDOW_MS = 60 * 1000;     // 60 sec window
  const CAPACITY  = 30;             // max 30 calls per window
  const BACKOFF   = 5000;          // if 429, wait 5 sec
  
  const cache = CacheService.getScriptCache();
  const lock  = LockService.getScriptLock();
  lock.waitLock(30*1000);  // wait up to 30 sec to get the lock
  
  try {
    // 1️⃣ Read & parse the timestamp list
    let raw = cache.get(CACHE_KEY);
    let times = raw ? JSON.parse(raw) : [];
    
    const now = Date.now();
    // 2️⃣ Discard anything older than WINDOW_MS
    times = times.filter(t => now - t < WINDOW_MS);
    
    if (times.length >= CAPACITY) {
      // 3️⃣ Too many calls in window ⇒ compute sleep needed
      const oldest = Math.min(...times);
      const waitMs = WINDOW_MS - (now - oldest);
      Utilities.sleep(waitMs);
      
      // update now & purge again
      const again = Date.now();
      times = times.filter(t => again - t < WINDOW_MS);
    }
    
    // 4️⃣ Record this call
    times.push(Date.now());
    cache.put(CACHE_KEY, JSON.stringify(times), 120);  // cache expires after 2 min
    
  } finally {
    lock.releaseLock();
  }
  
  // 5️⃣ Do the actual fetch
  const resp = UrlFetchApp.fetch(url, options);
  
  // 6️⃣ If we got a 429, back off and retry once
  if (resp.getResponseCode() === 429) {
    Utilities.sleep(BACKOFF);
    return UrlFetchApp.fetch(url, options);
  }
  return resp;
}

/**
 * NEW - Orchestrates a bulk, in-memory scrape and generates a full email for each flagged row.
 */
function createFullEmail() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig('SHEET_NAME'));
  if (!sheet) {
    console.error(`Sheet with name "${getConfig('SHEET_NAME')}" not found.`);
    ui.alert(`Error: Sheet "${getConfig('SHEET_NAME')}" not found.`);
    return;
  }
  const flagCol = letterToColumn(getColumnConfig('FIND_OWNER_INFO_COL_LETTER'));
  const outputCol = letterToColumn(getColumnConfig('OUTPUT_COL_LETTER'));
  const data = sheet.getDataRange().getValues();
  const startTime = new Date();

  // Find all rows flagged with "1"
  const flaggedRows = data.map((row, index) => ({ row, index }))
                         .filter(item => item.row[flagCol - 1] === 1);

  if (flaggedRows.length === 0) {
    ui.alert('🛑 No rows flagged with "1" to process.');
    return;
  }

  // Prompt user before starting, similar to custom info customization
  const n = flaggedRows.length;
  const alertMessage =
      `You are about to start the customization process for ${n} row(s).\n\n` +
      `This may take a few minutes. Please do not close or refresh this sheet until you see the final "Process Complete!" message.\n\n` +
      `Press OK to begin.`;

  const response = ui.alert('Start Customization?', alertMessage, ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
    ui.alert('Process cancelled.');
    return;
  }

  console.log(`Starting bulk email creation for ${flaggedRows.length} rows...`);

  // 1️⃣ & 2️⃣ Scrape all data in memory first
  const linkedInScrapes = ScrapeOwnerLinkedInOptimized(flaggedRows);
  const webpageScrapes = ScrapeWebpagesDuckduckgoOptimized(flaggedRows);
  const allSheetUpdates = [];
  let emailsCreated = 0;

  // 3️⃣ Process each flagged row to generate an email
  console.log('Generating emails using in-memory data...');
  flaggedRows.forEach(item => {
    const rowNum = item.index + 1;
    const baseRowData = item.row;

    // Assemble all scraped data for this row
    const rowWebScrapes = webpageScrapes[rowNum] || {};
    const scrapedTexts = {
      linkedinProfileText: linkedInScrapes[rowNum] || '',
      homepage: rowWebScrapes.homepage || '',
      aboutUs: rowWebScrapes.aboutUs || '',
      ownerNameCompany: rowWebScrapes.ownerNameCompany || '',
      pages: rowWebScrapes.pages || []
    };

    try {
      console.log(`Generating email for row ${rowNum}`);
      
      // Build the prompt for a full email
      const systemPrompt = buildSystemPrompt2();

      // Gather the row's inputs from both sheet and scraped data
      const get = key => baseRowData[letterToColumn(getColumnConfig(key)) - 1] || '';
      const inputs = {
        company: get('COMPANY_COL_LETTER'),
        ownerName: get('OWNER_COL_LETTER'),
        homepage: scrapedTexts.homepage,
        aboutUs: scrapedTexts.aboutUs,
        ownerComp: scrapedTexts.ownerNameCompany,
        linkedin: scrapedTexts.linkedinProfileText,
        page1: scrapedTexts.pages[0] || '',
        page2: scrapedTexts.pages[1] || '',
        page3: scrapedTexts.pages[2] || '',
        page4: scrapedTexts.pages[3] || '',
        page5: scrapedTexts.pages[4] || ''
      };

      const userMsg = JSON.stringify(inputs, null, 2);

      // Call Anthropic API
      const resp = rateLimitedAnthropicFetch(
        'https://api.anthropic.com/v1/messages', {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'x-api-key': getApiKey('ANTHROPIC_API_KEY'),
            'anthropic-version': '2023-06-01'
          },
          payload: JSON.stringify({
            model: MODEL,
            max_tokens: MAX_TOKENS,
            temperature: TEMPERATURE,
            system: systemPrompt,
            messages: [{ role: 'user', content: userMsg }]
          }),
          muteHttpExceptions: true
        }
      );

      if (resp.getResponseCode() === 200) {
        const email = JSON.parse(resp.getContentText()).content[0].text.trim();
        allSheetUpdates.push({ row: rowNum, col: outputCol, value: email });
        emailsCreated++;
        try {
          const cfg = ensureConfigSheet();
          cfg.getRange(LAST_EMAIL_ROW, 2).setValue(email);
        } catch (err) {
          console.error('Failed to save latest email text:', err);
        }
      } else {
        const errorMsg = `ERROR: HTTP ${resp.getResponseCode()}`;
        allSheetUpdates.push({ row: rowNum, col: outputCol, value: errorMsg });
        console.error(`Email generation failed for row ${rowNum}: ${resp.getContentText()}`);
      }
      
    } catch (e) {
      const errorMsg = `ERROR: ${e.message}`;
      allSheetUpdates.push({ row: rowNum, col: outputCol, value: errorMsg });
      console.error(`Email generation failed for row ${rowNum}:`, e);
    }
  });

  // 4️⃣ Add flag updates and apply everything to the sheet
  flaggedRows.forEach(item => {
    allSheetUpdates.push({
      row: item.index + 1,
      col: flagCol,
      value: ''
    });
  });
  
  applyBatchUpdates(sheet, allSheetUpdates);

  // 5️⃣ Final summary
  const duration = Math.round((Date.now() - startTime) / 1000);
  SpreadsheetApp.flush(); // Ensure UI is updated before showing the alert

  ui.alert(
    'Process Complete!',
    `Processed ${flaggedRows.length} rows.\n\n` +
    `• Emails created: ${emailsCreated}`,
    ui.ButtonSet.OK
  );
  logAction(fn, 'Completed');
}

/**
 * Prompt the admin for feedback to change the email optimization style.
 * Saves Claude's suggested prompt to row EMAIL_PROMPT_ROW of // Config.
 */
function changeEmailOptimizationStyle() {
  requireAdmin();
  const ui = SpreadsheetApp.getUi();
  const cfg = ensureConfigSheet();

  const resp = ui.prompt(
    'Change Email Optimization Style',
    'Enter your feedback for improving the email customization:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const feedback = resp.getResponseText().trim();
  if (!feedback) {
    ui.alert('No feedback entered.');
    return;
  }

  let guidelines = DEFAULT_EMAIL_GUIDELINES;

  const lastEmail = cfg.getRange(LAST_EMAIL_ROW, 2).getValue();

  const instruction =
    'Rewrite the email customization prompt based on the user feedback. ' +
    'Use the original guidelines and the previous email text as context if provided. ' +
    'IMPORTANT: Respond ONLY with the text of the new prompt and nothing else.';

  const parts = [
    instruction,
    'ORIGINAL_EMAIL_GUIDELINES:\n' + guidelines,
  ];
  if (lastEmail) parts.push('LATEST_EMAIL_CUSTOMIZATION:\n' + lastEmail);
  parts.push('USER_FEEDBACK:\n' + feedback);

  const prompt = parts.join('\n\n');

  const apiResp = rateLimitedAnthropicFetch(
    'https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': getApiKey('ANTHROPIC_API_KEY'),
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: MODEL,
        max_tokens: MAX_TOKENS,
        temperature: TEMPERATURE,
        system: '',
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    }
  );

  if (apiResp.getResponseCode() !== 200) {
    ui.alert('Claude error: ' + apiResp.getContentText());
    return;
  }

  let newPrompt;
  try {
    newPrompt = JSON.parse(apiResp.getContentText()).content[0].text.trim();
  } catch (e) {
    ui.alert('Failed to parse Claude response');
    return;
  }

  const currentPrompt = readConfig('EMAIL_GUIDELINES');
  if (currentPrompt) {
    cfg.getRange(PREVIOUS_PROMPT_ROW, 2).setValue(currentPrompt);
  }
  writeConfig('EMAIL_GUIDELINES', newPrompt);
  ui.alert('Email optimization prompt updated.');
}

function revertToPreviousCustomization() {
  const ui = SpreadsheetApp.getUi();
  const cfg = ensureConfigSheet();
  const prev = cfg.getRange(PREVIOUS_PROMPT_ROW, 2).getValue();
  if (!prev) {
    ui.alert('No previous customization style found.');
    return;
  }
  const current = cfg.getRange(EMAIL_PROMPT_ROW, 2).getValue();
  cfg.getRange(PREVIOUS_PROMPT_ROW, 2).setValue(current);
  cfg.getRange(EMAIL_PROMPT_ROW, 2).setValue(prev);
  ui.alert('Reverted to previous customization style.');
}

function revertToDefaultCustomization() {
  const cfg = ensureConfigSheet();
  const current = cfg.getRange(EMAIL_PROMPT_ROW, 2).getValue();
  cfg.getRange(PREVIOUS_PROMPT_ROW, 2).setValue(current);
  cfg.getRange(EMAIL_PROMPT_ROW, 2).setValue(DEFAULT_EMAIL_GUIDELINES);
  SpreadsheetApp.getUi().alert('Reverted to default email customization.');
}

function changeCustomInfoStyle() {
  const fn = 'changeCustomInfoStyle';
  try {
    requireAdmin();
    const ui = SpreadsheetApp.getUi();
    const cfg = ensureConfigSheet();

  const resp = ui.prompt(
    'Change Custom Info Style',
    'Enter your feedback for improving the custom info paragraph:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const feedback = resp.getResponseText().trim();
  if (!feedback) {
    ui.alert('No feedback entered.');
    return;
  }

  let guidelines = DEFAULT_CUSTOM_INFO_GUIDELINES;

  const lastInfo = cfg.getRange(LAST_INFO_ROW, 2).getValue();

  const instruction =
    'Rewrite the custom info prompt based on the user feedback. ' +
    'Use the original guidelines and the previous paragraph as context if provided. ' +
    'IMPORTANT: Respond ONLY with the text of the new prompt and nothing else.';

  const parts = [
    instruction,
    'ORIGINAL_CUSTOM_INFO_GUIDELINES:\n' + guidelines,
  ];
  if (lastInfo) parts.push('LATEST_CUSTOM_INFO:\n' + lastInfo);
  parts.push('USER_FEEDBACK:\n' + feedback);

  const prompt = parts.join('\n\n');

  const apiResp = rateLimitedAnthropicFetch(
    'https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': getApiKey('ANTHROPIC_API_KEY'),
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: MODEL,
        max_tokens: MAX_TOKENS,
        temperature: TEMPERATURE,
        system: '',
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    }
  );

  if (apiResp.getResponseCode() !== 200) {
    ui.alert('Claude error: ' + apiResp.getContentText());
    return;
  }

  let newPrompt;
  try {
    newPrompt = JSON.parse(apiResp.getContentText()).content[0].text.trim();
  } catch (e) {
    ui.alert('Failed to parse Claude response');
    return;
  }

  const currentPrompt = readConfig('CUSTOM_INFO_GUIDELINES');
  if (currentPrompt) {
    cfg.getRange(PREVIOUS_INFO_ROW, 2).setValue(currentPrompt);
  }
  writeConfig('CUSTOM_INFO_GUIDELINES', newPrompt);
  // Ensure the visible cell shows the updated prompt immediately
  cfg.getRange(INFO_PROMPT_ROW, 2).setValue(newPrompt);
  ui.alert('Custom info prompt updated.');
    logAction(fn, 'Completed');
  } catch (e) {
    logAction(fn, 'Failed: ' + e.message);
    throw e;
  }
}

function revertToPreviousInfoStyle() {
  const fn = 'revertToPreviousInfoStyle';
  try {
    const ui = SpreadsheetApp.getUi();
    const cfg = ensureConfigSheet();
  const prev = cfg.getRange(PREVIOUS_INFO_ROW, 2).getValue();
  if (!prev) {
    ui.alert('No previous custom info style found.');
    return;
  }
  const current = cfg.getRange(INFO_PROMPT_ROW, 2).getValue();
  cfg.getRange(PREVIOUS_INFO_ROW, 2).setValue(current);
  cfg.getRange(INFO_PROMPT_ROW, 2).setValue(prev);
  ui.alert('Reverted to previous custom info style.');
    logAction(fn, 'Completed');
  } catch (e) {
    logAction(fn, 'Failed: ' + e.message);
    throw e;
  }
}

function revertToDefaultInfoStyle() {
  const fn = 'revertToDefaultInfoStyle';
  try {
    const cfg = ensureConfigSheet();
    const current = cfg.getRange(INFO_PROMPT_ROW, 2).getValue();
    cfg.getRange(PREVIOUS_INFO_ROW, 2).setValue(current);
    cfg.getRange(INFO_PROMPT_ROW, 2).setValue(DEFAULT_CUSTOM_INFO_GUIDELINES);
    SpreadsheetApp.getUi().alert('Reverted to default custom info style.');
    logAction(fn, 'Completed');
  } catch (e) {
    logAction(fn, 'Failed: ' + e.message);
    throw e;
  }
}

/**
 * MODIFIED - Orchestrates the entire scrape-and-customize process in memory.
 * Scrapes LinkedIn and webpages, passes the data directly to the AI,
 * and writes only the final customization to the sheet.
 */
function runCombinedScrapesOptimized() {
  const fn = 'runCombinedScrapesOptimized';
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig('SHEET_NAME'));
    const flagCol = letterToColumn(getColumnConfig('FIND_OWNER_INFO_COL_LETTER'));
    const data = sheet.getDataRange().getValues();
    const startTime = new Date();

  const flaggedRows = data.map((row, index) => ({ row, index })) // Keep original index
                         .filter(item => item.row[flagCol - 1] === 1);

    if (flaggedRows.length === 0) {
      ui.alert('🛑 No rows flagged with "1" to process.');
      logAction(fn, 'Failed: no rows flagged');
      return;
    }
  // ---  START: ADDED UI ALERT ---
  const n = flaggedRows.length;
  const alertMessage = 
      `You are about to start the customization process for ${n} row(s).\n\n` +
      `This may take a few minutes. Please do not close or refresh this sheet until you see the final "Process Complete!" message.\n\n` +
      `Press OK to begin.`;
      
  const response = ui.alert('Start Customization?', alertMessage, ui.ButtonSet.OK_CANCEL);

  // If user clicks "Cancel" or closes the dialog, stop the script.
  if (response !== ui.Button.OK) {
    ui.alert('Process cancelled.');
    return;
  }
  // ---  END: ADDED UI ALERT ---

  console.log(`Starting in-memory scrape and customization for ${flaggedRows.length} rows...`);

  // 1️⃣ & 2️⃣ Scrape all data and hold it in memory
  const linkedInScrapes = ScrapeOwnerLinkedInOptimized(flaggedRows);
  const webpageScrapes = ScrapeWebpagesDuckduckgoOptimized(flaggedRows);

  // 3️⃣ Run AI customizations using in-memory data
  let customCount = 0;
  console.log('Starting customizations using in-memory data...');

  flaggedRows.forEach(item => {
    const rowNum = item.index + 1; // 1-based row number
    const baseRowData = item.row;
    
    // Assemble all scraped data for this row
    const rowWebScrapes = webpageScrapes[rowNum] || {};
    const scrapedTexts = {
      linkedinProfileText: linkedInScrapes[rowNum] || '',
      homepage: rowWebScrapes.homepage || '',
      aboutUs: rowWebScrapes.aboutUs || '',
      ownerNameCompany: rowWebScrapes.ownerNameCompany || '',
      pages: rowWebScrapes.pages || []
    };

    try {
      console.log(`Processing customization for row ${rowNum}`);
      // Pass the sheet data and the in-memory scraped data to the customization function
      customCount += runCustomization(rowNum, sheet, baseRowData, scrapedTexts);
    } catch (e) {
      console.error(`Customization failed for row ${rowNum}:`, e);
    }
  });
  console.log('Customizations completed.');

  // 4️⃣ Update flags in batch
  const flagUpdates = flaggedRows.map(item => ({
    row: item.index + 1,
    col: flagCol,
    value: ''
  }));
  applyBatchUpdates(sheet, flagUpdates);

  // 5️⃣ Final summary
  const duration = Math.round((Date.now() - startTime) / 1000);

  // 👇 *** ADD THIS LINE ***
  SpreadsheetApp.flush(); 

  ui.alert(
    'Process Complete!',
    `Processed ${flaggedRows.length} rows.\n\n` +
    `• Customizations created: ${customCount}`,
    ui.ButtonSet.OK
  );
    logAction(fn, 'Completed');
  } catch (e) {
    logAction(fn, 'Failed: ' + e.message);
    throw e;
  }
}

/**
 * MODIFIED - Scrapes LinkedIn profiles and returns the data as an object
 * instead of writing it to the sheet.
 * @param {Array} flaggedRows The array of rows to process.
 * @return {Object} An object mapping row number to the scraped LinkedIn text.
 */
function ScrapeOwnerLinkedInOptimized(flaggedRows) {
  console.log('Starting LinkedIn scraping (in-memory)...');
  // const outputCol = letterToColumn(getColumnConfig('OWNER_LINKEDIN_SCRAPE_COL_LETTER')); // <<< THIS LINE WAS REMOVED
  const inputCol = letterToColumn(getColumnConfig('OWNER_LINKEDIN_COL_LETTER'));
  const apiKey = getApiKey('BRIGHTDATA_API_KEY');
  const datasetId = getApiKey('BRIGHTDATA_DATASET_ID');
  
  const scrapeResults = {}; // { rowNum: "scraped text", ... }
  const snapshots = [];

  // Step 1: Trigger all scrapes
  flaggedRows.forEach(item => {
    const rowNum = item.index + 1;
    const linkedinUrl = item.row[inputCol - 1] || '';
    if (!linkedinUrl) return;

    try {
      const snapshotId = triggerBrightDataScrape(linkedinUrl, apiKey, datasetId);
      if (snapshotId) {
        snapshots.push({ row: rowNum, snapshotId: snapshotId });
      } else {
        scrapeResults[rowNum] = 'ERROR: Failed to get snapshot ID';
      }
    } catch (error) {
      console.error(`Error triggering scrape for row ${rowNum}: ${error.message}`);
      scrapeResults[rowNum] = `ERROR: ${error.message}`;
    }
    Utilities.sleep(2000);
  });

  // Step 2: Poll all snapshots
  const maxRetries = 40;
  let retryCount = 0;
  let completedSnapshots = new Set();

  while (retryCount < maxRetries && completedSnapshots.size < snapshots.length) {
    Utilities.sleep(3000);
    retryCount++;
    
    for (const snapshot of snapshots) {
      if (completedSnapshots.has(snapshot.snapshotId)) continue;
      
      try {
        const profileData = getBrightDataSnapshot(snapshot.snapshotId, apiKey);
        if (profileData && profileData.status !== 'running') {
          scrapeResults[snapshot.row] = formatProfileForMCP(profileData);
          completedSnapshots.add(snapshot.snapshotId);
        }
      } catch (error) {
        scrapeResults[snapshot.row] = `ERROR: ${error.message}`;
        completedSnapshots.add(snapshot.snapshotId);
      }
    }
  }

  // Handle timeouts
  snapshots.forEach(snapshot => {
    if (!completedSnapshots.has(snapshot.snapshotId)) {
      scrapeResults[snapshot.row] = `TIMEOUT: Snapshot still running.`;
    }
  });
  
  console.log('LinkedIn scraping (in-memory) completed.');
  return scrapeResults;
}


/**
 * MODIFIED - Scrapes webpages and returns the data as an object
 * instead of writing it to the sheet.
 * @param {Array} flaggedRows The array of rows to process.
 * @return {Object} A nested object mapping row number to its various webpage scrapes.
 * e.g., { 2: { homepage: "...", aboutUs: "...", pages: [...] } }
 */
function ScrapeWebpagesDuckduckgoOptimized(flaggedRows) {
  console.log('Starting webpage scraping (in-memory)...');
  const brightDataKey = getApiKey('BRIGHTDATA_API_KEY');
  const unlockerZone = getApiKey('BRIGHTDATA_UNLOCKER_ZONE');
  const endpoint = 'https://api.brightdata.com/request';
  const MAX_CELL_LEN = 8000;
  const clamp = str => str.length > MAX_CELL_LEN ? str.slice(0, MAX_CELL_LEN) + '…' : str;

  // Column indexes
  const cfg = k => letterToColumn(getColumnConfig(k));
  const websiteCol = cfg('WEBSITE_COL_LETTER');
  const ownerCol = cfg('OWNER_COL_LETTER');
  const companyCol = cfg('COMPANY_COL_LETTER');
  const pageCols = ['PAGE1_COL_LETTER', 'PAGE2_COL_LETTER', 'PAGE3_COL_LETTER', 'PAGE4_COL_LETTER', 'PAGE5_COL_LETTER'];

  const fetchRequests = [];
  const meta = []; // Holds { row, key } to map responses back
  
  flaggedRows.forEach(item => {
    const rowNum = item.index + 1;
    const rowData = item.row;
    const website = rowData[websiteCol - 1];
    if (!website) return;

    const makeReq = url => ({
      url: endpoint, method: 'post', contentType: 'application/json',
      headers: { 'Authorization': `Bearer ${brightDataKey}` },
      payload: JSON.stringify({ zone: unlockerZone, url: url, format: 'json' }),
      muteHttpExceptions: true
    });
    
    // 1. Homepage
    fetchRequests.push(makeReq(website));
    meta.push({ row: rowNum, key: 'homepage' });

    // 2. About Us
    const domain = normalizeDomain(website);
    const aboutQ = `site:${domain} -site:linkedin.com -site:m.linkedin.com ("about us" OR "about")`;
    const aboutUrl = searchWithBrightData(aboutQ, 1)[0] || null;
    fetchRequests.push(makeReq(aboutUrl || ''));
    meta.push({ row: rowNum, key: 'aboutUs' });

    // 3. Owner+Company
    const owner = rowData[ownerCol - 1];
    const company = rowData[companyCol - 1];
    if (owner && company) {
      const rawUrls = searchWithBrightData(`"${owner}" "${company}"`, 5);
      const ownerUrl = rawUrls.find(u => !/linkedin\.com/.test(u)) || null;
      fetchRequests.push(makeReq(ownerUrl || ''));
      meta.push({ row: rowNum, key: 'ownerNameCompany' });
    }

    // 4. General pages
    const generalUrls = searchWithBrightData(`site:${domain}`, 5);
    pageCols.forEach((colKey, j) => {
      fetchRequests.push(makeReq(generalUrls[j] || ''));
      meta.push({ row: rowNum, key: `page${j + 1}` });
    });
  });

  // Fire requests and process responses
  const responses = UrlFetchApp.fetchAll(fetchRequests);
  const scrapeResults = {}; // { rowNum: { key: "text", ... } }

  responses.forEach((resp, idx) => {
    const { row, key } = meta[idx];
    if (!scrapeResults[row]) {
      scrapeResults[row] = { pages: [] };
    }
    
    let txt = '';
    if (resp.getResponseCode() !== 200) {
      txt = `ERROR: HTTP ${resp.getResponseCode()}`;
    } else {
      let body = resp.getContentText();
      let obj;
      try {
        obj = JSON.parse(body);
        body = obj.html || obj.response?.html || obj.data?.html || obj.body || body;
      } catch (e) { /* leave body as-is */ }
      txt = clamp(stripHtml(body));
    }
    
    if (key.startsWith('page')) {
      const pageIndex = parseInt(key.replace('page', ''), 10) - 1;
      scrapeResults[row].pages[pageIndex] = txt;
    } else {
      scrapeResults[row][key] = txt;
    }
  });

  console.log('Webpage scraping (in-memory) completed.');
  return scrapeResults;
}


/**
 * MODIFIED - Generates customization from data passed in memory.
 * @param {number} rowNum 1-based sheet row to process.
 * @param {Sheet} sheet The Sheet object.
 * @param {Array} baseRowData The original data array for the row.
 * @param {Object} scrapedTexts An object with all scraped text for the row.
 * @return {number} 1 if customization succeeded, 0 on error.
 */
function runCustomization(rowNum, sheet, baseRowData, scrapedTexts) {
  // Resolve column indexes needed for base data and output
  const companyCol = letterToColumn(getColumnConfig('COMPANY_COL_LETTER'));
  const ownerCol = letterToColumn(getColumnConfig('OWNER_COL_LETTER'));
  const outputCol = letterToColumn(getColumnConfig('OUTPUT_COL_LETTER'));

  // Get base data from the provided row array
  const company = baseRowData[companyCol - 1] || '';
  const owner = baseRowData[ownerCol - 1] || '';

  // Get scraped data from the provided object
  const {
    homepage,
    aboutUs,
    ownerNameCompany,
    linkedinProfileText,
    pages
  } = scrapedTexts;

  // Build prompts
  const systemPrompt = buildSystemPrompt();
  const userMsg = [
    `COMPANY: ${company}`,
    `HOMEPAGE: ${homepage}`,
    `ABOUT_US: ${aboutUs}`,
    `OWNER: ${owner}`,
    `OWNER_NAME_COMPANY_TEXT: ${ownerNameCompany}`,
    `LINKEDIN_PROFILE_TEXT: ${linkedinProfileText}`,
    `SCRAPED_PAGES:\n${pages.join('\n\n')}`
  ].join('\n');

  // Call Anthropic API
  const response = rateLimitedAnthropicFetch(
    'https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': getApiKey('ANTHROPIC_API_KEY'),
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: MODEL,
        max_tokens: MAX_TOKENS,
        temperature: TEMPERATURE,
        system: systemPrompt,
        messages: [{ role: 'user', content: userMsg }]
      }),
      muteHttpExceptions: true
    }
  );

  // Process response and write to sheet
  const code = response.getResponseCode();
  const raw = response.getContentText();
  if (code !== 200) {
    console.error(`Anthropic API error ${code}: ${raw}`);
    sheet.getRange(rowNum, outputCol).setValue(`ERROR: HTTP ${code}`);
    return 0;
  }

  let text = '';
  try {
    const obj = JSON.parse(raw);
    if (Array.isArray(obj.content)) {
      text = obj.content[0]?.text;
    } else {
       throw new Error('Unexpected response format');
    }
    text = (text || '').trim();
  } catch (e) {
    console.error('Response parsing error:', e, raw);
    sheet.getRange(rowNum, outputCol).setValue('ERROR: Invalid response');
    return 0;
  }

  // Write final output back to sheet
  sheet.getRange(rowNum, outputCol)
    .setValue(text)
    .setFontColor('red')
    .setFontStyle('italic');

  // Store this customization for reference
  try {
    const cfg = ensureConfigSheet();
    cfg.getRange(LAST_INFO_ROW, 2).setValue(text);
  } catch (e) {
    console.error('Failed to save latest email text:', e);
  }

  Utilities.sleep(2000); // Respect rate limits
  return 1;
}

