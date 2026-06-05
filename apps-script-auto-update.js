/**
 * ═══════════════════════════════════════════════════════
 *  100 Ways to Die in India — Google Apps Script
 *  Paste this entire file into:
 *  Google Sheet → Extensions → Apps Script → Code.gs
 * ═══════════════════════════════════════════════════════
 *
 *  What it does every night at 11 PM:
 *  1. Searches Google News RSS for India tragedy/negligence stories
 *  2. Calls Claude API to extract structured case details
 *  3. Appends a new row to your Google Sheet
 *  4. Marks the new case as "latest" and clears the flag on old rows
 *
 *  One-time setup:
 *  1. Paste this into Apps Script
 *  2. Set your Anthropic API key in Script Properties:
 *     Project Settings → Script Properties → Add:
 *       Key: ANTHROPIC_API_KEY   Value: sk-ant-...
 *  3. Run setupTrigger() once manually to install the nightly cron
 *  4. Make sure your sheet has these headers in row 1:
 *     id | date | tag | cat | headline | body | source_url | image_url | is_latest
 */

// ── CONFIG ────────────────────────────────────────────
const SHEET_NAME = 'Sheet1';

// Search terms rotated daily to get diverse cases
const SEARCH_QUERIES = [
  'India death negligence fire accident',
  'India stampede death infrastructure collapse',
  'India road accident death bridge collapse',
  'India rape murder woman killed',
  'India factory explosion worker killed',
  'India flooding death drainage failure',
  'India hospital death negligence oxygen',
  'India manual scavenging sewer death',
  'India child death borewell school',
  'India pothole death accident road',
  'India train accident derailment death',
  'India building collapse death',
  'India farmer suicide debt crop failure',
  'India heatwave death casualty',
  'India custodial death police',
];

const UNSPLASH_BY_CATEGORY = {
  fire:        'https://images.unsplash.com/photo-1557804506-669a67965ba0?w=1600&q=80',
  road:        'https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?w=1600&q=80',
  women:       'https://images.unsplash.com/photo-1516975080664-ed2fc6a32937?w=1600&q=80',
  student:     'https://images.unsplash.com/photo-1503676260728-1c00da094a0b?w=1600&q=80',
  infra:       'https://images.unsplash.com/photo-1504280390367-361c6d9f38f4?w=1600&q=80',
  medical:     'https://images.unsplash.com/photo-1519494026892-80bbd2d6fd0d?w=1600&q=80',
  labour:      'https://images.unsplash.com/photo-1504307651254-35680f356dfd?w=1600&q=80',
  stampede:    'https://images.unsplash.com/photo-1533105079780-92b9be482077?w=1600&q=80',
  system:      'https://images.unsplash.com/photo-1469398715555-76331a6e7b53?w=1600&q=80',
  environment: 'https://images.unsplash.com/photo-1547683905-f686c993aae5?w=1600&q=80',
};
// ─────────────────────────────────────────────────────

/**
 * Install the nightly trigger. Run this once manually.
 */
function setupTrigger() {
  // Remove existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runNightlyUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(23)   // 11 PM in the script's timezone (set in Project Settings → Time zone → Asia/Kolkata)
    .create();

  Logger.log('✅ Nightly trigger installed for 11 PM');
}

/**
 * Main entry point — called by the nightly trigger.
 */
function runNightlyUpdate() {
  try {
    Logger.log('🔄 Starting nightly update: ' + new Date().toISOString());

    const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set in Script Properties');

    // Pick a search query based on day of year for variety
    const dayIndex = Math.floor(Date.now() / 86400000) % SEARCH_QUERIES.length;
    const query = SEARCH_QUERIES[dayIndex];
    Logger.log('🔍 Search query: ' + query);

    // 1. Fetch news RSS
    const articles = fetchNewsRSS(query);
    if (articles.length === 0) {
      Logger.log('⚠️ No articles found. Skipping update.');
      return;
    }
    Logger.log(`📰 Found ${articles.length} articles`);

    // 2. Pick the most relevant one (not already in sheet)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const existingUrls = getExistingSourceUrls(sheet);
    const fresh = articles.filter(a => !existingUrls.has(a.url));

    if (fresh.length === 0) {
      Logger.log('ℹ️ All articles already in sheet. Skipping.');
      return;
    }

    const article = fresh[0];
    Logger.log('📄 Processing: ' + article.title);

    // 3. Call Claude to extract structured data
    const caseData = extractCaseWithClaude(apiKey, article);
    if (!caseData) {
      Logger.log('❌ Claude extraction returned null. Skipping.');
      return;
    }

    // 4. Get next ID
    const lastRow = sheet.getLastRow();
    const nextId = lastRow <= 1 ? 1 : parseInt(sheet.getRange(lastRow, 1).getValue()) + 1;

    // 5. Clear all existing is_latest flags
    if (lastRow > 1) {
      sheet.getRange(2, 9, lastRow - 1, 1).setValue('false');
    }

    // 6. Append new row
    const imageUrl = UNSPLASH_BY_CATEGORY[caseData.cat] || UNSPLASH_BY_CATEGORY.system;
    const newRow = [
      nextId,
      caseData.date,
      caseData.tag,
      caseData.cat,
      caseData.headline,
      caseData.body,
      article.url,
      imageUrl,
      'true'   // is_latest
    ];
    sheet.appendRow(newRow);

    Logger.log(`✅ Added case #${nextId}: ${caseData.tag} — ${caseData.date}`);

  } catch (err) {
    Logger.log('❌ Error in runNightlyUpdate: ' + err.message);
    // Optionally send email alert:
    // MailApp.sendEmail('your@email.com', 'Script error', err.message);
  }
}

/**
 * Fetch top articles from Google News RSS for a query.
 */
function fetchNewsRSS(query) {
  const encoded = encodeURIComponent(query);
  const url = `https://news.google.com/rss/search?q=${encoded}+when:2d&hl=en-IN&gl=IN&ceid=IN:en`;

  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const xml = resp.getContentText();
    const doc = XmlService.parse(xml);
    const items = doc.getRootElement().getChild('channel').getChildren('item');

    return items.slice(0, 5).map(item => ({
      title:       item.getChildText('title')       || '',
      description: item.getChildText('description') || '',
      url:         item.getChildText('link')         || '',
      pubDate:     item.getChildText('pubDate')      || '',
    })).filter(a => a.url && a.title);

  } catch (err) {
    Logger.log('RSS fetch error: ' + err.message);
    return [];
  }
}

/**
 * Call Claude API to turn a raw article into a structured case.
 */
function extractCaseWithClaude(apiKey, article) {
  const today = new Date();
  const dateStr = today.toLocaleDateString('en-IN', { day:'numeric', month:'long', year:'numeric' });

  const prompt = `You are the editor of "100 Ways to Die in India" — a dark satire website documenting avoidable deaths caused by systemic negligence, government failure, and institutional rot.

A news article has been found. Extract and return ONLY a JSON object with these fields:

{
  "date": "date of incident as a short string e.g. '${dateStr}'",
  "tag": "short category label (e.g. 'Fire Safety', 'Infrastructure', 'Gender Violence', 'Road Safety', 'Medical Negligence', 'Caste Violence', 'Food Safety', 'Labour Safety', 'Stampede', 'Climate Negligence')",
  "cat": "one of: fire | road | women | student | infra | medical | labour | stampede | system | environment",
  "headline": "a stark, punchy 1-2 line headline in the site's voice — taunting the system. Use <em> tags around the most damning phrase. Be brutal. Examples: 'She filed 3 complaints. Police said <em>sort it out</em>.' or '32 people went for a day of fun. <em>9 were children.</em>'",
  "body": "2-3 sentence context paragraph. Include the location, what happened, what the systemic failure was, and what pathetic response followed (if any). Factual. Bitter. Under 80 words.",
  "is_avoidable": true
}

Only include deaths that were CLEARLY avoidable — caused by negligence, systemic failure, corruption, or institutional rot. NOT wars, terrorism, or purely natural causes.

If this article is not about an avoidable death in India, return: {"skip": true}

ARTICLE TITLE: ${article.title}
ARTICLE DESCRIPTION: ${article.description}
SOURCE URL: ${article.url}

Return ONLY the JSON. No explanation. No markdown. No backticks.`;

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',  // fast and cheap for this task
        max_tokens: 600,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());
    if (data.error) throw new Error(data.error.message);

    const text = data.content[0].text.trim();
    const parsed = JSON.parse(text);

    if (parsed.skip) {
      Logger.log('⏭️ Claude marked article as non-avoidable death. Skipping.');
      return null;
    }

    return parsed;

  } catch (err) {
    Logger.log('Claude API error: ' + err.message);
    return null;
  }
}

/**
 * Get all source URLs already in the sheet (column G = index 7).
 */
function getExistingSourceUrls(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Set();
  const values = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  return new Set(values.flat().filter(Boolean));
}

/**
 * Run this manually to test with a sample article.
 */
function testWithSampleArticle() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  const sample = {
    title: 'Three sanitation workers die in Hyderabad sewer cleaning accident',
    description: 'Three municipal sanitation workers died after inhaling toxic gases while cleaning a sewer in Hyderabad\'s old city area. The workers had no protective equipment. Officials confirmed none of the mandatory safety protocols were followed.',
    url: 'https://example.com/test-article',
    pubDate: new Date().toString()
  };

  const result = extractCaseWithClaude(apiKey, sample);
  Logger.log('Test result: ' + JSON.stringify(result, null, 2));
}

/**
 * One-time: Set up the sheet headers.
 */
function setupSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = ['id', 'date', 'tag', 'cat', 'headline', 'body', 'source_url', 'image_url', 'is_latest'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  // Set column widths for readability
  sheet.setColumnWidth(1, 50);   // id
  sheet.setColumnWidth(2, 120);  // date
  sheet.setColumnWidth(3, 150);  // tag
  sheet.setColumnWidth(4, 80);   // cat
  sheet.setColumnWidth(5, 400);  // headline
  sheet.setColumnWidth(6, 500);  // body
  sheet.setColumnWidth(7, 300);  // source_url
  sheet.setColumnWidth(8, 300);  // image_url
  sheet.setColumnWidth(9, 80);   // is_latest
  Logger.log('✅ Sheet headers set up');
}
