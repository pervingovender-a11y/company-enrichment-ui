export default function Home() {
  return (
    <div className="min-h-screen bg-slate-950 text-slate-100">
      <main className="mx-auto flex w-full max-w-4xl flex-col gap-8 px-6 py-16">
        <header className="flex flex-col gap-3">
          <p className="text-sm font-semibold uppercase tracking-[0.2em] text-cyan-300">
            Google Apps Script
          </p>
          <h1 className="text-3xl font-semibold text-white sm:text-4xl">
            Scrape key company data from website rows in Sheets
          </h1>
          <p className="text-base text-slate-300">
            This improved starter reads each website URL from your sheet, fetches
            the HTML, throttles requests to avoid rate limits, and captures a
            structured payload (title, description, contact links, phone, email).
            Results are written back to the same row.
          </p>
        </header>

        <section className="rounded-2xl border border-slate-800 bg-slate-900/60 p-6 shadow-lg">
          <h2 className="text-lg font-semibold text-white">Apps Script</h2>
          <p className="mt-2 text-sm text-slate-300">
            Paste this complete script into a single Apps Script file attached
            to your sheet and run
            <span className="font-semibold text-white"> enrichSheet()</span>.
          </p>
          <p className="mt-3 text-sm text-slate-400">
            The script adds an “Enrichment” menu with Start, Pause, Resume, and
            Stop controls to help manage long runs.
          </p>
          <p className="mt-3 text-sm text-slate-400">
            Results are appended to a separate “Output” sheet in batches of
            1000 rows while still writing progress and status in the source
            sheet.
          </p>
          <p className="mt-3 text-sm text-slate-400">
            Scraping respects robots.txt for each domain and captures fetch
            failures without crashing the run.
          </p>
          <p className="mt-3 text-sm text-slate-400">
            Runs auto-resume on long sheets using a timed trigger when nearing
            Apps Script execution limits.
          </p>
          <pre className="mt-4 overflow-x-auto rounded-xl bg-slate-950 p-4 text-sm text-slate-100">
            <code>{`/**
 * Reads a sheet with Website URLs and enriches each row.
 * Expected columns: first_name | last_name | email | Website
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Enrichment')
    .addItem('Start', 'startEnrichment')
    .addItem('Pause', 'pauseEnrichment')
    .addItem('Resume', 'resumeEnrichment')
    .addItem('Stop', 'stopEnrichment')
    .addToUi();
}

function startEnrichment() {
  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty('ENRICH_PAUSED');
  props.deleteProperty('ENRICH_STOPPED');
  props.setProperty('ENRICH_LAST_ROW', '1');
  clearResumeTrigger();
  enrichSheet();
}

function pauseEnrichment() {
  PropertiesService.getDocumentProperties().setProperty('ENRICH_PAUSED', 'true');
}

function resumeEnrichment() {
  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty('ENRICH_PAUSED');
  props.deleteProperty('ENRICH_STOPPED');
  setResumeFromLastProcessed();
  clearResumeTrigger();
  enrichSheet();
}

function stopEnrichment() {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('ENRICH_STOPPED', 'true');
  props.deleteProperty('ENRICH_PAUSED');
  clearResumeTrigger();
}

function enrichSheet() {
  var startTime = Date.now();
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var websiteIndex = headers.indexOf('Website');

  if (websiteIndex === -1) {
    throw new Error('Missing Website column in header row.');
  }

  var outputHeaders = [
    'Title',
    'Description',
    'Phone',
    'Email',
    'Contact Links',
    'Progress',
    'Status'
  ];
  var outputSheet = getOrCreateOutputSheet(sheet.getName() + ' Output');
  var outputStart = findOutputStart(headers, outputHeaders);

  if (!outputStart) {
    outputStart = headers.length + 1;
    sheet.getRange(1, outputStart, 1, outputHeaders.length).setValues([outputHeaders]);
  }

  var props = PropertiesService.getDocumentProperties();
  var lastRow = Number(props.getProperty('ENRICH_LAST_ROW') || '1');
  var totalUrls = countRowsWithUrl(values, websiteIndex);
  var processed = countFinalRows(values, outputStart, outputHeaders.length);
  var progressColumn = outputStart + outputHeaders.length - 2;
  var statusColumn = outputStart + outputHeaders.length - 1;
  var batchSize = 1000;
  var batchRows = [];
  var maxRuntimeMs = 5 * 60 * 1000;

  for (var i = Math.max(1, lastRow); i < values.length; i++) {
    var row = values[i];
    var url = row[websiteIndex];
    var existingStatus = row[statusColumn - 1];

    if (!url || isFinalStatus(existingStatus)) {
      continue;
    }

    if (props.getProperty('ENRICH_STOPPED') === 'true') {
      flushBatch(outputSheet, headers, outputHeaders, batchRows);
      batchRows = [];
      props.deleteProperty('ENRICH_LAST_ROW');
      sheet.getRange(i + 1, statusColumn).setValue('Stopped');
      SpreadsheetApp.getActiveSpreadsheet().toast('Enrichment stopped.', 'Enrichment progress', 3);
      return;
    }

    if (props.getProperty('ENRICH_PAUSED') === 'true') {
      flushBatch(outputSheet, headers, outputHeaders, batchRows);
      batchRows = [];
      props.setProperty('ENRICH_LAST_ROW', String(i));
      sheet.getRange(i + 1, statusColumn).setValue('Paused');
      SpreadsheetApp.getActiveSpreadsheet().toast('Enrichment paused.', 'Enrichment progress', 3);
      return;
    }

    sheet.getRange(i + 1, progressColumn).setValue(renderProgressBar(processed, totalUrls));
    sheet.getRange(i + 1, statusColumn).setValue('Fetching...');

    if (!canScrape(url)) {
      processed++;
      var blockedRow = [
        null,
        null,
        null,
        null,
        '',
        renderProgressBar(processed, totalUrls),
        'Blocked by robots.txt'
      ];

      sheet
        .getRange(i + 1, outputStart, 1, outputHeaders.length)
        .setValues([blockedRow]);

      batchRows.push(buildOutputRow(row, blockedRow));
      if (batchRows.length >= batchSize) {
        flushBatch(outputSheet, headers, outputHeaders, batchRows);
        batchRows = [];
      }

      props.setProperty('ENRICH_LAST_ROW', String(i + 1));
      continue;
    }

    var data = scrapeWebsite(url);
    processed++;

    var outputRow = [
      data.title,
      data.description,
      data.phone,
      data.email,
      data.contactLinks.join(', '),
      renderProgressBar(processed, totalUrls),
      data.statusLabel
    ];

    sheet
      .getRange(i + 1, outputStart, 1, outputHeaders.length)
      .setValues([outputRow]);

    batchRows.push(buildOutputRow(row, outputRow));
    if (batchRows.length >= batchSize) {
      flushBatch(outputSheet, headers, outputHeaders, batchRows);
      batchRows = [];
    }

    props.setProperty('ENRICH_LAST_ROW', String(i + 1));

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Processed ' + processed + ' of ' + totalUrls + ' websites',
      'Enrichment progress',
      3
    );

    Utilities.sleep(1200);

    if (Date.now() - startTime > maxRuntimeMs) {
      flushBatch(outputSheet, headers, outputHeaders, batchRows);
      batchRows = [];
      props.setProperty('ENRICH_LAST_ROW', String(i + 1));
      scheduleResume();
      sheet.getRange(i + 1, statusColumn).setValue('Queued');
      SpreadsheetApp.getActiveSpreadsheet().toast('Queued next batch to resume.', 'Enrichment progress', 3);
      return;
    }
  }

  flushBatch(outputSheet, headers, outputHeaders, batchRows);

  props.deleteProperty('ENRICH_LAST_ROW');
  clearResumeTrigger();
}

function setResumeFromLastProcessed() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var outputHeaders = [
    'Title',
    'Description',
    'Phone',
    'Email',
    'Contact Links',
    'Progress',
    'Status'
  ];
  var outputStart = findOutputStart(headers, outputHeaders);

  if (!outputStart) {
    return;
  }

  var statusColumnIndex = outputStart + outputHeaders.length - 2;
  var lastProcessed = 1;

  for (var i = 1; i < values.length; i++) {
    var status = values[i][statusColumnIndex];
    if (isFinalStatus(status)) {
      lastProcessed = i + 1;
    }
  }

  PropertiesService.getDocumentProperties().setProperty(
    'ENRICH_LAST_ROW',
    String(Math.min(lastProcessed, values.length - 1))
  );
}

function countRowsWithUrl(values, websiteIndex) {
  var count = 0;

  for (var i = 1; i < values.length; i++) {
    if (values[i][websiteIndex]) {
      count++;
    }
  }

  return count;
}

function countFinalRows(values, outputStart, outputHeaderLength) {
  var count = 0;
  var statusColumn = outputStart + outputHeaderLength - 1;

  for (var i = 1; i < values.length; i++) {
    if (isFinalStatus(values[i][statusColumn - 1])) {
      count++;
    }
  }

  return count;
}

function renderProgressBar(processed, total) {
  if (!total) {
    return 'Unknown';
  }

  var percentage = Math.round((processed / total) * 100);
  var totalBlocks = 10;
  var filledBlocks = Math.round((percentage / 100) * totalBlocks);
  var bar = Array(filledBlocks + 1).join('█') + Array(totalBlocks - filledBlocks + 1).join('░');

  return bar + ' ' + percentage + '%';
}

function isFinalStatus(status) {
  return status === 'Done' || status === 'Error' || status === 'Blocked by robots.txt';
}

function canScrape(url) {
  var baseUrl = getBaseUrl(url);
  var robotsUrl = baseUrl + '/robots.txt';
  var robotsResponse = safeFetch(robotsUrl);

  if (!robotsResponse.ok) {
    return true;
  }

  var rules = parseRobotsTxt(robotsResponse.text);
  return isPathAllowed(getPath(url), rules);
}

function parseRobotsTxt(text) {
  var lines = text.split('\n');
  var rules = {};
  var currentUserAgent = null;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].split('#')[0].trim();
    if (!line) {
      continue;
    }

    var parts = line.split(':');
    if (parts.length < 2) {
      continue;
    }

    var key = parts[0].trim().toLowerCase();
    var value = parts.slice(1).join(':').trim();

    if (key === 'user-agent') {
      currentUserAgent = value.toLowerCase();
      if (!rules[currentUserAgent]) {
        rules[currentUserAgent] = { allow: [], disallow: [] };
      }
    }

    if (!currentUserAgent) {
      continue;
    }

    if (key === 'allow') {
      rules[currentUserAgent].allow.push(value);
    }

    if (key === 'disallow') {
      rules[currentUserAgent].disallow.push(value);
    }
  }

  return rules;
}

function isPathAllowed(path, rules) {
  var botRules = rules['companyenrichmentbot'] || rules['*'];
  if (!botRules) {
    return true;
  }

  var matchedAllow = matchRobotsRule(path, botRules.allow);
  var matchedDisallow = matchRobotsRule(path, botRules.disallow);

  if (!matchedAllow && !matchedDisallow) {
    return true;
  }

  if (matchedAllow && !matchedDisallow) {
    return true;
  }

  if (!matchedAllow && matchedDisallow) {
    return false;
  }

  return matchedAllow.length >= matchedDisallow.length;
}

function matchRobotsRule(path, rules) {
  var winner = '';

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    if (!rule) {
      continue;
    }

    var pattern = rule.replace(/\*/g, '.*');
    if (rule.endsWith('$')) {
      pattern = pattern.slice(0, -1) + '$';
    }

    var regex = new RegExp('^' + pattern);
    if (regex.test(path) && rule.length > winner.length) {
      winner = rule;
    }
  }

  return winner;
}

function getBaseUrl(url) {
  var match = url.match(/^(https?:\/\/[^/]+)/i);
  return match ? match[1] : url;
}

function getPath(url) {
  var match = url.match(/^https?:\/\/[^/]+(\/.*)$/i);
  return match ? match[1] : '/';
}

function safeFetch(url) {
  return safeFetchWithRetry(url, 2);
}

function safeFetchWithRetry(url, attempts) {
  try {
    for (var i = 0; i <= attempts; i++) {
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (compatible; CompanyEnrichmentBot/1.0)'
        }
      });

      var status = response.getResponseCode();
      if (status < 400) {
        return {
          ok: true,
          status: status,
          text: response.getContentText()
        };
      }

      if (i < attempts) {
        Utilities.sleep(500 * (i + 1));
      }
    }

    return {
      ok: false,
      status: response.getResponseCode(),
      text: response.getContentText()
    };
  } catch (error) {
    return {
      ok: false,
      status: 0,
      text: '',
      error: error
    };
  }
}

function getOrCreateOutputSheet(name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  return sheet;
}

function appendOutputBatch(outputSheet, inputHeaders, outputHeaders, rows) {
  if (!rows.length) {
    return;
  }

  var fullHeaders = inputHeaders.concat(outputHeaders);
  var headerRange = outputSheet.getRange(1, 1, 1, fullHeaders.length);
  var existingHeaders = headerRange.getValues()[0];
  var needsHeaders = outputSheet.getLastRow() === 0 || !headersMatch(existingHeaders, fullHeaders);

  if (needsHeaders) {
    headerRange.setValues([fullHeaders]);
  }

  var startRow = outputSheet.getLastRow() + 1;
  outputSheet.getRange(startRow, 1, rows.length, fullHeaders.length).setValues(rows);
}

function flushBatch(outputSheet, inputHeaders, outputHeaders, rows) {
  if (!rows.length) {
    return;
  }

  appendOutputBatch(outputSheet, inputHeaders, outputHeaders, rows);
  rows.length = 0;
}

function buildOutputRow(inputRow, outputRow) {
  return inputRow.concat(outputRow);
}

function headersMatch(existingHeaders, expectedHeaders) {
  if (existingHeaders.length < expectedHeaders.length) {
    return false;
  }

  for (var i = 0; i < expectedHeaders.length; i++) {
    if (existingHeaders[i] !== expectedHeaders[i]) {
      return false;
    }
  }

  return true;
}

function findOutputStart(headers, outputHeaders) {
  var startIndex = headers.indexOf(outputHeaders[0]);

  if (startIndex === -1) {
    return null;
  }

  for (var i = 0; i < outputHeaders.length; i++) {
    if (headers[startIndex + i] !== outputHeaders[i]) {
      return null;
    }
  }

  return startIndex + 1;
}

/**
 * Fetch a website and return structured data.
 */
function scrapeWebsite(url) {
  var response = safeFetch(url);
  var status = response.status;
  var html = response.text;

  if (!response.ok) {
    return {
      url: url,
      status: status,
      statusLabel: 'Error',
      title: null,
      description: null,
      phone: null,
      email: null,
      contactLinks: []
    };
  }

  var titleMatch = html.match(/<title>([\s\S]*?)<\/title>/i);
  var descriptionMatch = html.match(
    /<meta\s+name=["']description["']\s+content=["']([^"']*)["']/i
  );

  var phoneMatch = html.match(/(\+?\d[\d\s\-\(\)]{7,}\d)/);
  var emailMatch = html.match(/([\w.+-]+@[\w-]+\.[\w.-]+)/);

  return {
    url: url,
    status: status,
    statusLabel: 'Done',
    title: titleMatch ? titleMatch[1].trim() : null,
    description: descriptionMatch ? descriptionMatch[1].trim() : null,
    phone: phoneMatch ? phoneMatch[1].trim() : null,
    email: emailMatch ? emailMatch[1].trim() : null,
    contactLinks: extractLinks(html, ['contact', 'about', 'support'], url)
  };
}

function scheduleResume() {
  clearResumeTrigger();
  ScriptApp.newTrigger('resumeEnrichment').timeBased().after(60 * 1000).create();
}

function clearResumeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'resumeEnrichment') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Extract links that contain relevant keywords.
 */
function extractLinks(html, keywords, baseUrl) {
  var links = [];
  var regex = /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  var match;

  while ((match = regex.exec(html)) !== null) {
    var href = normalizeUrl(match[1], baseUrl);
    var text = match[2].replace(/<[^>]+>/g, '').toLowerCase();

    if (keywords.some(function (keyword) { return text.includes(keyword); })) {
      links.push(href);
    }
  }

  return links;
}

/**
 * Convert relative URLs to absolute URLs.
 */
function normalizeUrl(href, baseUrl) {
  if (href.indexOf('http') === 0) {
    return href;
  }

  if (href.indexOf('//') === 0) {
    return 'https:' + href;
  }

  if (href.indexOf('/') === 0) {
    var originMatch = baseUrl.match(/^(https?:\/\/[^/]+)/i);
    return originMatch ? originMatch[1] + href : href;
  }

  return href;
}`}</code>
          </pre>
        </section>

        <section className="grid gap-4 rounded-2xl border border-slate-800 bg-slate-900/40 p-6 text-sm text-slate-200">
          <h2 className="text-lg font-semibold text-white">Sheet input</h2>
          <p className="text-slate-300">
            Your sheet can match the format below. The script will read the
            Website column and append enrichment results to the right.
          </p>
          <pre className="overflow-x-auto rounded-xl bg-slate-950 p-4 text-xs text-slate-200">
            <code>{`first_name\tlast_name\temail\tWebsite
Manish\tSoni\tmanish@rasoifoods.co.uk\thttps://rasoifoods.co.uk
Richard\tGladwin\trichard@gladwinbrothers.com\thttps://gladwinbrothers.com
Kelvin\tQuick\tkelvin@chinaredsheffield.com\thttps://chinaredsheffield.com`}</code>
          </pre>
        </section>

        <section className="grid gap-4 rounded-2xl border border-slate-800 bg-slate-900/40 p-6 text-sm text-slate-200">
          <h2 className="text-lg font-semibold text-white">Why this is better</h2>
          <ul className="list-disc space-y-2 pl-6 text-slate-300">
            <li>
              Skips rows that already have a title so reruns are safe.
            </li>
            <li>
              Adds throttling between requests to avoid rate limits.
            </li>
            <li>
              Reuses existing output headers so reruns do not append duplicate
              columns.
            </li>
            <li>
              Tracks per-row progress with a visual bar and status column, plus
              an overall toast while the script runs.
            </li>
            <li>
              Adds Start, Pause, Resume, and Stop menu actions to control the
              run and pick up where it left off.
            </li>
            <li>
              Resume scans the Status column to continue from the last completed
              row.
            </li>
            <li>
              Appends combined input + enrichment output to a dedicated output
              sheet in batches of 1000 rows for faster writes.
            </li>
            <li>
              Checks robots.txt before scraping and records fetch errors so the
              run continues safely.
            </li>
            <li>
              Auto-resumes long runs with a timed trigger to avoid execution
              timeouts.
            </li>
            <li>
              Normalizes relative contact links to absolute URLs.
            </li>
            <li>
              Returns empty data for 4xx/5xx responses instead of throwing.
            </li>
          </ul>
        </section>
      </main>
    </div>
  );
}
