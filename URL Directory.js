/** MAIN: Build grouped buttons sidebar with search (directory.gs) */
function showUrls() {
  const host = openHostSpreadsheet_(); // from config.gs
  const sheetName = (typeof getDirectorySheetName_ === 'function')
    ? getDirectorySheetName_()
    : 'Sheet Directory';

  const sheet = host.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found in the host spreadsheet.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No directory rows found.');
    return;
  }

  // Read C:D:F in one go (C=Group, D=URL, F=Name)
  const height = lastRow - 1;
  const data = sheet.getRange(2, 3, height, 4).getValues(); // C2:F

  // Group -> [{url,name}]
  const grouped = {};
  const urlOk = (u) => /^https?:\/\//i.test(String(u || '').trim());

  data.forEach(row => {
    const group = String(row[0] || 'Ungrouped').trim() || 'Ungrouped'; // C
    const url   = String(row[1] || '').trim();                          // D
    const name  = String(row[3] || '').trim() || 'Untitled';            // F
    if (!urlOk(url)) return;
  
    if (!grouped[group]) grouped[group] = [];
    grouped[group].push({ url, name });
  });
  if (!Object.keys(grouped).length) {
    SpreadsheetApp.getUi().alert('No valid links found in the directory.');
    return;
  }

  // Deduplicate per group (by URL) and sort
  Object.keys(grouped).forEach(g => {
    const seen = new Set();
    grouped[g] = grouped[g]
      .filter(it => (seen.has(it.url) ? false : (seen.add(it.url), true)))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  });
  const groupsSorted = Object.keys(grouped).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' })
  );

  // Safe encoders
  const escHtml = s => String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');

  // Build HTML with search UI
  let html = `
  <html>
    <head>
      <meta charset="utf-8">
      <style>
        body { font-family: Arial, sans-serif; background:#f7f7f7; padding:16px; }
        h1 { color:#333; text-align:center; margin:0 0 12px; font-size:18px; }
        h2 { color:#444; margin:12px 0 8px; font-size:13px; }
        .search { display:flex; gap:8px; margin:8px 0 12px; }
        .search input {
          flex:1; padding:8px; border:1px solid #ccc; border-radius:6px; font-size:13px;
        }
        .search button {
          padding:8px 10px; border:none; border-radius:6px; cursor:pointer; font-size:13px;
          background:#e0e0e0;
        }
        .search button:hover { background:#d5d5d5; }
        button.link {
          background:#4CAF50; border:none; color:#fff; padding:8px 12px;
          font-size:13px; margin:6px 4px; cursor:pointer; border-radius:6px;
          transition: background-color .2s;
        }
        button.link:hover { background:#45a049; }
        .group { margin-bottom:10px; }
        .hidden { display:none !important; }
        .no-results { margin-top:12px; color:#777; font-size:12px; text-align:center; }
      </style>
    </head>
    <body>
      <h1>Select File</h1>
      <div class="search">
        <input id="q" type="text" placeholder="Search group or name..." autocomplete="off" />
        <button id="clear">Clear</button>
      </div>
      <div id="container">
  `;

  groupsSorted.forEach(group => {
    const gId = 'g-' + escHtml(group).replace(/\W+/g, '_');
    html += `<div class="group" data-group="${escHtml(group)}" id="${gId}">
               <h2>${escHtml(group)}</h2>`;
    grouped[group].forEach(item => {
      html += `<button class="link" data-name="${escHtml(item.name)}"
                       onclick="window.open('${escHtml(item.url)}','_blank','noopener')">
                 ${escHtml(item.name)}
               </button><br>`;
    });
    html += `</div>`;
  });

  html += `
      </div>
      <div id="empty" class="no-results hidden">No matches.</div>

      <script>
        (function(){
          const q = document.getElementById('q');
          const clearBtn = document.getElementById('clear');
          const container = document.getElementById('container');
          const groups = Array.from(container.querySelectorAll('.group'));
          const emptyMsg = document.getElementById('empty');

          const norm = s => (s||'').toLowerCase().trim();
          let t; // debounce timer

          function applyFilter() {
            const term = norm(q.value);
            let shownGroups = 0;

            groups.forEach(g => {
              const gName = norm(g.getAttribute('data-group'));
              const buttons = Array.from(g.querySelectorAll('button.link'));
              let anyShow = false;

              buttons.forEach(b => {
                const n = norm(b.getAttribute('data-name'));
                const show = !term || gName.includes(term) || n.includes(term);
                b.classList.toggle('hidden', !show);
                if (show) anyShow = true;
              });

              g.classList.toggle('hidden', !anyShow);
              if (anyShow) shownGroups++;
            });

            emptyMsg.classList.toggle('hidden', shownGroups > 0);
          }

          q.addEventListener('input', function() {
            clearTimeout(t);
            t = setTimeout(applyFilter, 120); // debounce
          });

          clearBtn.addEventListener('click', function(){
            q.value = '';
            applyFilter();
            q.focus();
          });

          // initial render
          applyFilter();
        })();
      </script>
    </body>
  </html>`;

  const ui = HtmlService.createHtmlOutput(html)
    .setTitle('Directory')
    .setWidth(340)
    .setHeight(560);

  SpreadsheetApp.getUi().showSidebar(ui);
}
