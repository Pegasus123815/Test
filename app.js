const fileInput = document.getElementById('excelFiles');
const totalsEl = document.getElementById('totals');
const streamTableBody = document.querySelector('#streamTable tbody');
const transactionsBody = document.querySelector('#transactionsTable tbody');

const currency = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
});

fileInput.addEventListener('change', async (event) => {
  const files = Array.from(event.target.files || []);
  const rows = [];

  for (const file of files) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      json.forEach((entry) => {
        const incomeStream = entry['Income Stream'] || entry['Stream'] || 'Unspecified';
        const amount = Number(entry['Amount'] || entry['Income'] || 0);

        rows.push({
          fileName: file.name,
          incomeStream,
          amount: Number.isFinite(amount) ? amount : 0,
          advisor: entry['Advisor'] || '',
          date: entry['Date'] || '',
        });
      });
    }
  }

  render(rows);
});

function render(rows) {
  const totals = {
    totalIncome: rows.reduce((sum, row) => sum + row.amount, 0),
    totalRecords: rows.length,
    streamCount: new Set(rows.map((row) => row.incomeStream)).size,
  };

  totalsEl.innerHTML = `
    <div class="metric"><h3>Total Income</h3><p>${currency.format(totals.totalIncome)}</p></div>
    <div class="metric"><h3>Records</h3><p>${totals.totalRecords}</p></div>
    <div class="metric"><h3>Income Streams</h3><p>${totals.streamCount}</p></div>
  `;

  const grouped = rows.reduce((acc, row) => {
    if (!acc[row.incomeStream]) {
      acc[row.incomeStream] = { total: 0, count: 0 };
    }
    acc[row.incomeStream].total += row.amount;
    acc[row.incomeStream].count += 1;
    return acc;
  }, {});

  const streamEntries = Object.entries(grouped)
    .sort((a, b) => b[1].total - a[1].total)
    .map(([name, stats]) => `
      <tr>
        <td>${escapeHtml(name)}</td>
        <td>${currency.format(stats.total)}</td>
        <td>${stats.count}</td>
      </tr>
    `)
    .join('');

  streamTableBody.innerHTML = streamEntries || '<tr><td colspan="3">No data</td></tr>';

  const transactions = rows
    .map(
      (row) => `
      <tr>
        <td>${escapeHtml(row.fileName)}</td>
        <td>${escapeHtml(row.incomeStream)}</td>
        <td>${currency.format(row.amount)}</td>
        <td>${escapeHtml(row.advisor)}</td>
        <td>${escapeHtml(String(row.date))}</td>
      </tr>
    `,
    )
    .join('');

  transactionsBody.innerHTML = transactions || '<tr><td colspan="5">No data</td></tr>';
}

function escapeHtml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}
