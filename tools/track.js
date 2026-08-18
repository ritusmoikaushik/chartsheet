// Weekly metrics tracker. Appends one row per run to metrics.csv so there is a
// time series from day one rather than a reconstruction three months later.
//
//   node tools/track.js            append today's numbers
//   node tools/track.js --print    show the file
//
// Reads only public endpoints. No accounts, no keys.
'use strict'

const fs = require('fs')
const path = require('path')
const https = require('https')

const PACKAGE = process.env.TRACK_PACKAGE || 'xlchart'
const REPO = process.env.TRACK_REPO || 'ritusmoikaushik/xlchart'
// Benchmarks: the stale incumbent, and the base libraries this plugs into.
const COMPARE = ['xlsx-chart', 'exceljs']
const CSV = path.join(__dirname, '..', 'metrics.csv')

const get = url => new Promise(resolve => {
  https.get(url, { headers: { 'user-agent': 'xlchart-tracker' } }, res => {
    let body = ''
    res.on('data', c => { body += c })
    res.on('end', () => {
      try { resolve(JSON.parse(body)) } catch { resolve(null) }
    })
  }).on('error', () => resolve(null))
})

const downloads = async name => {
  const data = await get(`https://api.npmjs.org/downloads/point/last-week/${name}`)
  return data && typeof data.downloads === 'number' ? data.downloads : ''
}

async function main () {
  const today = new Date().toISOString().slice(0, 10)

  const own = await downloads(PACKAGE)
  const repo = await get(`https://api.github.com/repos/${REPO}`)
  const openIssues = await get(
    `https://api.github.com/search/issues?q=repo:${REPO}+is:issue+is:open&per_page=1`)

  const comparisons = {}
  for (const name of COMPARE) comparisons[name] = await downloads(name)

  const row = {
    date: today,
    downloads_week: own,
    stars: repo && repo.stargazers_count != null ? repo.stargazers_count : '',
    forks: repo && repo.forks_count != null ? repo.forks_count : '',
    watchers: repo && repo.subscribers_count != null ? repo.subscribers_count : '',
    open_issues: openIssues && openIssues.total_count != null ? openIssues.total_count : '',
    ...Object.fromEntries(Object.entries(comparisons).map(([k, v]) => [`cmp_${k}`, v])),
  }

  const headers = Object.keys(row)
  if (!fs.existsSync(CSV)) fs.writeFileSync(CSV, headers.join(',') + '\n')
  fs.appendFileSync(CSV, headers.map(h => row[h]).join(',') + '\n')

  console.log(`${today}  ${PACKAGE}: ${row.downloads_week || '—'}/wk` +
    `  stars ${row.stars || '—'}  issues ${row.open_issues || '—'}`)
  for (const [name, value] of Object.entries(comparisons)) {
    console.log(`          ${name}: ${value || '—'}/wk`)
  }
  console.log(`\nappended to ${CSV}`)

  // Thresholds registered in advance, on 2026-08-18, before any data existed.
  if (typeof own === 'number') {
    const verdict = own >= 1000 ? 'TRACTION — build the paid tier'
      : own >= 200 ? 'ALIVE BUT SLOW — keep writing docs, do not build paid yet'
        : 'BELOW THRESHOLD — no demand or no channel yet'
    console.log(`90-day rule: ${verdict}`)
  }
}

if (process.argv.includes('--print')) {
  console.log(fs.existsSync(CSV) ? fs.readFileSync(CSV, 'utf8') : 'no metrics.csv yet')
} else {
  main()
}
