// Generates the figures used on the docs site. Drawn from the same numbers the code
// examples use, so the pictures cannot drift away from the text.
'use strict'

const fs = require('fs')
const path = require('path')
const OUT = path.join(__dirname, '..', 'docs', 'img')

const DATA = [
  { label: 'Jan', sales: 120, costs: 90 },
  { label: 'Feb', sales: 150, costs: 95 },
  { label: 'Mar', sales: 180, costs: 110 },
]

const head = (w, h, title) =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" ` +
  `role="img" aria-label="${title}" font-family="ui-sans-serif, system-ui, Segoe UI, Roboto, sans-serif">`

// An SVG loaded through <img> cannot see the page's custom properties, so it carries its
// own dark-mode rules. The var() forms still win if the file is ever inlined instead.
const STYLE = `<style>
  .bg{fill:var(--fig-surface,#fff)} .rule{stroke:var(--fig-rule,#e3e4e8);stroke-width:1;fill:none}
  .ink{fill:var(--fig-ink,#16181d)} .soft{fill:var(--fig-soft,#4d525c)}
  .head{fill:var(--fig-head,#f1f2f4)} .a{fill:var(--fig-accent,#1d6b4f)} .b{fill:var(--fig-accent2,#8bbfa8)}
  .mono{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}
  .s-soft{stroke:var(--fig-soft,#4d525c)} .f-soft{fill:var(--fig-soft,#4d525c)}
  .s-accent{stroke:var(--fig-accent,#1d6b4f)} .f-accent{fill:var(--fig-accent,#1d6b4f)}
  .s-rule{stroke:var(--fig-rule,#e3e4e8)}
  @media (prefers-color-scheme: dark){
    .s-soft{stroke:var(--fig-soft,#a2a9b5)} .f-soft{fill:var(--fig-soft,#a2a9b5)}
    .s-accent{stroke:var(--fig-accent,#5fbf95)} .f-accent{fill:var(--fig-accent,#5fbf95)}
    .s-rule{stroke:var(--fig-rule,#2b303a)}
    .bg{fill:var(--fig-surface,#14171d)} .rule{stroke:var(--fig-rule,#2b303a)}
    .ink{fill:var(--fig-ink,#e9ebef)} .soft{fill:var(--fig-soft,#a2a9b5)}
    .head{fill:var(--fig-head,#1b1f27)} .a{fill:var(--fig-accent,#5fbf95)} .b{fill:var(--fig-accent2,#2f6f57)}
  }
</style>`

// ---------------------------------------------------------------- sheet + chart
function sheetAndChart () {
  const w = 720, h = 300
  const cw = [56, 76, 76], rh = 26, x0 = 16, y0 = 34
  let g = ''

  // grid
  const rows = [['', 'Month', 'Sales', 'Costs'], ...DATA.map((d, i) =>
    [String(i + 2), d.label, String(d.sales), String(d.costs)])]
  rows.unshift(['1', 'Month', 'Sales', 'Costs'])
  rows.splice(1, 1)

  const colX = [x0, x0 + 30, x0 + 30 + cw[0], x0 + 30 + cw[0] + cw[1]]
  const totalW = 30 + cw[0] + cw[1] + cw[2]

  rows.forEach((cells, r) => {
    const y = y0 + r * rh
    cells.forEach((text, c) => {
      const x = colX[c]
      const cwid = c === 0 ? 30 : cw[c - 1]
      const isHeader = r === 0 || c === 0
      g += `<rect x="${x}" y="${y}" width="${cwid}" height="${rh}" class="${isHeader ? 'head' : 'bg'}"/>`
      g += `<rect x="${x}" y="${y}" width="${cwid}" height="${rh}" class="rule"/>`
      const anchor = c === 0 ? 'middle' : (c === 1 ? 'start' : 'end')
      const tx = c === 0 ? x + cwid / 2 : (c === 1 ? x + 9 : x + cwid - 9)
      g += `<text x="${tx}" y="${y + 17}" text-anchor="${anchor}" font-size="12.5" ` +
        `class="${r === 0 || c === 0 ? 'soft' : 'ink'} mono"${r === 0 ? ' font-weight="600"' : ''}>${text}</text>`
    })
  })

  // chart panel
  const px = x0 + totalW + 34, pw = w - px - 16, ph = 232
  g += `<rect x="${px}" y="${y0}" width="${pw}" height="${ph}" rx="6" class="bg"/>`
  g += `<rect x="${px}" y="${y0}" width="${pw}" height="${ph}" rx="6" class="rule"/>`
  g += `<text x="${px + pw / 2}" y="${y0 + 24}" text-anchor="middle" font-size="13" font-weight="600" class="ink">Quarterly performance</text>`

  const plotL = px + 40, plotR = px + pw - 16, plotT = y0 + 40, plotB = y0 + ph - 34
  const max = 200
  for (let v = 0; v <= max; v += 50) {
    const y = plotB - (v / max) * (plotB - plotT)
    g += `<line x1="${plotL}" y1="${y}" x2="${plotR}" y2="${y}" class="rule"/>`
    g += `<text x="${plotL - 8}" y="${y + 4}" text-anchor="end" font-size="10" class="soft">${v}</text>`
  }

  const band = (plotR - plotL) / DATA.length
  DATA.forEach((d, i) => {
    const bw = 18, gap = 5
    const cx = plotL + band * i + band / 2
    const hs = (d.sales / max) * (plotB - plotT)
    const hc = (d.costs / max) * (plotB - plotT)
    g += `<rect x="${cx - bw - gap / 2}" y="${plotB - hs}" width="${bw}" height="${hs}" rx="2" class="a"/>`
    g += `<rect x="${cx + gap / 2}" y="${plotB - hc}" width="${bw}" height="${hc}" rx="2" class="b"/>`
    g += `<text x="${cx}" y="${plotB + 16}" text-anchor="middle" font-size="11" class="soft">${d.label}</text>`
  })

  // legend
  g += `<rect x="${plotL}" y="${y0 + ph - 16}" width="9" height="9" rx="2" class="a"/>`
  g += `<text x="${plotL + 14}" y="${y0 + ph - 8}" font-size="10" class="soft">Sales</text>`
  g += `<rect x="${plotL + 56}" y="${y0 + ph - 16}" width="9" height="9" rx="2" class="b"/>`
  g += `<text x="${plotL + 70}" y="${y0 + ph - 8}" font-size="10" class="soft">Costs</text>`

  g += `<text x="${x0}" y="20" font-size="11" class="soft">Data sheet</text>`
  g += `<text x="${px}" y="20" font-size="11" class="soft">Chart written into the same file</text>`

  return head(w, h, 'A worksheet of numbers beside the chart chartsheet writes into the same file') +
    STYLE + g + '</svg>'
}

// ---------------------------------------------------------------- round trip loss
function roundTrip () {
  const w = 720, h = 210
  let g = ''
  const card = (x, title, sub, hasChart, tone) => {
    let c = `<rect x="${x}" y="46" width="196" height="128" rx="7" class="bg"/>`
    c += `<rect x="${x}" y="46" width="196" height="128" rx="7" class="rule"/>`
    c += `<text x="${x + 98}" y="32" text-anchor="middle" font-size="12" font-weight="600" class="ink">${title}</text>`
    // little grid
    for (let r = 0; r < 3; r++) {
      for (let col = 0; col < 3; col++) {
        c += `<rect x="${x + 14 + col * 26}" y="${62 + r * 15}" width="24" height="13" rx="2" class="head"/>`
      }
    }
    // chart area
    if (hasChart) {
      const bars = [26, 40, 34]
      bars.forEach((bh, i) => {
        c += `<rect x="${x + 106 + i * 22}" y="${118 - bh}" width="14" height="${bh}" rx="2" class="${tone}"/>`
      })
      c += `<line x1="${x + 100}" y1="120" x2="${x + 182}" y2="120" class="rule"/>`
    } else {
      c += `<rect x="${x + 100}" y="66" width="82" height="54" rx="4" fill="none" class="s-rule" stroke-dasharray="4 4"/>`
      c += `<text x="${x + 141}" y="97" text-anchor="middle" font-size="11" class="soft">gone</text>`
    }
    c += `<text x="${x + 98}" y="164" text-anchor="middle" font-size="10.5" class="soft">${sub}</text>`
    return c
  }

  const arrow = (x, label) => {
    let a = `<line x1="${x}" y1="110" x2="${x + 44}" y2="110" class="s-soft" stroke-width="1.5"/>`
    a += `<path d="M${x + 44} 110 l-7 -4 v8 z" class="f-soft"/>`
    a += `<text x="${x + 22}" y="100" text-anchor="middle" font-size="10" class="soft">${label}</text>`
    return a
  }

  g += card(16, 'template.xlsx', 'one chart', true, 'a')
  g += arrow(220, 'load + edit')
  g += card(272, 'out.xlsx', 'edit applied, chart lost', false, 'a')
  g += arrow(476, 'preserveCharts')
  g += card(528, 'out.xlsx', 'edit applied, chart intact', true, 'a')

  return head(w, h, 'A workbook loses its chart when rewritten, and keeps it when the charts are preserved') + STYLE + g + '</svg>'
}

// ---------------------------------------------------------------- package anatomy
function anatomy () {
  const w = 720, h = 268
  let g = ''
  const box = (x, y, bw, bh, label, note, cls) => {
    let b = `<rect x="${x}" y="${y}" width="${bw}" height="${bh}" rx="6" class="${cls || 'bg'}"/>`
    b += `<rect x="${x}" y="${y}" width="${bw}" height="${bh}" rx="6" class="rule"/>`
    b += `<text x="${x + 12}" y="${y + 21}" font-size="12" class="ink mono">${label}</text>`
    if (note) b += `<text x="${x + 12}" y="${y + 38}" font-size="10.5" class="soft">${note}</text>`
    return b
  }
  const link = (x1, y1, x2, y2, label) => {
    let l = `<path d="M${x1} ${y1} C ${(x1 + x2) / 2} ${y1}, ${(x1 + x2) / 2} ${y2}, ${x2} ${y2}" fill="none" class="s-accent" stroke-width="1.4"/>`
    l += `<circle cx="${x2}" cy="${y2}" r="2.6" class="f-accent"/>`
    if (label) l += `<text x="${(x1 + x2) / 2}" y="${(y1 + y2) / 2 - 6}" text-anchor="middle" font-size="9.5" class="soft">${label}</text>`
    return l
  }

  g += `<text x="16" y="20" font-size="11" class="soft">report.xlsx — a zip of parts. chartsheet writes the three on the right and the links between them.</text>`
  g += box(16, 36, 210, 52, 'xl/worksheets/sheet1.xml', 'your data, written by ExcelJS', 'head')
  g += box(16, 116, 210, 52, '[Content_Types].xml', 'every part needs an Override', 'head')
  g += box(16, 196, 210, 52, 'xl/workbook.xml', 'lists the sheets', 'head')

  g += box(390, 36, 210, 52, 'xl/drawings/drawing1.xml', 'one per sheet, holds anchors')
  g += box(390, 116, 210, 52, 'xl/charts/chart1.xml', 'the chart and its series')
  g += box(390, 196, 210, 52, 'xl/pivotTables/pivotTable1.xml', 'plus its cache parts')

  g += link(226, 62, 390, 62, 'r:id')
  g += link(226, 142, 390, 142, 'Override')
  g += link(600, 88, 600, 116, '')
  g += `<text x="612" y="106" font-size="9.5" class="soft">rel</text>`
  g += link(226, 222, 390, 222, 'cacheId')

  return head(w, h, 'The parts of an xlsx package and the relationships between them') + STYLE + g + '</svg>'
}

fs.writeFileSync(path.join(OUT, 'round-trip.svg'), roundTrip())
console.log('docs/img/round-trip.svg')
fs.writeFileSync(path.join(OUT, 'anatomy.svg'), anatomy())
console.log('docs/img/anatomy.svg')

fs.mkdirSync(OUT, { recursive: true })
fs.writeFileSync(path.join(OUT, 'sheet-and-chart.svg'), sheetAndChart())
console.log('docs/img/sheet-and-chart.svg', Math.round(fs.statSync(path.join(OUT, 'sheet-and-chart.svg')).size / 1024) + 'KB')
