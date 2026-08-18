// Test suite. Every case builds a real workbook and validates the package,
// because the only failure that matters is Excel refusing to open the file.
'use strict'

const assert = require('assert')
const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const { addChart, addCharts, validate } = require('../src')

const OUT = path.join(__dirname, 'output')
let passed = 0
let failed = 0

async function test (name, fn) {
  try {
    await fn()
    console.log(`  ok    ${name}`)
    passed++
  } catch (err) {
    console.log(`  FAIL  ${name}\n        ${err.message}`)
    failed++
  }
}

async function baseWorkbook (sheets = ['Data']) {
  const wb = new ExcelJS.Workbook()
  for (const name of sheets) {
    const ws = wb.addWorksheet(name)
    ws.addRow(['Month', 'Sales', 'Costs'])
    ws.addRow(['Jan', 120, 90])
    ws.addRow(['Feb', 150, 95])
    ws.addRow(['Mar', 180, 110])
    ws.addRow(['Apr', 140, 105])
  }
  return Buffer.from(await wb.xlsx.writeBuffer())
}

const barSpec = (over = {}) => ({
  type: 'bar',
  title: 'Quarterly',
  categories: "'Data'!$A$2:$A$5",
  series: [
    { nameRef: "'Data'!$B$1", ref: "'Data'!$B$2:$B$5" },
    { nameRef: "'Data'!$C$1", ref: "'Data'!$C$2:$C$5" },
  ],
  ...over,
})

async function assertValid (buffer, label) {
  const result = await validate(buffer)
  assert.deepStrictEqual(result.errors, [], `${label}: ${result.errors.join(' | ')}`)
  assert.ok(result.valid, `${label} should be valid`)
  return result
}

async function main () {
  fs.mkdirSync(OUT, { recursive: true })
  console.log('excel-chart tests\n')

  for (const type of ['bar', 'line', 'pie', 'doughnut', 'area', 'radar']) {
    await test(`${type} chart produces a valid package`, async () => {
      const out = await addChart(await baseWorkbook(), barSpec({ type }))
      await assertValid(out, type)
      fs.writeFileSync(path.join(OUT, `${type}.xlsx`), out)
    })
  }

  await test('scatter chart produces a valid package', async () => {
    const out = await addChart(await baseWorkbook(), {
      type: 'scatter',
      title: 'Sales vs costs',
      series: [{ name: 'obs', xRef: "'Data'!$B$2:$B$5", ref: "'Data'!$C$2:$C$5" }],
    })
    await assertValid(out, 'scatter')
    fs.writeFileSync(path.join(OUT, 'scatter.xlsx'), out)
  })

  await test('several charts share one drawing part with unique shape ids', async () => {
    const out = await addCharts(await baseWorkbook(), [
      barSpec({ anchor: { col: 4, row: 1 } }),
      barSpec({ type: 'line', title: 'Trend', anchor: { col: 4, row: 20 } }),
      barSpec({ type: 'pie', title: 'Share', anchor: { col: 13, row: 1 } }),
    ])
    await assertValid(out, 'multi')

    const zip = await JSZip.loadAsync(out)
    const drawings = Object.keys(zip.files).filter(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n))
    assert.strictEqual(drawings.length, 1, 'expected exactly one drawing part')
    const xml = await zip.file(drawings[0]).async('string')
    const anchors = (xml.match(/<xdr:oneCellAnchor>/g) || []).length
    assert.strictEqual(anchors, 3, `expected 3 anchors, got ${anchors}`)
    const ids = [...xml.matchAll(/<xdr:cNvPr\b[^>]*id="(\d+)"/g)].map(m => m[1])
    assert.strictEqual(new Set(ids).size, 3, `shape ids must be unique, got ${ids}`)
    fs.writeFileSync(path.join(OUT, 'multi.xlsx'), out)
  })

  await test('charts land on the named sheet, not always the first', async () => {
    const out = await addChart(await baseWorkbook(['Summary', 'Data']),
      barSpec({ sheet: 'Data' }))
    await assertValid(out, 'named sheet')

    const zip = await JSZip.loadAsync(out)
    const first = await zip.file('xl/worksheets/sheet1.xml').async('string')
    const second = await zip.file('xl/worksheets/sheet2.xml').async('string')
    assert.ok(!/<drawing\b/.test(first), 'sheet1 (Summary) should have no drawing')
    assert.ok(/<drawing\b/.test(second), 'sheet2 (Data) should carry the drawing')
  })

  await test('charts on different sheets get separate drawing parts', async () => {
    const out = await addCharts(await baseWorkbook(['One', 'Two']), [
      barSpec({ sheet: 'One', categories: "'One'!$A$2:$A$5", series: [{ ref: "'One'!$B$2:$B$5" }] }),
      barSpec({ sheet: 'Two', categories: "'Two'!$A$2:$A$5", series: [{ ref: "'Two'!$B$2:$B$5" }] }),
    ])
    await assertValid(out, 'two sheets')
    const zip = await JSZip.loadAsync(out)
    const drawings = Object.keys(zip.files).filter(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n))
    assert.strictEqual(drawings.length, 2, 'expected one drawing part per sheet')
  })

  await test('an unknown sheet name fails loudly', async () => {
    const wb = await baseWorkbook()
    await assert.rejects(
      () => addChart(wb, barSpec({ sheet: 'Nope' })),
      /no worksheet named "Nope"/)
  })

  await test('a spec with no series is rejected', async () => {
    await assert.rejects(
      async () => addChart(await baseWorkbook(), { type: 'bar', series: [] }),
      /at least one entry in series/)
  })

  await test('an unknown chart type is rejected', async () => {
    await assert.rejects(
      async () => addChart(await baseWorkbook(), barSpec({ type: 'sunburst' })),
      /unknown chart type/)
  })

  await test('titles containing XML metacharacters are escaped', async () => {
    const out = await addChart(await baseWorkbook(),
      barSpec({ title: 'Profit & "loss" <2026>' }))
    await assertValid(out, 'escaping')
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file('xl/charts/chart1.xml').async('string')
    assert.ok(xml.includes('&amp;'), 'ampersand should be escaped')
    assert.ok(!/<a:t>[^<]*<2026>/.test(xml), 'angle brackets should be escaped')
  })

  await test('options reach the chart XML', async () => {
    const out = await addChart(await baseWorkbook(), barSpec({
      stacked: true, dataLabels: true, legend: 'b', numberFormat: '#,##0.00',
      series: [{ nameRef: "'Data'!$B$1", ref: "'Data'!$B$2:$B$5", colour: '#3366CC' }],
    }))
    await assertValid(out, 'options')
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file('xl/charts/chart1.xml').async('string')
    assert.ok(xml.includes('<c:grouping val="stacked"/>'), 'stacked grouping')
    assert.ok(xml.includes('<c:overlap val="100"/>'), 'stacked needs overlap')
    assert.ok(xml.includes('<c:showVal val="1"/>'), 'data labels')
    assert.ok(xml.includes('<c:legendPos val="b"/>'), 'legend position')
    assert.ok(xml.includes('#,##0.00'), 'number format')
    assert.ok(xml.includes('3366CC'), 'series colour')
  })

  await test('the validator catches a broken package', async () => {
    const out = await addChart(await baseWorkbook(), barSpec())
    const zip = await JSZip.loadAsync(out)
    // strip the chart's content type: exactly what makes Excel refuse a file
    const ct = await zip.file('[Content_Types].xml').async('string')
    zip.file('[Content_Types].xml', ct.replace(/<Override PartName="\/xl\/charts[^>]*>/, ''))
    const broken = await zip.generateAsync({ type: 'nodebuffer' })
    const result = await validate(broken)
    assert.strictEqual(result.valid, false, 'validator should reject it')
    assert.ok(result.errors.some(e => /content type/.test(e)), result.errors.join(' | '))
  })

  await test('works on workbooks written by SheetJS, not just ExcelJS', async () => {
    const XLSX = require('xlsx')
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 120, 90], ['Feb', 150, 95], ['Mar', 180, 110], ['Apr', 140, 105],
    ]), 'Data')
    const base = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
    const out = await addChart(base, barSpec({ title: 'From SheetJS' }))
    await assertValid(out, 'sheetjs')
    fs.writeFileSync(path.join(OUT, 'sheetjs.xlsx'), out)
  })

  await test('a workbook with no charts is still valid', async () => {
    await assertValid(await baseWorkbook(), 'plain workbook')
  })

  console.log(`\n${passed} passed, ${failed} failed`)
  process.exit(failed ? 1 : 0)
}

main()
