// Test suite. Every case builds a real workbook and validates the package,
// because the only failure that matters is Excel refusing to open the file.
'use strict'

const assert = require('assert')
const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const {
  addChart, addCharts, validate, captureCharts, restoreCharts, chartCount,
  addPivotTable, addPivotTables,
} = require('../src')

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
  console.log('chartsheet tests\n')

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

  await test('ExcelJS destroys charts on read-write, and we put them back', async () => {
    const original = await addChart(await baseWorkbook(), barSpec({ title: 'Template' }))
    assert.strictEqual(chartCount(await captureCharts(original)), 1, 'template should have a chart')

    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(original)
    wb.getWorksheet('Data').getCell('B2').value = 999
    const rewritten = Buffer.from(await wb.xlsx.writeBuffer())
    assert.strictEqual(chartCount(await captureCharts(rewritten)), 0,
      'ExcelJS is expected to drop the chart — if this fails, ExcelJS fixed the bug')

    const restored = await restoreCharts(rewritten, await captureCharts(original))
    assert.strictEqual(chartCount(await captureCharts(restored)), 1, 'chart should be back')
    await assertValid(restored, 'restored')

    const check = new ExcelJS.Workbook()
    await check.xlsx.load(restored)
    assert.strictEqual(check.getWorksheet('Data').getCell('B2').value, 999,
      'the edit must survive alongside the chart')
    fs.writeFileSync(path.join(OUT, 'preserved.xlsx'), restored)
  })

  await test('restoring onto a workbook whose sheet was renamed skips safely', async () => {
    const original = await addChart(await baseWorkbook(), barSpec())
    const record = await captureCharts(original)
    const other = await baseWorkbook(['Renamed'])
    const out = await restoreCharts(other, record)
    await assertValid(out, 'renamed sheet')
    assert.strictEqual(chartCount(await captureCharts(out)), 0, 'nothing to restore onto')
  })

  await test('a workbook with no charts is still valid', async () => {
    await assertValid(await baseWorkbook(), 'plain workbook')
  })


  // ---------------------------------------------------------------- pivot tables

  async function pivotSource () {
    const wb = new ExcelJS.Workbook()
    const src = wb.addWorksheet('Data')
    src.addRow(['Region', 'Product', 'Sales'])
    for (const row of [
      ['East', 'Alpha', 100], ['East', 'Beta', 150], ['West', 'Alpha', 200],
      ['West', 'Beta', 120], ['East', 'Alpha', 80], ['North', 'Beta', 60],
    ]) src.addRow(row)
    wb.addWorksheet('Pivot')
    return Buffer.from(await wb.xlsx.writeBuffer())
  }

  const pivotSpec = (over = {}) => ({
    sourceSheet: 'Data',
    sourceRef: 'A1:C7',
    targetSheet: 'Pivot',
    anchor: 'A3',
    rows: ['Region'],
    columns: ['Product'],
    values: [{ field: 'Sales', fn: 'sum' }],
    ...over,
  })

  await test('a pivot table produces a valid package', async () => {
    const out = await addPivotTable(await pivotSource(), pivotSpec())
    await assertValid(out, 'pivot')
    fs.writeFileSync(path.join(OUT, 'pivot.xlsx'), out)
  })

  await test('every pivot part is present and wired', async () => {
    const zip = await JSZip.loadAsync(await addPivotTable(await pivotSource(), pivotSpec()))
    for (const part of [
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/pivotTables/pivotTable1.xml',
      'xl/pivotTables/_rels/pivotTable1.xml.rels',
    ]) assert.ok(zip.file(part), `missing ${part}`)
  })

  await test('recordCount matches the records actually written', async () => {
    const zip = await JSZip.loadAsync(await addPivotTable(await pivotSource(), pivotSpec()))
    const def = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string')
    const rec = await zip.file('xl/pivotCache/pivotCacheRecords1.xml').async('string')
    assert.strictEqual(Number(def.match(/recordCount="(\d+)"/)[1]), (rec.match(/<r>/g) || []).length)
  })

  await test('shared items are distinct and the declared count matches', async () => {
    const zip = await JSZip.loadAsync(await addPivotTable(await pivotSource(), pivotSpec()))
    const def = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string')
    const region = def.match(/name="Region"[\s\S]*?<\/cacheField>/)[0]
    assert.strictEqual((region.match(/<s v="/g) || []).length, 3, 'East, West, North')
    assert.ok(/count="3"/.test(region), 'sharedItems count must match the items listed')
  })

  await test('the cacheId is bound to a pivotCache in the workbook', async () => {
    const zip = await JSZip.loadAsync(await addPivotTable(await pivotSource(), pivotSpec()))
    const pt = await zip.file('xl/pivotTables/pivotTable1.xml').async('string')
    const wb = await zip.file('xl/workbook.xml').async('string')
    const id = pt.match(/cacheId="(\d+)"/)[1]
    assert.ok(new RegExp(`<pivotCache[^>]*cacheId="${id}"`).test(wb), 'cacheId not declared')
  })

  await test('a pivot with no cache relationship is reported', async () => {
    const zip = await JSZip.loadAsync(await addPivotTable(await pivotSource(), pivotSpec()))
    zip.remove('xl/pivotTables/_rels/pivotTable1.xml.rels')
    const result = await validate(await zip.generateAsync({ type: 'nodebuffer' }))
    assert.ok(!result.valid, 'validator must not pass a pivot with no cache relationship')
    assert.ok(result.errors.some(e => /pivotCacheDefinition/.test(e)), result.errors.join('; '))
  })

  await test('a pivot whose cacheId is undeclared is reported', async () => {
    const zip = await JSZip.loadAsync(await addPivotTable(await pivotSource(), pivotSpec()))
    const wb = await zip.file('xl/workbook.xml').async('string')
    zip.file('xl/workbook.xml', wb.replace(/<pivotCaches>[\s\S]*?<\/pivotCaches>/, ''))
    const result = await validate(await zip.generateAsync({ type: 'nodebuffer' }))
    assert.ok(!result.valid, 'validator must not pass an unbound cacheId')
  })

  await test('charts and a pivot table coexist in one workbook', async () => {
    let out = await addPivotTable(await pivotSource(), pivotSpec())
    out = await addChart(out, {
      type: 'bar',
      title: 'Sales',
      categories: "'Data'!$A$2:$A$7",
      series: [{ nameRef: "'Data'!$C$1", ref: "'Data'!$C$2:$C$7" }],
      sheet: 'Data',
    })
    await assertValid(out, 'chart plus pivot')
    fs.writeFileSync(path.join(OUT, 'pivot-and-chart.xlsx'), out)
  })

  await test('two pivot tables in one workbook', async () => {
    const out = await addPivotTables(await pivotSource(), [
      pivotSpec(),
      pivotSpec({ anchor: 'H3', name: 'PivotTable2', rows: ['Product'], columns: [] }),
    ])
    await assertValid(out, 'two pivots')
    const zip = await JSZip.loadAsync(out)
    assert.ok(zip.file('xl/pivotTables/pivotTable2.xml'), 'second pivot part missing')
  })

  await test('an unknown field name names the fields that do exist', async () => {
    const src = await pivotSource()
    await assert.rejects(() => addPivotTable(src, pivotSpec({ rows: ['Nope'] })),
      /Nope[\s\S]*Region, Product, Sales/)
  })

  await test('an unknown aggregate is rejected', async () => {
    const src = await pivotSource()
    await assert.rejects(
      () => addPivotTable(src, pivotSpec({ values: [{ field: 'Sales', fn: 'median' }] })),
      /unknown aggregate/)
  })

  await test('a pivot with no rows and no columns is rejected', async () => {
    const src = await pivotSource()
    await assert.rejects(() => addPivotTable(src, pivotSpec({ rows: [], columns: [] })),
      /at least one field in rows or columns/)
  })

  await test('a workbook with no pivot tables is still valid', async () => {
    await assertValid(await pivotSource(), 'plain workbook with a spare sheet')
  })

  // ------------------------------------------- producer-independent attribute order

  // Excel writes ContentType before PartName and may write Target before Id. Assuming
  // one order made the validator report false problems on files Excel itself produced.
  const flipAttrs = xml => xml
    .replace(/<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"\s*\/>/g,
      (m, part, type) => `<Override ContentType="${type}" PartName="${part}"/>`)
    .replace(/<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"\s*\/>/g,
      (m, id, type, target) => `<Relationship Target="${target}" Type="${type}" Id="${id}"/>`)

  async function reorderAttributes (buffer) {
    const zip = await JSZip.loadAsync(buffer)
    for (const name of Object.keys(zip.files)) {
      if (zip.files[name].dir) continue
      if (name !== '[Content_Types].xml' && !name.endsWith('.rels')) continue
      zip.file(name, flipAttrs(await zip.file(name).async('string')))
    }
    return zip.generateAsync({ type: 'nodebuffer' })
  }

  await test('a file whose attributes are in Excel order still validates', async () => {
    const out = await addChart(await baseWorkbook(), barSpec())
    const flipped = await reorderAttributes(out)
    const ct = await (await JSZip.loadAsync(flipped)).file('[Content_Types].xml').async('string')
    assert.ok(/<Override ContentType="[^"]+" PartName=/.test(ct), 'test did not actually reorder')
    const result = await validate(flipped)
    assert.ok(result.valid, 'false positives on Excel attribute order: ' + result.errors.join('; '))
  })

  await test('charts can be added to a file whose rels are in Excel order', async () => {
    const flipped = await reorderAttributes(await addChart(await baseWorkbook(), barSpec()))
    const out = await addChart(flipped, barSpec({ type: 'line' }))
    await assertValid(out, 'second chart onto reordered rels')
    const zip = await JSZip.loadAsync(out)
    assert.ok(zip.file('xl/charts/chart2.xml'), 'second chart not written')
    assert.ok(!zip.file('xl/drawings/drawing2.xml'), 'must reuse the sheet drawing part')
  })

  await test('a genuinely missing content type is still caught after reordering', async () => {
    const zip = await JSZip.loadAsync(await reorderAttributes(
      await addChart(await baseWorkbook(), barSpec())))
    const ct = await zip.file('[Content_Types].xml').async('string')
    zip.file('[Content_Types].xml', ct.replace(/<Override[^>]*chart1\.xml[^>]*\/>/, ''))
    const result = await validate(await zip.generateAsync({ type: 'nodebuffer' }))
    assert.ok(!result.valid, 'reordering must not blind the validator to real defects')
  })

  await test('pivot tables work on SheetJS output, not just ExcelJS', async () => {
    const XLSX = require('xlsx')
    const rows = [['Region', 'Product', 'Sales']]
    for (const r of ['East', 'West', 'North']) {
      for (const p of ['Alpha', 'Beta']) rows.push([r, p, 100 + rows.length * 7])
    }
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), 'Data')
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[]]), 'Report')

    const out = await addPivotTable(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }), {
      sourceSheet: 'Data', sourceRef: 'A1:C7', targetSheet: 'Report', anchor: 'A3',
      rows: ['Region'], columns: ['Product'], values: [{ field: 'Sales', fn: 'sum' }],
    })
    await assertValid(out, 'pivot on SheetJS output')

    // the cache has to hold the values SheetJS wrote, however it stored them
    const zip = await JSZip.loadAsync(out)
    const def = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string')
    for (const value of ['East', 'West', 'North', 'Alpha', 'Beta']) {
      assert.ok(def.includes(`<s v="${value}"/>`), `cache is missing ${value}`)
    }
    assert.strictEqual(Number(def.match(/recordCount="(\d+)"/)[1]), 6)
  })
  console.log(`\n${passed} passed, ${failed} failed`)
  process.exit(failed ? 1 : 0)
}

main()
