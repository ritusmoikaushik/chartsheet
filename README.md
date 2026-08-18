# chartsheet

Native Excel charts for `.xlsx` files built with **ExcelJS** or **SheetJS** — the chart support
neither library has.

ExcelJS is installed over 11 million times a week and cannot write a chart. SheetJS Community
leaves charts to its paid tier. So every spreadsheet exported by every app using them arrives as
bare numbers, and the person who opens it has to select the data and insert the chart by hand.

This adds the chart. A real, native Excel chart — click it, change a number, it redraws. Not an
image.

```bash
npm install chartsheet
```

## Use

```js
const ExcelJS = require('exceljs')
const { addChart } = require('chartsheet')

const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('Data')
ws.addRow(['Month', 'Sales', 'Costs'])
ws.addRow(['Jan', 120, 90])
ws.addRow(['Feb', 150, 95])
ws.addRow(['Mar', 180, 110])

let buffer = await wb.xlsx.writeBuffer()      // ExcelJS writes the sheet

buffer = await addChart(buffer, {             // excel-chart adds the chart
  type: 'bar',
  title: 'Quarterly performance',
  categories: "'Data'!$A$2:$A$4",
  series: [
    { nameRef: "'Data'!$B$1", ref: "'Data'!$B$2:$B$4" },
    { nameRef: "'Data'!$C$1", ref: "'Data'!$C$2:$C$4" },
  ],
  anchor: { col: 4, row: 1 },
})

require('fs').writeFileSync('report.xlsx', buffer)
```

Works the same on a SheetJS workbook:

```js
const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
const withChart = await addChart(buffer, spec)
```

## Chart types

`bar` · `column` · `line` · `pie` · `doughnut` · `area` · `scatter` · `radar`

## Options

| Option | Meaning |
|---|---|
| `type` | one of the types above; defaults to `bar` |
| `title` | chart title; omit for no title |
| `series` | `[{ ref, nameRef, name, colour }]` — `ref` is required |
| `categories` | cell range for the category labels |
| `sheet` | worksheet name to place the chart on; defaults to the first |
| `anchor` | `{ col, row }`, zero-based top-left cell |
| `width` / `height` | pixels; defaults 600 × 340 |
| `stacked` | stack the series |
| `horizontal` | horizontal bars (`bar` only) |
| `dataLabels` | print values on the chart |
| `legend` | `'r'`, `'l'`, `'t'`, `'b'`, or `false` to hide |
| `numberFormat` | value-axis format, e.g. `'#,##0.00'` |
| `gridlines` | `false` to remove value-axis gridlines |
| `xTitle` / `yTitle` | axis titles |

Several charts at once:

```js
const { addCharts } = require('chartsheet')
buffer = await addCharts(buffer, [barSpec, lineSpec, pieSpec])
```

## Validating a workbook

Excel reports a damaged file only as *"we found a problem with some content"*, with no detail. This
tells you what is actually wrong:

```js
const { validate } = require('chartsheet')

const { valid, errors, warnings } = await validate(buffer)
// errors: [ 'xl/charts/chart1.xml: no <Override> content type — ...' ]
```

It checks the things that make Excel refuse a file: unresolved relationships, missing content-type
overrides, duplicate shape ids, drawing parts that are related but never referenced, and `<drawing>`
placed out of schema order. Useful on any xlsx, not only ones this library touched.

## Why this exists

Adding a chart to an xlsx means writing a chart part, a drawing part, two sets of relationships,
and content-type overrides — and getting any of it slightly wrong produces a file Excel refuses to
open, with no useful error. The most-upvoted feature request on ExcelJS asked for chart support in
**2016** and is still open.

## Notes

- Cell references use standard Excel syntax and should be absolute: `'Sheet name'!$B$2:$B$10`
- A worksheet holds one drawing part; every chart on that sheet becomes another anchor inside it,
  which this handles for you
- Output opens in Excel, Google Sheets and LibreOffice


## Keeping charts through a read-write cycle

ExcelJS does not model chart parts, so opening a workbook that has charts, editing it and writing it
back **silently deletes every chart**. No error, no warning. Reported on ExcelJS in 2020, 2021 and
2023; still open.

```js
const { preserveCharts } = require('chartsheet')

const original = fs.readFileSync('template.xlsx')   // has charts

const wb = new ExcelJS.Workbook()
await wb.xlsx.load(original)
wb.getWorksheet('Data').getCell('B2').value = 999
const rewritten = await wb.xlsx.writeBuffer()       // charts are gone here

const output = await preserveCharts(original, rewritten)   // and back again
```

Or in two steps, if the edit happens elsewhere:

```js
const record = await captureCharts(original)
// ...
const output = await restoreCharts(rewritten, record)
```

Sheets are matched by **name**, so charts land back where they belong even if sheet order changed.
A sheet that was renamed or removed is skipped rather than guessed at.

## Licence

MIT
