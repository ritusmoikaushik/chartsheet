// Chart preservation across a read-write cycle.
//
// ExcelJS does not model chart or drawing parts, so a workbook that is opened,
// edited and written back loses every chart it had — silently, with no warning.
// Four issues have been open about this on ExcelJS since 2020. The fix is to lift
// the chart parts off the original file and put them back on the output.
'use strict'

const JSZip = require('jszip')
const {
  CT, relsPathFor, nextRelId, addRelationship, findRelTarget,
  readPart, resolveTarget, resolveSheetPath, declareContentType, insertDrawing,
  NS, EMPTY_RELS,
} = require('./package')

const toBuffer = value => Buffer.isBuffer(value) ? value : Buffer.from(value)

const sheetNames = async zip => {
  const workbook = await readPart(zip, 'xl/workbook.xml')
  return [...workbook.matchAll(/<sheet\b[^>]*\/?>/g)]
    .map(m => (m[0].match(/name="([^"]*)"/) || [])[1])
    .filter(Boolean)
}

// A chart part may itself relate to colours, style and user-shape parts. Excel
// writes those routinely, so follow the graph rather than copying one file.
async function collectPart (zip, partPath, seen) {
  if (!partPath || seen.has(partPath)) return
  const file = zip.file(partPath)
  if (!file) return
  seen.set(partPath, await file.async('nodebuffer'))

  const relsPath = relsPathFor(partPath)
  const relsFile = zip.file(relsPath)
  if (!relsFile) return
  const relsXml = await relsFile.async('string')
  seen.set(relsPath, Buffer.from(relsXml))

  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    if (/TargetMode="External"/.test(m[0])) continue
    const target = (m[0].match(/Target="([^"]+)"/) || [])[1]
    if (target) await collectPart(zip, resolveTarget(partPath, target), seen)
  }
}

/**
 * Read every chart on every sheet of a workbook, with the parts they depend on.
 * @returns {Promise<object>} an opaque record to hand to `restoreCharts`.
 */
async function captureCharts (workbook) {
  const zip = await JSZip.loadAsync(toBuffer(workbook))
  const captured = []

  for (const name of await sheetNames(zip)) {
    const sheetPath = await resolveSheetPath(zip, name)
    const sheetXml = await readPart(zip, sheetPath)
    const ref = sheetXml.match(/<drawing[^>]*r:id="(rId\d+)"/)
    if (!ref) continue

    const sheetRels = await readPart(zip, relsPathFor(sheetPath), EMPTY_RELS)
    const target = findRelTarget(sheetRels, ref[1])
    if (!target) continue
    const drawingPath = resolveTarget(sheetPath, target)

    const drawingXml = await readPart(zip, drawingPath, null)
    if (drawingXml == null) continue

    // only carry anchors that actually hold a chart; images belong to ExcelJS
    const drawingRels = await readPart(zip, relsPathFor(drawingPath), EMPTY_RELS)
    const parts = new Map()
    const charts = []
    for (const m of drawingXml.matchAll(/<c:chart\b[^>]*r:id="(rId\d+)"/g)) {
      const chartTarget = findRelTarget(drawingRels, m[1])
      if (!chartTarget) continue
      const chartPath = resolveTarget(drawingPath, chartTarget)
      await collectPart(zip, chartPath, parts)
      charts.push({ relId: m[1], path: chartPath })
    }
    if (!charts.length) continue

    captured.push({
      sheet: name,
      drawingXml,
      charts,
      parts: [...parts].map(([path, data]) => ({ path, data })),
    })
  }

  return { version: 1, sheets: captured }
}

const chartCount = record =>
  (record && record.sheets || []).reduce((n, s) => n + s.charts.length, 0)

/**
 * Put previously captured charts back onto a workbook, matching sheets by name.
 * @returns {Promise<Buffer>} the workbook with its charts restored.
 */
async function restoreCharts (workbook, record) {
  if (!record || !Array.isArray(record.sheets) || !record.sheets.length) {
    return toBuffer(workbook)
  }
  const zip = await JSZip.loadAsync(toBuffer(workbook))
  const present = new Set(await sheetNames(zip))

  let nextChart = Object.keys(zip.files)
    .filter(n => /^xl\/charts\/chart\d+\.xml$/.test(n)).length + 1
  let nextDrawing = Object.keys(zip.files)
    .filter(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).length + 1

  const skipped = []

  for (const entry of record.sheets) {
    if (!present.has(entry.sheet)) {
      // the sheet was renamed or removed; its charts have nowhere to go
      skipped.push(entry.sheet)
      continue
    }
    const sheetPath = await resolveSheetPath(zip, entry.sheet)
    let sheetXml = await readPart(zip, sheetPath)
    if (/<drawing\b/.test(sheetXml)) {
      // the writer kept a drawing of its own; leave it alone rather than corrupt it
      skipped.push(entry.sheet)
      continue
    }

    const drawingIndex = nextDrawing++
    const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`
    let drawingRels = EMPTY_RELS
    let drawingXml = entry.drawingXml

    // copy chart parts under fresh numbers so nothing collides with existing ones
    for (const chart of entry.charts) {
      const index = nextChart++
      const newChartPath = `xl/charts/chart${index}.xml`
      const rename = new Map([[chart.path, newChartPath]])

      for (const part of entry.parts) {
        if (!part.path.startsWith(chart.path.replace(/\.xml$/, '')) &&
            part.path !== chart.path && !part.path.includes(`chart${chart.path.match(/chart(\d+)/)?.[1]}`)) {
          continue
        }
        const mapped = rename.get(part.path) ||
          part.path.replace(/chart(\d+)/g, `chart${index}`)
        rename.set(part.path, mapped)
        zip.file(mapped, part.data)
        if (/^xl\/charts\/chart\d+\.xml$/.test(mapped)) {
          await declareContentType(zip, `/${mapped}`, CT.chart)
        }
      }

      const relId = nextRelId(drawingRels)
      drawingRels = addRelationship(drawingRels, relId, 'chart', `../charts/chart${index}.xml`)
      // repoint this anchor at the relationship id it now has
      drawingXml = drawingXml.replace(
        new RegExp(`(<c:chart\\b[^>]*r:id=")${chart.relId}(")`), `$1${relId}$2`)
    }

    zip.file(drawingPath, drawingXml)
    zip.file(relsPathFor(drawingPath), drawingRels)
    await declareContentType(zip, `/${drawingPath}`, CT.drawing)

    const sheetRelsPath = relsPathFor(sheetPath)
    let sheetRels = await readPart(zip, sheetRelsPath, EMPTY_RELS)
    const relId = nextRelId(sheetRels)
    sheetRels = addRelationship(
      sheetRels, relId, 'drawing', `../drawings/drawing${drawingIndex}.xml`)
    zip.file(sheetRelsPath, sheetRels)

    zip.file(sheetPath, insertDrawing(sheetXml, relId))
  }

  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  if (skipped.length) out.skippedSheets = skipped
  return out
}

/**
 * Convenience wrapper: carry the charts from `original` onto `rewritten`.
 *
 *   const original = fs.readFileSync('template.xlsx')
 *   // ... ExcelJS reads it, edits it, writes `rewritten` ...
 *   const output = await preserveCharts(original, rewritten)
 */
async function preserveCharts (original, rewritten) {
  return restoreCharts(rewritten, await captureCharts(original))
}

module.exports = { captureCharts, restoreCharts, preserveCharts, chartCount }
