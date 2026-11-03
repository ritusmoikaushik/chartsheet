// Chart preservation across a read-write cycle.
//
// ExcelJS does not model chart or drawing parts, so a workbook that is opened,
// edited and written back loses every chart it had — silently, with no warning.
// Four issues have been open about this on ExcelJS since 2020. The fix is to lift
// the chart parts off the original file and put them back on the output.
'use strict'

const JSZip = require('jszip')
const {
  CT, relsPathFor, nextRelId, nextPartIndex, addRelationship, findRelTarget,
  readPart, resolveTarget, relativeTarget, resolveSheetPath, declareContentType,
  contentTypeOf, insertDrawing, EMPTY_RELS,
} = require('./package')

const toBuffer = value => Buffer.isBuffer(value) ? value : Buffer.from(value)

const sheetNames = async zip => {
  const workbook = await readPart(zip, 'xl/workbook.xml')
  return [...workbook.matchAll(/<sheet\b[^>]*\/?>/g)]
    .map(m => (m[0].match(/name="([^"]*)"/) || [])[1])
    .filter(Boolean)
}

// The internal relationships in a rels part, as the tag and its Target attribute.
const internalRels = relsXml => [...relsXml.matchAll(/<Relationship\b[^>]*>/g)]
  .map(m => ({ tag: m[0], target: (m[0].match(/\bTarget\s*=\s*"([^"]+)"/) || [])[1] }))
  .filter(rel => rel.target && !/TargetMode\s*=\s*"External"/.test(rel.tag))

// A chart part may itself relate to colours, style and user-shape parts. Excel
// writes those routinely, so follow the graph rather than copying one file.
async function collectPart (zip, partPath, seen, typesXml) {
  if (!partPath || seen.has(partPath)) return
  const file = zip.file(partPath)
  if (!file) return
  seen.set(partPath, {
    data: await file.async('nodebuffer'),
    contentType: contentTypeOf(typesXml, partPath),
  })

  const relsPath = relsPathFor(partPath)
  const relsFile = zip.file(relsPath)
  if (!relsFile) return
  const relsXml = await relsFile.async('string')
  seen.set(relsPath, { data: Buffer.from(relsXml) })

  for (const { target } of internalRels(relsXml)) {
    await collectPart(zip, resolveTarget(partPath, target), seen, typesXml)
  }
}

// Every part reachable from one chart, the chart first, through the captured rels.
function chartGraph (chartPath, parts) {
  const order = []
  const queue = [chartPath]
  while (queue.length) {
    const path = queue.shift()
    if (order.includes(path) || !parts.has(path)) continue
    order.push(path)
    const rels = parts.get(relsPathFor(path))
    if (!rels) continue
    for (const { target } of internalRels(rels.data.toString())) {
      queue.push(resolveTarget(path, target))
    }
  }
  return order
}

const escapeRegExp = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

// A free name for a copied part: same folder and stem, the next number not in use.
// colors1.xml becomes colorsN.xml; the drawing that holds a chart's user shapes takes
// a drawingN.xml from the same sequence as worksheet drawings, so they cannot collide.
function freshPartPath (zip, oldPath) {
  const [, dir, stem, ext] = oldPath.match(/^(.*\/)?([^/]*?)\d*(\.[^./]+)$/)
  const prefix = (dir || '') + stem
  const index = nextPartIndex(zip, new RegExp(`^${escapeRegExp(prefix)}(\\d+)${escapeRegExp(ext)}$`))
  return `${prefix}${index}${ext}`
}

/**
 * Read every chart on every sheet of a workbook, with the parts they depend on.
 * @returns {Promise<object>} an opaque record to hand to `restoreCharts`.
 */
async function captureCharts (workbook) {
  const zip = await JSZip.loadAsync(toBuffer(workbook))
  const typesXml = await readPart(zip, '[Content_Types].xml', '')
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
      await collectPart(zip, chartPath, parts, typesXml)
      charts.push({ relId: m[1], path: chartPath })
    }
    if (!charts.length) continue

    captured.push({
      sheet: name,
      drawingXml,
      charts,
      parts: [...parts].map(([path, part]) => ({ path, ...part })),
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

    let drawingRels = EMPTY_RELS
    let drawingXml = entry.drawingXml
    const parts = new Map(entry.parts.map(part => [part.path, part]))

    // Copy each chart with everything it relates to, under fresh numbers so nothing
    // collides with parts already in the workbook, then repoint the rels at the copies.
    for (const chart of entry.charts) {
      const graph = chartGraph(chart.path, parts)
      const rename = new Map()
      for (const oldPath of graph) {
        const newPath = freshPartPath(zip, oldPath)
        rename.set(oldPath, newPath)
        zip.file(newPath, parts.get(oldPath).data)  // claims the name before the next lookup
      }

      for (const oldPath of graph) {
        const newPath = rename.get(oldPath)
        const contentType = parts.get(oldPath).contentType ||
          (/^xl\/charts\/chart\d+\.xml$/.test(newPath) ? CT.chart : null)
        if (contentType) await declareContentType(zip, `/${newPath}`, contentType)

        const rels = parts.get(relsPathFor(oldPath))
        if (!rels) continue
        let relsXml = rels.data.toString()
        for (const { tag, target } of internalRels(relsXml)) {
          const moved = rename.get(resolveTarget(oldPath, target))
          if (!moved) continue
          const next = target.startsWith('/') ? `/${moved}` : relativeTarget(newPath, moved)
          relsXml = relsXml.replace(tag, tag.replace(/\bTarget\s*=\s*"[^"]*"/, `Target="${next}"`))
        }
        zip.file(relsPathFor(newPath), relsXml)
      }

      const relId = nextRelId(drawingRels)
      drawingRels = addRelationship(drawingRels, relId, 'chart',
        relativeTarget('xl/drawings/drawing.xml', rename.get(chart.path)))
      // repoint this anchor at the relationship id it now has
      drawingXml = drawingXml.replace(
        new RegExp(`(<c:chart\\b[^>]*r:id=")${chart.relId}(")`), `$1${relId}$2`)
    }

    // numbered after the charts, whose user-shape drawings draw from the same sequence
    const drawingIndex = nextPartIndex(zip, /^xl\/drawings\/drawing(\d+)\.xml$/)
    const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`
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
