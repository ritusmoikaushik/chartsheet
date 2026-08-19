// excel-chart — native Excel charts for xlsx files built with ExcelJS or SheetJS.
'use strict'

const JSZip = require('jszip')
const {
  NS, CT, escapeXml, relsPathFor, nextRelId, addRelationship,
  readPart, resolveSheetPath, ensureDrawing, declareContentType,
} = require('./package')
const { chartXml, validateSpec } = require('./chart-xml')
const { validate } = require('./validate')
const preserve = require('./preserve')
const pivot = require('./pivot')

const EMU_PER_PIXEL = 9525

// One anchor per chart. cNvPr/@id must be unique within the drawing part or Excel
// declares the file corrupt and offers to repair it.
function anchorXml (spec, chartRelId, shapeId) {
  const { col = 5, row = 1, colOffset = 0, rowOffset = 0 } = spec.anchor || {}
  const cx = Math.round((spec.width || 600) * EMU_PER_PIXEL)
  const cy = Math.round((spec.height || 340) * EMU_PER_PIXEL)
  const name = escapeXml(spec.name || `Chart ${shapeId}`)

  return '<xdr:oneCellAnchor>' +
    `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOffset}</xdr:colOff>` +
    `<xdr:row>${row}</xdr:row><xdr:rowOff>${rowOffset}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>' +
    `<xdr:cNvPr id="${shapeId}" name="${name}"/><xdr:cNvGraphicFramePr/>` +
    '</xdr:nvGraphicFramePr>' +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${NS.chart}">` +
    `<c:chart xmlns:c="${NS.chart}" xmlns:r="${NS.rel}" r:id="${chartRelId}"/>` +
    '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>'
}

const toBuffer = value => {
  if (Buffer.isBuffer(value)) return value
  if (value instanceof Uint8Array || value instanceof ArrayBuffer) return Buffer.from(value)
  throw new TypeError('expected an xlsx Buffer, Uint8Array or ArrayBuffer')
}

/**
 * Add one chart to an xlsx workbook.
 *
 * @param {Buffer|Uint8Array|ArrayBuffer} workbook  xlsx bytes, e.g. from
 *        `await wb.xlsx.writeBuffer()` (ExcelJS) or `XLSX.write(wb, {type:'buffer'})`.
 * @param {object} spec  chart definition; see README.
 * @returns {Promise<Buffer>} the workbook with the chart added.
 */
async function addChart (workbook, spec) {
  const validated = validateSpec(spec)
  const zip = await JSZip.loadAsync(toBuffer(workbook))

  const sheetPath = validated.sheetPath || await resolveSheetPath(zip, validated.sheet)
  const drawing = await ensureDrawing(zip, sheetPath)

  const chartIndex = Object.keys(zip.files)
    .filter(name => /^xl\/charts\/chart\d+\.xml$/.test(name)).length + 1
  const chartPath = `xl/charts/chart${chartIndex}.xml`
  zip.file(chartPath, chartXml(validated))

  // the chart relationship belongs to the drawing part, not the worksheet
  const drawingRelsPath = relsPathFor(drawing.path)
  const drawingRels = await readPart(zip, drawingRelsPath)
  const chartRelId = nextRelId(drawingRels)
  zip.file(drawingRelsPath, addRelationship(
    drawingRels, chartRelId, 'chart', `../charts/chart${chartIndex}.xml`))

  const drawingXml = await readPart(zip, drawing.path)
  const usedIds = [...drawingXml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(m => Number(m[1]))
  const shapeId = (usedIds.length ? Math.max(...usedIds) : 1) + 1
  zip.file(drawing.path, drawingXml.replace('</xdr:wsDr>',
    anchorXml(validated, chartRelId, shapeId) + '</xdr:wsDr>'))

  await declareContentType(zip, `/${drawing.path}`, CT.drawing)
  await declareContentType(zip, `/${chartPath}`, CT.chart)

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/** Add several charts in one pass. */
async function addCharts (workbook, specs) {
  let out = workbook
  for (const spec of specs) out = await addChart(out, spec)
  return out
}

module.exports = {
  addChart,
  addCharts,
  validate,
  addPivotTable: pivot.addPivotTable,
  addPivotTables: pivot.addPivotTables,
  captureCharts: preserve.captureCharts,
  restoreCharts: preserve.restoreCharts,
  preserveCharts: preserve.preserveCharts,
  chartCount: preserve.chartCount,
}
