// OOXML package plumbing: locating worksheet parts, reading and writing
// relationships, and keeping [Content_Types].xml in step.
'use strict'

const NS = {
  chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  main: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  sheetDrawing: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  pkgRel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  spreadsheet: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}

const CT = {
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  pivotCacheDef:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRec:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  pivotTable:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
}

const escapeXml = value => String(value)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;')

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<Relationships xmlns="${NS.pkgRel}"></Relationships>`

// Relationship targets are relative to the *folder* of the owning part.
function resolveTarget (ownerPart, target) {
  if (target.startsWith('/')) return target.slice(1)
  const segments = ownerPart.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') segments.pop()
    else if (segment !== '.') segments.push(segment)
  }
  return segments.join('/')
}

const relsPathFor = part => {
  const bits = part.split('/')
  return `${bits.slice(0, -1).join('/')}/_rels/${bits[bits.length - 1]}.rels`
}

const nextRelId = relsXml => {
  const used = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1]))
  return `rId${(used.length ? Math.max(...used) : 0) + 1}`
}

const addRelationship = (relsXml, id, type, target) => relsXml.replace(
  '</Relationships>',
  `<Relationship Id="${id}" Type="${NS.rel}/${type}" Target="${target}"/></Relationships>`)

// Attribute order inside a <Relationship> is not fixed and producers disagree, so match
// the tag first and read its attributes independently rather than assuming Id comes first.
const findRelTarget = (relsXml, id) => {
  for (const tag of relsXml.match(/<Relationship\b[^>]*?\/?>/g) || []) {
    const found = tag.match(/\bId\s*=\s*"([^"]*)"/)
    if (!found || found[1] !== id) continue
    const target = tag.match(/\bTarget\s*=\s*"([^"]*)"/)
    return target ? target[1] : null
  }
  return null
}

async function readPart (zip, path, fallback) {
  const file = zip.file(path)
  if (file) return file.async('string')
  if (fallback !== undefined) return fallback
  throw new Error(`missing part: ${path}`)
}

/**
 * Map a worksheet name to its part path, via workbook.xml and its relationships.
 * Falls back to the first sheet when no name is given.
 */
async function resolveSheetPath (zip, sheetName) {
  const workbook = await readPart(zip, 'xl/workbook.xml')
  const rels = await readPart(zip, 'xl/_rels/workbook.xml.rels')

  const sheets = [...workbook.matchAll(/<sheet\b[^>]*\/?>/g)].map(m => {
    const tag = m[0]
    const name = tag.match(/name="([^"]*)"/)
    const rid = tag.match(/r:id="([^"]+)"/)
    return { name: name ? name[1] : null, rid: rid ? rid[1] : null }
  })
  if (!sheets.length) throw new Error('workbook declares no worksheets')

  let sheet
  if (sheetName == null) {
    sheet = sheets[0]
  } else {
    // sheet names in workbook.xml are XML-escaped
    const wanted = escapeXml(sheetName)
    sheet = sheets.find(s => s.name === wanted || s.name === sheetName)
    if (!sheet) {
      const known = sheets.map(s => s.name).join(', ')
      throw new Error(`no worksheet named "${sheetName}" (found: ${known})`)
    }
  }

  const target = findRelTarget(rels, sheet.rid)
  if (!target) throw new Error(`worksheet "${sheet.name}" has no resolvable relationship`)
  return resolveTarget('xl/workbook.xml', target)
}

/**
 * Put a <drawing> reference into a worksheet, in the place the schema requires.
 *
 * CT_Worksheet is a fixed sequence, and <drawing> is NOT last in it: legacyDrawing,
 * picture, oleObjects, controls, webPublishItems, tableParts and extLst all come
 * after it. Appending before </worksheet> corrupts any sheet that has an Excel
 * table on it, which is exactly the kind of sheet that also has a chart.
 */
const AFTER_DRAWING = [
  'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects',
  'controls', 'webPublishItems', 'tableParts', 'extLst',
]

function insertDrawing (sheetXml, relId) {
  const element = '<drawing xmlns:r="' + NS.rel + '" r:id="' + relId + '"/>'
  let at = -1
  for (const name of AFTER_DRAWING) {
    for (const opener of ['<' + name + '>', '<' + name + ' ', '<' + name + '/']) {
      const i = sheetXml.indexOf(opener)
      if (i !== -1 && (at === -1 || i < at)) at = i
    }
  }
  return at === -1
    ? sheetXml.replace('</worksheet>', element + '</worksheet>')
    : sheetXml.slice(0, at) + element + sheetXml.slice(at)
}

/**
 * A worksheet owns at most ONE drawing part; every chart or image on the sheet is
 * another anchor inside it. Return that part, creating it and wiring it up if absent.
 */
async function ensureDrawing (zip, sheetPath) {
  let sheetXml = await readPart(zip, sheetPath)
  const sheetRelsPath = relsPathFor(sheetPath)
  let sheetRels = await readPart(zip, sheetRelsPath, EMPTY_RELS)

  const existing = sheetXml.match(/<drawing[^>]*r:id="(rId\d+)"/)
  if (existing) {
    const target = findRelTarget(sheetRels, existing[1])
    if (!target) {
      throw new Error(`${sheetPath} references ${existing[1]} but its rels part does not define it`)
    }
    return { path: resolveTarget(sheetPath, target) }
  }

  const index = Object.keys(zip.files)
    .filter(name => /^xl\/drawings\/drawing\d+\.xml$/.test(name)).length + 1
  const path = `xl/drawings/drawing${index}.xml`

  zip.file(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<xdr:wsDr xmlns:xdr="${NS.sheetDrawing}" xmlns:a="${NS.main}"></xdr:wsDr>`)
  zip.file(relsPathFor(path), EMPTY_RELS)

  const relId = nextRelId(sheetRels)
  sheetRels = addRelationship(sheetRels, relId, 'drawing', `../drawings/drawing${index}.xml`)
  zip.file(sheetRelsPath, sheetRels)

  zip.file(sheetPath, insertDrawing(sheetXml, relId))

  return { path }
}

/**
 * Bind a pivot cache definition to the workbook: a relationship, plus the
 * <pivotCaches> entry that gives the cache the id every pivot table refers to.
 *
 * Both halves are required. A cache part with no entry here is invisible to Excel,
 * and a table whose cacheId is undeclared makes Excel offer to repair the file.
 *
 * @returns {Promise<{cacheId: number, relId: string}>}
 */
async function declarePivotCache (zip, cacheDefTarget) {
  const relsPath = 'xl/_rels/workbook.xml.rels'
  let rels = await readPart(zip, relsPath, EMPTY_RELS)
  const relId = nextRelId(rels)
  rels = addRelationship(rels, relId, 'pivotCacheDefinition', cacheDefTarget)
  zip.file(relsPath, rels)

  const workbookXml = await readPart(zip, 'xl/workbook.xml')
  const used = [...workbookXml.matchAll(/cacheId="(\d+)"/g)].map(m => Number(m[1]))
  const cacheId = (used.length ? Math.max(...used) : 0) + 1
  const entry = `<pivotCache cacheId="${cacheId}" r:id="${relId}"/>`

  let next
  if (/<pivotCaches>/.test(workbookXml)) {
    next = workbookXml.replace('</pivotCaches>', `${entry}</pivotCaches>`)
  } else {
    // pivotCaches sits late in the CT_Workbook sequence, but ahead of extLst
    const block = `<pivotCaches>${entry}</pivotCaches>`
    next = /<extLst\b/.test(workbookXml)
      ? workbookXml.replace(/<extLst\b/, `${block}<extLst`)
      : workbookXml.replace('</workbook>', `${block}</workbook>`)
  }
  zip.file('xl/workbook.xml', next)

  return { cacheId, relId }
}

async function declareContentType (zip, partName, contentType) {
  const path = '[Content_Types].xml'
  let xml = await readPart(zip, path)
  if (xml.includes(`PartName="${partName}"`)) return
  zip.file(path, xml.replace('</Types>',
    `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`))
}

module.exports = {
  NS, CT, EMPTY_RELS,
  escapeXml, resolveTarget, relsPathFor, nextRelId, addRelationship, findRelTarget,
  readPart, resolveSheetPath, ensureDrawing, declareContentType, declarePivotCache,
  insertDrawing,
}
