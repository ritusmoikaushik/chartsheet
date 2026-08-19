// Package validator. Excel reports a broken workbook only as "we found a problem
// with some content", with no detail, so this reports the specific defects that
// cause it. Deliberately dependency-free and regex-based: it checks package
// wiring, which is where these files actually break.
'use strict'

const JSZip = require('jszip')
const { resolveTarget, relsPathFor } = require('./package')

const RE = {
  relationship: /<Relationship\b[^>]*>/g,
  attr: (tag, name) => (tag.match(new RegExp(`${name}="([^"]*)"`)) || [])[1],
  worksheet: /^xl\/worksheets\/sheet\d+\.xml$/,
  chartPart: /^xl\/charts\/chart\d+\.xml$/,
  pivotPart: /^xl\/pivotTables\/pivotTable\d+\.xml$/,
  cacheDefPart: /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/,
  cacheRecPart: /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/,
}

async function validate (workbook) {
  // JSZip takes Buffer, ArrayBuffer, Uint8Array or Blob directly. Only wrap when we
  // have a Buffer to wrap — this file runs in browsers too.
  const zip = await JSZip.loadAsync(
    typeof Buffer !== 'undefined' && !Buffer.isBuffer(workbook) &&
    !(workbook instanceof ArrayBuffer) && !ArrayBuffer.isView(workbook)
      ? Buffer.from(workbook)
      : workbook)
  const names = Object.keys(zip.files).filter(n => !zip.files[n].dir)
  const errors = []
  const warnings = []
  const read = async path => zip.file(path) ? zip.file(path).async('string') : null

  // 1. content types cover every part
  const contentTypes = await read('[Content_Types].xml')
  if (!contentTypes) {
    errors.push('[Content_Types].xml is missing; the package is not a valid OOXML file')
  } else {
    const defaults = new Set([...contentTypes.matchAll(/<Default\b[^>]*Extension="([^"]+)"/g)]
      .map(m => m[1].toLowerCase()))
    const overrides = new Map([...contentTypes.matchAll(
      /<Override\b[^>]*PartName="([^"]+)"[^>]*ContentType="([^"]+)"/g)].map(m => [m[1], m[2]]))

    for (const name of names) {
      if (name === '[Content_Types].xml') continue
      const ext = name.split('.').pop().toLowerCase()
      if (!overrides.has(`/${name}`) && !defaults.has(ext)) {
        errors.push(`no content type declared for ${name} — Excel will refuse the file`)
      }
    }

    // The catch-all <Default Extension="xml"> satisfies the check above but NOT Excel:
    // chart, drawing and worksheet parts each need their own Override, and a missing
    // one is a leading cause of "we found a problem with some content".
    const REQUIRED = [
      [RE.chartPart, 'drawingml.chart+xml', 'chart'],
      [/^xl\/drawings\/drawing\d+\.xml$/, 'officedocument.drawing+xml', 'drawing'],
      [RE.worksheet, 'spreadsheetml.worksheet+xml', 'worksheet'],
      [RE.pivotPart, 'spreadsheetml.pivotTable+xml', 'pivot table'],
      [RE.cacheDefPart, 'spreadsheetml.pivotCacheDefinition+xml', 'pivot cache definition'],
      [RE.cacheRecPart, 'spreadsheetml.pivotCacheRecords+xml', 'pivot cache records'],
    ]
    for (const name of names) {
      for (const [pattern, expected, label] of REQUIRED) {
        if (!pattern.test(name)) continue
        const declared = overrides.get(`/${name}`)
        if (!declared) {
          errors.push(`${name}: no <Override> content type — a ${label} part needs one ` +
            'declared explicitly; the generic xml Default does not satisfy Excel')
        } else if (!declared.includes(expected)) {
          errors.push(`${name}: content type is "${declared}" but a ${label} part must ` +
            `declare ${expected}`)
        }
      }
    }
  }

  // 2. every relationship resolves to a part that exists
  const relsByPath = {}
  for (const name of names.filter(n => n.endsWith('.rels'))) {
    const owner = name.replace('_rels/', '').replace(/\.rels$/, '')
    const xml = await read(name)
    const map = {}
    for (const [tag] of [...xml.matchAll(RE.relationship)].map(m => [m[0]])) {
      if (RE.attr(tag, 'TargetMode') === 'External') continue
      const id = RE.attr(tag, 'Id')
      const target = RE.attr(tag, 'Target')
      const type = RE.attr(tag, 'Type') || ''
      const resolved = resolveTarget(owner, target)
      map[id] = { target: resolved, type }
      if (!names.includes(resolved)) {
        errors.push(`${name}: ${id} points at ${resolved}, which is not in the package`)
      }
    }
    relsByPath[name] = map
  }

  // 3. worksheet -> drawing -> chart wiring
  for (const sheetPath of names.filter(n => RE.worksheet.test(n))) {
    const xml = await read(sheetPath)
    const drawingRefs = [...xml.matchAll(/<drawing\b[^>]*r:id="(rId\d+)"/g)].map(m => m[1])

    if (drawingRefs.length > 1) {
      errors.push(`${sheetPath}: ${drawingRefs.length} <drawing> elements; ` +
        'the schema allows exactly one per worksheet')
    }
    if (drawingRefs.length && !/<\/sheetData>[\s\S]*<drawing\b/.test(xml)) {
      errors.push(`${sheetPath}: <drawing> appears before </sheetData>; ` +
        'Excel requires it as the last element')
    }
    if (!drawingRefs.length) continue

    const sheetRels = relsByPath[relsPathFor(sheetPath)]
    if (!sheetRels) {
      errors.push(`${sheetPath}: references a drawing but has no relationships part`)
      continue
    }
    const entry = sheetRels[drawingRefs[0]]
    if (!entry) {
      errors.push(`${sheetPath}: ${drawingRefs[0]} is not defined in its relationships part`)
      continue
    }

    // a drawing part that is related but not referenced renders nothing at all
    for (const [id, rel] of Object.entries(sheetRels)) {
      if (rel.type.endsWith('/drawing') && rel.target !== entry.target) {
        errors.push(`${sheetPath}: drawing part ${rel.target} (${id}) is related but not ` +
          'referenced by <drawing>; its charts will silently not appear')
      }
    }

    const drawingXml = await read(entry.target)
    if (drawingXml == null) continue
    const shapeIds = [...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*id="(\d+)"/g)].map(m => m[1])
    if (new Set(shapeIds).size !== shapeIds.length) {
      errors.push(`${entry.target}: duplicate shape ids [${shapeIds.join(', ')}]; ` +
        'Excel rejects duplicate cNvPr ids')
    }

    const drawingRels = relsByPath[relsPathFor(entry.target)] || {}
    for (const m of drawingXml.matchAll(/<c:chart\b[^>]*r:id="(rId\d+)"/g)) {
      if (!drawingRels[m[1]]) {
        errors.push(`${entry.target}: chart relationship ${m[1]} is not defined`)
      }
    }
  }

  // 4. chart parts carry at least one series and real cell references
  for (const chartPath of names.filter(n => RE.chartPart.test(n))) {
    const xml = await read(chartPath)
    if (!/<c:chartSpace\b/.test(xml)) {
      errors.push(`${chartPath}: root element is not <c:chartSpace>`)
    }
    const seriesCount = (xml.match(/<c:ser>/g) || []).length
    if (seriesCount === 0) errors.push(`${chartPath}: contains no <c:ser>; the chart will be blank`)
    if (!/<c:f>/.test(xml)) warnings.push(`${chartPath}: no cell references; the chart has no data`)
    if (/<(?!\/?[a-z]+:)[a-zA-Z]+[\s>]/.test(xml.replace(/<\?xml[^>]*\?>/, ''))) {
      errors.push(`${chartPath}: contains unprefixed elements; ` +
        'every element in a chart part must be namespaced')
    }
  }

  // 5. pivot tables: cache wiring is what breaks, and it breaks silently
  const pivotPaths = names.filter(n => RE.pivotPart.test(n))
  if (pivotPaths.length) {
    const workbookXml = await read('xl/workbook.xml')
    const declaredCaches = new Set(
      [...(workbookXml || '').matchAll(/<pivotCache\b[^>]*cacheId="(\d+)"/g)].map(m => m[1]))

    for (const pivotPath of pivotPaths) {
      const xml = await read(pivotPath)
      if (!/<pivotTableDefinition\b/.test(xml)) {
        errors.push(`${pivotPath}: root element is not <pivotTableDefinition>`)
        continue
      }
      const cacheId = (xml.match(/\bcacheId="(\d+)"/) || [])[1]
      if (!cacheId) {
        errors.push(`${pivotPath}: no cacheId; Excel cannot bind the table to its data`)
      } else if (!declaredCaches.has(cacheId)) {
        errors.push(`${pivotPath}: cacheId ${cacheId} is not declared in ` +
          '<pivotCaches> in xl/workbook.xml — the table has no cache to read')
      }
      if (!/<dataFields\b/.test(xml)) {
        errors.push(`${pivotPath}: no <dataFields>; the table has nothing to aggregate`)
      }
      if (!/<rowFields\b/.test(xml) && !/<colFields\b/.test(xml)) {
        errors.push(`${pivotPath}: neither <rowFields> nor <colFields>; the table is empty`)
      }
      // the table part must relate to its cache definition; cacheId alone is not enough
      const rels = await read(relsPathFor(pivotPath))
      if (!rels || !/pivotCacheDefinition/.test(rels)) {
        errors.push(`${pivotPath}: no relationship to a pivotCacheDefinition part — ` +
          'Excel reports a problem with the content')
      }
    }
  }

  // cache records must match what the definition claims, or Excel silently drops rows
  for (const defPath of names.filter(n => RE.cacheDefPart.test(n))) {
    const xml = await read(defPath)
    const claimed = (xml.match(/recordCount="(\d+)"/) || [])[1]
    const relTarget = await read(relsPathFor(defPath))
    if (!relTarget || !/pivotCacheRecords/.test(relTarget)) {
      errors.push(`${defPath}: no relationship to a pivotCacheRecords part`)
      continue
    }
    const recPath = names.find(n => RE.cacheRecPart.test(n))
    if (recPath && claimed != null) {
      const actual = ((await read(recPath)).match(/<r>/g) || []).length
      if (Number(claimed) !== actual) {
        errors.push(`${defPath}: recordCount="${claimed}" but ${recPath} holds ` +
          `${actual} records — Excel will show the wrong totals`)
      }
    }
  }

  return { valid: errors.length === 0, errors, warnings }
}

module.exports = { validate }
