// Pivot table preservation across a read-write cycle.
//
// The same hole as charts, and the oldest open issue about it — ExcelJS #261, filed
// 2017. ExcelJS models no pivot part, so opening a workbook and writing it back
// returns a file with the pivot tables, the caches and the workbook's <pivotCaches>
// entry all gone. Nothing is corrupt; they were simply never written, which is why
// there is no error to catch. Lift them off the original and put them back.
'use strict'

const JSZip = require('jszip')
const {
  CT, relsPathFor, nextRelId, addRelationship, resolveTarget,
  readPart, resolveSheetPath, declareContentType, declarePivotCache, EMPTY_RELS,
} = require('./package')

const toBuffer = value => Buffer.isBuffer(value) ? value : Buffer.from(value)

const PIVOT_PART = /^xl\/pivotTables\/pivotTable\d+\.xml$/
const CACHE_DIR = /^xl\/pivotCache\//

const sheetNames = async zip => {
  const workbook = await readPart(zip, 'xl/workbook.xml')
  return [...workbook.matchAll(/<sheet\b[^>]*\/?>/g)]
    .map(m => (m[0].match(/name="([^"]*)"/) || [])[1])
    .filter(Boolean)
}

// Relationships of one type, in declaration order. Matching on the tail of the type
// keeps pivotTable from also matching pivotCacheDefinition.
const relTargets = (relsXml, type) => (relsXml.match(/<Relationship\b[^>]*?\/?>/g) || [])
  .filter(tag => new RegExp(`Type="[^"]*/${type}"`).test(tag))
  .map(tag => (tag.match(/\bTarget\s*=\s*"([^"]*)"/) || [])[1])
  .filter(Boolean)

// A cache definition relates to its records part, and Excel occasionally writes more
// beside it. Follow the graph, but never outside xl/pivotCache/ — a cache can relate
// back to the source worksheet, and copying that over the edited one would undo the
// very edit this exists to protect.
async function collectCache (zip, partPath, seen) {
  if (!partPath || seen.has(partPath) || !CACHE_DIR.test(partPath)) return
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
    if (target) await collectCache(zip, resolveTarget(partPath, target), seen)
  }
}

/**
 * Read every pivot table on every sheet of a workbook, with the caches behind them.
 * @returns {Promise<object>} an opaque record to hand to `restorePivotTables`.
 */
async function capturePivotTables (workbook) {
  const zip = await JSZip.loadAsync(toBuffer(workbook))
  const sheets = []
  const caches = new Map()

  for (const name of await sheetNames(zip)) {
    const sheetPath = await resolveSheetPath(zip, name)
    const sheetRels = await readPart(zip, relsPathFor(sheetPath), EMPTY_RELS)

    const tables = []
    for (const target of relTargets(sheetRels, 'pivotTable')) {
      const pivotPath = resolveTarget(sheetPath, target)
      const xml = await readPart(zip, pivotPath, null)
      if (xml == null) continue

      // the table reaches its cache through its own rels, not through cacheId
      const pivotRels = await readPart(zip, relsPathFor(pivotPath), EMPTY_RELS)
      const cacheTarget = relTargets(pivotRels, 'pivotCacheDefinition')[0]
      if (!cacheTarget) continue
      const cachePath = resolveTarget(pivotPath, cacheTarget)

      if (!caches.has(cachePath)) {
        const parts = new Map()
        await collectCache(zip, cachePath, parts)
        if (!parts.has(cachePath)) continue
        caches.set(cachePath, [...parts].map(([path, data]) => ({ path, data })))
      }
      tables.push({ xml, cachePath })
    }
    if (tables.length) sheets.push({ sheet: name, tables })
  }

  return {
    version: 1,
    sheets,
    caches: [...caches].map(([path, parts]) => ({ path, parts })),
  }
}

const pivotTableCount = record =>
  (record && record.sheets || []).reduce((n, s) => n + s.tables.length, 0)

// The data under a restored pivot has just been edited by whatever rewrote the
// workbook, so the captured records are stale by definition. refreshOnLoad makes
// Excel rebuild every total from the sheet when the file opens.
const forceRefreshOnLoad = xml => {
  const root = xml.match(/<pivotCacheDefinition\b[^>]*>/)
  if (!root) return xml
  const updated = /\brefreshOnLoad\s*=/.test(root[0])
    ? root[0].replace(/\brefreshOnLoad\s*=\s*"[^"]*"/, 'refreshOnLoad="1"')
    : root[0].replace(/^<pivotCacheDefinition\b/, '<pivotCacheDefinition refreshOnLoad="1"')
  return xml.replace(root[0], updated)
}

const countMatching = (zip, re) => Object.keys(zip.files).filter(n => re.test(n)).length

/**
 * Put previously captured pivot tables back onto a workbook, matching sheets by name.
 * @returns {Promise<Buffer>} the workbook with its pivot tables restored.
 */
async function restorePivotTables (workbook, record) {
  if (!record || !Array.isArray(record.sheets) || !record.sheets.length) {
    return toBuffer(workbook)
  }
  const zip = await JSZip.loadAsync(toBuffer(workbook))
  const present = new Set(await sheetNames(zip))
  const cacheParts = new Map((record.caches || []).map(c => [c.path, c.parts]))

  let nextPivot = countMatching(zip, PIVOT_PART) + 1
  let nextCache = countMatching(zip, /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/) + 1

  // pivot table names are workbook-wide; two called PivotTable1 is a repair prompt
  const usedNames = new Set()
  for (const name of Object.keys(zip.files).filter(n => PIVOT_PART.test(n))) {
    const xml = await readPart(zip, name, '')
    const found = xml.match(/<pivotTableDefinition\b[^>]*\bname="([^"]*)"/)
    if (found) usedNames.add(found[1])
  }

  const written = new Map()   // original cache path -> the numbers it was given here
  const skipped = []

  for (const entry of record.sheets) {
    if (!present.has(entry.sheet)) {
      // the sheet was renamed or removed; its pivot tables have nowhere to go
      skipped.push(entry.sheet)
      continue
    }
    const sheetPath = await resolveSheetPath(zip, entry.sheet)
    const sheetRelsPath = relsPathFor(sheetPath)
    let sheetRels = await readPart(zip, sheetRelsPath, EMPTY_RELS)
    if (relTargets(sheetRels, 'pivotTable').length) {
      // the writer kept a pivot of its own; leave it alone rather than duplicate it
      skipped.push(entry.sheet)
      continue
    }

    for (const table of entry.tables) {
      const parts = cacheParts.get(table.cachePath)
      if (!parts) continue

      let cache = written.get(table.cachePath)
      if (!cache) {
        const index = nextCache++
        const defPath = `xl/pivotCache/pivotCacheDefinition${index}.xml`
        const recPath = `xl/pivotCache/pivotCacheRecords${index}.xml`

        for (const part of parts) {
          // every part of this cache is renumbered together, rels included, so the
          // targets inside them keep pointing at the parts they came with
          const mapped = part.path.replace(
            /(pivotCache(?:Definition|Records))\d+\.xml/, `$1${index}.xml`)
          if (mapped === defPath || /\/_rels\//.test(mapped)) {
            let xml = part.data.toString('utf8').replace(
              /(pivotCache(?:Definition|Records))\d+\.xml/g, `$1${index}.xml`)
            if (mapped === defPath) xml = forceRefreshOnLoad(xml)
            zip.file(mapped, xml)
          } else {
            zip.file(mapped, part.data)
          }
          if (mapped === defPath) await declareContentType(zip, `/${defPath}`, CT.pivotCacheDef)
          if (mapped === recPath) await declareContentType(zip, `/${recPath}`, CT.pivotCacheRec)
        }

        const { cacheId } = await declarePivotCache(
          zip, `pivotCache/pivotCacheDefinition${index}.xml`)
        cache = { cacheId, index }
        written.set(table.cachePath, cache)
      }

      const index = nextPivot++
      const pivotPath = `xl/pivotTables/pivotTable${index}.xml`

      let xml = table.xml.replace(
        /(<pivotTableDefinition\b[^>]*?\bcacheId=")\d+(")/, `$1${cache.cacheId}$2`)
      const name = (xml.match(/<pivotTableDefinition\b[^>]*\bname="([^"]*)"/) || [])[1]
      if (name && usedNames.has(name)) {
        const fresh = `${name}_${index}`
        xml = xml.replace(/(<pivotTableDefinition\b[^>]*?\bname=")[^"]*(")/, `$1${fresh}$2`)
        usedNames.add(fresh)
      } else if (name) {
        usedNames.add(name)
      }
      zip.file(pivotPath, xml)

      // cacheId alone is not enough; the table needs its own rel to the definition
      zip.file(relsPathFor(pivotPath), addRelationship(EMPTY_RELS, 'rId1',
        'pivotCacheDefinition', `../pivotCache/pivotCacheDefinition${cache.index}.xml`))
      await declareContentType(zip, `/${pivotPath}`, CT.pivotTable)

      const relId = nextRelId(sheetRels)
      sheetRels = addRelationship(sheetRels, relId, 'pivotTable',
        `../pivotTables/pivotTable${index}.xml`)
    }

    zip.file(sheetRelsPath, sheetRels)
  }

  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  if (skipped.length) out.skippedSheets = skipped
  return out
}

/**
 * Convenience wrapper: carry the pivot tables from `original` onto `rewritten`.
 *
 *   const original = fs.readFileSync('template.xlsx')
 *   // ... ExcelJS reads it, edits it, writes `rewritten` ...
 *   const output = await preservePivotTables(original, rewritten)
 */
async function preservePivotTables (original, rewritten) {
  return restorePivotTables(rewritten, await capturePivotTables(original))
}

module.exports = {
  capturePivotTables, restorePivotTables, preservePivotTables, pivotTableCount,
}
