// Runs the documentation. Every complete example on the site is extracted and executed
// in a scratch directory against the real package, because a code sample that does not
// run is worse than no code sample — the reader blames themselves.
//
// A block counts as a complete program when it requires something and calls main().
// Fragments are left alone; the pages carrying them say so in prose.
'use strict'

const fs = require('fs')
const os = require('os')
const path = require('path')
const { execFileSync } = require('child_process')

const DOCS = path.join(__dirname, '..', 'docs')
const ROOT = path.join(__dirname, '..')

const unescapeHtml = s => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
  .replace(/&amp;/g, '&')

function examples () {
  const out = []
  for (const file of fs.readdirSync(DOCS).filter(f => f.endsWith('.html'))) {
    const html = fs.readFileSync(path.join(DOCS, file), 'utf8')
    const blocks = html.match(/<pre><code>[\s\S]*?<\/code><\/pre>/g) || []
    blocks.forEach((block, i) => {
      const code = unescapeHtml(block.replace(/<\/?pre>|<\/?code>/g, '').replace(/<[^>]*>/g, ''))
      const complete = /require\(/.test(code) && /^main\(\)\s*$/m.test(code)
      out.push({ file, index: i + 1, code, complete })
    })
  }
  return out
}

function main () {
  const all = examples()
  const runnable = all.filter(e => e.complete)
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'chartsheet-docs-'))

  // resolve `require('chartsheet')` to the working tree, not a published copy
  fs.mkdirSync(path.join(scratch, 'node_modules'), { recursive: true })
  fs.writeFileSync(path.join(scratch, 'package.json'), JSON.stringify({ name: 'docs-check' }))
  fs.symlinkSync(ROOT, path.join(scratch, 'node_modules', 'chartsheet'), 'junction')
  for (const dep of ['exceljs', 'xlsx', 'jszip']) {
    const from = path.join(ROOT, 'node_modules', dep)
    if (fs.existsSync(from)) {
      fs.symlinkSync(from, path.join(scratch, 'node_modules', dep), 'junction')
    }
  }

  let passed = 0
  let failed = 0
  console.log(`\ndocumentation examples  (${runnable.length} complete of ${all.length} blocks)\n`)

  for (const example of runnable) {
    const name = `${example.file}#${example.index}`
    const script = path.join(scratch, `${example.file.replace(/\W/g, '_')}_${example.index}.js`)
    // examples write into the cwd; give each one the scratch directory
    fs.writeFileSync(script, example.code)
    try {
      execFileSync(process.execPath, [script], { cwd: scratch, stdio: 'pipe', timeout: 60000 })
      console.log(`  ok    ${name}`)
      passed++
    } catch (err) {
      const detail = (err.stderr ? err.stderr.toString() : err.message)
        .split('\n').filter(Boolean).slice(0, 4).join('\n        ')
      console.log(`  FAIL  ${name}\n        ${detail}`)
      failed++
    }
  }

  const fragments = all.filter(e => !e.complete).length
  console.log(`\n${passed} passed, ${failed} failed, ${fragments} fragments not executed`)
  process.exit(failed ? 1 : 0)
}

main()
