// Bundles the validator for the browser page in docs/. The published npm package is
// unaffected — package.json ships src/ only.
'use strict'

const esbuild = require('esbuild')
const path = require('path')

esbuild.build({
  entryPoints: [path.join(__dirname, 'web-entry.js')],
  bundle: true,
  minify: true,
  format: 'iife',
  globalName: 'chartsheetWeb',
  target: ['es2019'],
  platform: 'browser',
  outfile: path.join(__dirname, '..', 'docs', 'validator.js'),
  legalComments: 'none',
}).then(() => {
  const { statSync } = require('fs')
  const kb = Math.round(statSync(path.join(__dirname, '..', 'docs', 'validator.js')).size / 1024)
  console.log(`docs/validator.js  ${kb}KB`)
}).catch(err => { console.error(err); process.exit(1) })
