// Generation of the c:chartSpace part. Every element here is in the chart
// namespace; a: elements are DrawingML. Both prefixes are declared once on the
// root, so children never redeclare them.
'use strict'

const { NS, escapeXml } = require('./package')

const CAT_AX = '111111111'
const VAL_AX = '222222222'
const SER_AX = '333333333'

const TYPES = new Set([
  'bar', 'column', 'line', 'pie', 'doughnut', 'area', 'scatter', 'radar',
])

const richText = (tag, text) => text == null || text === '' ? '' :
  `<c:${tag}><c:tx><c:rich><a:bodyPr/><a:lstStyle/>` +
  `<a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p>` +
  `</c:rich></c:tx><c:overlay val="0"/></c:${tag}>`

const solidFill = colour => !colour ? '' :
  `<c:spPr><a:solidFill><a:srgbClr val="${String(colour).replace('#', '').toUpperCase()}"/>` +
  '</a:solidFill></c:spPr>'

const dataLabels = spec => !spec.dataLabels ? '' :
  '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
  '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'

const seriesName = series =>
  series.nameRef ? `<c:tx><c:strRef><c:f>${escapeXml(series.nameRef)}</c:f></c:strRef></c:tx>`
    : series.name ? `<c:tx><c:v>${escapeXml(series.name)}</c:v></c:tx>`
      : ''

function renderSeries (series, index, spec) {
  const head = `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    seriesName(series) + solidFill(series.colour || series.color)

  // scatter carries x/y refs rather than categories and values
  if (spec.type === 'scatter') {
    return head +
      `<c:xVal><c:numRef><c:f>${escapeXml(series.xRef || spec.categories)}</c:f></c:numRef></c:xVal>` +
      `<c:yVal><c:numRef><c:f>${escapeXml(series.ref)}</c:f></c:numRef></c:yVal>` +
      '<c:smooth val="0"/></c:ser>'
  }

  const categories = spec.categories
    ? `<c:cat><c:strRef><c:f>${escapeXml(spec.categories)}</c:f></c:strRef></c:cat>` : ''
  const tail = spec.type === 'line' ? '<c:smooth val="0"/>' : ''
  return head + dataLabels(spec) + categories +
    `<c:val><c:numRef><c:f>${escapeXml(series.ref)}</c:f></c:numRef></c:val>` + tail + '</c:ser>'
}

const categoryAxis = spec =>
  `<c:catAx><c:axId val="${CAT_AX}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="b"/>${richText('title', spec.xTitle)}` +
  '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
  `<c:crossAx val="${VAL_AX}"/><c:crosses val="autoZero"/><c:auto val="1"/>` +
  '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'

const valueAxis = (spec, axId = VAL_AX, crossAx = CAT_AX) =>
  `<c:valAx><c:axId val="${axId}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="l"/>` +
  (spec.gridlines === false ? '' : '<c:majorGridlines/>') +
  richText('title', spec.yTitle) +
  `<c:numFmt formatCode="${escapeXml(spec.numberFormat || 'General')}" sourceLinked="0"/>` +
  '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
  `<c:crossAx val="${crossAx}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`

// scatter uses two value axes, not a category axis
const scatterAxes = spec =>
  `<c:valAx><c:axId val="${CAT_AX}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="b"/>${richText('title', spec.xTitle)}` +
  '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
  `<c:crossAx val="${VAL_AX}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>` +
  valueAxis(spec)

function plotArea (spec) {
  const series = spec.series.map((s, i) => renderSeries(s, i, spec)).join('')
  const axisIds = `<c:axId val="${CAT_AX}"/><c:axId val="${VAL_AX}"/>`

  switch (spec.type) {
    case 'pie':
      return `<c:pieChart><c:varyColors val="1"/>${series}${dataLabels(spec)}` +
        '<c:firstSliceAng val="0"/></c:pieChart>'

    case 'doughnut':
      return `<c:doughnutChart><c:varyColors val="1"/>${series}${dataLabels(spec)}` +
        '<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>'

    case 'line':
      return `<c:lineChart><c:grouping val="${spec.stacked ? 'stacked' : 'standard'}"/>` +
        `<c:varyColors val="0"/>${series}<c:marker val="1"/>${axisIds}</c:lineChart>` +
        categoryAxis(spec) + valueAxis(spec)

    case 'area':
      return `<c:areaChart><c:grouping val="${spec.stacked ? 'stacked' : 'standard'}"/>` +
        `<c:varyColors val="0"/>${series}${axisIds}</c:areaChart>` +
        categoryAxis(spec) + valueAxis(spec)

    case 'scatter':
      return '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' +
        `${series}${axisIds}</c:scatterChart>` + scatterAxes(spec)

    case 'radar':
      return '<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>' +
        `${series}${axisIds}</c:radarChart>` + categoryAxis(spec) + valueAxis(spec)

    default: { // bar / column
      const horizontal = spec.type === 'bar' && spec.horizontal !== false && spec.horizontal
      return `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/>` +
        `<c:grouping val="${spec.stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>` +
        `${series}<c:gapWidth val="${spec.gapWidth == null ? 150 : spec.gapWidth}"/>` +
        (spec.stacked ? '<c:overlap val="100"/>' : '') + axisIds + '</c:barChart>' +
        categoryAxis(spec) + valueAxis(spec)
    }
  }
}

function validateSpec (spec) {
  if (!spec || typeof spec !== 'object') throw new Error('a chart spec object is required')
  const type = spec.type || 'bar'
  if (!TYPES.has(type)) {
    throw new Error(`unknown chart type "${type}" (expected one of: ${[...TYPES].join(', ')})`)
  }
  if (!Array.isArray(spec.series) || spec.series.length === 0) {
    throw new Error('chart spec needs at least one entry in series')
  }
  spec.series.forEach((s, i) => {
    if (!s || typeof s.ref !== 'string' || !s.ref) {
      throw new Error(`series[${i}] needs a ref, e.g. "'Sheet1'!$B$2:$B$10"`)
    }
  })
  if (type === 'scatter' && !spec.categories && !spec.series.every(s => s.xRef)) {
    throw new Error('a scatter chart needs categories or an xRef on every series')
  }
  return { ...spec, type: type === 'column' ? 'bar' : type }
}

function chartXml (rawSpec) {
  const spec = validateSpec(rawSpec)
  const legend = spec.legend === false ? ''
    : `<c:legend><c:legendPos val="${typeof spec.legend === 'string' ? spec.legend : 'r'}"/>` +
      '<c:overlay val="0"/></c:legend>'

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<c:chartSpace xmlns:c="${NS.chart}" xmlns:a="${NS.main}" xmlns:r="${NS.rel}">` +
    '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>' +
    `<c:chart>${richText('title', spec.title)}` +
    `<c:autoTitleDeleted val="${spec.title ? 0 : 1}"/>` +
    `<c:plotArea><c:layout/>${plotArea(spec)}</c:plotArea>` +
    legend +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>'
}

module.exports = { chartXml, validateSpec, TYPES, CAT_AX, VAL_AX, SER_AX }
