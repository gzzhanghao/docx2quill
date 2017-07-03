import Delta from 'quill-delta'
import { posix as path } from 'path'

export default class Parser {

  pAttrs = {}

  rAttrs = {}

  baseRAttrs = {}

  defaultRAttrs = {}

  defaultPAttrs = {}

  link = null

  fieldStack = []

  delta = new Delta

  constructor(docx, options) {
    this.docx = docx
    this.options = options
  }

  parse = async () => {
    await this.parseDefaults(this.docx.styles['w:docDefaults'])
    await this.parseDocument(this.docx.document)
  }

  parseDefaults = item => {
    if (item['w:pPrDefault'] && item['w:pPrDefault']['w:pPr']) {
      this.parseParagraphProps(item['w:pPrDefault']['w:pPr'])
      this.defaultPAttrs = this.pAttrs
      this.defaultRAttrs = this.baseRAttrs
    }
    if (item['w:rPrDefault'] && item['w:rPrDefault']['w:rPr']) {
      Object.assign(this.defaultRAttrs, this.parseRunProps(item['w:rPrDefault']['w:rPr']))
    }
  }

  parseDocument = async item => {
    await each(item['w:body'].$children, {
      'w:p': this.parseParagraph,
    })
  }

  parseParagraph = async item => {
    this.pAttrs = Object.assign({}, this.defaultPAttrs)
    this.baseRAttrs = Object.assign({}, this.defaultRAttrs)

    if (item['w:pPr']) {
      await this.parseParagraphProps(item['w:pPr'])
    }

    await each(item.$children, {
      'w:r': this.parseRun,
      'w:hyperlink': this.parseHyperlink,
    })

    this.delta = this.delta.insert('\n', this.pAttrs)
  }

  parseRun = async item => {
    this.rAttrs = this.baseRAttrs

    if (item['w:rPr']) {
      this.rAttrs = Object.assign({}, this.rAttrs, this.parseRunProps(item['w:rPr']))
    }

    if (this.link) {
      this.rAttrs.link = this.link
    } else {
      const HYPERLINK_REGEX = /^HYPERLINK "(.+)"$/
      const field = this.fieldStack.find(field => field.content.match(HYPERLINK_REGEX))

      if (field) {
        this.rAttrs.link = field.content.match(HYPERLINK_REGEX)[1]
      }
    }

    await each(item.$children, {
      'w:br': this.parseBreak,
      'w:t': this.parseText,
      'w:tab': this.parseTab,
      'w:fldChar': this.parseFieldChar,
      'w:instrText': this.parseFieldText,
      'w:pict': this.parsePicture,
      'w:drawing': this.parseDrawing,
    })
  }

  parseHyperlink = async item => {
    const origin = this.link

    this.link = this.getRelationById(item['@r:id'])

    await each(item.$children, { 'w:r': this.parseRun })

    this.link = origin
  }

  parseBreak = item => {
    // @todo handle page break
    this.delta = this.delta.insert('\n', this.pAttrs)
  }

  parseText = item => {
    let content = item.$content
    if (!content && item['@xml:space'] === 'preserve') {
      content = ' '
    }
    this.delta = this.delta.insert(content, this.rAttrs)
  }

  parseTab = item => {
    this.delta = this.delta.insert('\t', this.rAttrs)
  }

  parseFieldChar = item => {
    switch (item['@w:fldCharType']) {

      case 'begin': {
        this.fieldStack.unshift({
          content: '',
          state: 'begin',
        })
        break
      }

      case 'separate': {
        this.fieldStack[0].state = 'separate'
        break
      }

      case 'end': {
        this.fieldStack.shift()
        break
      }
    }
  }

  parseFieldText = item => {
    this.fieldStack[0].content = item.$content.trim()
  }

  parsePicture = async item => {
    let imgPath = null
    let size = null

    try {

      const shape = item['v:shape']
      const style = shape['@style']

      imgPath = this.getRelationById(shape['v:imagedata']['@r:id'])
      size = {
        width: style.match(/\bwidth\s*:\s*([^;]+)/)[1],
        height: style.match(/\bheight\s*:\s*([^;]+)/)[1],
      }

    } catch (error) {

      return
    }

    await this.processImage(imgPath, size)
  }

  parseDrawing = async item => {
    let imgPath = null
    let size = null

    try {

      const inline = item['wp:inline']

      imgPath = this.getRelationById(inline['a:graphic']['a:graphicData']['pic:pic']['pic:blipFill']['a:blip']['@r:embed'])
      size = {
        width: inline['wp:extent']['@cx'],
        height: inline['wp:extent']['@cy'],
      }

    } catch (error) {

      return
    }

    await this.processImage(imgPath, size)
  }

  parseParagraphProps = item => {

    const styleId = val(item['w:pStyle'])
    if (styleId) {
      const style = this.getStyleById(styleId)

      if (/^heading [1-6]$/i.test(val(style['w:name']))) {

        this.pAttrs.header = val(style['w:name']).slice(8) | 0

      } else {

        if (style['w:pPr']) {
          this.parseParagraphProps(style['w:pPr'])
        }
        if (style['w:rPr']) {
          Object.assign(this.baseRAttrs, this.parseRunProps(style['w:rPr']))
        }
      }
    }

    if (item['w:rPr']) {
      Object.assign(this.baseRAttrs, this.parseRunProps(item['w:rPr']))
    }

    const align = val(item['w:jc'])
    if (['left', 'center', 'right'].includes(align)) {
      this.pAttrs.align = align
    }

    if (item['w:numPr']) {
      const number = this.docx.numbering.numbers.find(num => num['@w:numId'] === val(item['w:numPr']['w:numId']))
      const abstract = this.docx.numbering.abstracts.find(abstract => abstract['@w:abstractNumId'] === val(number['w:abstractNumId']))
      const level = abstract.$children.find(level => level['@w:ilvl'] === val(item['w:numPr']['w:ilvl']))

      this.pAttrs.indent = level['@w:ilvl'] | 0
      this.pAttrs.list = {
        ordered: val(level['w:numFmt']) !== 'bullet',
      }
    }
  }

  parseRunProps = item => {
    const attrs = {}

    const styleId = val(item['w:rStyle'])
    if (styleId) {
      const style = this.getStyleById(styleId)
      if (style['w:rPr']) {
        Object.assign(attrs, this.parseRunProps(style['w:rPr']))
      }
    }

    if (bool(item['w:b'])) {
      attrs.bold = true
    }

    if (bool(item['w:i'])) {
      attrs.italic = true
    }

    if (bool(item['w:strike'])) {
      attrs.strike = true
    }

    const underline = val(item['w:u'])
    if (underline && underline !== 'none') {
      attrs.underline = true
    }

    const size = val(item['w:sz'])
    if (size) {
      attrs.size = size | 0
    }

    const color = val(item['w:color'])
    if (color) {
      attrs.color = color
    }

    const script = val(item['w:vertAlign'])
    const scriptMap = { superscript: 'super', subscript: 'sub' }
    if (script && scriptMap[script]) {
      attrs.script = scriptMap[script]
    }

    return attrs
  }

  async processImage(imgPath, size) {
    const fullPath = path.join('word', imgPath)
    await this.options.processImage(this, fullPath, size)
  }

  getRelationById(rId) {
    return this.docx.relations.$children.find(relation => relation['@Id'] === rId)['@Target']
  }

  getStyleById(styleId) {
    return this.docx.styles.$children.find(style => style['@w:styleId'] === styleId)
  }
}

async function each(list, handlers) {
  for (const item of list) {
    if (handlers[item.$type]) {
      await handlers[item.$type](item)
    }
  }
}

function val(item) {
  return item && item['@w:val']
}

function bool(item) {
  if (!item) {
    return false
  }
  const value = val(item)
  return !value || ['1', 'true'].includes(value)
}
