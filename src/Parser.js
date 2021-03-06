import units from 'units-css'
import Delta from 'quill-delta'
import { posix as path } from 'path'

const DEFAULT_OPTIONS = {
  processImage() {},
  notfoundError() {},
}

export default class Parser {

  /**
   * p > pPr
   */
  paragraphAttrs = []

  /**
   * pPr > rPr
   */
  baseRunAttrs = []

  /**
   * r > rPr
   */
  runAttrs = []

  /**
   * pPrDefault > rPr
   */
  defaultRunAttrs = {}

  /**
   * pPrDefault > pPr > rPr
   * rPrDefault > rPr
   */
  defaultParagraphAttrs = {}

  /**
   * hyperlink
   * instrText: HYPERLINK
   */
  link = null

  /**
   * fldChar
   */
  fieldStack = []

  /**
   * result
   */
  delta = new Delta

  constructor(docx, options) {
    this.docx = docx
    this.options = Object.assign({}, DEFAULT_OPTIONS, options)
  }

  parse = async () => {
    if (this.docx.styles && this.docx.styles['w:docDefaults']) {
      await this.parseDefaults(this.docx.styles['w:docDefaults'])
    }
    await this.parseBody(this.docx.document['w:body'])
  }

  parseDefaults = item => {

    if (item['w:pPrDefault'] && item['w:pPrDefault']['w:pPr']) {

      this.paragraphAttrs[0] = {}
      this.baseRunAttrs[0] = {}

      this.parseParagraphProps(item['w:pPrDefault']['w:pPr'])

      this.defaultParagraphAttrs = this.paragraphAttrs[0]
      this.defaultRunAttrs = this.baseRunAttrs[0]
    }

    if (item['w:rPrDefault'] && item['w:rPrDefault']['w:rPr']) {
      Object.assign(this.defaultRunAttrs, this.parseRunProps(item['w:rPrDefault']['w:rPr']))
    }
  }

  parseBody = async item => {
    await each(item.$children, {
      'w:p': this.parseParagraph,
    })
  }

  parseParagraph = async item => {
    this.paragraphAttrs.unshift(Object.assign({}, this.defaultParagraphAttrs))
    this.baseRunAttrs.unshift(Object.assign({}, this.defaultRunAttrs))

    if (item['w:pPr']) {
      await this.parseParagraphProps(item['w:pPr'])
    }

    await each(item.$children, {
      'w:r': this.parseRun,
      'w:hyperlink': this.parseHyperlink,
    })

    this.delta = this.delta.insert('\n', this.paragraphAttrs[0])

    this.paragraphAttrs.shift()
    this.baseRunAttrs.shift()
  }

  parseRun = async item => {
    this.runAttrs.unshift(Object.assign({}, this.baseRunAttrs[0]))

    if (item['w:rPr']) {
      Object.assign(this.runAttrs[0], this.parseRunProps(item['w:rPr']))
    }

    if (this.link) {

      this.runAttrs[0].link = this.link

    } else {

      const HYPERLINK_REGEX = /^HYPERLINK "(.+)"$/

      for (const field of this.fieldStack) {
        const match = field.content.trim().match(HYPERLINK_REGEX)
        if (match) {
          this.runAttrs[0].link = match[1]
        }
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
      'w:noBreakHyphen': this.parseNoBreakHypen,
      'mc:AlternateContent': this.parseAlternateContent,
    })

    this.runAttrs.shift()
  }

  parseHyperlink = async item => {
    const origin = this.link

    this.link = this.getRelationById(item['@r:id'])

    await each(item.$children, {
      'w:r': this.parseRun,
      'w:hyperlink': this.parseHyperlink,
    })

    this.link = origin
  }

  parseBreak = item => {
    // @todo generate soft break
    // @todo handle page break
    this.delta = this.delta.insert('\n', this.paragraphAttrs[0])
  }

  parseText = item => {
    let content = item.$content
    if (!content && item['@xml:space'] === 'preserve') {
      content = ' '
    }
    this.delta = this.delta.insert(content, this.runAttrs[0])
  }

  parseTab = item => {
    this.delta = this.delta.insert('\t', this.runAttrs[0])
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
        if (this.fieldStack[0]) {
          this.fieldStack[0].state = 'separate'
        }
        break
      }

      case 'end': {
        this.fieldStack.shift()
        break
      }
    }
  }

  parseFieldText = item => {
    if (this.fieldStack[0]) {
      this.fieldStack[0].content += item.$content
    }
  }

  parsePicture = async item => {
    await each(item.$children, {
      'v:shape': this.parseShape,
      'v:group': this.parseGroup,
    })
  }

  parseGroup = async item => {
    await each(item.$children, {
      'v:shape': this.parseShape,
      'v:group': this.parseGroup,
    })
  }

  parseShape = async item => {

    await each(item.$children, {

      'v:textbox': this.parseTextBox,

      'v:imagedata': async imgData => {
        try {

          const style = item['@style'] || ''

          const imgPath = this.getRelationById(item['v:imagedata']['@r:id'])
          const size = {
            width: style.match(/\bwidth\s*:\s*([^;]+)/),
            height: style.match(/\bheight\s*:\s*([^;]+)/),
          }

          if (size.width) {
            size.width = units.convert('px', size.width[1])
          }

          if (size.height) {
            size.height = units.convert('px', size.height[1])
          }

          return this.processImage(imgPath, size)

        } catch (error) {

          // noop
        }
      },
    })
  }

  parseDrawing = async item => {
    try {

      const inline = item['wp:inline'] || item['wp:anchor']

      const imgPath = this.getRelationById(inline['a:graphic']['a:graphicData']['pic:pic']['pic:blipFill']['a:blip']['@r:embed'])
      const size = {
        width: inline['wp:extent']['@cx'],
        height: inline['wp:extent']['@cy'],
      }

      // sizes are specified with EMUs
      // see: https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.wordprocessing.extent(v=office.14).aspx

      if (size.width) {
        size.width = (size.width | 0) / 914400 * 96
      }

      if (size.height) {
        size.height = (size.height | 0) / 914400 * 96
      }

      return this.processImage(imgPath, size)

    } catch (error) {

      // noop
    }
  }

  parseTextBox = async item => {
    try {

      await this.parseBody(item['w:txbxContent'])

    } catch (error) {

      // noop
    }
  }

  parseNoBreakHypen = item => {
    this.delta = this.delta.insert('\u2011', this.runAttrs[0])
  }

  parseAlternateContent = item => {
    if (item['mc:Fallback'] && item['mc:Fallback']['w:pict']) {
      return this.parsePicture(item['mc:Fallback']['w:pict'])
    }
  }

  parseParagraphProps = item => {

    if (!item) {
      return
    }

    if (val(item['w:pStyle'])) {
      this.handleParagraphStyle(val(item['w:pStyle']))
    }

    const align = val(item['w:jc'])
    if (['left', 'center', 'right'].includes(align)) {
      this.paragraphAttrs[0].align = align
    }

    if (item['w:numPr']) {
      this.handleParagraphNumbering(item['w:numPr'])
    }
  }

  handleParagraphStyle = styleId => {
    const style = this.getStyleById(styleId)

    if (!style) {
      return
    }

    if (/^heading [1-6]$/i.test(val(style['w:name']))) {
      this.paragraphAttrs[0].header = val(style['w:name']).slice(8) | 0
      return
    }

    this.parseParagraphProps(style['w:pPr'])
    Object.assign(this.baseRunAttrs, this.parseRunProps(style['w:rPr']))
  }

  handleParagraphNumbering = item => {
    const number = this.docx.numbering.numbers.find(num => num['@w:numId'] === val(item['w:numId']))

    if (!number) {
      this.options.notfoundError('number', val(item['w:numId']))
      return
    }

    const abstract = this.docx.numbering.abstracts.find(abstract => abstract['@w:abstractNumId'] === val(number['w:abstractNumId']))

    if (!abstract) {
      this.options.notfoundError('abstract number', val(number['w:abstractNumId']))
      return
    }

    const level = abstract.$children.find(level => level['@w:ilvl'] === val(item['w:ilvl']))

    if (!level) {
      this.options.notfoundError('level', val(item['w:ilvl']))
      return
    }

    this.paragraphAttrs[0].indent = level['@w:ilvl'] | 0

    this.paragraphAttrs[0].list = 'ordered'
    if (val(level['w:numFmt']) === 'bullet') {
      this.paragraphAttrs[0].list = 'bullet'
    }
  }

  parseRunProps = item => {
    const attrs = {}

    if (!item) {
      return attrs
    }

    if (val(item['w:rStyle'])) {
      const style = this.getStyleById(val(item['w:rStyle']))
      if (style) {
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

    if (item['w:rFonts'] && item['w:rFonts']['@w:ascii']) {
      attrs.font = item['w:rFonts']['@w:ascii']
    }

    // font sizes are specified with half-point
    // see: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    const size = val(item['w:sz'])
    if (size) {
      attrs.size = units.convert('px', `${(size | 0) / 2}pt`)
    }

    const color = val(item['w:color'])
    if (color) {
      attrs.color = `#${color}`
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
    const relation = this.docx.relations.$children.find(relation => relation['@Id'] === rId)
    if (relation) {
      return relation['@Target']
    }
    this.options.notfoundError('relationship', rId)
  }

  getStyleById(styleId) {
    if (this.docx.styles) {
      const style = this.docx.styles.$children.find(style => style['@w:styleId'] === styleId)
      if (style) {
        return style
      }
    }
    this.options.notfoundError('style', styleId)
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
