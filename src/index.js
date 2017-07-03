import parseDocx from '@gzzhanghao/docx'
import Parser from './Parser'

export { Parser }

export default async function parse(buffer, options) {
  const parser = new Parser(await parseDocx(buffer), options)

  await parser.parse()

  return parser.delta
}
