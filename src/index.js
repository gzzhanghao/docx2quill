import parseDocx from 'docx'
import Parser from './Parser'

export default async function parse(buffer, options) {
  const parser = new Parser(await parseDocx(buffer), options)

  await parser.parse()

  return parser.delta
}
