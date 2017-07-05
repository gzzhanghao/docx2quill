# docx2quill

Convert DocX to Quill Delta.

__This package is under heavy development, use it at your own risk__

## Usage

```javascript
import docx2Quill from '@gzzhanghao/docx2quill'

docx2Quill(docxBuffer, {

  /**
   * Custom image processor
   */
  async processImage(parser, filePath, size) {
    parser.delta = parser.delta.insert({
      image: await uploadImage(await parser.zip.file(filePath).async('nodebuffer'))
    }, {
      width: size.width,
      height: size.height,
    })
  },

  /**
   * Report for item not found error
   */
  notfoundError(type, id) {
    // type: number | abstract number | level | relationship | style
  },

}).then(delta => {

  // delta: Quill Delta
})
```
