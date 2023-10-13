
'use strict'

if (process.env.NODE_ENV === 'production') {
  module.exports = require('./editor-to-word.cjs.production.min.js')
} else {
  module.exports = require('./editor-to-word.cjs.development.js')
}
