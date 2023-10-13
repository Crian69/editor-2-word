'use strict';

Object.defineProperty(exports, '__esModule', { value: true });

function _interopDefault (ex) { return (ex && (typeof ex === 'object') && 'default' in ex) ? ex['default'] : ex; }

var docx = require('docx');
var tinycolor = _interopDefault(require('tinycolor2'));
var JSZip = _interopDefault(require('jszip'));
var htmlToAst = require('html-to-ast');
var fileSaver = require('file-saver');

function _regeneratorRuntime() {
  _regeneratorRuntime = function () {
    return e;
  };
  var t,
    e = {},
    r = Object.prototype,
    n = r.hasOwnProperty,
    o = Object.defineProperty || function (t, e, r) {
      t[e] = r.value;
    },
    i = "function" == typeof Symbol ? Symbol : {},
    a = i.iterator || "@@iterator",
    c = i.asyncIterator || "@@asyncIterator",
    u = i.toStringTag || "@@toStringTag";
  function define(t, e, r) {
    return Object.defineProperty(t, e, {
      value: r,
      enumerable: !0,
      configurable: !0,
      writable: !0
    }), t[e];
  }
  try {
    define({}, "");
  } catch (t) {
    define = function (t, e, r) {
      return t[e] = r;
    };
  }
  function wrap(t, e, r, n) {
    var i = e && e.prototype instanceof Generator ? e : Generator,
      a = Object.create(i.prototype),
      c = new Context(n || []);
    return o(a, "_invoke", {
      value: makeInvokeMethod(t, r, c)
    }), a;
  }
  function tryCatch(t, e, r) {
    try {
      return {
        type: "normal",
        arg: t.call(e, r)
      };
    } catch (t) {
      return {
        type: "throw",
        arg: t
      };
    }
  }
  e.wrap = wrap;
  var h = "suspendedStart",
    l = "suspendedYield",
    f = "executing",
    s = "completed",
    y = {};
  function Generator() {}
  function GeneratorFunction() {}
  function GeneratorFunctionPrototype() {}
  var p = {};
  define(p, a, function () {
    return this;
  });
  var d = Object.getPrototypeOf,
    v = d && d(d(values([])));
  v && v !== r && n.call(v, a) && (p = v);
  var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p);
  function defineIteratorMethods(t) {
    ["next", "throw", "return"].forEach(function (e) {
      define(t, e, function (t) {
        return this._invoke(e, t);
      });
    });
  }
  function AsyncIterator(t, e) {
    function invoke(r, o, i, a) {
      var c = tryCatch(t[r], t, o);
      if ("throw" !== c.type) {
        var u = c.arg,
          h = u.value;
        return h && "object" == typeof h && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) {
          invoke("next", t, i, a);
        }, function (t) {
          invoke("throw", t, i, a);
        }) : e.resolve(h).then(function (t) {
          u.value = t, i(u);
        }, function (t) {
          return invoke("throw", t, i, a);
        });
      }
      a(c.arg);
    }
    var r;
    o(this, "_invoke", {
      value: function (t, n) {
        function callInvokeWithMethodAndArg() {
          return new e(function (e, r) {
            invoke(t, n, e, r);
          });
        }
        return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg();
      }
    });
  }
  function makeInvokeMethod(e, r, n) {
    var o = h;
    return function (i, a) {
      if (o === f) throw new Error("Generator is already running");
      if (o === s) {
        if ("throw" === i) throw a;
        return {
          value: t,
          done: !0
        };
      }
      for (n.method = i, n.arg = a;;) {
        var c = n.delegate;
        if (c) {
          var u = maybeInvokeDelegate(c, n);
          if (u) {
            if (u === y) continue;
            return u;
          }
        }
        if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) {
          if (o === h) throw o = s, n.arg;
          n.dispatchException(n.arg);
        } else "return" === n.method && n.abrupt("return", n.arg);
        o = f;
        var p = tryCatch(e, r, n);
        if ("normal" === p.type) {
          if (o = n.done ? s : l, p.arg === y) continue;
          return {
            value: p.arg,
            done: n.done
          };
        }
        "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg);
      }
    };
  }
  function maybeInvokeDelegate(e, r) {
    var n = r.method,
      o = e.iterator[n];
    if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y;
    var i = tryCatch(o, e.iterator, r.arg);
    if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y;
    var a = i.arg;
    return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y);
  }
  function pushTryEntry(t) {
    var e = {
      tryLoc: t[0]
    };
    1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e);
  }
  function resetTryEntry(t) {
    var e = t.completion || {};
    e.type = "normal", delete e.arg, t.completion = e;
  }
  function Context(t) {
    this.tryEntries = [{
      tryLoc: "root"
    }], t.forEach(pushTryEntry, this), this.reset(!0);
  }
  function values(e) {
    if (e || "" === e) {
      var r = e[a];
      if (r) return r.call(e);
      if ("function" == typeof e.next) return e;
      if (!isNaN(e.length)) {
        var o = -1,
          i = function next() {
            for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next;
            return next.value = t, next.done = !0, next;
          };
        return i.next = i;
      }
    }
    throw new TypeError(typeof e + " is not iterable");
  }
  return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", {
    value: GeneratorFunctionPrototype,
    configurable: !0
  }), o(GeneratorFunctionPrototype, "constructor", {
    value: GeneratorFunction,
    configurable: !0
  }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) {
    var e = "function" == typeof t && t.constructor;
    return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name));
  }, e.mark = function (t) {
    return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t;
  }, e.awrap = function (t) {
    return {
      __await: t
    };
  }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () {
    return this;
  }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) {
    void 0 === i && (i = Promise);
    var a = new AsyncIterator(wrap(t, r, n, o), i);
    return e.isGeneratorFunction(r) ? a : a.next().then(function (t) {
      return t.done ? t.value : a.next();
    });
  }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () {
    return this;
  }), define(g, "toString", function () {
    return "[object Generator]";
  }), e.keys = function (t) {
    var e = Object(t),
      r = [];
    for (var n in e) r.push(n);
    return r.reverse(), function next() {
      for (; r.length;) {
        var t = r.pop();
        if (t in e) return next.value = t, next.done = !1, next;
      }
      return next.done = !0, next;
    };
  }, e.values = values, Context.prototype = {
    constructor: Context,
    reset: function (e) {
      if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t);
    },
    stop: function () {
      this.done = !0;
      var t = this.tryEntries[0].completion;
      if ("throw" === t.type) throw t.arg;
      return this.rval;
    },
    dispatchException: function (e) {
      if (this.done) throw e;
      var r = this;
      function handle(n, o) {
        return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o;
      }
      for (var o = this.tryEntries.length - 1; o >= 0; --o) {
        var i = this.tryEntries[o],
          a = i.completion;
        if ("root" === i.tryLoc) return handle("end");
        if (i.tryLoc <= this.prev) {
          var c = n.call(i, "catchLoc"),
            u = n.call(i, "finallyLoc");
          if (c && u) {
            if (this.prev < i.catchLoc) return handle(i.catchLoc, !0);
            if (this.prev < i.finallyLoc) return handle(i.finallyLoc);
          } else if (c) {
            if (this.prev < i.catchLoc) return handle(i.catchLoc, !0);
          } else {
            if (!u) throw new Error("try statement without catch or finally");
            if (this.prev < i.finallyLoc) return handle(i.finallyLoc);
          }
        }
      }
    },
    abrupt: function (t, e) {
      for (var r = this.tryEntries.length - 1; r >= 0; --r) {
        var o = this.tryEntries[r];
        if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) {
          var i = o;
          break;
        }
      }
      i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null);
      var a = i ? i.completion : {};
      return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a);
    },
    complete: function (t, e) {
      if ("throw" === t.type) throw t.arg;
      return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y;
    },
    finish: function (t) {
      for (var e = this.tryEntries.length - 1; e >= 0; --e) {
        var r = this.tryEntries[e];
        if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y;
      }
    },
    catch: function (t) {
      for (var e = this.tryEntries.length - 1; e >= 0; --e) {
        var r = this.tryEntries[e];
        if (r.tryLoc === t) {
          var n = r.completion;
          if ("throw" === n.type) {
            var o = n.arg;
            resetTryEntry(r);
          }
          return o;
        }
      }
      throw new Error("illegal catch attempt");
    },
    delegateYield: function (e, r, n) {
      return this.delegate = {
        iterator: values(e),
        resultName: r,
        nextLoc: n
      }, "next" === this.method && (this.arg = t), y;
    }
  }, e;
}
function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) {
  try {
    var info = gen[key](arg);
    var value = info.value;
  } catch (error) {
    reject(error);
    return;
  }
  if (info.done) {
    resolve(value);
  } else {
    Promise.resolve(value).then(_next, _throw);
  }
}
function _asyncToGenerator(fn) {
  return function () {
    var self = this,
      args = arguments;
    return new Promise(function (resolve, reject) {
      var gen = fn.apply(self, args);
      function _next(value) {
        asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value);
      }
      function _throw(err) {
        asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err);
      }
      _next(undefined);
    });
  };
}
function _extends() {
  _extends = Object.assign ? Object.assign.bind() : function (target) {
    for (var i = 1; i < arguments.length; i++) {
      var source = arguments[i];
      for (var key in source) {
        if (Object.prototype.hasOwnProperty.call(source, key)) {
          target[key] = source[key];
        }
      }
    }
    return target;
  };
  return _extends.apply(this, arguments);
}
function _unsupportedIterableToArray(o, minLen) {
  if (!o) return;
  if (typeof o === "string") return _arrayLikeToArray(o, minLen);
  var n = Object.prototype.toString.call(o).slice(8, -1);
  if (n === "Object" && o.constructor) n = o.constructor.name;
  if (n === "Map" || n === "Set") return Array.from(o);
  if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen);
}
function _arrayLikeToArray(arr, len) {
  if (len == null || len > arr.length) len = arr.length;
  for (var i = 0, arr2 = new Array(len); i < len; i++) arr2[i] = arr[i];
  return arr2;
}
function _createForOfIteratorHelperLoose(o, allowArrayLike) {
  var it = typeof Symbol !== "undefined" && o[Symbol.iterator] || o["@@iterator"];
  if (it) return (it = it.call(o)).next.bind(it);
  if (Array.isArray(o) || (it = _unsupportedIterableToArray(o)) || allowArrayLike && o && typeof o.length === "number") {
    if (it) o = it;
    var i = 0;
    return function () {
      if (i >= o.length) return {
        done: true
      };
      return {
        done: false,
        value: o[i++]
      };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}

var Splitter_Colon = ':';
var Splitter_Semicolon = ';';
// px by twips
var PXbyTWIPS = 16;
// px by pt
var PXbyPT = 3 / 4;
// default font size in px
var D_FontSizePX = 16.3;
// default font size in pt
var D_FontSizePT = D_FontSizePX * PXbyPT;
// default line height
var D_LineHeight = 1.5;
// default table full width in DXA
var D_TableFullWidth = 9035;
var D_TableBorderColor = '444444';
// table cell margin in twips
var D_CELL_MARGIN = 1 * PXbyTWIPS;
// table border width in px
var D_TableBorderSize = 2;
// table cell height in px
var D_TableCellHeightPx = 18;
// text-align
var AlignMap = {
  left: docx.AlignmentType.LEFT,
  center: docx.AlignmentType.CENTER,
  right: docx.AlignmentType.RIGHT
};
var hyperlinkColor = '#007AFF';
// style with tag
var D_TagStyleMap = {
  p: "line-height: " + D_LineHeight + ";",
  strong: 'font-weight: bold;',
  em: 'font-style: italic;',
  u: 'text-decoration: underline;',
  del: 'text-decoration: line-through;',
  h1: "font-weight: bold; font-size: 40px; line-height: " + D_LineHeight + ";",
  h2: "font-weight: bold; font-size: 36px; line-height: " + D_LineHeight + ";",
  h3: "font-weight: bold; font-size: 24px; line-height: " + D_LineHeight + ";",
  h4: "font-weight: bold; font-size: 18px; line-height: " + D_LineHeight + ";",
  h5: "font-weight: bold; font-size: 15px; line-height: " + D_LineHeight + ";",
  h6: "font-weight: bold; font-size: 13px; line-height: " + D_LineHeight + ";",
  sub: 'subscript: true;',
  sup: 'superscript: true;',
  a: "text-decoration: underline; color: " + hyperlinkColor + ";"
};
// default paper layout
var D_Layout = {
  bottomMargin: '2.54cm',
  leftMargin: '3.18cm',
  rightMargin: '3.18cm',
  topMargin: '2.54cm',
  orientation: docx.PageOrientation.PORTRAIT
};
// Direction
var Direction = {
  left: 'left',
  right: 'right',
  firstLine: 'firstLine',
  start: 'start',
  end: 'end',
  hanging: 'hanging'
};
var PaddingDirection = {
  'padding-left': Direction.left,
  'padding-right': Direction.right,
  'padding-top': Direction.start,
  'padding-bottom': Direction.end
};
// Size
var Size = {
  em: 'em',
  px: 'px',
  pt: 'pt'
};
// single line
var SingleLine = {
  type: 'single',
  color: '3d4757'
};
var TagType = {
  table: 'table',
  link: 'a',
  text: 'text',
  img: 'img',
  ordered_list: 'ol',
  unordered_list: 'ul'
};
// default border style
var DefaultBorder = {
  style: docx.BorderStyle.SINGLE,
  size: 0,
  color: '#fff'
};
// table cell vertical align map
var verticalAlignMap = {
  top: docx.VerticalAlign.TOP,
  middle: docx.VerticalAlign.CENTER,
  bottom: docx.VerticalAlign.BOTTOM
};

function typeOf(obj) {
  var toString = Object.prototype.toString;
  var map = {
    '[object Boolean]': 'boolean',
    '[object Number]': 'number',
    '[object String]': 'string',
    '[object Function]': 'function',
    '[object Array]': 'array',
    '[object Date]': 'date',
    '[object RegExp]': 'regExp',
    '[object Undefined]': 'undefined',
    '[object Null]': 'null',
    '[object Object]': 'object'
  };
  // @ts-ignore
  return map[toString.call(obj)];
}
var isFilledArray = function isFilledArray(arr) {
  return Array.isArray(arr) && arr.length > 0;
};
// unique array by given key
var getUniqueArrayByKey = function getUniqueArrayByKey(arr, uniqueKey) {
  if (uniqueKey === void 0) {
    uniqueKey = 'id';
  }
  var isEveryObject = arr.every(function (item) {
    return typeOf(item) === 'object';
  });
  if (!isFilledArray(arr) || arr.length === 1 || !isEveryObject) return arr;
  var hash = [];
  return arr.reduce(function (item, next) {
    var k = next[uniqueKey];
    if (k && !hash.includes(k)) {
      hash.push(k);
      item.push(next);
    }
    return item;
  }, []);
};
var removeTagDIV = function removeTagDIV(str) {
  var reg = /<div[^>]*?>|<\/div>/gi;
  return str.replace(reg, '');
};
var escape2Html = function escape2Html(str) {
  var arrEntities = {
    lt: '<',
    gt: '>',
    nbsp: ' ',
    amp: '&',
    quot: '"'
  };
  return str.replace(/&(lt|gt|nbsp|amp|quot);/gi, function (_, t) {
    // @ts-ignore
    return arrEntities[t];
  });
};
var trimHtml = function trimHtml(str) {
  return removeTagDIV(escape2Html(str));
};
var deepCopyByJSON = function deepCopyByJSON(obj) {
  return JSON.parse(JSON.stringify(obj));
};
var isValidColor = function isValidColor(color) {
  return tinycolor(color).isValid();
};
var toHex = function toHex(color) {
  return tinycolor(color).toHexString();
};
/**
 * parse size
 */
var handleSizeNumber = function handleSizeNumber(val) {
  var m = val.match(/\d+(.\d+)?/g);
  if (val.match(/\d+(.\d+)?/g) && m && Array.isArray(m) && m[0]) {
    var target = m[0];
    var type = target ? val.replace(new RegExp(target, 'g'), '') : '';
    return {
      value: parseFloat(target),
      type: type
    };
  }
  return {
    type: 'UNKNOWN',
    value: 0
  };
};
// parse '2.54cm' to 2.54
var numberCM = function numberCM(size) {
  return parseFloat(size == null ? void 0 : size.toUpperCase().replace(/CM/i, ''));
};
// calc margin in twip
var calcMargin = function calcMargin(margin) {
  return docx.convertMillimetersToTwip(10 * numberCM(margin));
};
var optimizeBlankSpace = function optimizeBlankSpace(content, ratio) {
  if (ratio === void 0) {
    ratio = 1;
  }
  var textWithoutBlank = content.trimEnd();
  var blank = content.slice(textWithoutBlank.length);
  var optimizedBlank = ratio === 1 ? blank : new Array(blank.length * ratio).fill(' ').join('');
  var text = blank.length > 1 ? "" + textWithoutBlank + optimizedBlank + '\t' : content;
  return text;
};
var getImageBlob = /*#__PURE__*/function () {
  var _ref = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(src) {
    var blob;
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          _context.next = 2;
          return fetch(src).then(function (res) {
            return res.blob();
          });
        case 2:
          blob = _context.sent;
          return _context.abrupt("return", blob);
        case 4:
        case "end":
          return _context.stop();
      }
    }, _callee);
  }));
  return function getImageBlob(_x) {
    return _ref.apply(this, arguments);
  };
}();

// style map
var StyleMap = {
  fontFamily: 'font-family',
  textAlign: 'text-align',
  paddingRight: 'padding-right',
  paddingLeft: 'padding-left',
  lineHeight: 'line-height',
  fontSize: 'font-size',
  color: 'color',
  backgroundColor: 'background-color',
  textDecoration: 'text-decoration',
  textIndent: 'text-indent',
  borderColor: 'border-color',
  height: 'height',
  width: 'width',
  fontWeight: 'font-weight',
  verticalAlign: 'vertical-align',
  lineThrough: 'line-through',
  underline: 'underline',
  fontStyle: 'font-style',
  subScript: 'subscript',
  superScript: 'superscript'
};

var backgroundHandler = function backgroundHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.shading = {
    type: docx.ShadingType.CLEAR,
    fill: val.replace(/#/g, '')
  };
  return styleOption;
};

var superScriptHandler = function superScriptHandler(_, styleOp) {
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.superScript = true;
  return styleOption;
};

var subScriptHandler = function subScriptHandler(_, styleOp) {
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.subScript = true;
  return styleOption;
};

var colorHandler = function colorHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.color = val.replace(/#/g, '');
  return styleOption;
};

var widthHandler = function widthHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  var w = parseFloat(val.replace(/%/i, ''));
  styleOption.tWidth = w;
  return styleOption;
};

var verticalAlignHandler = function verticalAlignHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.verticalAlign = verticalAlignMap[val];
  return styleOption;
};

var textDecorationHandler = function textDecorationHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  if (val === StyleMap.lineThrough) {
    styleOption.strike = true;
  } else if (val === StyleMap.underline) {
    styleOption.underline = SingleLine;
  }
  return styleOption;
};

var paddingHandler = function paddingHandler(_ref, styleOp) {
  var key = _ref.key,
    val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  var dire = PaddingDirection[key] || Direction.left;
  var _handleSizeNumber = handleSizeNumber(val),
    value = _handleSizeNumber.value,
    type = _handleSizeNumber.type;
  // handle indent
  var indent = {};
  var size = styleOption.size || D_FontSizePX;
  var oneCharSizePT = size / PXbyPT / 2 * PXbyTWIPS;
  var isEM = type.match(Size.em);
  var isPX = type.match(Size.px);
  var isPT = type.match(Size.pt);
  var indentValue = 0;
  if (isEM) {
    indentValue = value * oneCharSizePT;
  } else if (isPX) {
    indentValue = value / 20 * oneCharSizePT;
  } else if (isPT) {
    indentValue = value / D_FontSizePT * oneCharSizePT;
  }
  indent[dire] = indentValue;
  styleOption.indent = indent;
  return styleOption;
};

var lineHeightHandler = function lineHeightHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  var spacing = {};
  var _handleSizeNumber = handleSizeNumber(val),
    value = _handleSizeNumber.value,
    type = _handleSizeNumber.type;
  var lineHeightToSpace = 240;
  var isPx = type.toLowerCase() === 'px';
  var isPr = type.toLowerCase() == '%';
  var lineHeightVal = value;
  if (isPx && value) {
    lineHeightVal = value / 16;
  } else if (isPr) {
    lineHeightVal = value / 100;
  }
  // when line-height is 1.0 these is no need to set spacing
  var isNoSpacing = lineHeightVal == 1;
  if (value && !isNoSpacing) {
    var s = lineHeightVal * lineHeightToSpace;
    spacing.line = s;
    spacing.lineRule = docx.LineRuleType.AUTO;
  }
  styleOption.spacing = spacing;
  return styleOption;
};

var heightHandler = function heightHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  var h = parseFloat(val.replace(/px/i, ''));
  styleOption.tHeight = h;
  return styleOption;
};

var alignHandler = function alignHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.alignment = AlignMap[val] || docx.AlignmentType.LEFT;
  return styleOption;
};

var boldHandler = function boldHandler(_, styleOp) {
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.bold = true;
  return styleOption;
};

var borderColorHandler = function borderColorHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.borderColor = val.replace(/#/i, '');
  return styleOption;
};

var fontFamilyHandler = function fontFamilyHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  if (val.indexOf(',') === -1 && val.indexOf(' ') === -1) {
    styleOption.font = val;
  }
  return styleOption;
};

var fontStyleHandler = function fontStyleHandler(_, styleOp) {
  var styleOption = deepCopyByJSON(styleOp);
  styleOption.italics = true;
  return styleOption;
};

var textIndentHandler = function textIndentHandler(_ref, styleOp) {
  var val = _ref.val;
  var styleOption = deepCopyByJSON(styleOp);
  var _handleSizeNumber = handleSizeNumber(val),
    value = _handleSizeNumber.value,
    type = _handleSizeNumber.type;
  var indent = {};
  var size = styleOption.size || D_FontSizePX;
  var oneCharSizePT = size / PXbyPT / 2 * PXbyTWIPS;
  var isEM = type.match(Size.em);
  var isPX = type.match(Size.px);
  var isPT = type.match(Size.pt);
  var indentValue = 0;
  if (isEM) {
    indentValue = value * oneCharSizePT;
  } else if (isPX) {
    indentValue = value / 20 * oneCharSizePT;
  } else if (isPT) {
    indentValue = value / D_FontSizePT * oneCharSizePT;
  }
  // for now only support firstLine for the reason that it is the only one in web
  indent.firstLine = indentValue;
  styleOption.indent = indent;
  return styleOption;
};

var isColor = function isColor(_ref) {
  var key = _ref.key;
  return key === StyleMap.color;
};
var isBackgroundColor = function isBackgroundColor(_ref2) {
  var key = _ref2.key;
  return key === StyleMap.backgroundColor;
};
var isTextDecoration = function isTextDecoration(_ref3) {
  var key = _ref3.key;
  return key === StyleMap.textDecoration;
};
var isPadding = function isPadding(_ref4) {
  var key = _ref4.key;
  return key.indexOf('padding-') > -1;
};
var isTextAlign = function isTextAlign(_ref5) {
  var key = _ref5.key;
  return key.indexOf(StyleMap.textAlign) > -1;
};
var isLineHeight = function isLineHeight(_ref6) {
  var key = _ref6.key;
  return key === StyleMap.lineHeight;
};
var isFontFamily = function isFontFamily(_ref7) {
  var key = _ref7.key;
  return key === StyleMap.fontFamily;
};
var isVerticalAlign = function isVerticalAlign(_ref8) {
  var key = _ref8.key;
  return key === StyleMap.verticalAlign;
};
var isBorderColor = function isBorderColor(_ref9) {
  var key = _ref9.key;
  return key === StyleMap.borderColor;
};
var isWidth = function isWidth(_ref10) {
  var key = _ref10.key;
  return key === StyleMap.width;
};
var isHeight = function isHeight(_ref11) {
  var key = _ref11.key;
  return key === StyleMap.height;
};
var isTextIndent = function isTextIndent(_ref12) {
  var key = _ref12.key;
  return key === StyleMap.textIndent;
};
var isBold = function isBold(_ref14) {
  var key = _ref14.key,
    val = _ref14.val;
  return key === StyleMap.fontWeight && val.toLowerCase() === 'bold';
};
var isFontStyle = function isFontStyle(_ref15) {
  var key = _ref15.key;
  return key === StyleMap.fontStyle;
};
var isSubScript = function isSubScript(_ref17) {
  var key = _ref17.key,
    val = _ref17.val;
  return key === StyleMap.subScript && val.toLowerCase() === 'true';
};
var isSuperScript = function isSuperScript(_ref18) {
  var key = _ref18.key,
    val = _ref18.val;
  return key === StyleMap.superScript && val.toLowerCase() === 'true';
};

var tokens = [{
  name: 'color',
  judge: isColor,
  handler: colorHandler
}, {
  name: 'backgroundColor',
  judge: isBackgroundColor,
  handler: backgroundHandler
}, {
  name: 'bold',
  judge: isBold,
  handler: boldHandler
}, {
  name: 'align',
  judge: isTextAlign,
  handler: alignHandler
}, {
  name: 'borderColor',
  judge: isBorderColor,
  handler: borderColorHandler
}, {
  name: 'fontFamily',
  judge: isFontFamily,
  handler: fontFamilyHandler
}, {
  name: 'fontStyle',
  judge: isFontStyle,
  handler: fontStyleHandler
}, {
  name: 'height',
  judge: isHeight,
  handler: heightHandler
}, {
  name: 'lineHeight',
  judge: isLineHeight,
  handler: lineHeightHandler
}, {
  name: 'padding',
  judge: isPadding,
  handler: paddingHandler
}, {
  name: 'textDecoration',
  judge: isTextDecoration,
  handler: textDecorationHandler
}, {
  name: 'textIndent',
  judge: isTextIndent,
  handler: textIndentHandler
}, {
  name: 'verticalAlign',
  judge: isVerticalAlign,
  handler: verticalAlignHandler
}, {
  name: 'width',
  judge: isWidth,
  handler: widthHandler
}, {
  name: 'subScript',
  judge: isSubScript,
  handler: subScriptHandler
}, {
  name: 'superScript',
  judge: isSuperScript,
  handler: superScriptHandler
}];
var provideStyle = function provideStyle(styles) {
  var styleOption = {};
  styles.forEach(function (style) {
    var token = tokens.find(function (token) {
      return token.judge(style);
    });
    if (token) {
      styleOption = token.handler(style, styleOption);
    }
  });
  return styleOption;
};

// text node
// text node with content
var isFillTextNode = function isFillTextNode(node) {
  return node && node.type === 'text' && node.content;
};

// convert styles to flat array
var toFlatStyleList = function toFlatStyleList(styleStringList) {
  var inlined = styleStringList.filter(Boolean).map(function (str) {
    return str.split("" + Splitter_Semicolon);
  }).flat().filter(function (str) {
    return str.indexOf("" + Splitter_Colon) > -1;
  }).map(function (attr) {
    var _attr$trim$split = attr.trim().split(Splitter_Colon),
      key = _attr$trim$split[0],
      val = _attr$trim$split[1];
    var v = typeOf(val) === 'string' ? val.trim().replace(/;/i, '') : val;
    var value = isValidColor(v) ? toHex(v) : v;
    return {
      key: key.trim(),
      val: value
    };
  });
  return getUniqueArrayByKey(inlined, 'key');
};
// text creator
var calcTextRunStyle = function calcTextRunStyle(styleList, tagStyleMap) {
  if (tagStyleMap === void 0) {
    tagStyleMap = D_TagStyleMap;
  }
  var styleOption = {};
  if (!styleList || styleList.length === 0) return styleOption;
  var tagList = Object.keys(tagStyleMap);
  // handle tag style like: em del strong...
  var tagStyleList = styleList.filter(function (str) {
    return tagList.includes(str);
  });
  var inlined = tagStyleList.map(function (str) {
    return tagStyleMap[str];
  }).filter(Boolean);
  // flat inline styles
  var styles = toFlatStyleList([].concat(styleList, inlined));
  var fontSizeSty = styles.find(function (sty) {
    return sty.key === StyleMap.fontSize;
  });
  var fontSize = fontSizeSty && fontSizeSty.val ? handleSizeNumber(fontSizeSty.val) : null;
  /**
   * size(halfPts): Set the font size, measured in half-points
   */
  if (fontSize) {
    var value = fontSize.value,
      type = fontSize.type;
    var size = type === 'pt' ? value * 2 : value * PXbyPT * 2;
    styleOption.size = size;
  } else {
    styleOption.size = D_FontSizePT * 2;
  }
  var inlinedStyleOption = provideStyle(styles);
  return _extends({}, styleOption, inlinedStyleOption);
};
var textCreator = function textCreator(node, tagStyleMap) {
  if (tagStyleMap === void 0) {
    tagStyleMap = D_TagStyleMap;
  }
  var shape = node.shape,
    content = node.content;
  var textBuildParam = {
    text: optimizeBlankSpace(content)
  };
  var styleOption = shape && shape.length ? calcTextRunStyle(shape, tagStyleMap) : {};
  return new docx.TextRun(_extends({}, textBuildParam, styleOption));
};
// map children as ParagraphChild
var getChildrenByTextRun = /*#__PURE__*/function () {
  var _ref = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2(nodeList, tagStyleMap) {
    var texts, concatText;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          if (tagStyleMap === void 0) {
            tagStyleMap = D_TagStyleMap;
          }
          texts = [];
          concatText = /*#__PURE__*/function () {
            var _ref2 = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(list, arr) {
              var _iterator, _step, node, attrs, shape, src, _attrs$width, width, _attrs$height, height, styleOp, imgBlob, image, _attrs, text;
              return _regeneratorRuntime().wrap(function _callee$(_context) {
                while (1) switch (_context.prev = _context.next) {
                  case 0:
                    _iterator = _createForOfIteratorHelperLoose(list);
                  case 1:
                    if ((_step = _iterator()).done) {
                      _context.next = 42;
                      break;
                    }
                    node = _step.value;
                    if (!isFillTextNode(node)) {
                      _context.next = 7;
                      break;
                    }
                    arr.push(textCreator(node, tagStyleMap));
                    _context.next = 40;
                    break;
                  case 7:
                    if (!(node.name === TagType.img)) {
                      _context.next = 25;
                      break;
                    }
                    attrs = node.attrs, shape = node.shape;
                    src = attrs.src, _attrs$width = attrs.width, width = _attrs$width === void 0 ? 100 : _attrs$width, _attrs$height = attrs.height, height = _attrs$height === void 0 ? 100 : _attrs$height;
                    styleOp = calcTextRunStyle(shape);
                    if (!src) {
                      _context.next = 23;
                      break;
                    }
                    _context.prev = 12;
                    _context.next = 15;
                    return getImageBlob(String(src));
                  case 15:
                    imgBlob = _context.sent;
                    image = new docx.ImageRun({
                      data: imgBlob,
                      transformation: {
                        width: styleOp.tWidth || Number(width),
                        height: styleOp.tHeight || Number(height)
                      }
                    });
                    arr.push(image);
                    _context.next = 23;
                    break;
                  case 20:
                    _context.prev = 20;
                    _context.t0 = _context["catch"](12);
                    console.log('download image error', _context.t0);
                  case 23:
                    _context.next = 40;
                    break;
                  case 25:
                    if (!isFilledArray(node.children)) {
                      _context.next = 40;
                      break;
                    }
                    if (!(node.name === TagType.link)) {
                      _context.next = 38;
                      break;
                    }
                    _attrs = node.attrs;
                    _context.t1 = docx.ExternalHyperlink;
                    _context.next = 31;
                    return getChildrenByTextRun(node.children, tagStyleMap);
                  case 31:
                    _context.t2 = _context.sent;
                    _context.t3 = _attrs.href ? String(_attrs.href) : '';
                    _context.t4 = {
                      children: _context.t2,
                      link: _context.t3
                    };
                    text = new _context.t1(_context.t4);
                    arr.push(text);
                    _context.next = 40;
                    break;
                  case 38:
                    _context.next = 40;
                    return concatText(node.children, arr);
                  case 40:
                    _context.next = 1;
                    break;
                  case 42:
                  case "end":
                    return _context.stop();
                }
              }, _callee, null, [[12, 20]]);
            }));
            return function concatText(_x3, _x4) {
              return _ref2.apply(this, arguments);
            };
          }();
          _context2.next = 5;
          return concatText(nodeList, texts);
        case 5:
          return _context2.abrupt("return", texts);
        case 6:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return function getChildrenByTextRun(_x, _x2) {
    return _ref.apply(this, arguments);
  };
}();

var calcTableWidth = function calcTableWidth(colsArr) {
  return colsArr.reduce(function (prev, cur) {
    return prev + cur;
  }, 0);
};
var getTableBorderStyleSingle = function getTableBorderStyleSingle(size, color) {
  return {
    style: docx.BorderStyle.SINGLE,
    size: size * 10,
    color: color
  };
};
var tablePxByXDA = D_TableFullWidth / 553;
var getColGroupWidth = function getColGroupWidth(cols) {
  var count = cols.length;
  var defaultWidth = count ? D_TableFullWidth / tablePxByXDA / count : 0;
  return cols.filter(function (c) {
    return c.name === 'col';
  }).map(function (col) {
    var _handleSizeNumber;
    var attrs = col.attrs;
    return tablePxByXDA * (((_handleSizeNumber = handleSizeNumber(String(attrs.width))) == null ? void 0 : _handleSizeNumber.value) || defaultWidth);
  });
};
var handleCellWidthFromColgroup = function handleCellWidthFromColgroup(cols, index, colspan) {
  return cols.slice(index, index + colspan).reduce(function (prev, cur) {
    return prev + cur;
  }, 0);
};
var getCellWidthInDXA = function getCellWidthInDXA(size) {
  return size * tablePxByXDA;
};
// table node to docx ITableOptions
var tableNodeToITableOptions = /*#__PURE__*/function () {
  var _ref = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(tableNode, tagStyleMap) {
    var tc, attrs, shape, isTBody, tbody, colGroup, cols, tableParam, styleOp, border, borderSize, borderColor, borders, isTr, isTd, firstRowColumnSize, hasColGroup, trs, rows, _iterator, _step, _calcTextRunStyle, _step$value, tr, idx, children, _attrs, trHeight, tds, cellChildren, _iterator2, _step2, tdObj, td, index, _attrs2, _shape, styles, tdStyleOption, texts, _iterator3, _step3, t, _shape2, content, _children, c, cellParam, colspan, rowspan, width, cellWidth, i, margins, tableCellOptions, para, h, tableWidths;
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          if (tagStyleMap === void 0) {
            tagStyleMap = D_TagStyleMap;
          }
          tc = tableNode.children, attrs = tableNode.attrs, shape = tableNode.shape;
          isTBody = function isTBody(n) {
            return n.name === 'tbody';
          };
          tbody = tc.find(isTBody);
          if (tbody) {
            _context.next = 6;
            break;
          }
          return _context.abrupt("return", null);
        case 6:
          // deal colgroup for cell width
          colGroup = tc.find(function (n) {
            return n.name === 'colgroup';
          });
          cols = colGroup ? getColGroupWidth(colGroup.children) : []; // const colsTotalWidth = cols.reduce((prev, cur) => prev + cur, 0);
          // Google DOCS does not support start and end borders, instead they use left and right borders.
          // So to set left and right borders for Google DOCS you should use
          // see https://docx.js.org/#/usage/tables
          tableParam = {
            layout: docx.TableLayoutType.FIXED,
            borders: {
              top: DefaultBorder,
              left: DefaultBorder,
              right: DefaultBorder,
              bottom: DefaultBorder
            },
            rows: []
          };
          styleOp = calcTextRunStyle(shape, tagStyleMap);
          border = attrs.border;
          borderSize = border ? parseFloat(border) : D_TableBorderSize;
          borderColor = styleOp.borderColor || D_TableBorderColor;
          borders = {
            top: getTableBorderStyleSingle(borderSize, borderColor),
            right: getTableBorderStyleSingle(borderSize, borderColor),
            bottom: getTableBorderStyleSingle(borderSize, borderColor),
            left: getTableBorderStyleSingle(borderSize, borderColor)
          };
          tableParam.borders = borders;
          isTr = function isTr(n) {
            return n.name === 'tr';
          };
          isTd = function isTd(n) {
            return n.name === 'td';
          };
          firstRowColumnSize = [];
          hasColGroup = false;
          trs = tbody.children.filter(isTr);
          rows = [];
          _iterator = _createForOfIteratorHelperLoose(trs.map(function (tr, idx) {
            return {
              tr: tr,
              idx: idx
            };
          }));
        case 22:
          if ((_step = _iterator()).done) {
            _context.next = 69;
            break;
          }
          _step$value = _step.value, tr = _step$value.tr, idx = _step$value.idx;
          children = tr.children, _attrs = tr.attrs;
          trHeight = _attrs != null && _attrs.style ? ((_calcTextRunStyle = calcTextRunStyle([_attrs == null ? void 0 : _attrs.style], tagStyleMap)) == null ? void 0 : _calcTextRunStyle.tHeight) || D_TableCellHeightPx : D_TableCellHeightPx;
          tds = children.filter(isTd);
          cellChildren = [];
          _iterator2 = _createForOfIteratorHelperLoose(tds.map(function (item, index) {
            return {
              item: item,
              index: index
            };
          }));
        case 29:
          if ((_step2 = _iterator2()).done) {
            _context.next = 63;
            break;
          }
          tdObj = _step2.value;
          td = tdObj.item, index = tdObj.index;
          _attrs2 = td.attrs, _shape = td.shape; // table paragraph use line-height 1.0 for default
          styles = _extends({}, tagStyleMap);
          delete styles.p;
          tdStyleOption = calcTextRunStyle(_shape, styles); // TODO: support Nested Tables and other elements
          // use `contentBuilder` maybe better
          texts = [];
          _iterator3 = _createForOfIteratorHelperLoose(td.children);
        case 38:
          if ((_step3 = _iterator3()).done) {
            _context.next = 49;
            break;
          }
          t = _step3.value;
          _shape2 = t.shape, content = t.content, _children = t.children;
          if (!(_children != null && _children.length)) {
            _context.next = 46;
            break;
          }
          _context.next = 44;
          return getChildrenByTextRun(_children || [], styles);
        case 44:
          c = _context.sent;
          texts.push(new docx.Paragraph(_extends({
            children: c
          }, calcTextRunStyle(_shape2, styles))));
        case 46:
          texts.push(new docx.Paragraph(_extends({
            text: content
          }, calcTextRunStyle(_shape2, styles))));
        case 47:
          _context.next = 38;
          break;
        case 49:
          cellParam = {
            children: texts
          };
          colspan = _attrs2.colspan, rowspan = _attrs2.rowspan;
          if (colspan && Number(colspan) !== 0) {
            cellParam.columnSpan = Number(colspan);
          }
          if (rowspan && Number(rowspan) !== 0) {
            cellParam.rowSpan = Number(rowspan);
          }
          hasColGroup = !!cols.length && cols.every(function (c) {
            return c !== 0;
          });
          if (hasColGroup) {
            width = handleCellWidthFromColgroup(cols, index, cellParam.columnSpan || 1);
            tdStyleOption.tWidth = width;
          }
          cellWidth = hasColGroup ? tdStyleOption.tWidth || D_TableFullWidth / cols.length : getCellWidthInDXA(tdStyleOption.tWidth || 185);
          cellParam.width = {
            size: cellWidth,
            type: docx.WidthType.DXA
          };
          if (idx === 0) {
            if (cellParam.columnSpan) {
              for (i = 0; i < cellParam.columnSpan; i++) {
                firstRowColumnSize.push(cellWidth / cellParam.columnSpan);
              }
            } else {
              firstRowColumnSize.push(cellWidth);
            }
          }
          margins = {
            marginUnitType: docx.WidthType.DXA,
            top: D_CELL_MARGIN,
            bottom: D_CELL_MARGIN,
            left: D_CELL_MARGIN,
            right: D_CELL_MARGIN
          };
          tableCellOptions = _extends({}, cellParam, calcTextRunStyle(_shape, styles), {
            margins: margins
          });
          cellChildren.push(new docx.TableCell(tableCellOptions));
        case 61:
          _context.next = 29;
          break;
        case 63:
          para = {
            children: cellChildren,
            height: {
              value: 0,
              rule: docx.HeightRule.EXACT
            }
          };
          h = (trHeight != null ? trHeight : D_TableCellHeightPx) * tablePxByXDA;
          para.height = {
            value: h,
            rule: docx.HeightRule.EXACT
          };
          rows.push(new docx.TableRow(para));
        case 67:
          _context.next = 22;
          break;
        case 69:
          tableWidths = hasColGroup ? cols : firstRowColumnSize;
          tableParam.columnWidths = tableWidths;
          tableParam.width = {
            size: calcTableWidth(tableWidths),
            type: docx.WidthType.DXA
          };
          tableParam.rows = rows;
          return _context.abrupt("return", tableParam);
        case 74:
        case "end":
          return _context.stop();
      }
    }, _callee);
  }));
  return function tableNodeToITableOptions(_x, _x2) {
    return _ref.apply(this, arguments);
  };
}();
// create docx table from table node
var tableCreator = /*#__PURE__*/function () {
  var _ref2 = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2(tableNode, tagStyleMap) {
    var tableParam;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          if (tagStyleMap === void 0) {
            tagStyleMap = D_TagStyleMap;
          }
          _context2.next = 3;
          return tableNodeToITableOptions(tableNode, tagStyleMap);
        case 3:
          tableParam = _context2.sent;
          if (tableParam) {
            _context2.next = 6;
            break;
          }
          return _context2.abrupt("return", null);
        case 6:
          return _context2.abrupt("return", new docx.Table(tableParam));
        case 7:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return function tableCreator(_x3, _x4) {
    return _ref2.apply(this, arguments);
  };
}();

var contentBuilder = /*#__PURE__*/function () {
  var _ref = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(node, tagStyleMap) {
    var type, name, children, content, shape, para, isText, isLink, isTable, isList, isNormalParagraphWithChildren, paragraphOption, _paragraphOption, table;
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          if (tagStyleMap === void 0) {
            tagStyleMap = D_TagStyleMap;
          }
          type = node.type, name = node.name, children = node.children, content = node.content, shape = node.shape;
          para = {
            text: content,
            children: []
          };
          isText = type === TagType.text && content;
          isLink = name === TagType.link;
          isTable = name === TagType.table;
          isList = name === TagType.ordered_list || name === TagType.unordered_list;
          isNormalParagraphWithChildren = !isLink && !isTable && !isList && children && isFilledArray(children) && children.length > 0;
          if (!isText) {
            _context.next = 13;
            break;
          }
          paragraphOption = _extends({}, para, calcTextRunStyle(shape, tagStyleMap));
          return _context.abrupt("return", new docx.Paragraph(paragraphOption));
        case 13:
          if (!isNormalParagraphWithChildren) {
            _context.next = 21;
            break;
          }
          _context.next = 16;
          return getChildrenByTextRun(children, tagStyleMap);
        case 16:
          para.children = _context.sent;
          _paragraphOption = _extends({}, para, calcTextRunStyle(shape, tagStyleMap));
          return _context.abrupt("return", new docx.Paragraph(_paragraphOption));
        case 21:
          if (!isTable) {
            _context.next = 28;
            break;
          }
          _context.next = 24;
          return tableCreator(node, tagStyleMap);
        case 24:
          table = _context.sent;
          return _context.abrupt("return", table);
        case 28:
          if (!isList) {
            _context.next = 32;
            break;
          }
          return _context.abrupt("return", null);
        case 32:
          return _context.abrupt("return", null);
        case 33:
        case "end":
          return _context.stop();
      }
    }, _callee);
  }));
  return function contentBuilder(_x, _x2) {
    return _ref.apply(this, arguments);
  };
}();

var getInnerTextNode = function getInnerTextNode(node) {
  var inner = node;
  while (inner && inner.children && inner.children.length === 1) {
    inner = inner.children[0];
  }
  return inner;
};
// recursion chain style
var chainStyle = function chainStyle(nodeList, style, tagStyleMap) {
  if (style === void 0) {
    style = [];
  }
  if (!nodeList || !isFilledArray(nodeList)) return;
  nodeList.forEach(function (node) {
    var attrs = node.attrs,
      children = node.children,
      name = node.name;
    var STYLE = typeof (attrs == null ? void 0 : attrs.style) === 'string' ? [attrs.style].concat(style) : style;
    var shape = name ? [name].concat(STYLE) : [].concat(STYLE);
    node.shape = shape;
    if (isFilledArray(children)) {
      chainStyle(children, shape);
    }
  });
};
// style builder
var StyleBuilder = function StyleBuilder(list, tagStyleMap) {
  var nList = [].concat(list);
  chainStyle(nList, []);
  return nList;
};
// element creator
var ElementCreator = /*#__PURE__*/function () {
  var _ref = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee(astList, tagStyleMap) {
    var tags, ps, _iterator, _step, node, p;
    return _regeneratorRuntime().wrap(function _callee$(_context) {
      while (1) switch (_context.prev = _context.next) {
        case 0:
          if (tagStyleMap === void 0) {
            tagStyleMap = D_TagStyleMap;
          }
          if (!(!astList || astList.length === 0)) {
            _context.next = 3;
            break;
          }
          return _context.abrupt("return", []);
        case 3:
          tags = StyleBuilder(astList.filter(function (n) {
            return n.type === 'tag';
          }));
          if (tags) {
            _context.next = 6;
            break;
          }
          return _context.abrupt("return", []);
        case 6:
          ps = [];
          _iterator = _createForOfIteratorHelperLoose(tags);
        case 8:
          if ((_step = _iterator()).done) {
            _context.next = 16;
            break;
          }
          node = _step.value;
          _context.next = 12;
          return contentBuilder(node, tagStyleMap);
        case 12:
          p = _context.sent;
          if (p) {
            ps.push(p);
          }
        case 14:
          _context.next = 8;
          break;
        case 16:
          return _context.abrupt("return", [].concat(ps));
        case 17:
        case "end":
          return _context.stop();
      }
    }, _callee);
  }));
  return function ElementCreator(_x, _x2) {
    return _ref.apply(this, arguments);
  };
}();
// parse html string into Node list
var htmlToAST = function htmlToAST(html) {
  return htmlToAst.parse(html);
};
// generate Document
var genDocument = /*#__PURE__*/function () {
  var _ref2 = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2(html, options) {
    var layoutOp, layout, styleMap, ast, paragraphs, orientation, topMargin, leftMargin, rightMargin, bottomMargin, header, footer, margin, page, section, _ast, _ast2, doc;
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          layoutOp = (options == null ? void 0 : options.layout) || {};
          layout = _extends({}, D_Layout, layoutOp);
          styleMap = (options == null ? void 0 : options.tagStyleMap) || D_TagStyleMap;
          ast = htmlToAST(html);
          _context2.next = 6;
          return ElementCreator(ast, styleMap);
        case 6:
          paragraphs = _context2.sent;
          orientation = layout.orientation, topMargin = layout.topMargin, leftMargin = layout.leftMargin, rightMargin = layout.rightMargin, bottomMargin = layout.bottomMargin, header = layout.header, footer = layout.footer;
          margin = {
            top: calcMargin(topMargin),
            left: calcMargin(leftMargin),
            right: calcMargin(rightMargin),
            bottom: calcMargin(bottomMargin)
          };
          page = {
            margin: margin,
            size: {
              orientation: orientation
            }
          };
          section = {
            properties: {
              page: page
            },
            children: paragraphs,
            headers: {},
            footers: {}
          };
          if (!header) {
            _context2.next = 20;
            break;
          }
          _ast = htmlToAst.parse(header);
          _context2.t0 = docx.Header;
          _context2.next = 16;
          return ElementCreator(_ast, styleMap);
        case 16:
          _context2.t1 = _context2.sent;
          _context2.t2 = {
            children: _context2.t1
          };
          _context2.t3 = new _context2.t0(_context2.t2);
          section.headers = {
            "default": _context2.t3
          };
        case 20:
          if (!footer) {
            _context2.next = 29;
            break;
          }
          _ast2 = htmlToAst.parse(footer);
          _context2.t4 = docx.Footer;
          _context2.next = 25;
          return ElementCreator(_ast2, styleMap);
        case 25:
          _context2.t5 = _context2.sent;
          _context2.t6 = {
            children: _context2.t5
          };
          _context2.t7 = new _context2.t4(_context2.t6);
          section.footers = {
            "default": _context2.t7
          };
        case 29:
          doc = new docx.Document({
            styles: {
              paragraphStyles: []
            },
            sections: [section]
          });
          return _context2.abrupt("return", doc);
        case 31:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return function genDocument(_x3, _x4) {
    return _ref2.apply(this, arguments);
  };
}();
// export html as docx file
var exportAsDocx = /*#__PURE__*/function () {
  var _ref3 = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee3(doc, docName) {
    return _regeneratorRuntime().wrap(function _callee3$(_context3) {
      while (1) switch (_context3.prev = _context3.next) {
        case 0:
          if (docName === void 0) {
            docName = '';
          }
          docx.Packer.toBlob(doc).then(function (blob) {
            fileSaver.saveAs(blob, docName + ".docx");
          });
        case 2:
        case "end":
          return _context3.stop();
      }
    }, _callee3);
  }));
  return function exportAsDocx(_x5, _x6) {
    return _ref3.apply(this, arguments);
  };
}();
// html -> docx
var exportHtmlToDocx = /*#__PURE__*/function () {
  var _ref4 = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee4(html, docName, options) {
    var doc;
    return _regeneratorRuntime().wrap(function _callee4$(_context4) {
      while (1) switch (_context4.prev = _context4.next) {
        case 0:
          if (docName === void 0) {
            docName = 'doc';
          }
          _context4.next = 3;
          return genDocument(trimHtml(html), options);
        case 3:
          doc = _context4.sent;
          exportAsDocx(doc, docName);
          return _context4.abrupt("return", doc);
        case 6:
        case "end":
          return _context4.stop();
      }
    }, _callee4);
  }));
  return function exportHtmlToDocx(_x7, _x8, _x9) {
    return _ref4.apply(this, arguments);
  };
}();
// export multi files as .zip
var exportMultiDocsAsZip = /*#__PURE__*/function () {
  var _ref5 = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee5(docList, fileName, export_option) {
    var zip, len, d, html, name, option, file, _iterator2, _step2, docFile, _html, _name, _option, doc, _file;
    return _regeneratorRuntime().wrap(function _callee5$(_context5) {
      while (1) switch (_context5.prev = _context5.next) {
        case 0:
          if (fileName === void 0) {
            fileName = 'docs';
          }
          zip = new JSZip();
          len = docList.length;
          if (!(len === 1)) {
            _context5.next = 11;
            break;
          }
          d = docList[0];
          html = d.html, name = d.name, option = d.option;
          _context5.next = 8;
          return genDocument(trimHtml(html), option || export_option);
        case 8:
          file = _context5.sent;
          exportAsDocx(file, name);
          return _context5.abrupt("return");
        case 11:
          _iterator2 = _createForOfIteratorHelperLoose(docList);
        case 12:
          if ((_step2 = _iterator2()).done) {
            _context5.next = 24;
            break;
          }
          docFile = _step2.value;
          _html = docFile.html, _name = docFile.name, _option = docFile.option;
          _context5.next = 17;
          return genDocument(trimHtml(_html), _option || export_option);
        case 17:
          doc = _context5.sent;
          _context5.next = 20;
          return docx.Packer.toBlob(doc);
        case 20:
          _file = _context5.sent;
          zip.file(_name + ".docx", _file);
        case 22:
          _context5.next = 12;
          break;
        case 24:
          zip.generateAsync({
            type: 'blob'
          }).then(function (content) {
            fileSaver.saveAs(content, fileName + ".zip");
          });
        case 25:
        case "end":
          return _context5.stop();
      }
    }, _callee5);
  }));
  return function exportMultiDocsAsZip(_x10, _x11, _x12) {
    return _ref5.apply(this, arguments);
  };
}();
var exportAsZip = exportMultiDocsAsZip;

Object.defineProperty(exports, 'parse', {
  enumerable: true,
  get: function () {
    return htmlToAst.parse;
  }
});
exports.D_Layout = D_Layout;
exports.D_TagStyleMap = D_TagStyleMap;
exports.ElementCreator = ElementCreator;
exports.StyleBuilder = StyleBuilder;
exports.chainStyle = chainStyle;
exports.exportAsDocx = exportAsDocx;
exports.exportAsZip = exportAsZip;
exports.exportHtmlToDocx = exportHtmlToDocx;
exports.exportMultiDocsAsZip = exportMultiDocsAsZip;
exports.genDocument = genDocument;
exports.getInnerTextNode = getInnerTextNode;
exports.htmlToAST = htmlToAST;
exports.tableNodeToITableOptions = tableNodeToITableOptions;
//# sourceMappingURL=editor-to-word.cjs.development.js.map
