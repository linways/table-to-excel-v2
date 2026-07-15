var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/tableToExcel.js
var tableToExcel_exports = {};
__export(tableToExcel_exports, {
  default: () => tableToExcel_default
});
module.exports = __toCommonJS(tableToExcel_exports);

// src/parser.js
var TTEParser = (function() {
  let methods = {};
  methods.parseDomToTable = function(ws, table, opts) {
    let _r, _c, cs, rs, r, c;
    let rows = [...table.getElementsByTagName("tr")];
    let widths = table.getAttribute("data-cols-width");
    if (widths)
      widths = widths.split(",").map(function(item) {
        return parseInt(item);
      });
    let merges = [];
    for (_r = 0; _r < rows.length; ++_r) {
      let row = rows[_r];
      r = _r + 1;
      c = 1;
      if (row.getAttribute("data-exclude") === "true") {
        rows.splice(_r, 1);
        _r--;
        continue;
      }
      if (row.getAttribute("data-height")) {
        let exRow = ws.getRow(r);
        exRow.height = parseFloat(row.getAttribute("data-height"));
      }
      let tds = [...row.children];
      for (_c = 0; _c < tds.length; ++_c) {
        let td = tds[_c];
        if (td.getAttribute("data-exclude") === "true") {
          tds.splice(_c, 1);
          _c--;
          continue;
        }
        for (let _m = 0; _m < merges.length; ++_m) {
          var m = merges[_m];
          if (m.s.c == c && m.s.r <= r && r <= m.e.r) {
            c = m.e.c + 1;
            _m = -1;
          }
        }
        let exCell = ws.getCell(getColumnAddress(c, r));
        cs = parseInt(td.getAttribute("colspan")) || 1;
        rs = parseInt(td.getAttribute("rowspan")) || 1;
        if (cs > 1 || rs > 1) {
          merges.push({
            s: { c, r },
            e: { c: c + cs - 1, r: r + rs - 1 }
          });
        }
        c += cs;
        exCell.value = getValue(td);
        if (!opts.autoStyle) {
          let styles = getStylesDataAttr(td);
          exCell.font = styles.font || null;
          exCell.alignment = styles.alignment || null;
          exCell.border = styles.border || null;
          exCell.fill = styles.fill || null;
          exCell.numFmt = styles.numFmt || null;
        }
      }
    }
    if (widths)
      widths.forEach((width, _i) => {
        ws.columns[_i].width = width;
      });
    applyMerges(ws, merges);
    return ws;
  };
  let applyMerges = function(ws, merges) {
    merges.forEach((m) => {
      ws.mergeCells(
        getExcelColumnName(m.s.c) + m.s.r + ":" + getExcelColumnName(m.e.c) + m.e.r
      );
    });
  };
  let htmldecode = (function() {
    let entities = [
      ["nbsp", " "],
      ["middot", "\xB7"],
      ["quot", '"'],
      ["apos", "'"],
      ["gt", ">"],
      ["lt", "<"],
      ["amp", "&"]
    ].map(function(x) {
      return [new RegExp("&" + x[0] + ";", "g"), x[1]];
    });
    return function htmldecode2(str) {
      let o = str.trim().replace(/\s+/g, " ").replace(/<\s*[bB][rR]\s*\/?>/g, "\n").replace(/<[^>]*>/g, "");
      for (let i = 0; i < entities.length; ++i)
        o = o.replace(entities[i][0], entities[i][1]);
      return o;
    };
  })();
  let getExcelColumnName = function(num) {
    for (var ret = "", a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
      ret = String.fromCharCode(parseInt(num % b / a) + 65) + ret;
    }
    return ret;
  };
  let getColumnAddress = function(col, row) {
    return getExcelColumnName(col) + row;
  };
  let getValue = function(td) {
    let dataType = td.getAttribute("data-t");
    let rawVal = htmldecode(td.innerHTML);
    if (dataType) {
      let val;
      switch (dataType) {
        case "n":
          val = Number(rawVal);
          break;
        case "d":
          let date = new Date(rawVal);
          val = new Date(
            Date.UTC(
              date.getFullYear(),
              date.getMonth(),
              date.getDate(),
              date.getHours(),
              date.getMinutes(),
              date.getSeconds()
            )
          );
          break;
        case "b":
          val = rawVal.toLowerCase() === "true" ? true : rawVal.toLowerCase() === "false" ? false : Boolean(parseInt(rawVal));
          break;
        default:
          val = rawVal;
      }
      return val;
    } else if (td.getAttribute("data-hyperlink")) {
      return {
        text: rawVal,
        hyperlink: td.getAttribute("data-hyperlink")
      };
    } else if (td.getAttribute("data-error")) {
      return { error: td.getAttribute("data-error") };
    }
    return rawVal;
  };
  let getStylesDataAttr = function(td) {
    let font = {};
    if (td.getAttribute("data-f-name"))
      font.name = td.getAttribute("data-f-name");
    if (td.getAttribute("data-f-sz")) font.size = td.getAttribute("data-f-sz");
    if (td.getAttribute("data-f-color"))
      font.color = { argb: td.getAttribute("data-f-color") };
    if (td.getAttribute("data-f-bold") === "true") font.bold = true;
    if (td.getAttribute("data-f-italic") === "true") font.italic = true;
    if (td.getAttribute("data-f-underline") === "true") font.underline = true;
    if (td.getAttribute("data-f-strike") === "true") font.strike = true;
    let alignment = {};
    if (td.getAttribute("data-a-h"))
      alignment.horizontal = td.getAttribute("data-a-h");
    if (td.getAttribute("data-a-v"))
      alignment.vertical = td.getAttribute("data-a-v");
    if (td.getAttribute("data-a-wrap") === "true") alignment.wrapText = true;
    if (td.getAttribute("data-a-text-rotation"))
      alignment.textRotation = td.getAttribute("data-a-text-rotation");
    if (td.getAttribute("data-a-indent"))
      alignment.indent = td.getAttribute("data-a-indent");
    if (td.getAttribute("data-a-rtl") === "true")
      alignment.readingOrder = "rtl";
    let border = {
      top: {},
      left: {},
      bottom: {},
      right: {}
    };
    if (td.getAttribute("data-b-a-s")) {
      let style = td.getAttribute("data-b-a-s");
      border.top.style = style;
      border.left.style = style;
      border.bottom.style = style;
      border.right.style = style;
    }
    if (td.getAttribute("data-b-a-c")) {
      let color = { argb: td.getAttribute("data-b-a-c") };
      border.top.color = color;
      border.left.color = color;
      border.bottom.color = color;
      border.right.color = color;
    }
    if (td.getAttribute("data-b-t-s")) {
      border.top.style = td.getAttribute("data-b-t-s");
      if (td.getAttribute("data-b-t-c"))
        border.top.color = { argb: td.getAttribute("data-b-t-c") };
    }
    if (td.getAttribute("data-b-l-s")) {
      border.left.style = td.getAttribute("data-b-l-s");
      if (td.getAttribute("data-b-l-c"))
        border.left.color = { argb: td.getAttribute("data-b-t-c") };
    }
    if (td.getAttribute("data-b-b-s")) {
      border.bottom.style = td.getAttribute("data-b-b-s");
      if (td.getAttribute("data-b-b-c"))
        border.bottom.color = { argb: td.getAttribute("data-b-t-c") };
    }
    if (td.getAttribute("data-b-r-s")) {
      border.right.style = td.getAttribute("data-b-r-s");
      if (td.getAttribute("data-b-r-c"))
        border.right.color = { argb: td.getAttribute("data-b-t-c") };
    }
    let fill;
    if (td.getAttribute("data-fill-color")) {
      fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: td.getAttribute("data-fill-color") }
      };
    }
    let numFmt;
    if (td.getAttribute("data-num-fmt"))
      numFmt = td.getAttribute("data-num-fmt");
    return {
      font,
      alignment,
      border,
      fill,
      numFmt
    };
  };
  return methods;
})();
var parser_default = TTEParser;

// src/tableToExcel.js
var import_file_saver = __toESM(require("file-saver"), 1);
var import_exceljs = __toESM(require("exceljs"), 1);
var TableToExcel = (function(Parser) {
  let methods = {};
  methods.initWorkBook = function() {
    let wb = new import_exceljs.default.Workbook();
    return wb;
  };
  methods.initSheet = function(wb, sheetName) {
    let ws = wb.addWorksheet(sheetName);
    return ws;
  };
  methods.save = function(wb, fileName) {
    wb.xlsx.writeBuffer().then(function(buffer) {
      const blob = new Blob([buffer], { type: "application/octet-stream" });
      const reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onloadend = function() {
        const base64data = reader.result;
        let data = {
          isFileOpenFromBase64Request: true,
          isMetaDataIncluded: true,
          fileName,
          base64Data: base64data
        };
        data = JSON.stringify(data);
        if (window.CustomWebviewMessageChannel) {
          window.CustomWebviewMessageChannel.postMessage(data);
        }
      };
      (0, import_file_saver.default)(
        blob,
        fileName
      );
    });
  };
  methods.tableToSheet = function(wb, table, opts) {
    let ws = this.initSheet(wb, opts.sheet.name);
    ws = Parser.parseDomToTable(ws, table, opts);
    return wb;
  };
  methods.tableToBook = function(table, opts) {
    let wb = this.initWorkBook();
    wb = this.tableToSheet(wb, table, opts);
    return wb;
  };
  methods.convert = function(table, opts = {}) {
    let defaultOpts = {
      name: "export.xlsx",
      autoStyle: false,
      sheet: {
        name: "Sheet 1"
      }
    };
    opts = { ...defaultOpts, ...opts };
    let wb = this.tableToBook(table, opts);
    this.save(wb, opts.name);
  };
  return methods;
})(parser_default);
var tableToExcel_default = TableToExcel;
window.TableToExcel = TableToExcel;
//# sourceMappingURL=tableToExcel.cjs.map