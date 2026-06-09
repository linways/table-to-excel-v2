import Parser from "./parser";
import saveAs from "file-saver";
import ExcelJS from "../node_modules/exceljs/dist/es5/exceljs.browser";

function getBase64FromSrc(src) {
  return new Promise(function(resolve, reject) {
    if (src && src.indexOf('data:') === 0) {
      let m = src.match(/^data:image\/(\w+);base64,(.+)$/);
      if (m) return resolve({ extension: m[1], base64: m[2] });
    }
    fetch(src)
      .then(function(r) { return r.blob(); })
      .then(function(blob) {
        let reader = new FileReader();
        reader.onload = function() {
          let m = reader.result.match(/^data:image\/(\w+);base64,(.+)$/);
          if (m) resolve({ extension: m[1], base64: m[2] });
          else reject(new Error('Cannot convert image to base64'));
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      })
      .catch(reject);
  });
}

const TableToExcel = (function(Parser) {
  let methods = {};

  methods.initWorkBook = function() {
    let wb = new ExcelJS.Workbook();
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
        // FOR FILE CREATION IN MOBILE APP.
        let data = {
          isFileOpenFromBase64Request: true,
          isMetaDataIncluded: true,
          fileName: fileName,
          base64Data: base64data
        }
        data = JSON.stringify(data);
        if (window.CustomWebviewMessageChannel) {
          window.CustomWebviewMessageChannel.postMessage(data);
        }
      };
      
      saveAs(
        blob,
        fileName
      );
    });
  };

  methods.tableToSheet = function(wb, table, opts) {
    let ws = this.initSheet(wb, opts.sheet.name);
    ws = Parser.parseDomToTable(ws, table, opts);
    if (!ws._images || !ws._images.length) {
      return Promise.resolve(wb);
    }
    let promises = ws._images.map(function(imgData) {
      return getBase64FromSrc(imgData.src)
        .then(function(res) {
          let imageId = wb.addImage({ base64: res.base64, extension: res.extension });
          ws.addImage(imageId, {
            tl: { col: imgData.col, row: imgData.row },
            ext: { width: imgData.width, height: imgData.height }
          });
        })
        .catch(function(err) {
          console.warn('TableToExcel: Could not embed image', err);
        });
    });
    return Promise.all(promises).then(function() { return wb; });
  };

  methods.tableToBook = function(table, opts) {
    let wb = this.initWorkBook();
    return this.tableToSheet(wb, table, opts);
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
    let self = this;
    this.tableToBook(table, opts).then(function(wb) {
      self.save(wb, opts.name);
    });
  };

  return methods;
})(Parser);

export default TableToExcel;
window.TableToExcel = TableToExcel;
