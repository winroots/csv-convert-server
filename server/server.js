const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx-js-style");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
app.use(cors());

const upload = multer({ dest: "uploads/" });

const PRICE_FILE = path.join(__dirname, "price-list.xlsx");

const NAME_COL = 1;
const BARCODE_COL = 4;
const NUMBER_COLS = [6, 7, 10, 13];

// ✅ 防止乱匹配
function isValidProductName(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  if (text.length < 4) return false;
  if (/^[0-9]+$/.test(text)) return false;
  if (/^[0-9]+,[0-9]+$/.test(text)) return false;
  if (/^[0-9]+\.[0-9]+$/.test(text)) return false;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return false;
  if (!/[a-zA-Z\u4e00-\u9fa5]/.test(text)) return false;
  return true;
}

function normalizeName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[’']/g, "")
    .replace(/[^a-z0-9\u4e00-\u9fa5]/g, "")
    .trim();
}

function convertNumber(value) {
  if (typeof value === "number") return value;
  if (!value) return value;

  const text = String(value).trim();

  if (/^-?[0-9]+,[0-9]+$/.test(text)) {
    return Number(text.replace(",", "."));
  }

  if (text.startsWith(",") && /^[0-9]+$/.test(text.slice(1))) {
    return Number("0." + text.slice(1));
  }

  if (/^[0-9]{1,3}(,[0-9]{3})+$/.test(text)) {
    return Number(text.replace(/,/g, "")) / 1000;
  }

  return value;
}

function loadPriceList() {
  const wb = XLSX.readFile(PRICE_FILE);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  return data
    .map(row => ({
      key: normalizeName(row["名称"]),
      barcode: String(row["条码"] || "").trim()
    }))
    .filter(i => i.key && i.barcode);
}

function findBarcode(name, priceList) {
  const key = normalizeName(name);
  if (!key) return null;

  const exact = priceList.find(i => i.key === key);
  if (exact) return exact.barcode;

  const contains = priceList.find(i =>
    key.includes(i.key) || i.key.includes(key)
  );
  if (contains) return contains.barcode;

  return null;
}

app.post("/api/convert", upload.single("csv"), (req, res) => {

  const apiKey = req.headers["x-api-key"] || req.body.key;

  if (apiKey !== "dmo2025") {
    return res.status(403).json({ message: "禁止访问" });
  }

  const csvPath = req.file?.path;

  try {
    if (!csvPath) {
      return res.status(400).json({ message: "没有收到 CSV 文件" });
    }

    const text = fs.readFileSync(csvPath, "utf-8");
    const rows = text.split("\n").map(r => r.replace(/\r$/, "").split(";"));

    const priceList = loadPriceList();

    let matched = 0;
    let changed = 0;
    const updatedRows = [];

    rows.forEach((row, i) => {
      // 数字处理
      NUMBER_COLS.forEach(idx => {
        if (row[idx]) row[idx] = convertNumber(row[idx]);
      });

      if (i === 0) return;

      const name = row[NAME_COL];
      if (!isValidProductName(name)) return;

      const oldBarcode = String(row[BARCODE_COL] || "").trim();
      const newBarcode = findBarcode(name, priceList);

      if (newBarcode) {
        matched++;

        if (newBarcode !== oldBarcode) {
          changed++;
          updatedRows.push(i);
        }

        row[BARCODE_COL] = String(newBarcode);
      }
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // ✅ 表头
    const headers = {
      B1: "名称",
      E1: "条码",
      G1: "价格",
      H1: "数量",
      K1: "建议卖价格",
      N1: "IVA"
    };

    Object.entries(headers).forEach(([cell, value]) => {
      ws[cell] = { v: value, t: "s" };
    });

    // ✅ 条码列防科学计数法
    for (let r = 0; r < rows.length; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: BARCODE_COL });
      if (ws[addr]) {
        ws[addr].t = "s";
        ws[addr].v = String(ws[addr].v || "");
      }
    }

    // ✅ 数字列格式
    for (let r = 1; r < rows.length; r++) {
      NUMBER_COLS.forEach(c => {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (ws[addr] && typeof ws[addr].v === "number") {
          ws[addr].t = "n";
          ws[addr].z = "0.###";
        }
      });
    }

    // ✅ 🟢 绿色标记修改条码
    updatedRows.forEach(r => {
      const addr = XLSX.utils.encode_cell({ r, c: BARCODE_COL });
      if (ws[addr]) {
        ws[addr].s = {
          fill: {
            patternType: "solid",
            fgColor: { rgb: "C6EFCE" }
          }
        };
      }
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    const buffer = XLSX.write(wb, {
      type: "buffer",
      bookType: "xlsx",
      cellStyles: true
    });

    // ✅ ⭐ 核心修复：用用户上传文件名
    const originalName = req.file.originalname || "file.csv";
    const baseName = originalName.replace(/\.csv$/i, "");
    const safeFileName = encodeURIComponent(`${baseName}.xlsx`);
    
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${safeFileName}`
    );

    res.setHeader("X-Matched-Count", matched);
    res.setHeader("X-Changed-Barcode-Count", changed);
    res.setHeader("X-Row-Count", rows.length - 1);
    res.setHeader("X-Unmatched-Count", Math.max(rows.length - 1 - matched, 0));
    res.setHeader("X-Delimiter", ";");


    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ message: "处理失败" });
  } finally {
    if (csvPath && fs.existsSync(csvPath)) {
      fs.unlinkSync(csvPath);
    }
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ 后端已启动：http://localhost:${PORT}`);
});