const express = require("express");
const router = express.Router();
const multer = require("multer");
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

router.post("/process-excel", upload.single("file"), (req, res) => {
  try {
    console.log("Dosya işleme başlıyor...");

    const file = req.file;

    if (!file) return res.status(400).json({ error: "Dosya yüklenmedi" });

    const buffer = file.buffer;

    // Excel dosyasını okuyun
    const workbook = XLSX.read(buffer, { type: "buffer" });
    console.log("Çalışma kitabı işlendi.");
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    if (!data || data.length === 0)
      return res
        .status(400)
        .json({ error: "Excel dosyası boş veya geçersiz formatta" });
    console.log(
      "Veri çıkarıldı:"
      // data
    );

    // Verileri işleyin
    const processedData = [];
    const baseCodeCounts = {};
    const categoriesData = {};

    data.forEach((row) => {
      const salePrice = row["Piyasa Satış Fiyatı (KDV Dahil)"];
      const trendyolPrice = row["Trendyol'da Satılacak Fiyat (KDV Dahil)"];

      const originalProductCode = row["Model Kodu"];
      const baseProductCode = extractBaseProductCode(originalProductCode);

      // Base code için sayaç oluştur veya artır
      if (!baseCodeCounts[baseProductCode]) {
        baseCodeCounts[baseProductCode] = 1;
      } else {
        baseCodeCounts[baseProductCode] += 1;
      }

      const variantNumber = baseCodeCounts[baseProductCode];
      const isFirstVariant = variantNumber === 1;
      const variantProductCode = `${baseProductCode}-${variantNumber}`;
      const category = row["Kategori İsmi"];

      const stockAmount = (salePrice == 0 || trendyolPrice == 0) ? "0" : "5";


      // Veri nesnesini oluşturun
      const dataObject = {
        "Kategori No": row["Kategori No"],
        "Kategori İsmi": row["Kategori İsmi"],
        "Ürün Adı": row["Ürün Adı"],
        "Ürün Kodu": baseProductCode, // Tek ürün kodu
        Marka: row["Marka"],
        "Varyant - Ürün Kodu": variantProductCode,
        "Varyant - Renk": row["Ürün Rengi"],
        "Varyant - Beden": row["Beden"],
        "Ürün Stok Miktarı": stockAmount,
        "Liste Fiyatı (Kdv Dahil)": row["Piyasa Satış Fiyatı (KDV Dahil)"],
        "Goturc İndirimli Satış Fiyatı (Kdv Dahil)":
          row["Trendyol'da Satılacak Fiyat (KDV Dahil)"] || "",
        "Para Birimi": row["Para birimi"] || "TL",
        Görsel1: row["Görsel 1"] || "",
        Görsel2: row["Görsel 2"] || "",
        Görsel3: row["Görsel 3"] || "",
        Görsel4: row["Görsel 4"] || "",
        Görsel5: row["Görsel 5"] || "",
        Görsel6: row["Görsel 6"] || "",
        Görsel7: row["Görsel 7"] || "",
        Görsel8: row["Görsel 8"] || "",
        Görsel9: row["Görsel 9"] || "",
        "Ürün Açıklama": row["Ürün Açıklaması"] || row["Ürün Adı"],
        Beden: "",
        "Garanti Süresi": "",
        Materyal: "",
        "Uyumlu Marka": "",
        "Hazırlık Süresi": row["Sevkiyat Süresi"] || 2,
        "Kargo Şablonu": row["Kargo Şablonu"] || "",
      };

      // // Eğer bu ürünün ilk varyantıysa, ek alanları ekleyin
      // if (isFirstVariant) {
      //   dataObject["Çerçeve Tipi"] = row["Çerçeve Tipi"] || "";
      //   dataObject["Materyal"] = row["Materyal"] || "";
      //   dataObject["Parça Sayısı"] = row["Parça Sayısı"] || "";
      //   dataObject["Tema / Stil"] = row["Tema / Stil"] || "";
      // }

      // processedData.push(dataObject);

      if (!categoriesData[category]) {
        categoriesData[category] = [];
      }
      categoriesData[category].push(dataObject);
      // console.log(`Orijinal Ürün Kodu: ${originalProductCode}, Base Product Code: ${baseProductCode}, Varyant No: ${variantNumber}`);
    });

    function extractBaseProductCode(productCode) {
      // Güncellenmiş regex ile boyut bilgisini eşleştir ve kaldır
      const regex = /^(.*?)[0-9]+x[0-9]+(.*?)-?$/;
      const match = productCode.match(regex);
      if (match) {
        return `${match[1]}${match[2]}`.replace(/-$/, "");
      } else {
        // Eşleşme yoksa orijinal ürün kodundan sonundaki '-' işaretini kaldır
        return productCode.replace(/-$/, "");
      }
    }

    const outputDir = path.join(__dirname, "../../excels");

    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`Çıktı dizini oluşturuldu: ${outputDir}`);
    }

    const outputPaths = [];

    for (const category in categoriesData) {
      const outputData = categoriesData[category];
      const outputSheet = XLSX.utils.json_to_sheet(outputData);

      const headers = Object.keys(outputData[0]);
      XLSX.utils.sheet_add_aoa(outputSheet, [headers], { origin: "A1" });

      const outputWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "Sonuç");

      const outputPath = path.join(
        outputDir,
        `XXXXXXXXXXXXXXXXXXXXXXXXXXXX${category}_YeniUrun.xlsx`
      );
      const wbout = XLSX.write(outputWorkbook, {
        bookType: "xlsx",
        type: "buffer",
      });
      fs.writeFileSync(outputPath, wbout);

      outputPaths.push(outputPath);
      console.log(`Dosya başarıyla yazıldı: ${outputPath}`);
    }

    res.json({
      message: "Dosya başarıyla işlendi ve kaydedildi.",
      filePath: outputPaths,
    });
  } catch (error) {
    console.error("Bir hata oluştu:", error);
    res.status(500).json({ error: "Bir hata oluştu", details: error.message });
  }
});

module.exports = router;
