const express = require('express');
const router = express.Router();
const multer = require('multer');
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

router.post('/process-excel', upload.single('file'), (req, res) => {
  try {
    console.log('Dosya işleme başlıyor...');

    const file = req.file;

    if (!file) return res.status(400).json({ error: 'Dosya yüklenmedi' });

    // Dosya buffer'ını okuyun
    const buffer = file.buffer;

    // Excel dosyasını okuyun
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    console.log('Çalışma kitabı işlendi.');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    if (!data || data.length === 0)
      return res.status(400).json({ error: 'Excel dosyası boş veya geçersiz formatta' });
    console.log('Veri çıkarıldı:', data);

    // Verileri işleyin
    const processedData = [];
    const baseCodeCounts = {};

    data.forEach((row) => {
      const productCode = row["Ürün Kodu"];

      // Base code'u çıkar
      const baseCode = extractBaseCode(productCode);

      // Base code için sayaç oluştur veya artır
      if (!baseCodeCounts[baseCode]) {
        baseCodeCounts[baseCode] = 1;
      } else {
        baseCodeCounts[baseCode] += 1;
      }

      const variantNumber = baseCodeCounts[baseCode];

      const isFirstVariant = variantNumber === 1;

      // Ürün Kodu'nun sonunda varsa ekstra tireyi kaldır
      const cleanedProductCode = productCode.replace(/-$/, '');

      // Veri nesnesini oluşturun
      const dataObject = {
        "Kategori No": row["Kategori No"],
        "Kategori Açıklama": row["Kategori Açıklama"],
        "Ürün Adı": row["Ürün Adı"],
        "Ürün Kodu": row["Ürün Kodu"],
        "Marka": row["Marka"],
        "Varyant - Ürün Kodu": `${cleanedProductCode}-${variantNumber}`,
        "Varyant - Boyut": row["Boyut/Ebat"],
        "Ürün Stok Miktarı": "5",
        "Liste Fiyatı (Kdv Dahil)": row["Liste Fiyatı (Kdv Dahil)"],
        "Goturc İndirimli Satış Fiyatı (Kdv Dahil)":
          row["Goturc İndirimli Satış Fiyatı (Kdv Dahil)"],
        "Para Birimi": row["Para birimi"],
        "Görsel1": row["Görsel1"] || "",
        "Görsel2": row["Görsel2"] || "",
        "Görsel3": row["Görsel3"] || "",
        "Görsel4": row["Görsel4"] || "",
        "Görsel5": row["Görsel5"] || "",
        "Görsel6": row["Görsel6"] || "",
        "Görsel7": row["Görsel7"] || "",
        "Görsel8": row["Görsel8"] || "",
        "Görsel9": row["Görsel9"] || "",
        "Ürün Açıklama": row["Ürün Açıklama"],
        "Boyut/Ebat": "",
        "Hazırlık Süresi": row["Hazırlık Süresi (Gün)"] || 1,
        "Kargo Şablonu": row["Kargo Şablonu"] || "Yurtiçi",
      };

      // Eğer bu ürünün ilk varyantıysa, ek alanları ekleyin
      if (isFirstVariant) {
        dataObject["Çerçeve Tipi"] = row["Çerçeve Tipi"] || "";
        dataObject["Materyal"] = row["Materyal"] || "";
        dataObject["Parça Sayısı"] = row["Parça Sayısı"] || "";
        dataObject["Tema / Stil"] = row["Tema / Stil"] || "";
      }

      processedData.push(dataObject);
    });

    // Base code'u çıkarmak için fonksiyon
    function extractBaseCode(productCode) {
      const regex = /([A-Z]+\d+)-?$/;
      const match = productCode.match(regex);
      if (match) {
        return match[1];
      } else {
        return productCode;
      }
    }

    // Çıktı dizinini ve yolunu belirleyin
    const outputDir = path.join(__dirname, '../../excels');
    const outputPath = path.join(outputDir, 'deneme53.xlsx');

    // Çıktı dizini yoksa oluşturun
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`Çıktı dizini oluşturuldu: ${outputDir}`);
    }

    // Yeni Excel dosyasını oluşturun
    const outputData = processedData; // İşlenen veriler
    const outputSheet = XLSX.utils.json_to_sheet(outputData);

    // Başlıkları eklemek için
    const headers = Object.keys(outputData[0]);
    XLSX.utils.sheet_add_aoa(outputSheet, [headers], { origin: 'A1' });

    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, 'Sonuç');

    // Dosyayı yazın
    const wbout = XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'buffer' });
    fs.writeFileSync(outputPath, wbout);
    console.log('Dosya başarıyla yazıldı:', outputPath);

    // İstemciye başarılı mesajı gönderin
    res.json({
      message: 'Dosya başarıyla işlendi ve kaydedildi.',
      filePath: outputPath,
    });
  } catch (error) {
    console.error('Bir hata oluştu:', error);
    res.status(500).json({ error: 'Bir hata oluştu', details: error.message });
  }
});

module.exports = router;
