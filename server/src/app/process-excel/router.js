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

    if (!data || data.length === 0) return res.status(400).json({ error: 'Excel dosyası boş veya geçersiz formatta' });
    console.log('Veri çıkarıldı:', data);

    // Verileri işleyin
    const processedData = [];
    const productsProcessed = new Set();
    
    data.forEach((row, index) => {
      const productCode = row["Ürün Kodu"];
      const isFirstVariant = !productsProcessed.has(productCode);
    
      // Ürün kodunu işlenen ürünler listesine ekleyin
      productsProcessed.add(productCode);
    
      // Veri nesnesini oluşturun
      const dataObject = {
        "Kategori No": row["Kategori No"],
        "Kategori Açıklama": row["Kategori Açıklama"],
        "Ürün Adı": row["Ürün Adı"],
        "Ürün Kodu": row["Ürün Kodu"],
        "Marka": row["Marka"],
        "Varyant - Ürün Kodu": `${row["Ürün Kodu"]}-${index + 1}`,
        "Varyant - Boyut": row["Boyut/Ebat"],
        "Ürün Stok Miktarı": "5",
        "Liste Fiyatı(Kdv Dahil)": row["Liste Fiyatı (Kdv Dahil)"],
        "Goturc İndirimli Satış Fiyatı (Kdv Dahil)": row["Goturc İndirimli Satış Fiyatı (Kdv Dahil)"],
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
        "Boyut/Ebat": row["Boyut/Ebat"] || "",
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

    // Çıktı dizinini ve yolunu belirleyin
    const outputDir = path.join(__dirname, '../../excels');
    const outputPath = path.join(outputDir, 'deneme52.xlsx');

    // Çıktı dizini yoksa oluşturun
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`Çıktı dizini oluşturuldu: ${outputDir}`);
    }

    // Yeni Excel dosyasını oluşturun
    const outputData = processedData; // İşlenen veriler
    const outputSheet = XLSX.utils.json_to_sheet(outputData);
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

// const router = require("express").Router()

// const XLSX = require("xlsx");
// const fs = require("fs");
// const path = require("path");

// // İşlem yapmak istediğiniz Excel dosyasının yolunu belirtin
// const filePath1 = path.join(__dirname, "zidekordeneme-1.xlsx");

// console.log(__dirname)
// // Çıktı dosyasının kaydedileceği dizin
// const outputDir = path.join(__dirname, "excels");

// // Çıktı dosyasının tam yolu
// const outputPath = path.join(outputDir, "deneme2.xlsx");

// router.get("/me", tokenCheck, me)

// try {
//   console.log("Dosya işleme başlıyor...");

//   // 1. İlk dosyanın okunması
//   if (!fs.existsSync(filePath1)) {
//     throw new Error(`Dosya bulunamadı: ${filePath1}`);
//   }

//   const workbook1 = XLSX.readFile(filePath1);
//   console.log("İlk çalışma kitabı işlendi.");
//   const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
//   const data1 = XLSX.utils.sheet_to_json(sheet1);

//   if (!data1 || data1.length === 0) {
//     throw new Error("İlk dosya boş veya geçersiz formatta.");
//   }
//   console.log("Data1 çıkarıldı:", data1);

//   const categoryNumber = data1[0]["Kategori No"];
//   const categoryDescription = data1[0]["Kategori Açıklama"];
//   const productName = data1[0]["Ürün Adı"];
//   const brand = data1[0]["Marka"];

//   console.log(
//     "Kategori detayları çıkarıldı:",
//     categoryNumber,
//     categoryDescription,
//     productName
//   );

//   // 2. data2'nin başlatılması
//   // Gerekli başlıkları tanımlayın
//   const headers = [
//     "Kategori No",
//     "Kategori Açıklama",
//     "Ürün Adı",
//     "Ürün Kodu",
//     "Marka",
//     "Varyant - Ürün Kodu",
//     "Varyant - Boyut",
//     "Ürün Stok Mikta",
//     "Liste Fiyatı(Kdv Dahil)",
//     "Goturc İndirimli Satış Fiyatı (Kdv Dahil)",
//     "Para Birimi",
//     "Görsel1",
//     "Görsel2",
//     "Görsel3",
//     "Görsel4",
//     "Görsel5",
//     "Görsel6",
//     "Görsel7",
//     "Görsel8",
//     "Görsel9",
//     "Ürün Açıklama",
//     "Boyut/Ebat",
//     "Hazırlık Süresi",
//     "Kargo Şablonu",
//   ];

//   // data2'yi başlıklarla birlikte başlatın
//   let data2 = [headers]; // Başlık satırı eklemek istiyorsanız: let data2 = [headers];

//   // 3. Verilerin işlenmesi
//   const populatedData = data1.map((row, index) => ({
//     "Kategori No": categoryNumber,
//     "Kategori Açıklama": categoryDescription,
//     "Ürün Adı": row["Ürün Adı"],
//     "Ürün Kodu": row["Ürün Kodu"],
//     "Marka": brand,
//     "Varyant - Ürün Kodu": `${row["Ürün Kodu"]}-${index + 1}`,
//     "Varyant - Boyut": row["Boyut/Ebat"],
//     "Ürün Stok Mikta": "5",
//     "Liste Fiyatı(Kdv Dahil)": row["Liste Fiyatı (Kdv Dahil)"],
//     "Goturc İndirimli Satış Fiyatı (Kdv Dahil)":
//       row["Goturc İndirimli Satış Fiyatı (Kdv Dahil)"],
//     "Para Birimi": row["Para birimi"],
//     "Görsel1": row["Görsel1"],
//     "Görsel2": row["Görsel2"],
//     "Görsel3": row["Görsel3"],
//     "Görsel4": row["Görsel4"],
//     "Görsel5": row["Görsel5"],
//     "Görsel6": row["Görsel6"] || "",
//     "Görsel7": row["Görsel7"] || "",
//     "Görsel8": row["Görsel8"] || "",
//     "Görsel9": row["Görsel9"] || "",
//     "Ürün Açıklama": row["Ürün Açıklama"],
//     "Boyut/Ebat": row["Boyut/Ebat"] || "",
//     "Hazırlık Süresi": row["Hazırlık Süresi (Gün)"] || 1,
//     "Kargo Şablonu": row["Kargo Şablonu"] || "Yurtiçi",
//   }));

//   // populatedData'yı data2'ye ekleyin
//   data2.push(...populatedData);

//   // 4. Çıktı dosyasının oluşturulması
//   console.log("Çıktı dosyası oluşturuluyor...");
//   const outputSheet = XLSX.utils.json_to_sheet(data2);

//   // Başlıkları eklemek için (eğer data2'yi boş başlattıysanız)
//   XLSX.utils.sheet_add_aoa(outputSheet, [headers], { origin: "A1" });

//   const outputWorkbookNew = XLSX.utils.book_new();
//   XLSX.utils.book_append_sheet(outputWorkbookNew, outputSheet, "Sonuç");

//   // 5. Çıktı dizininin oluşturulması (varsa)
//   if (!fs.existsSync(outputDir)) {
//     fs.mkdirSync(outputDir, { recursive: true });
//     console.log(`Çıktı dizini oluşturuldu: ${outputDir}`);
//   }

//   // 6. Çıktı dosyasının yazılması
//   try {
//     const wbout = XLSX.write(outputWorkbookNew, { bookType: "xlsx", type: "buffer" });
//     fs.writeFileSync(outputPath, wbout);
//     console.log("Dosya başarıyla yazıldı:", outputPath);
//   } catch (writeError) {
//     console.error("Dosya yazma hatası:", writeError);
//     process.exit(1);
//   }

//   console.log("İşlem başarıyla tamamlandı.");
// } catch (error) {
//   console.error("Bir hata oluştu:", error.message);
//   process.exit(1);
// }

// module.exports = router