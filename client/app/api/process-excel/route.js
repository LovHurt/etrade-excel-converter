export const runtime = "nodejs";

import * as XLSX from "xlsx";
import path from "path";
import { NextResponse } from "next/server";

export async function POST(req) {
  try {
    console.log("Starting file processing...");

    const formData = await req.formData();
    const file1 = formData.get("file1");
    const file2 = formData.get("file2");
    const productCode = formData.get("productCode");
    const startNumber = parseInt(formData.get("startNumber"), 10);

    if (!file1 || !file2) {
      console.error("Missing file uploads in formData.");
      return NextResponse.json(
        { error: "File uploads are required" },
        { status: 400 }
      );
    }
    console.log("Files received:", file1, file2);

    const buffer1 = Buffer.from(await file1.arrayBuffer());
    console.log("First file buffer created.");

    const workbook1 = XLSX.read(buffer1, { type: "buffer" });
    console.log("First workbook processed.");
    const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
    const data1 = XLSX.utils.sheet_to_json(sheet1);

    // Validate data1 content
    if (!data1 || data1.length === 0) {
      console.error("First file contains no data or is invalid.");
      return NextResponse.json(
        { error: "First file is empty or incorrectly formatted" },
        { status: 400 }
      );
    }
    console.log("Data1 extracted:", data1);

    const categoryNumber = data1[0]["Kategori No"];
    const categoryDescription = data1[0]["Kategori Açıklama"];
    const productName = data1[0]["Ürün Adı"];
    const brand = data1[0]["Marka"];

    console.log(
      "Category details extracted:",
      categoryNumber,
      categoryDescription,
      productName
    );

    const buffer2 = Buffer.from(await file2.arrayBuffer());
    console.log("Second file buffer created.");

    const workbook2 = XLSX.read(buffer2, { type: "buffer" });
    console.log("Second workbook processed.");
    const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];
    let data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

    if (data2.length <= 1) {
      const headers = data2[0];
      data2 = [];
      console.log("Data2 initialized with headers:", headers);
    } else {
      console.log("Data2 extracted with data:", data2);
    }

    const populatedData = data1.map((row, index) => ({
      "Kategori No": categoryNumber,
      "Kategori Açıklama": categoryDescription,
      "Ürün Adı": row["productName"],
      "Ürün Kodu": row["Ürün Kodu"],
      Marka: brand,
      "Varyant - Ürün Kodu": `${productCode}-${startNumber + index}`,
      "Varyant - Boyut": row["Boyut"],
      "Ürün Stok Mikta": "5",
      "Liste Fiyatı(Kdv Dahil)": row["Liste Fiyatı"],
      "Goturc İndirimli Satış Fiyatı (Kdv Dahil)": row["İndirimli Fiyat"],
      "Para Birimi": "TL",
      Görsel1: row["Görsel1"],
      Görsel2: row["Görsel2"],
      Görsel3: row["Görsel3"],
      Görsel4: row["Görsel4"],
      Görsel5: row["Görsel5"],
      Görsel6: row["Görsel6"],
      Görsel7: row["Görsel7"],
      Görsel8: row["Görsel8"],
      Görsel9: row["Görsel9"],
      "Ürün Açıklama": row["Ürün Açıklama"],
      "Boyut/Ebat": "",
      "Hazırlık Süresi": 1,
      "Kargo Şablonu": "Yurt İçi",
    }));

    data2.push(...populatedData);

    console.log("Creating output file with populated data...");
    const outputSheet = XLSX.utils.json_to_sheet(data2);
    const outputWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "Sonuç");

    const outputDir = path.join(process.cwd(), "excels");
    console.log("107", outputDir);

    const outputPath = path.join(outputDir, "deneme2.xlsx");
    console.log("110",outputPath)

    try {
      XLSX.writeFile(outputWorkbook, outputPath);
      console.log("File successfully written to:", outputPath);
    } catch (writeError) {
      console.error("Error writing file:", writeError);
      return NextResponse.json(
        {
          error: "Failed to write Excel file",
          details: writeError.message,
        },
        { status: 500 }
      );
    }

    return NextResponse.json({
      message: "File processed and saved successfully.",
      filePath: outputPath,
    });
  } catch (error) {
    console.error("An error occurred:", error);
    return NextResponse.json(
      { error: "An error occurred", details: error.message },
      { status: 500 }
    );
  }
}
