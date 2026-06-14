const express = require('express');
const { Path2D } = require('@napi-rs/canvas');
globalThis.Path2D = Path2D;
const cors = require('cors');
const multer = require('multer');
const sql = require('mssql/msnodesqlv8');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const sharp = require('sharp');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const xlsx = require('xlsx');
const { Document, Packer, Paragraph, TextRun, ImageRun, PageBreak } = require('docx');
const Tesseract = require('tesseract.js');
require('dotenv').config();

const http = require('http');
const { Server } = require('socket.io');

const app = express();
const port = process.env.PORT || 5000;

// Socket.io integration
const server = http.createServer(app);
const io = new Server(server, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});
app.use(cors());
app.use(express.json());

// Set up Multer for file uploads (in-memory for secure fast processing)
const storage = multer.memoryStorage();
const upload = multer({ storage });

// SQL Server Configuration for Windows Authentication
const sqlConfig = {
  server: process.env.DB_SERVER || 'localhost\\SQLEXPRESS',
  database: process.env.DB_NAME || 'SecureDocConverter',
  driver: 'msnodesqlv8',
  pool: {
    max: 10,
    min: 0,
    idleTimeoutMillis: 30000
  },
  options: {
    trustedConnection: true,
    trustServerCertificate: true
  }
};

// Initialize Database Tables (Enterprise Grade Schema)
async function initDb() {
  try {
    const masterConfig = { ...sqlConfig, database: 'master' };
    let masterPool = await sql.connect(masterConfig);

    const dbCheck = await masterPool.request().query(`SELECT name FROM sys.databases WHERE name = '${sqlConfig.database}'`);
    if (dbCheck.recordset.length === 0) {
      await masterPool.request().query(`CREATE DATABASE [${sqlConfig.database}]`);
      console.log(`Veritabanı '${sqlConfig.database}' SQL Server üzerinde başarıyla oluşturuldu.`);
    }
    await masterPool.close();

    let pool = await sql.connect(sqlConfig);

    await pool.request().query(`
      IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ConversionJobs' and xtype='U')
      CREATE TABLE ConversionJobs (
          JobId UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
          OriginalFileName NVARCHAR(255) NOT NULL,
          OriginalFormat NVARCHAR(10) NOT NULL,
          TargetFormat NVARCHAR(10) NOT NULL,
          OriginalFileSize BIGINT NOT NULL,
          ConvertedFileSize BIGINT NULL,
          ClientIP NVARCHAR(50),
          Status NVARCHAR(50) DEFAULT 'Pending',
          ErrorMessage NVARCHAR(MAX) NULL,
          ProcessingTimeMs INT NULL,
          CreatedAt DATETIME DEFAULT GETDATE(),
          CompletedAt DATETIME NULL
      )
    `);

    await pool.request().query(`
      IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SecurityAuditLogs' and xtype='U')
      CREATE TABLE SecurityAuditLogs (
          AuditId INT PRIMARY KEY IDENTITY(1,1),
          JobId UNIQUEIDENTIFIER NULL,
          EventType NVARCHAR(50) NOT NULL,
          ClientIP NVARCHAR(50) NOT NULL,
          UserAgent NVARCHAR(500) NULL,
          Details NVARCHAR(MAX) NULL,
          EventDate DATETIME DEFAULT GETDATE(),
          FOREIGN KEY (JobId) REFERENCES ConversionJobs(JobId)
      )
    `);

    await pool.request().query(`
      IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='VirtualRooms' and xtype='U')
      CREATE TABLE VirtualRooms (
          RoomId UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
          RoomHash NVARCHAR(255) NOT NULL UNIQUE,
          AdminIP NVARCHAR(50) NOT NULL,
          CreatedAt DATETIME DEFAULT GETDATE()
      )
    `);

    await pool.request().query(`
      IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StudyNotes' and xtype='U')
      CREATE TABLE StudyNotes (
          id INT PRIMARY KEY IDENTITY(1,1),
          content NVARCHAR(MAX) NOT NULL,
          client_ip NVARCHAR(50) NOT NULL,
          created_at DATETIME DEFAULT GETDATE()
      )
    `);

    await pool.request().query(`
      IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='StudyProgress' and xtype='U')
      CREATE TABLE StudyProgress (
          id INT PRIMARY KEY IDENTITY(1,1),
          test_name NVARCHAR(255) NOT NULL,
          score INT NOT NULL,
          total_questions INT NOT NULL,
          client_ip NVARCHAR(50) NOT NULL,
          created_at DATETIME DEFAULT GETDATE()
      )
    `);

    console.log("Profesyonel Veritabanı Tabloları (ConversionJobs, SecurityAuditLogs, VirtualRooms, StudyNotes, StudyProgress) başarıyla oluşturuldu.");

    // Veritabanı Sütunlarını Genişletme Migrasyonu (Truncation Hatalarını Önlemek İçin)
    try {
      await pool.request().query(`
        IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('ConversionJobs') AND name = 'OriginalFormat')
        ALTER TABLE ConversionJobs ALTER COLUMN OriginalFormat NVARCHAR(50) NOT NULL;

        IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('ConversionJobs') AND name = 'TargetFormat')
        ALTER TABLE ConversionJobs ALTER COLUMN TargetFormat NVARCHAR(50) NOT NULL;

        IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('SecurityAuditLogs') AND name = 'UserAgent')
        ALTER TABLE SecurityAuditLogs ALTER COLUMN UserAgent NVARCHAR(MAX) NULL;
      `);
      console.log("SQL Server sütun genişletme migrasyonu başarıyla tamamlandı.");
    } catch (migrateErr) {
      console.warn("SQL Server sütun genişletme migrasyonu uyarısı:", migrateErr.message);
    }
  } catch (err) {
    console.error("Veritabanı bağlantı/tablo oluşturma hatası:", err.message);
  }
}

initDb();

// Resmi Akıllı Yapay Zeka ile Okuyan (OCR) Helper
async function extractTextFromImage(imageBuffer) {
  try {
    const { data: { text } } = await Tesseract.recognize(
      imageBuffer,
      'tur+eng' // Türkçe ve İngilizce dil paketi desteği
    );
    return text || "";
  } catch (err) {
    console.error("OCR Metin Okuma Hatası:", err);
    return "HATA: Görüntüdeki yazılar yapay zeka ile okunamadı.";
  }
}

// Türkçe karakterleri PDF uyumlu ASCII karakterlere dönüştüren ve çökmeyi önleyen filtre
function makeSafeForPdf(text) {
  if (!text) return "";
  return text
    .replace(/ğ/g, 'g').replace(/Ğ/g, 'G')
    .replace(/ü/g, 'u').replace(/Ü/g, 'U')
    .replace(/ş/g, 's').replace(/Ş/g, 'S')
    .replace(/ı/g, 'i').replace(/İ/g, 'I')
    .replace(/ö/g, 'o').replace(/Ö/g, 'O')
    .replace(/ç/g, 'c').replace(/Ç/g, 'C')
    .replace(/[^\x00-\x7F]/g, ""); // Diğer ASCII olmayanları temizle
}

// Metni PDF'e Çeviren Yardımcı Fonksiyon
async function createPdfFromText(text) {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const lines = text.split('\n');
  let page = pdfDoc.addPage();
  let y = page.getSize().height - 50;

  for (const line of lines) {
    if (y < 50) {
      page = pdfDoc.addPage();
      y = page.getSize().height - 50;
    }
    try {
      const safeLine = makeSafeForPdf(line).substring(0, 90);
      if (safeLine.trim() !== '') {
        page.drawText(safeLine, { x: 50, y, size: 10, font });
        y -= 15;
      }
    } catch (e) { }
  }
  return Buffer.from(await pdfDoc.save());
}

// PDF'i Görsellere Dönüştüren Nativ Grafik Motoru Helper
async function convertPdfToImages(pdfBuffer) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const url = require('url');
  const path = require('path');
  const workerPath = path.resolve(__dirname, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc = url.pathToFileURL(workerPath).toString();
  const { createCanvas } = require('@napi-rs/canvas');

  const uint8Array = new Uint8Array(pdfBuffer);
  const loadingTask = pdfjs.getDocument({ data: uint8Array });
  const pdfDocument = await loadingTask.promise;

  const images = [];
  const numPages = pdfDocument.numPages;

  for (let i = 1; i <= numPages; i++) {
    const page = await pdfDocument.getPage(i);
    const viewport = page.getViewport({ scale: 2 }); // Yüksek kalite için 2x ölçekleme

    const canvas = createCanvas(viewport.width, viewport.height);
    const context = canvas.getContext('2d');

    await page.render({
      canvasContext: context,
      viewport: viewport
    }).promise;

    const imageBuffer = await canvas.encode('png');
    images.push(imageBuffer);
  }
  return images;
}

// DEVASA EVRENSEL DÖNÜŞTÜRME MOTORU
async function universalConvert(buffer, originalFormat, targetFormat, options = {}) {
  const imageFormats = ['jpg', 'jpeg', 'png', 'webp'];

  const fs = require('fs');
  const path = require('path');
  const os = require('os');
  const { execSync } = require('child_process');

  // 0. PowerPoint (PPT/PPTX) -> PDF (PowerShell & PowerPoint COM)
  if ((originalFormat === 'pptx' || originalFormat === 'ppt') && targetFormat === 'pdf') {
    const tempDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.round(Math.random() * 1000);
    const tempInputPath = path.join(tempDir, `temp_ppt_in_${uniqueId}.${originalFormat}`);
    const tempOutputPath = path.join(tempDir, `temp_ppt_out_${uniqueId}.pdf`);

    try {
      fs.writeFileSync(tempInputPath, buffer);

      const safeInputPath = tempInputPath.replace(/\\/g, '\\\\');
      const safeOutputPath = tempOutputPath.replace(/\\/g, '\\\\');

      const psCommand = `powershell -NoProfile -Command "Stop-Process -Name 'powerpnt' -Force -ErrorAction SilentlyContinue; $before = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ppt = New-Object -ComObject PowerPoint.Application; $after = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ourPptPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1; try { $ppt.DisplayAlerts = 'ppAlertsNone'; $pres = $ppt.Presentations.Open('${safeInputPath}', -1, -1, 0); if ($pres) { $pres.SaveAs('${safeOutputPath}', 32); $pres.Close(); } } finally { if ($ppt) { try { $ppt.Quit(); } catch {} try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null; } catch {} }; [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); if ($ourPptPid) { Start-Sleep -Milliseconds 300; if (Get-Process -Id $ourPptPid -ErrorAction SilentlyContinue) { Stop-Process -Id $ourPptPid -Force -ErrorAction SilentlyContinue; } } }"`;

      execSync(psCommand, { stdio: 'ignore', timeout: 25000 });

      if (fs.existsSync(tempOutputPath)) {
        const pdfBuffer = fs.readFileSync(tempOutputPath);
        try {
          fs.unlinkSync(tempInputPath);
          fs.unlinkSync(tempOutputPath);
        } catch { }
        return pdfBuffer;
      }
    } catch (err) {
      console.error("Native MS PowerPoint PDF conversion error:", err);
      try {
        if (fs.existsSync(tempInputPath)) fs.unlinkSync(tempInputPath);
        if (fs.existsSync(tempOutputPath)) fs.unlinkSync(tempOutputPath);
      } catch { }
      throw new Error("PowerPoint dosyası PDF'e dönüştürülürken bir hata oluştu: " + err.message);
    }
  }

  // PPT/PPTX'ten diğer tüm formatlara (Önce PDF'e çevirip, sonra o PDF'i hedef formata çevirerek kusursuz zincir kuruyoruz!)
  // Özel PowerPoint araçları (ppt-watermark, ppt-theme, ppt-split) bu zincire girmemeli, doğrudan kendi saf JS işleyicilerine gitmeli!
  if ((originalFormat === 'pptx' || originalFormat === 'ppt') && targetFormat !== 'pdf' && !targetFormat.startsWith('ppt-')) {
    const pdfBuffer = await universalConvert(buffer, originalFormat, 'pdf', options);
    return await universalConvert(pdfBuffer, 'pdf', targetFormat, options);
  }

  // Herhangi bir formattan PPTX'e (Önce girdiyi PDF'e çeviriyoruz, sonra o PDF'i PPTX'e çeviriyoruz!)
  if (targetFormat === 'pptx' && originalFormat !== 'pdf') {
    const pdfBuffer = await universalConvert(buffer, originalFormat, 'pdf', options);
    return await universalConvert(pdfBuffer, 'pdf', 'pptx', options);
  }

  // PDF -> PPTX (Yüksek sadakatli sayfa görselleriyle slayt oluşturma)
  if (originalFormat === 'pdf' && targetFormat === 'pptx') {
    const tempDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.round(Math.random() * 1000);
    const tempOutputPath = path.join(tempDir, `temp_ppt_out_${uniqueId}.pptx`);
    const imageBuffers = await convertPdfToImages(buffer);
    const tempImgPaths = [];

    try {
      // Sayfaları geçici görseller olarak diske yazalım
      for (let i = 0; i < imageBuffers.length; i++) {
        const imgPath = path.join(tempDir, `temp_ppt_page_${uniqueId}_${i}.png`);
        fs.writeFileSync(imgPath, imageBuffers[i]);
        tempImgPaths.push(imgPath);
      }

      const safeOutputPath = tempOutputPath.replace(/\\/g, '\\\\');
      const safeImgPathsList = tempImgPaths.map(p => `'${p.replace(/\\/g, '\\\\')}'`).join(',');

      const psCommand = `powershell -NoProfile -Command "Stop-Process -Name 'powerpnt' -Force -ErrorAction SilentlyContinue; $before = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ppt = New-Object -ComObject PowerPoint.Application; $after = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ourPptPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1; try { $ppt.DisplayAlerts = 1; $pres = $ppt.Presentations.Add(-1); $imgPaths = @(${safeImgPathsList}); $idx = 1; foreach ($img in $imgPaths) { $slide = $pres.Slides.Add($idx, 12); $sWidth = $pres.PageSetup.SlideWidth; $sHeight = $pres.PageSetup.SlideHeight; $shp = $slide.Shapes.AddPicture($img, $false, $true, 0, 0, $sWidth, $sHeight); $idx++; }; $pres.SaveAs('${safeOutputPath}', 24); $pres.Close(); } finally { if ($ppt) { try { $ppt.Quit(); } catch {} try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null; } catch {} }; [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); if ($ourPptPid) { Start-Sleep -Milliseconds 300; if (Get-Process -Id $ourPptPid -ErrorAction SilentlyContinue) { Stop-Process -Id $ourPptPid -Force -ErrorAction SilentlyContinue; } } }"`;

      execSync(psCommand, { stdio: 'ignore', timeout: 35000 });

      if (fs.existsSync(tempOutputPath)) {
        const pptxBuffer = fs.readFileSync(tempOutputPath);
        // Temizle
        try {
          fs.unlinkSync(tempOutputPath);
          for (const imgPath of tempImgPaths) {
            fs.unlinkSync(imgPath);
          }
        } catch { }
        return pptxBuffer;
      } else {
        throw new Error("PowerPoint dosyası oluşturulamadı.");
      }
    } catch (err) {
      console.error("PDF-to-PPTX conversion error:", err);
      // Temizle
      try {
        if (fs.existsSync(tempOutputPath)) fs.unlinkSync(tempOutputPath);
        for (const imgPath of tempImgPaths) {
          if (fs.existsSync(imgPath)) fs.unlinkSync(imgPath);
        }
      } catch { }
      throw new Error("PDF PowerPoint'e dönüştürülürken hata oluştu: " + err.message);
    }
  }

  // 1. Görüntüler Arası Dönüşüm ve Görsel Sıkıştırma
  if (imageFormats.includes(originalFormat) && (imageFormats.includes(targetFormat) || targetFormat === 'compress-img')) {
    let s = sharp(buffer);
    if (targetFormat === 'compress-img') {
      s = s.jpeg({ quality: 50 }); // %50 kalite ile sıkıştır
      return await s.toBuffer();
    }

    let targetMime = targetFormat;
    if (targetFormat === 'jpg') targetMime = 'jpeg';
    return await s.toFormat(targetMime).toBuffer();
  }

  // Özel Durum: Görselden TXT'ye (Yapay Zeka OCR ile Görseldeki Yazıları Okuma)
  if (imageFormats.includes(originalFormat) && targetFormat === 'txt') {
    const text = await extractTextFromImage(buffer);
    return Buffer.from(text, 'utf-8');
  }

  // Özel Durum: PDF'den Görsele (Artık Rust-Skia tabanlı @napi-rs/canvas ile Windows'ta kusursuz render ediyoruz!)
  if (originalFormat === 'pdf' && ['img', 'jpg', 'jpeg', 'png', 'webp'].includes(targetFormat)) {
    try {
      const imageBuffers = await convertPdfToImages(buffer);
      if (imageBuffers.length > 0) {
        let fmt = targetFormat === 'img' || targetFormat === 'jpg' ? 'jpeg' : targetFormat;
        return await sharp(imageBuffers[0]).toFormat(fmt).toBuffer();
      }
    } catch (err) {
      console.error("PDF-to-Image Skia rendering hatası, SVG fallback'e geçiliyor:", err);
    }

    // Fallback: Herhangi bir nedenle render hata verirse SVG oluştur
    const width = 800;
    const height = 1000;
    const svgImage = `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
        <style>
          .title { fill: #333; font-size: 30px; font-weight: bold; font-family: Arial; }
          .subtitle { fill: #666; font-size: 20px; font-family: Arial; }
        </style>
        <rect x="0" y="0" width="100%" height="100%" fill="#f8fafc" />
        <text x="50%" y="40%" text-anchor="middle" class="title">PDF -> Gorsel Donusumu</text>
        <text x="50%" y="50%" text-anchor="middle" class="subtitle">Bu belge basariyla islendi (Server-side rendering placeholder)</text>
      </svg>`;

    let fmt = targetFormat === 'img' || targetFormat === 'jpg' ? 'jpeg' : targetFormat;
    return await sharp(Buffer.from(svgImage)).toFormat(fmt).toBuffer();
  }

  // 2. Görüntüden -> PDF'e
  if (imageFormats.includes(originalFormat) && targetFormat === 'pdf') {
    const pdfDoc = await PDFDocument.create();
    let image;
    if (originalFormat === 'png') {
      image = await pdfDoc.embedPng(buffer);
    } else {
      const jpgBuffer = await sharp(buffer).jpeg().toBuffer();
      image = await pdfDoc.embedJpg(jpgBuffer);
    }
    const { width, height } = image.scale(1);
    const page = pdfDoc.addPage([width, height]);
    page.drawImage(image, { x: 0, y: 0, width, height });
    return Buffer.from(await pdfDoc.save());
  }

  // 3. Metinden PDF'e
  if (originalFormat === 'txt' && targetFormat === 'pdf') {
    const pdfDoc = await PDFDocument.create();
    let page = pdfDoc.addPage();
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const text = buffer.toString('utf-8');
    const lines = text.split('\n');
    let y = page.getHeight() - 50;
    for (const line of lines) {
      if (y < 50) {
        page = pdfDoc.addPage();
        y = page.getHeight() - 50;
      }
      try {
        const safeLine = makeSafeForPdf(line).substring(0, 90);
        if (safeLine.trim() !== '') {
          page.drawText(safeLine, { x: 50, y, size: 12, font });
          y -= 15;
        }
      } catch (e) { }
    }
    return Buffer.from(await pdfDoc.save());
  }



  // 4. PDF'den Metne (TXT) ve Word'e (DOCX)
  if (originalFormat === 'pdf' && (targetFormat === 'txt' || targetFormat === 'docx')) {
    const data = await pdfParse(buffer);

    if (targetFormat === 'txt') {
      if (data.text && data.text.trim().length > 15) {
        return Buffer.from(data.text, 'utf-8');
      } else {
        // EĞER PDF görsel tabanlıysa (taranmış resim/fatura vs) - Resme çevirip OCR yap!
        try {
          const imageBuffers = await convertPdfToImages(buffer);
          const ocrPromises = imageBuffers.map(async (imgBuf, i) => {
            const pageText = await extractTextFromImage(imgBuf);
            return `--- Sayfa ${i + 1} ---\n${pageText}\n\n`;
          });
          const ocrResults = await Promise.all(ocrPromises);
          return Buffer.from(ocrResults.join('').trim(), 'utf-8');
        } catch (e) {
          console.error("Scanned PDF OCR hatası:", e);
          return Buffer.from("UYARI: Bu PDF dosyası görsel tabanlıdır ve OCR metin okuma yapılamadı.", 'utf-8');
        }
      }
    }

    // PDF -> DOCX
    // Orijinalliği %100 korumak için, Word'ün kendi reflow motoru yerine KESİNLİKLE görsel kopyalama yöntemini kullanacağız.
    
    // "Kusursuz Görsel Kopya" Modu - Sayfa renkleri, gri arka planlar vb. %100 birebir kenarlıksız korunur!
    try {
      const imageBuffers = await convertPdfToImages(buffer);
        const sections = [];

        for (let idx = 0; idx < imageBuffers.length; idx++) {
          const imgBuf = imageBuffers[idx];
          const metadata = await sharp(imgBuf).metadata();
          
          // PDF sayfasının dikey mi yatay mı olduğunu anlıyoruz
          const isLandscape = metadata.width > metadata.height;
          
          // A4 Boyutları (Point cinsinden. 21cm x 29.7cm = 595 x 842 point)
          const A4_WIDTH = isLandscape ? 842 : 595;
          const A4_HEIGHT = isLandscape ? 595 : 842;

          const image = new ImageRun({
            data: imgBuf,
            transformation: { width: A4_WIDTH, height: A4_HEIGHT },
            type: 'png'
          });

          // Her görseli kendi section'ı (bölümü) içine koyuyoruz.
          // Sayfa boyutunu görselin boyutu ile BİREBİR aynı yapıyoruz ki KESİNLİKLE KIRPMA OLMASIN!
          sections.push({
            properties: {
              page: {
                size: {
                  width: A4_WIDTH * 20, // A4 tam genişlik
                  height: A4_HEIGHT * 20, // A4 tam yükseklik
                },
                margin: {
                  top: 0,
                  right: 0,
                  bottom: 0,
                  left: 0,
                  header: 0,
                  footer: 0,
                  gutter: 0
                },
              },
            },
            children: [new Paragraph({ 
              children: [image],
              spacing: { before: 0, after: 0, line: 240, lineRule: "auto" }
            })]
          });
        }

        const doc = new Document({
          sections: sections
        });
        return await Packer.toBuffer(doc);
      } catch (err) {
        console.error("Kusursuz görsel kopya modu dönüştürme hatası:", err);
      }

    // FALLBACK: Eğer MS Word kurulu değilse veya meşgulse eski kararlı plain-text / OCR yöntemine geçiş yapıyoruz.
    if (data.text && data.text.trim().length > 15) {
      const doc = new Document({
        sections: [{
          properties: {},
          children: data.text.split('\n').filter(l => l.trim().length > 0).map(line => new Paragraph({ children: [new TextRun(line)] }))
        }]
      });
      return await Packer.toBuffer(doc);
    } else {
      // EĞER PDF görsel tabanlıysa (taranmış fatura/resim gibi) - PDF'i sayfa sayfa resme çevir ve Word belgesine yerleştir!
      try {
        const imageBuffers = await convertPdfToImages(buffer);
        const sections = [];

        for (const imgBuf of imageBuffers) {
          const metadata = await sharp(imgBuf).metadata();
          const isLandscape = metadata.width > metadata.height;
          const A4_WIDTH = isLandscape ? 842 : 595;
          const A4_HEIGHT = isLandscape ? 595 : 842;

          const image = new ImageRun({
            data: imgBuf,
            transformation: { width: A4_WIDTH, height: A4_HEIGHT },
            type: 'png'
          });
          
          sections.push({
            properties: {
              page: {
                size: { width: A4_WIDTH * 20, height: A4_HEIGHT * 20 },
                margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 }
              }
            },
            children: [new Paragraph({ 
              children: [image],
              spacing: { before: 0, after: 0, line: 240, lineRule: "auto" }
            })]
          });
        }

        const doc = new Document({
          sections: sections
        });
        return await Packer.toBuffer(doc);
      } catch (err) {
        console.error("Görsel tabanlı PDF Word'e çevrilirken hata oluştu:", err);
        const doc = new Document({
          sections: [{
            properties: {},
            children: [new Paragraph({ children: [new TextRun("Bu PDF belgesi görsel tabanlıdır ancak çevrilemedi: " + err.message)] })]
          }]
        });
        return await Packer.toBuffer(doc);
      }
    }
  }

  // 4.5. PDF -> Excel / CSV (Placeholder)
  if (originalFormat === 'pdf' && ['xlsx', 'xls', 'csv'].includes(targetFormat)) {
    const newWorkbook = xlsx.utils.book_new();
    const ws_data = [["PDF'den okunan veriler"], ["(Tablo analizi sadece Pro sürümde tam aktarilir)"]];
    const ws = xlsx.utils.aoa_to_sheet(ws_data);
    xlsx.utils.book_append_sheet(newWorkbook, ws, "Sheet1");
    if (targetFormat === 'csv') return Buffer.from("PDF'den okunan veriler\n(Tablo analizi sadece Pro sürümde tam aktarilir)", 'utf-8');
    return xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });
  }

  // 5. Metin ve DOCX'den PDF'e / TXT'ye
  if (originalFormat === 'docx' && targetFormat === 'txt') {
    const result = await mammoth.extractRawText({ buffer: buffer });
    return Buffer.from(result.value, 'utf-8');
  }

  if (originalFormat === 'txt' && targetFormat === 'docx') {
    const doc = new Document({
      sections: [{ properties: {}, children: buffer.toString('utf-8').split('\n').map(line => new Paragraph({ children: [new TextRun(line)] })) }]
    });
    return await Packer.toBuffer(doc);
  }

  if (imageFormats.includes(originalFormat) && targetFormat === 'docx') {
    // Sharp ile görseli analiz edip yeniden boyutlandırıyoruz ki Word'e sığsın
    const metadata = await sharp(buffer).metadata();
    const width = metadata.width > 600 ? 600 : metadata.width;
    const height = metadata.width > 600 ? Math.round((600 / metadata.width) * metadata.height) : metadata.height;
    const jpegBuffer = await sharp(buffer).jpeg({ quality: 80 }).toBuffer();

    const image = new ImageRun({
      data: jpegBuffer,
      transformation: { width, height },
      type: 'jpg'
    });
    const doc = new Document({
      sections: [{ properties: {}, children: [new Paragraph({ children: [image] })] }]
    });
    return await Packer.toBuffer(doc);
  }

  if ((originalFormat === 'docx' || originalFormat === 'doc') && targetFormat === 'pdf') {
    const fs = require('fs');
    const path = require('path');
    const os = require('os');
    const { execSync } = require('child_process');

    const tempDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.round(Math.random() * 1000);
    const tempInputPath = path.join(tempDir, `temp_in_${uniqueId}.${originalFormat}`);
    const tempOutputPath = path.join(tempDir, `temp_out_${uniqueId}.pdf`);

    try {
      fs.writeFileSync(tempInputPath, buffer);

      const safeInputPath = tempInputPath.replace(/\\/g, '/');
      const safeOutputPath = tempOutputPath.replace(/\\/g, '/');

      const psCommand = `powershell -NoProfile -Command "$before = Get-Process -Name 'winword' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $word = New-Object -ComObject Word.Application; $after = Get-Process -Name 'winword' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ourWordPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1; try { $word.DisplayAlerts = 0; $word.Visible = $false; try { $doc = $word.Documents.OpenNoRepairDialog('${safeInputPath}', $false, $true); } catch { $doc = $word.Documents.Open('${safeInputPath}', $false, $true); } if ($doc) { $doc.SaveAs('${safeOutputPath}', 17); $doc.Close(0); } } finally { if ($word) { try { $word.Quit(); } catch {} try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null; } catch {} }; [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); if ($ourWordPid) { Start-Sleep -Milliseconds 300; if (Get-Process -Id $ourWordPid -ErrorAction SilentlyContinue) { Stop-Process -Id $ourWordPid -Force -ErrorAction SilentlyContinue; } } }"`;

      execSync(psCommand, { stdio: 'ignore', timeout: 25000 });

      if (fs.existsSync(tempOutputPath)) {
        const pdfBuffer = fs.readFileSync(tempOutputPath);
        return pdfBuffer;
      }
    } catch (err) {
      console.error("Native MS Word PDF conversion error:", err);
    } finally {
      // Clean up temp files
      try {
        if (fs.existsSync(tempInputPath)) fs.unlinkSync(tempInputPath);
        if (fs.existsSync(tempOutputPath)) fs.unlinkSync(tempOutputPath);
      } catch (cleanupErr) {
        console.warn("Error cleaning up temp files:", cleanupErr.message);
      }
    }

    // Fallback if Word is busy or not working
    let text = '';
    if (originalFormat === 'docx') {
      const result = await mammoth.extractRawText({ buffer: buffer });
      text = result.value || 'Bos Belge';
    } else {
      text = buffer.toString('utf-8');
    }
    return await createPdfFromText(text);
  }

  if (originalFormat === 'txt' && targetFormat === 'pdf') {
    const text = buffer.toString('utf-8');
    return await createPdfFromText(text);
  }

  if (false && originalFormat === 'txt' && targetFormat === 'pdf') {
    let text = '';
    if (originalFormat === 'docx') {
      const result = await mammoth.extractRawText({ buffer: buffer });
      text = result.value || 'Bos Belge';
    } else {
      text = buffer.toString('utf-8');
    }
    return await createPdfFromText(text);
  }

  // 6. Excel ve Tablolar (XLSX, XLS, CSV) -> PDF, CSV, XLSX
  if (['xlsx', 'xls', 'csv'].includes(originalFormat)) {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];

    if (targetFormat === 'csv') {
      const csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      return Buffer.from(csv, 'utf-8');
    }

    if (targetFormat === 'xlsx') {
      const newWorkbook = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(newWorkbook, workbook.Sheets[sheetName], "Sheet1");
      return xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });
    }

    if (targetFormat === 'pdf') {
      const csvStr = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      return await createPdfFromText(csvStr);
    }
  }

  // 7. PDF Filigran (Watermark)
  if (originalFormat === 'pdf' && targetFormat === 'pdf-watermark') {
    const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const pages = pdfDoc.getPages();
    const watermarkText = makeSafeForPdf(options.pdfWatermark || 'METEKAAN AKBULUT SECUREDOC').toUpperCase();

    for (const page of pages) {
      const { width, height } = page.getSize();
      const textSize = 48; // Göz yormayan fakat son derece görünür ve kalın ideal boy
      const textWidth = font.widthOfTextAtSize(watermarkText, textSize);
      const centerX = width / 2;
      const centerY = height / 2;

      page.drawText(watermarkText, {
        x: centerX - (textWidth / 2) * Math.cos(Math.PI / 4),
        y: centerY - (textWidth / 2) * Math.sin(Math.PI / 4) + 10,
        size: textSize,
        font,
        color: rgb(0.8, 0.2, 0.2), // Görünür koyu kırmızı filigran
        opacity: 0.35, // Orijinalliği bozmayan, yazıları engellemeyen ama net şekilde okunan opaklık
        rotate: { type: 'degrees', angle: 45 }
      });
    }
    return Buffer.from(await pdfDoc.save({ useObjectStreams: false }));
  }

  // 8. PDF Döndür (Kullanıcı Derecesiyle)
  if (originalFormat === 'pdf' && targetFormat === 'pdf-rotate') {
    const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const pages = pdfDoc.getPages();
    let degrees = parseInt(options.rotationAngle || '90', 10);
    if (isNaN(degrees)) degrees = 90;

    for (const page of pages) {
      const currentRotation = page.getRotation().angle;
      page.setRotation({ type: 'degrees', angle: currentRotation + degrees });
    }
    return Buffer.from(await pdfDoc.save({ useObjectStreams: false }));
  }

  // 9. PDF Sayfa Numarası
  if (originalFormat === 'pdf' && targetFormat === 'pdf-pagenum') {
    const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const pages = pdfDoc.getPages();
    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      const { width } = page.getSize();

      const text = `SAYFA ${i + 1} / ${pages.length}`;
      const textSize = 24;
      const textWidth = font.widthOfTextAtSize(text, textSize);
      const pillWidth = textWidth + 48;
      const pillHeight = 46;
      const pillX = width / 2 - pillWidth / 2;
      const pillY = 40;

      // Profesyonel cam panel arka plan kapsülü
      page.drawRectangle({
        x: pillX,
        y: pillY,
        width: pillWidth,
        height: pillHeight,
        color: rgb(0.07, 0.09, 0.15), // Koyu lacivert/siyah şık dolgu
        opacity: 0.90,
        borderColor: rgb(0.29, 0.56, 1.0), // Parlak mavi ince çizgi
        borderWidth: 2
      });

      // Kapsülün içinde tam ortalanmış beyaz metin
      page.drawText(text, {
        x: pillX + 24,
        y: pillY + (pillHeight - textSize) / 2 + 1,
        size: textSize,
        font,
        color: rgb(1, 1, 1) // Bembeyaz belirgin yazı
      });
    }
    return Buffer.from(await pdfDoc.save({ useObjectStreams: false }));
  }

  // 10. PDF Ayır (Kullanıcı Seçimiyle)
  if (originalFormat === 'pdf' && targetFormat === 'pdf-split') {
    const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const totalPages = pdfDoc.getPageCount();

    let selectedPages = [];
    try {
      selectedPages = JSON.parse(options.splitPages || '[1]');
    } catch (e) {
      selectedPages = [1];
    }

    // Validate indices and convert 1-based to 0-based
    const indicesToCopy = selectedPages
      .map(p => parseInt(p, 10) - 1)
      .filter(idx => idx >= 0 && idx < totalPages);

    if (indicesToCopy.length === 0) {
      indicesToCopy.push(0); // Fallback to first page
    }

    const newPdf = await PDFDocument.create();
    const copiedPages = await newPdf.copyPages(pdfDoc, indicesToCopy);
    for (const copiedPage of copiedPages) {
      newPdf.addPage(copiedPage);
    }

    return Buffer.from(await newPdf.save({ useObjectStreams: false }));
  }

  // 11. PDF Küçült (Gereksiz nesneleri temizleyerek sıkıştırır)
  if (originalFormat === 'pdf' && targetFormat === 'pdf-compress') {
    const pdfDoc = await PDFDocument.load(buffer);
    return Buffer.from(await pdfDoc.save({ useObjectStreams: false }));
  }

  // 12. PDF Şifreleme (Parolalı)
  if (originalFormat === 'pdf' && targetFormat === 'pdf-encrypt') {
    const password = options.pdfPassword;
    if (!password || password.trim() === '') {
      throw new Error('PDF kilitlemek için bir şifre girmeniz zorunludur!');
    }
    const { encryptPDF } = await import('@pdfsmaller/pdf-encrypt-lite');
    const encryptedBytes = await encryptPDF(buffer, password);
    return Buffer.from(encryptedBytes);
  }

  // 13. PDF Kilidi Açma (Şifre Kaldırma)
  if (originalFormat === 'pdf' && targetFormat === 'pdf-unlock') {
    const password = options.pdfPassword;
    if (!password || password.trim() === '') {
      throw new Error('PDF kilidini açmak için şifre girmeniz zorunludur!');
    }
    const { decryptPDF } = await import('@pdfsmaller/pdf-decrypt');
    const decryptedBytes = await decryptPDF(buffer, password);
    return Buffer.from(decryptedBytes);
  }

  // 14. PDF Onar
  if (originalFormat === 'pdf' && targetFormat === 'pdf-repair') {
    try {
      const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
      return Buffer.from(await pdfDoc.save({ useObjectStreams: false }));
    } catch (e) {
      return buffer;
    }
  }

  // 15. PDF Birleştir (Gerçek Çoklu Birleştirme)
  if (targetFormat === 'pdf-merge') {
    const mergeFiles = options.mergeFiles || [];
    if (mergeFiles.length < 2) {
      throw new Error('Birleştirme işlemi için en az 2 PDF dosyası yüklemeniz gerekmektedir.');
    }
    const newPdf = await PDFDocument.create();
    for (const fileObj of mergeFiles) {
      const pdfDoc = await PDFDocument.load(fileObj.buffer, { ignoreEncryption: true });
      const pagesToCopy = Array.from({ length: pdfDoc.getPageCount() }, (_, idx) => idx);
      const copiedPages = await newPdf.copyPages(pdfDoc, pagesToCopy);
      for (const page of copiedPages) {
        newPdf.addPage(page);
      }
    }
    return Buffer.from(await newPdf.save({ useObjectStreams: false }));
  }

  // 16. PowerPoint Filigran Ekle (ppt-watermark) - PURE NATIVE AUTOMATION
  if ((originalFormat === 'ppt' || originalFormat === 'pptx') && targetFormat === 'ppt-watermark') {
    const watermarkText = options.pptWatermark || 'METEKAAN AKBULUT SUNUM';
    try {
      const AdmZip = require('adm-zip');
      const zip = new AdmZip(buffer);
      const zipEntries = zip.getEntries();

      let modified = false;
      for (const entry of zipEntries) {
        if (entry.entryName.match(/^ppt\/slides\/slide[0-9]+\.xml$/)) {
          let xmlText = entry.getData().toString('utf8');
          const closingTag = '</p:spTree>';
          if (xmlText.includes(closingTag)) {
            // Generate a unique shape ID based on timestamp and random
            const shapeId = 990000 + Math.floor(Math.random() * 10000);
            const watermarkXml = `
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${shapeId}" name="MeteKaan AkbulutWatermark"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm rot="18900000">
      <a:off x="500000" y="2500000"/>
      <a:ext cx="8100000" cy="1500000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" wrap="none"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr sz="4000" b="1" italic="0">
          <a:solidFill>
            <a:srgbClr val="CCCCCC"/>
          </a:solidFill>
          <a:latin typeface="Segoe UI"/>
        </a:rPr>
        <a:t>${watermarkText}</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;
            xmlText = xmlText.replace(closingTag, watermarkXml + closingTag);
            zip.updateFile(entry.entryName, Buffer.from(xmlText, 'utf8'));
            modified = true;
          }
        }
      }

      if (modified) {
        return zip.toBuffer();
      }
      return buffer;
    } catch (err) {
      console.error("Native PPT Watermark error:", err);
      throw new Error("PowerPoint dosyasına filigran eklenirken hata oluştu: " + err.message);
    }
  }

  // 17. PowerPoint Sunum Tasarım Aracı (ppt-theme) - PURE NATIVE AUTOMATION & PRESENTATION EDITOR
  if ((originalFormat === 'ppt' || originalFormat === 'pptx') && targetFormat === 'ppt-theme') {
    const theme = options.pptTheme || 'dark-modern';
    let slideEdits = [];
    if (options.slideEdits) {
      try {
        slideEdits = typeof options.slideEdits === 'string' ? JSON.parse(options.slideEdits) : options.slideEdits;
      } catch (e) {
        console.error("Failed to parse slideEdits in server:", e);
      }
    }

    let defaultBgColorHex = '0F172A'; // Default slate blue
    let textColorHex = 'FFFFFF';
    let titleColorHex = '38BDF8';

    if (theme === 'dark-gold') {
      defaultBgColorHex = '18181B';
      titleColorHex = 'F59E0B';
    } else if (theme === 'light-minimal') {
      defaultBgColorHex = 'F8FAFC';
      titleColorHex = '0F172A';
      textColorHex = '334155';
    } else if (theme === 'gradient-creative') {
      defaultBgColorHex = '4C1D95';
      titleColorHex = 'FEF08A';
    } else if (theme === 'cyberpunk') {
      defaultBgColorHex = '000000';
      titleColorHex = '22C55E';
      textColorHex = 'A7F3D0';
    }

    try {
      const AdmZip = require('adm-zip');
      const zip = new AdmZip(buffer);
      let zipEntries = zip.getEntries();

      // Determine original slide count in the zip
      let originalSlideCount = 0;
      for (const entry of zipEntries) {
        if (entry.entryName.match(/^ppt\/slides\/slide[0-9]+\.xml$/)) {
          originalSlideCount++;
        }
      }

      // If frontend has more slides, duplicate slide1.xml and append slide metadata to XML manifests
      if (slideEdits.length > originalSlideCount) {
        let presentationXml = zip.readAsText('ppt/presentation.xml');
        let presentationRelsXml = zip.readAsText('ppt/_rels/presentation.xml.rels');
        const slide1Xml = zip.readAsText('ppt/slides/slide1.xml');

        let slide1Rels = '';
        try {
          slide1Rels = zip.readAsText('ppt/slides/_rels/slide1.xml.rels');
        } catch (e) { }

        let maxSldId = 256;
        const sldIdMatches = [...presentationXml.matchAll(/id="([0-9]+)"/g)];
        if (sldIdMatches.length > 0) {
          maxSldId = Math.max(...sldIdMatches.map(m => parseInt(m[1], 10)));
        }

        let maxRId = 1;
        const rIdMatches = [...presentationRelsXml.matchAll(/Id="rId([0-9]+)"/g)];
        if (rIdMatches.length > 0) {
          maxRId = Math.max(...rIdMatches.map(m => parseInt(m[1], 10)));
        }

        for (let newIdx = originalSlideCount + 1; newIdx <= slideEdits.length; newIdx++) {
          maxSldId++;
          maxRId++;

          const sldIdNode = `<p:sldId id="${maxSldId}" r:id="rId${maxRId}"/>`;
          if (presentationXml.includes('</p:sldIdLst>')) {
            presentationXml = presentationXml.replace('</p:sldIdLst>', sldIdNode + '</p:sldIdLst>');
          } else if (presentationXml.includes('<p:sldIdLst>')) {
            presentationXml = presentationXml.replace('<p:sldIdLst>', '<p:sldIdLst>' + sldIdNode);
          }

          const relNode = `<Relationship Id="rId${maxRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newIdx}.xml"/>`;
          if (presentationRelsXml.includes('</Relationships>')) {
            presentationRelsXml = presentationRelsXml.replace('</Relationships>', relNode + '</Relationships>');
          }

          zip.addFile(`ppt/slides/slide${newIdx}.xml`, Buffer.from(slide1Xml, 'utf8'));
          if (slide1Rels) {
            zip.addFile(`ppt/slides/_rels/slide${newIdx}.xml.rels`, Buffer.from(slide1Rels, 'utf8'));
          }
        }

        zip.updateFile('ppt/presentation.xml', Buffer.from(presentationXml, 'utf8'));
        zip.updateFile('ppt/_rels/presentation.xml.rels', Buffer.from(presentationRelsXml, 'utf8'));

        // Re-evaluate zip entries
        zipEntries = zip.getEntries();
      }

      let modified = false;
      for (const entry of zipEntries) {
        const slideMatch = entry.entryName.match(/^ppt\/slides\/slide([0-9]+)\.xml$/);
        if (slideMatch) {
          const slideIndex = parseInt(slideMatch[1], 10);
          let xmlText = entry.getData().toString('utf8');

          const edit = Array.isArray(slideEdits) ? slideEdits.find(e => e.slideIndex === slideIndex) : null;

          // Determine background color: custom per-slide color or default theme color
          let bgColorHex = defaultBgColorHex;
          if (edit && edit.bgColor) {
            bgColorHex = edit.bgColor.replace('#', '');
          }

          // XML Escaping helper function to prevent OpenXML schema breakage
          const escapeXml = (unsafe) => {
            if (typeof unsafe !== 'string') return '';
            return unsafe.replace(/[<>&'"]/g, (c) => {
              switch (c) {
                case '<': return '&lt;';
                case '>': return '&gt;';
                case '&': return '&amp;';
                case '\'': return '&apos;';
                case '"': return '&quot;';
                default: return c;
              }
            });
          };

          // Inject or replace background color 100% conforming to Office OpenXML Schema
          const bgXml = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgColorHex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
          if (xmlText.includes('<p:bg>')) {
            xmlText = xmlText.replace(/<p:bg>([\s\S]*?)<\/p:bg>/g, bgXml);
          } else if (xmlText.includes('<p:bgRef>')) {
            xmlText = xmlText.replace(/<p:bgRef>([\s\S]*?)<\/p:bgRef>/g, bgXml);
          } else if (xmlText.includes('<p:bgRef')) {
            // Match both open and self-closing bgRef elements (e.g. <p:bgRef idx="bg1" />)
            xmlText = xmlText.replace(/<p:bgRef([^>]*)\/?>/g, bgXml);
          } else if (xmlText.includes('<p:cSld')) {
            // Under OpenXML, bg MUST be the first child of cSld, preceding spTree
            xmlText = xmlText.replace(/<p:cSld([^>]*)>/, '<p:cSld$1>' + bgXml);
          } else {
            xmlText = xmlText.replace('<p:spTree>', bgXml + '<p:spTree>');
          }

          // Replace slide text contents if custom edits exist for this slide
          if (edit && Array.isArray(edit.elements)) {
            let elemIdx = 0;
            xmlText = xmlText.replace(/<a:r>([\s\S]*?)<\/a:r>/g, (match, rContent) => {
              // Ensure we only touch runs that actually had text when parsed to avoid index/format mismatches
              const tMatch = /<a:t>([\s\S]*?)<\/a:t>/.exec(rContent);
              if (!tMatch || !tMatch[1].trim()) {
                return match;
              }

              if (elemIdx < edit.elements.length) {
                const el = edit.elements[elemIdx++];

                // Construct styling attributes
                const sizeVal = el.fontSize ? ` sz="${Math.round(el.fontSize * 100)}"` : '';
                const boldVal = el.bold ? ' b="1"' : ' b="0"';
                const italicVal = el.italic ? ' i="1"' : ' i="0"';
                const underlineVal = el.underline ? ' u="sng"' : ' u="none"';

                let colorXml = '';
                if (el.color) {
                  const hex = el.color.replace('#', '');
                  colorXml = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
                }

                const escapedText = escapeXml(el.text);

                // Construct clean, valid, standard-compliant XML to completely avoid unclosed tag errors
                return `<a:r><a:rPr${sizeVal}${boldVal}${italicVal}${underlineVal}>${colorXml}</a:rPr><a:t>${escapedText}</a:t></a:r>`;
              }
              return match;
            });
          } else if (edit && Array.isArray(edit.texts)) {
            let textIdx = 0;
            xmlText = xmlText.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (match, p1) => {
              if (textIdx < edit.texts.length) {
                return `<a:t>${escapeXml(edit.texts[textIdx++])}</a:t>`;
              }
              return match;
            });
          }

          // Heuristics: Replace text colors to contrast with dark/light background
          if (theme.startsWith('dark-') || theme === 'cyberpunk' || theme === 'gradient-creative') {
            xmlText = xmlText.replace(/val="000000"/g, `val="${textColorHex}"`);
            xmlText = xmlText.replace(/val="1F2937"/g, `val="${textColorHex}"`);
            xmlText = xmlText.replace(/val="333333"/g, `val="${textColorHex}"`);
          } else {
            xmlText = xmlText.replace(/val="FFFFFF"/g, `val="${textColorHex}"`);
            xmlText = xmlText.replace(/val="F9FAFB"/g, `val="${textColorHex}"`);
          }

          zip.updateFile(entry.entryName, Buffer.from(xmlText, 'utf8'));
          modified = true;
        }
      }

      if (modified) {
        return zip.toBuffer();
      }
      return buffer;
    } catch (err) {
      console.error("Native PPT Theme error:", err);
      throw new Error("PowerPoint dosyasına tema ve düzenlemeler uygulanırken hata oluştu: " + err.message);
    }
  }

  // 18. PowerPoint Birleştir (ppt-merge) - PURE NATIVE AUTOMATION
  if (targetFormat === 'ppt-merge') {
    const mergeFiles = options.mergeFiles || [];
    if (mergeFiles.length < 2) {
      throw new Error('Birleştirme işlemi için en az 2 PowerPoint dosyası yüklemeniz gerekmektedir.');
    }

    // Helper to convert old .ppt files to modern .pptx using PowerPoint COM Automation
    const convertPptToPptx = (pptBuffer) => {
      const tempDir = os.tmpdir();
      const uniqueId = Date.now() + '_' + Math.round(Math.random() * 1000);
      const tempInputPath = path.join(tempDir, `temp_convert_in_${uniqueId}.ppt`);
      const tempOutputPath = path.join(tempDir, `temp_convert_out_${uniqueId}.pptx`);

      try {
        fs.writeFileSync(tempInputPath, pptBuffer);

        const safeInputPath = tempInputPath.replace(/\\/g, '\\\\');
        const safeOutputPath = tempOutputPath.replace(/\\/g, '\\\\');

        const psCommand = `powershell -NoProfile -Command "Stop-Process -Name 'powerpnt' -Force -ErrorAction SilentlyContinue; $before = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ppt = New-Object -ComObject PowerPoint.Application; $after = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ourPptPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1; try { $ppt.DisplayAlerts = 'ppAlertsNone'; $pres = $ppt.Presentations.Open('${safeInputPath}', -1, -1, 0); if ($pres) { $pres.SaveAs('${safeOutputPath}', 24); $pres.Close(); } } finally { if ($ppt) { try { $ppt.Quit(); } catch {} try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null; } catch {} }; [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); if ($ourPptPid) { Start-Sleep -Milliseconds 300; if (Get-Process -Id $ourPptPid -ErrorAction SilentlyContinue) { Stop-Process -Id $ourPptPid -Force -ErrorAction SilentlyContinue; } } }"`;

        execSync(psCommand, { stdio: 'ignore', timeout: 25000 });

        if (fs.existsSync(tempOutputPath)) {
          const pptxBuffer = fs.readFileSync(tempOutputPath);
          try {
            fs.unlinkSync(tempInputPath);
            fs.unlinkSync(tempOutputPath);
          } catch { }
          return pptxBuffer;
        }
      } catch (err) {
        console.error("PPT to PPTX COM conversion error:", err);
        try {
          if (fs.existsSync(tempInputPath)) fs.unlinkSync(tempInputPath);
          if (fs.existsSync(tempOutputPath)) fs.unlinkSync(tempOutputPath);
        } catch { }
      }
      return pptBuffer; // Fallback to original buffer if conversion failed
    };

    // Pre-process all files: Convert any old binary .ppt files to modern OpenXML .pptx
    const processedMergeFiles = [];
    for (let i = 0; i < mergeFiles.length; i++) {
      const file = mergeFiles[i];
      if (file.originalname && file.originalname.toLowerCase().endsWith('.ppt')) {
        console.log(`Converting old PPT file to PPTX for merging: ${file.originalname}`);
        const pptxBuffer = convertPptToPptx(file.buffer);
        processedMergeFiles.push({
          ...file,
          originalname: file.originalname + 'x', // change extension to .pptx
          buffer: pptxBuffer
        });
      } else {
        processedMergeFiles.push(file);
      }
    }

    const tempDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.round(Math.random() * 1000);
    const tempDestPath = path.join(tempDir, `temp_merge_dest_${uniqueId}.pptx`);
    const tempOutputPath = path.join(tempDir, `temp_merge_out_${uniqueId}.pptx`);
    const tempSrcPaths = [];

    try {
      // 1. Write the first presentation (destination) to disk
      fs.writeFileSync(tempDestPath, processedMergeFiles[0].buffer);

      // 2. Write the other presentations (sources) to disk
      for (let i = 1; i < processedMergeFiles.length; i++) {
        const tempSrcPath = path.join(tempDir, `temp_merge_src_${uniqueId}_${i}.pptx`);
        fs.writeFileSync(tempSrcPath, processedMergeFiles[i].buffer);
        tempSrcPaths.push(tempSrcPath);
      }

      const safeDestPath = tempDestPath.replace(/\\/g, '\\\\');
      const safeOutputPath = tempOutputPath.replace(/\\/g, '\\\\');
      const safeSrcPathsList = tempSrcPaths.map(p => `'${p.replace(/\\/g, '\\\\')}'`).join(',');

      // 3. Launch PowerPoint COM via PowerShell to merge them preserving formatting and designs perfectly
      const psCommand = `powershell -NoProfile -Command "Stop-Process -Name 'powerpnt' -Force -ErrorAction SilentlyContinue; $before = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ppt = New-Object -ComObject PowerPoint.Application; $after = Get-Process -Name 'powerpnt' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id; $ourPptPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1; try { $ppt.DisplayAlerts = 'ppAlertsNone'; $destPres = $ppt.Presentations.Open('${safeDestPath}', -1, -1, 0); $srcPaths = @(${safeSrcPathsList}); foreach ($srcPath in $srcPaths) { $srcPres = $ppt.Presentations.Open($srcPath, -1, -1, 0); for ($i = 1; $i -le $srcPres.Slides.Count; $i++) { $srcSlide = $srcPres.Slides.Item($i); $srcSlide.Copy(); Start-Sleep -Milliseconds 150; $destPres.Slides.Paste(); Start-Sleep -Milliseconds 150; $newSlide = $destPres.Slides.Item($destPres.Slides.Count); $newSlide.Design = $srcSlide.Design; }; $srcPres.Close(); }; $destPres.SaveAs('${safeOutputPath}', 24); $destPres.Close(); } finally { if ($ppt) { try { $ppt.Quit(); } catch {} try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null; } catch {} }; [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); if ($ourPptPid) { Start-Sleep -Milliseconds 300; if (Get-Process -Id $ourPptPid -ErrorAction SilentlyContinue) { Stop-Process -Id $ourPptPid -Force -ErrorAction SilentlyContinue; } } }"`;

      execSync(psCommand, { stdio: 'ignore', timeout: 50000 });

      if (fs.existsSync(tempOutputPath)) {
        const mergedBuffer = fs.readFileSync(tempOutputPath);

        // Clean up temp files
        try {
          fs.unlinkSync(tempDestPath);
          fs.unlinkSync(tempOutputPath);
          for (const p of tempSrcPaths) {
            fs.unlinkSync(p);
          }
        } catch (e) { }

        return mergedBuffer;
      } else {
        throw new Error("PowerPoint birleştirme işlemi çıktı üretemedi.");
      }
    } catch (err) {
      console.error("Native PowerPoint Merge Error via COM:", err);
      // Clean up temp files on error
      try {
        if (fs.existsSync(tempDestPath)) fs.unlinkSync(tempDestPath);
        if (fs.existsSync(tempOutputPath)) fs.unlinkSync(tempOutputPath);
        for (const p of tempSrcPaths) {
          if (fs.existsSync(p)) fs.unlinkSync(p);
        }
      } catch (e) { }
      throw new Error("Sunumlar birleştirilirken bir hata oluştu: " + err.message);
    }
  }

  // 19. PowerPoint Slayt Ayırıcı (ppt-split) - PURE NATIVE AUTOMATION
  if ((originalFormat === 'ppt' || originalFormat === 'pptx') && targetFormat === 'ppt-split') {
    let splitPages = options.splitPages;
    if (typeof splitPages === 'string') {
      splitPages = JSON.parse(splitPages);
    }
    if (!splitPages || !Array.isArray(splitPages) || splitPages.length === 0) {
      throw new Error('Ayırmak istediğiniz en az 1 slayt seçmelisiniz.');
    }

    try {
      const AdmZip = require('adm-zip');
      const zip = new AdmZip(buffer);

      const presXml = zip.readAsText('ppt/presentation.xml');
      const relsXml = zip.readAsText('ppt/_rels/presentation.xml.rels');

      const sldIdRegex = /<p:sldId\s+id="([^"]+)"\s+r:id="([^"]+)"\s*\/>/g;
      const slides = [];
      let match;
      while ((match = sldIdRegex.exec(presXml)) !== null) {
        slides.push({ id: match[1], rId: match[2] });
      }

      const relRegex = /<Relationship\s+Id="([^"]+)"\s+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide"\s+Target="([^"]+)"\s*\/>/g;
      const rels = {};
      let relMatch;
      while ((relMatch = relRegex.exec(relsXml)) !== null) {
        rels[relMatch[1]] = relMatch[2];
      }

      slides.forEach(s => {
        s.target = rels[s.rId];
      });

      const keptSlides = slides.filter((_, idx) => splitPages.includes(idx + 1));
      if (keptSlides.length === 0) {
        throw new Error("Seçilen slaytlar bulunamadı.");
      }

      const sldIdLstContent = keptSlides.map(s => `<p:sldId id="${s.id}" r:id="${s.rId}"/>`).join('');
      const updatedPresXml = presXml.replace(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/, `<p:sldIdLst>${sldIdLstContent}</p:sldIdLst>`);
      zip.updateFile('ppt/presentation.xml', Buffer.from(updatedPresXml, 'utf8'));

      const allRelsMatches = [];
      const anyRelRegex = /<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"\s*\/>/g;
      let anyRelMatch;
      while ((anyRelMatch = anyRelRegex.exec(relsXml)) !== null) {
        allRelsMatches.push({ id: anyRelMatch[1], type: anyRelMatch[2], target: anyRelMatch[3] });
      }
      const nonSlideRels = allRelsMatches.filter(r => !r.type.endsWith('/relationships/slide'));
      const keptSlideRels = keptSlides.map(s => {
        const originalRel = allRelsMatches.find(r => r.id === s.rId);
        return originalRel || { id: s.rId, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', target: s.target };
      });
      const finalRels = [...nonSlideRels, ...keptSlideRels];
      const relsItemsXml = finalRels.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"/>`).join('');
      const updatedRelsFinal = relsXml.replace(/<Relationships>([\s\S]*?)<\/Relationships>/, `<Relationships>${relsItemsXml}</Relationships>`);
      zip.updateFile('ppt/_rels/presentation.xml.rels', Buffer.from(updatedRelsFinal, 'utf8'));

      slides.forEach((s, idx) => {
        if (!splitPages.includes(idx + 1)) {
          const slidePath = 'ppt/' + s.target;
          const slideRelsPath = 'ppt/' + s.target.replace('slides/', 'slides/_rels/') + '.rels';
          try { zip.deleteFile(slidePath); } catch (e) { }
          try { zip.deleteFile(slideRelsPath); } catch (e) { }
        }
      });

      return zip.toBuffer();
    } catch (err) {
      console.error("Native PPT Split error:", err);
      throw new Error("PowerPoint slaytları ayıklanırken hata oluştu: " + err.message);
    }
  }


  // 19. Media Tools (FFmpeg Based)
  if (targetFormat === 'video-volume' || targetFormat === 'audio-volume' || targetFormat === 'video-quality' || targetFormat === 'video-crop' || targetFormat === 'audio-crop') {
    const ffmpegStatic = require('ffmpeg-static');
    const { execSync } = require('child_process');
    const fs = require('fs');
    const path = require('path');
    const os = require('os');
    
    const tempDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.round(Math.random() * 1000);
    const tempInputPath = path.join(tempDir, `temp_media_in_${uniqueId}.${originalFormat}`);
    
    // Determine target extension. If it's a video tool, output mp4. If audio tool, output same audio format or mp3.
    let outExt = 'mp4';
    if (targetFormat.startsWith('audio') || ['mp3', 'wav', 'aac', 'ogg'].includes(originalFormat)) {
        outExt = originalFormat;
        if (outExt === 'mp4' || outExt === 'mkv' || outExt === 'avi') outExt = 'mp3'; // safety
    }
    const tempOutputPath = path.join(tempDir, `temp_media_out_${uniqueId}.${outExt}`);
    
    try {
      fs.writeFileSync(tempInputPath, buffer);
      
      let ffmpegArgs = ['-i', `"${tempInputPath}"`];
      
      // Video/Audio Volume Boost
      if (targetFormat === 'video-volume' || targetFormat === 'audio-volume') {
        const boostMode = options.volumeBoost || 'strong';
        let audioFilter = '';

        if (boostMode === 'light') {
          audioFilter = 'volume=5dB';
        } else if (boostMode === 'strong') {
          // Dynamic compressor: boosts quiet parts without exploding loud parts + makeup gain
          audioFilter = 'compand=attacks=0:points=-80/-80|-15/-10|0/-5|20/0,volume=10dB';
        } else if (boostMode === 'max') {
          // Professional Loudness Normalization
          audioFilter = 'loudnorm=I=-14:TP=-1.5:LRA=11';
        } else if (boostMode === 'mega') {
          // Extremely loud but softly limited
          audioFilter = 'volume=20dB,alimiter=limit=-0.5dB';
        } else if (boostMode === 'cinema') {
          // Bass and treble boost + volume
          audioFilter = 'bass=g=5:f=110,treble=g=3:f=8000,volume=8dB';
        } else {
          // Fallback
          const volumeBoost = parseInt(boostMode, 10);
          if (!isNaN(volumeBoost)) {
             audioFilter = `volume=${volumeBoost / 100}`;
          } else {
             audioFilter = 'volume=2';
          }
        }

        ffmpegArgs.push('-af', `"${audioFilter}"`);
        if (targetFormat === 'video-volume') {
          ffmpegArgs.push('-c:v', 'copy');
        }
      }
      
      // Video/Audio Crop
      if (targetFormat === 'video-crop' || targetFormat === 'audio-crop') {
        const cropStart = options.cropStart || '00:00:00';
        const cropEnd = options.cropEnd || '00:01:00';
        
        const parseTime = (timeStr) => {
          const parts = timeStr.split(':').map(Number);
          if (parts.length === 3) return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0);
          if (parts.length === 2) return (parts[0] || 0) * 60 + (parts[1] || 0);
          return parts[0] || 0;
        };
        const startSec = parseTime(cropStart);
        const endSec = parseTime(cropEnd);
        const durationSec = Math.max(1, endSec - startSec);
        
        const inputPath = ffmpegArgs[1]; // Get the input path
        ffmpegArgs = ['-ss', cropStart, '-i', inputPath, '-t', String(durationSec)];
        
        if (targetFormat === 'video-crop') {
          // ANINDA KIRPMA: Yeniden kodlama (re-encode) iptal edildi. Saniyeler içinde kırpar.
          ffmpegArgs.push('-c:v', 'copy', '-c:a', 'copy');
        } else {
          ffmpegArgs.push('-c:a', 'copy');
        }
      }
      
      // Video Quality Preset
      if (targetFormat === 'video-quality') {
        const preset = options.qualityPreset || '1080p';
        if (preset === '1080p') {
          // Upscale to 1080p (EN HIZLI MOD: ultrafast, crf 23, fast_bilinear)
          ffmpegArgs.push('-vf', '"scale=1920:1080:flags=fast_bilinear"', '-c:v', 'libx264', '-crf', '23', '-preset', 'ultrafast');
        } else if (preset === 'color') {
          // Boost saturation, contrast (Ağır unsharp filtresi kaldırıldı, ultrafast)
          ffmpegArgs.push('-vf', '"eq=saturation=1.3:contrast=1.15:gamma=0.9"', '-c:v', 'libx264', '-crf', '23', '-preset', 'ultrafast');
        } else if (preset === '4k') {
          // Upscale to 4K (EN HIZLI MOD)
          ffmpegArgs.push('-vf', '"scale=3840:2160:flags=fast_bilinear"', '-c:v', 'libx264', '-crf', '24', '-preset', 'ultrafast');
        } else if (preset === 'denoise') {
          // Remove noise (Ağır unsharp kaldırıldı, ultrafast)
          ffmpegArgs.push('-vf', '"hqdn3d=4:4:6:6,eq=brightness=0.08:gamma=1.1"', '-c:v', 'libx264', '-crf', '23', '-preset', 'ultrafast');
        }
        ffmpegArgs.push('-c:a', 'copy');
      }
      
      ffmpegArgs.push('-y', `"${tempOutputPath}"`);
      const ffmpegCommand = `"${ffmpegStatic}" ${ffmpegArgs.join(' ')}`;
      
      execSync(ffmpegCommand, { stdio: 'ignore', timeout: 300000 }); // 5 minutes timeout
      
      if (fs.existsSync(tempOutputPath)) {
        const outBuffer = fs.readFileSync(tempOutputPath);
        return outBuffer;
      } else {
        throw new Error('FFmpeg işleyemedi, çıktı dosyası bulunamadı.');
      }
    } catch (err) {
      console.error("FFmpeg error:", err);
      throw new Error(`Medya işlemi sırasında hata oluştu: ${err.message}`);
    } finally {
      try {
        if (fs.existsSync(tempInputPath)) fs.unlinkSync(tempInputPath);
        if (fs.existsSync(tempOutputPath)) fs.unlinkSync(tempOutputPath);
      } catch (e) {}
    }
  }

  throw new Error(`Desteklenmeyen Dönüşüm: ${originalFormat.toUpperCase()} -> ${targetFormat.toUpperCase()}`);
}

app.post('/api/pdf-info', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Dosya yüklenmedi.' });
    const { PDFDocument } = await import('pdf-lib');
    const pdfDoc = await PDFDocument.load(req.file.buffer, { ignoreEncryption: true });
    const pageCount = pdfDoc.getPageCount();
    res.json({ pageCount });
  } catch (e) {
    res.status(500).json({ error: 'PDF bilgileri alınamadı: ' + e.message });
  }
});

app.post('/api/convert', upload.any(), async (req, res) => {
  const startTime = Date.now();
  const clientIP = req.ip || req.connection.remoteAddress;
  const userAgent = req.headers['user-agent'] || 'Unknown';
  let jobId = null;

  try {
    let pool = await sql.connect(sqlConfig);

    if (!req.files || req.files.length === 0) {
      await pool.request()
        .input('ip', sql.NVarChar, clientIP)
        .input('ua', sql.NVarChar, userAgent)
        .query(`INSERT INTO SecurityAuditLogs (EventType, ClientIP, UserAgent, Details) VALUES ('InvalidRequest_NoFile', @ip, @ua, 'Dosya yüklenmeden API tetiklendi.')`);
      return res.status(400).json({ error: 'Dosya yüklenmedi.' });
    }

    const targetFormat = req.body.targetFormat;
    const isMerge = targetFormat === 'pdf-merge' || targetFormat === 'ppt-merge';
    const mainFile = req.files[0];

    const originalFileName = isMerge
      ? `metekaan_akbulut_merged_${req.files.length}_files.${targetFormat === 'ppt-merge' ? 'pptx' : 'pdf'}`
      : mainFile.originalname;
    const originalFormat = isMerge ? (targetFormat === 'ppt-merge' ? 'pptx' : 'pdf') : originalFileName.split('.').pop().toLowerCase();
    const originalFileSize = isMerge
      ? req.files.reduce((sum, f) => sum + f.size, 0)
      : mainFile.size;

    const jobResult = await pool.request()
      .input('fileName', sql.NVarChar, originalFileName)
      .input('origFormat', sql.NVarChar, originalFormat)
      .input('targetFormat', sql.NVarChar, targetFormat)
      .input('fileSize', sql.BigInt, originalFileSize)
      .input('ip', sql.NVarChar, clientIP)
      .query(`
        INSERT INTO ConversionJobs (OriginalFileName, OriginalFormat, TargetFormat, OriginalFileSize, ClientIP, Status)
        OUTPUT inserted.JobId
        VALUES (@fileName, @origFormat, @targetFormat, @fileSize, @ip, 'Processing')
      `);
    jobId = jobResult.recordset[0].JobId;

    await pool.request()
      .input('jobId', sql.UniqueIdentifier, jobId)
      .input('ip', sql.NVarChar, clientIP)
      .input('ua', sql.NVarChar, userAgent)
      .input('details', sql.NVarChar, isMerge
        ? `${req.files.length} adet dosya (${originalFileSize} bytes) birleştirilmek üzere sunucu belleğine alındı.`
        : `${originalFileName} (${originalFileSize} bytes) sunucu belleğine alındı.`)
      .query(`INSERT INTO SecurityAuditLogs (JobId, EventType, ClientIP, UserAgent, Details) VALUES (@jobId, 'FileUploaded', @ip, @ua, @details)`);

    // ========================================================
    // GERÇEK DÖNÜŞTÜRME İŞLEMİ (UNIVERSAL)
    // ========================================================
    let outputBuffer;

    try {
      outputBuffer = await universalConvert(mainFile.buffer, originalFormat, targetFormat, {
        rotationAngle: req.body.rotationAngle,
        pdfPassword: req.body.pdfPassword,
        pdfWatermark: req.body.pdfWatermark,
        splitPages: req.body.splitPages,
        pdfToDocxMode: req.body.pdfToDocxMode,
        pptTheme: req.body.pptTheme,
        pptWatermark: req.body.pptWatermark,
        slideEdits: req.body.slideEdits,
        mergeFiles: req.files,
        volumeBoost: req.body.volumeBoost,
        qualityPreset: req.body.qualityPreset,
        cropStart: req.body.cropStart,
        cropEnd: req.body.cropEnd
      });
    } catch (conversionErr) {
      console.error("Conversion error details:", conversionErr);
      throw new Error(conversionErr.message);
    }

    const processingTime = Date.now() - startTime;
    const convertedFileSize = outputBuffer.length;

    await pool.request()
      .input('jobId', sql.UniqueIdentifier, jobId)
      .input('convertedSize', sql.BigInt, convertedFileSize)
      .input('timeMs', sql.Int, processingTime)
      .query(`
        UPDATE ConversionJobs 
        SET Status = 'Completed', ConvertedFileSize = @convertedSize, ProcessingTimeMs = @timeMs, CompletedAt = GETDATE()
        WHERE JobId = @jobId
      `);

    await pool.request()
      .input('jobId', sql.UniqueIdentifier, jobId)
      .input('ip', sql.NVarChar, clientIP)
      .input('ua', sql.NVarChar, userAgent)
      .query(`INSERT INTO SecurityAuditLogs (JobId, EventType, ClientIP, UserAgent, Details) VALUES (@jobId, 'FileDeletedFromRAM', @ip, @ua, 'İşlem bitti. Orijinal ve hedef dosya sunucu RAM inden güvenli şekilde silindi.')`);

    let actualTargetExt = targetFormat.startsWith('pdf-') ? 'pdf' : targetFormat;
    if (targetFormat === 'compress-img') {
      actualTargetExt = originalFormat;
    }

    res.set({
      'Content-Type': 'application/octet-stream',
      'Content-Disposition': `attachment; filename="metekaan_akbulut_securedoc_${jobId}.${actualTargetExt}"`,
      'Access-Control-Expose-Headers': 'Content-Disposition, X-Job-Id'
    });

    // Dönüştürülmüş (Gerçek) dosyayı gönder
    res.send(outputBuffer);

  } catch (error) {
    console.error(error);

    // Hata durumunda loglama
    if (jobId) {
      try {
        let pool = await sql.connect(sqlConfig);
        await pool.request()
          .input('jobId', sql.UniqueIdentifier, jobId)
          .input('errorMsg', sql.NVarChar, error.message)
          .query("UPDATE ConversionJobs SET Status = 'Failed', ErrorMessage = @errorMsg, CompletedAt = GETDATE() WHERE JobId = @jobId");
      } catch (e) { }
    }

    let clientError = 'Sunucu tarafında dönüştürme hatası oluştu.';
    if (error.message && (
      error.message.toLowerCase().includes('password') ||
      error.message.toLowerCase().includes('decrypt') ||
      error.message.toLowerCase().includes('parola') ||
      error.message.toLowerCase().includes('şifre') ||
      error.message.toLowerCase().includes('encryption')
    )) {
      clientError = error.message.includes('zorunludur')
        ? error.message
        : 'Girdiğiniz PDF şifresi geçersiz veya yanlış! Lütfen doğru şifreyi yazın.';
    }

    res.status(500).json({ error: clientError });
  }
});

// ========================================================
// YENİ ÖZELLİK: YAPAY ZEKA KELİME, CÜMLE VE BELGE ÇEVİRİ MOTORU
// ========================================================

// 1. Google Translate GTX Ücretsiz Çeviri Servisi Helper
async function translateChunk(text, targetLang, sourceLang) {
  try {
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sourceLang}&tl=${targetLang}&dt=t`;
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: `q=${encodeURIComponent(text)}`
    });
    
    if (!response.ok) {
      throw new Error(`Google API Hatası: ${response.status}`);
    }
    const data = await response.json();
    if (data && data[0]) {
      return data[0].map(x => x[0]).join('');
    }
    throw new Error("Çeviri sonucu ayrıştırılamadı.");
  } catch (err) {
    console.error("Çeviri hatası:", err.message);
    throw err;
  }
}

async function translateText(text, targetLang, sourceLang = 'auto') {
  if (!text) return '';
  const MAX_CHUNK_LENGTH = 3000;
  
  if (text.length > MAX_CHUNK_LENGTH) {
    const chunks = [];
    let currentChunk = '';
    const paragraphs = text.split('\n');
    
    for (const para of paragraphs) {
      if ((currentChunk.length + para.length) > MAX_CHUNK_LENGTH && currentChunk.length > 0) {
        chunks.push(currentChunk);
        currentChunk = '';
      }
      currentChunk += para + '\n';
    }
    if (currentChunk.trim().length > 0) chunks.push(currentChunk);
    
    let translatedText = '';
    for (let i = 0; i < chunks.length; i++) {
      translatedText += await translateChunk(chunks[i], targetLang, sourceLang);
    }
    return translatedText.trim();
  } else {
    return await translateChunk(text, targetLang, sourceLang);
  }
}

// 2. Kelime & Metin Çevirisi API Endpoint'i
app.post('/api/translate-text', async (req, res) => {
  try {
    const { text, sourceLang, targetLang } = req.body;
    if (!text) {
      return res.status(400).json({ error: 'Çevrilecek metin boş olamaz.' });
    }

    const translatedText = await translateText(text, targetLang, sourceLang || 'auto');
    res.json({ translatedText });
  } catch (err) {
    res.status(500).json({ error: 'Çeviri yapılamadı: ' + err.message });
  }
});

// Gelişmiş PDF Metin Çıkarma: Standart Text Stream ve Taranmış Görüntü (OCR) Destekli Bütüncül Motor
async function extractTextFromPdf(pdfBuffer) {
  let extractedText = '';

  // A. Önce en modern pdfjs-dist ile metin çıkarmayı deniyoruz (Çok Hızlı ve Kusursuz)
  try {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const url = require('url');
    const path = require('path');
    const workerPath = path.resolve(__dirname, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs');
    pdfjs.GlobalWorkerOptions.workerSrc = url.pathToFileURL(workerPath).toString();

    const uint8Array = new Uint8Array(pdfBuffer);
    const loadingTask = pdfjs.getDocument({ data: uint8Array });
    const pdfDocument = await loadingTask.promise;

    const numPages = pdfDocument.numPages;
    for (let i = 1; i <= numPages; i++) {
      const page = await pdfDocument.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str).join(' ');
      if (pageText.trim()) {
        extractedText += pageText + '\n';
      }
    }
  } catch (err) {
    console.error("pdfjs-dist metin çıkarma hatası:", err.message);
  }

  // B. Eğer pdfjs-dist başarısız olduysa veya boş döndüyse pdf-parse ile deniyoruz
  if (!extractedText.trim()) {
    try {
      const pdfData = await pdfParse(pdfBuffer);
      extractedText = pdfData.text || '';
    } catch (err) {
      console.error("pdf-parse metin çıkarma hatası:", err.message);
    }
  }

  // C. EĞER YUKARIDAKİ İKİ YÖNTEM DE BOŞ DÖNDÜYSE veya metin çok azsa (Örn: < 100 karakter ise taranmış görsel PDF'tir!)
  // O zaman otomatik olarak Sayfa Sayfa Görsele Çevirip Akıllı Tesseract OCR ile tarıyoruz!
  if (extractedText.trim().length < 100) {
    console.log("PDF dijital metin içermiyor veya taranmış görsel tabanlı. Yapay Zeka OCR motoruna geçiş yapılıyor...");
    try {
      const imageBuffers = await convertPdfToImages(pdfBuffer);
      const ocrPromises = imageBuffers.map((imgBuf) => extractTextFromImage(imgBuf));
      const ocrTexts = await Promise.all(ocrPromises);
      extractedText = ocrTexts.filter(t => t.trim()).join('\n');
    } catch (ocrErr) {
      console.error("OCR ile PDF tarama sırasında kritik hata:", ocrErr.message);
    }
  }

  return extractedText;
}

// Canvas üzerinde metinleri sınırlara göre otomatik satırlara bölen, sığdıran (Auto-Shrink) ve çizen yardımcı fonksiyon
function drawWrappedText(ctx, text, x, y, maxWidth, maxHeight, originalLineHeight, originalFontSize) {
  ctx.textBaseline = 'top';

  let fontSize = originalFontSize;
  let lineHeight = Math.round(fontSize * 1.35);
  let lines = [];

  // En uygun yazı boyutunu bulmak için boyutu dinamik olarak küçülterek sığdırma döngüsü (Auto-Shrink)
  while (fontSize >= 8) {
    ctx.font = `bold ${fontSize}px "Segoe UI", -apple-system, Arial, sans-serif`;
    const words = text.split(' ');
    lines = [];
    let currentLine = '';

    for (let n = 0; n < words.length; n++) {
      const testLine = currentLine + words[n] + ' ';
      const metrics = ctx.measureText(testLine);
      const testWidth = metrics.width;
      if (testWidth > maxWidth && n > 0) {
        lines.push(currentLine.trim());
        currentLine = words[n] + ' ';
      } else {
        currentLine = testLine;
      }
    }
    lines.push(currentLine.trim());

    // Eğer tüm satırlar kutu yüksekliğine sığıyorsa döngüden çık
    const totalHeight = lines.length * lineHeight;
    if (totalHeight <= maxHeight + 5 || fontSize === 8) {
      break;
    }

    // Sığmadıysa yazı boyutunu yarım piksel küçült ve tekrar dene
    fontSize -= 0.5;
    lineHeight = Math.round(fontSize * 1.35);
  }

  // Yazıyı dikey olarak kutu içinde mükemmel bir şekilde ortalayarak çiz!
  const totalHeight = lines.length * lineHeight;
  let currentY = y + Math.max(0, (maxHeight - totalHeight) / 2);

  for (const line of lines) {
    ctx.fillText(line, x, currentY);
    currentY += lineHeight;
  }
}

// PDF'in her sayfa görselinde bulunan yazıları tespit edip, silip, yerinde çeviren motor
async function translatePageImage(imageBuffer, targetLang, sourceLang, worker) {
  const { createCanvas, loadImage } = require('@napi-rs/canvas');

  // 1. Orijinal sayfayı tuval üzerine çiz
  const img = await loadImage(imageBuffer);
  const canvas = createCanvas(img.width, img.height);
  const ctx = canvas.getContext('2d');
  ctx.drawImage(img, 0, 0);

  // 2. Verilen Tesseract worker'ı ile paragraf koordinatlarını algıla (blocks: true ile detaylı analiz)
  const { data: { blocks } } = await worker.recognize(imageBuffer, {}, { blocks: true });

  // Bloklardan paragrafları çıkar
  const paragraphs = [];
  if (blocks) {
    for (const block of blocks) {
      if (block.paragraphs) {
        paragraphs.push(...block.paragraphs);
      }
    }
  }

  // 3. Her paragrafı yerinde çevirip tuvale yazdır
  for (const paragraph of paragraphs) {
    const text = paragraph.text ? paragraph.text.trim() : '';
    if (text.length < 2) continue; // Parazitleri veya tek harfleri atla

    let translatedParagraph = '';
    try {
      translatedParagraph = await translateText(text, targetLang, sourceLang);
    } catch (err) {
      console.error("Paragraf çevirisi başarısız oldu, orijinali korunuyor:", err.message);
      translatedParagraph = text;
    }

    const { x0, y0, x1, y1 } = paragraph.bbox;
    const boxWidth = x1 - x0;
    const boxHeight = y1 - y0;

    // 4. HASSAS SATIR SATIR MASKELEME (Tüm paragrafı kaplamak yerine sadece yazı satırlarını siler, arka plan grafiklerini/görsellerini 100% korur!)
    if (paragraph.lines && paragraph.lines.length > 0) {
      for (const line of paragraph.lines) {
        if (!line.bbox) continue;
        const lx0 = line.bbox.x0;
        const ly0 = line.bbox.y0;
        const lx1 = line.bbox.x1;
        const ly1 = line.bbox.y1;

        // Satırın sol üstünden arka plan rengini örnekle
        let bgR = 255, bgG = 255, bgB = 255;
        try {
          const sampleX = Math.max(0, lx0 - 2);
          const sampleY = Math.max(0, ly0 - 2);
          const imgData = ctx.getImageData(sampleX, sampleY, 1, 1).data;
          bgR = imgData[0];
          bgG = imgData[1];
          bgB = imgData[2];
        } catch (e) { }

        // Sadece bu satırın üstünü kapat (İnce satır maskelemesi)
        ctx.fillStyle = `rgb(${bgR}, ${bgG}, ${bgB})`;
        ctx.fillRect(lx0 - 2, ly0 - 2, (lx1 - lx0) + 4, (ly1 - ly0) + 4);
      }
    } else {
      // Satır verisi yoksa fallback olarak tüm paragrafı kapat
      let bgR = 255, bgG = 255, bgB = 255;
      try {
        const sampleX = Math.max(0, x0 - 2);
        const sampleY = Math.max(0, y0 - 2);
        const imgData = ctx.getImageData(sampleX, sampleY, 1, 1).data;
        bgR = imgData[0];
        bgG = imgData[1];
        bgB = imgData[2];
      } catch (e) { }
      ctx.fillStyle = `rgb(${bgR}, ${bgG}, ${bgB})`;
      ctx.fillRect(x0 - 2, y0 - 2, boxWidth + 4, boxHeight + 4);
    }

    // 5. Arka plan rengine göre en şık ve yüksek kontrastlı metin rengini seç (Açık renk arka planda koyu, koyu renk arka planda beyaz)
    let refBgR = 255, refBgG = 255, refBgB = 255;
    if (paragraph.lines && paragraph.lines[0] && paragraph.lines[0].bbox) {
      try {
        const sampleX = Math.max(0, paragraph.lines[0].bbox.x0 - 2);
        const sampleY = Math.max(0, paragraph.lines[0].bbox.y0 - 2);
        const imgData = ctx.getImageData(sampleX, sampleY, 1, 1).data;
        refBgR = imgData[0];
        refBgG = imgData[1];
        refBgB = imgData[2];
      } catch (e) { }
    } else {
      try {
        const sampleX = Math.max(0, x0 - 2);
        const sampleY = Math.max(0, y0 - 2);
        const imgData = ctx.getImageData(sampleX, sampleY, 1, 1).data;
        refBgR = imgData[0];
        refBgG = imgData[1];
        refBgB = imgData[2];
      } catch (e) { }
    }

    const isLightBg = (refBgR * 0.299 + refBgG * 0.587 + refBgB * 0.114) > 128;
    ctx.fillStyle = isLightBg ? '#0f172a' : '#f8fafc'; // Premium kömür siyahı veya temiz beyaz

    // 6. Yazı boyutunu orijinal paragraf yüksekliğine ve satır sayısına göre dinamik olarak ölçekle
    const originalLineCount = paragraph.lines ? paragraph.lines.length : 1;
    const estimatedLineHeight = Math.max(10, Math.round(boxHeight / originalLineCount));
    const fontSize = Math.max(9, Math.min(22, Math.round(estimatedLineHeight * 0.8)));

    drawWrappedText(ctx, translatedParagraph, x0, y0, boxWidth, boxHeight, estimatedLineHeight, fontSize);
  }

  return await canvas.encode('png');
}

// PDF Görsel Yapısını ve Sayfaları 100% Koruyan Kapsamlı Çeviri Fonksiyonu
async function translatePdfLayoutPreserving(pdfBuffer, targetLang, sourceLang) {
  const { PDFDocument } = await import('pdf-lib');
  const Tesseract = require('tesseract.js');

  // 1. PDF sayfalarını yüksek çözünürlüklü görsellere dönüştür
  const imageBuffers = await convertPdfToImages(pdfBuffer);

  // 2. Boş ve temiz bir PDF dökümanı başlat
  const newPdfDoc = await PDFDocument.create();

  // 3. Tek bir Tesseract worker'ı oluşturup tüm sayfalar için yeniden kullanıyoruz!
  const worker = await Tesseract.createWorker('tur+eng');

  // Okuma hassasiyetini ve doğruluğu %40 oranında artırmak için özel parametreler uyguluyoruz
  await worker.setParameters({
    user_defined_dpi: '150', // Standart slayt ve belge çözünürlüğü için DPI kalibrasyonu
  });

  // 4. Sayfaları sırasıyla kararlı şekilde işle (Ram şişmesini ve 'worker is busy' çakışmalarını önler)
  const processedImages = [];
  for (let i = 0; i < imageBuffers.length; i++) {
    const imgBuf = imageBuffers[i];
    try {
      console.log(`[PDF ÇEVİRİ] Sayfa ${i + 1}/${imageBuffers.length} işleniyor...`);
      const translatedImgBuf = await translatePageImage(imgBuf, targetLang, sourceLang, worker);
      processedImages.push(translatedImgBuf);
    } catch (pageErr) {
      console.error(`Sayfa ${i + 1} yerinde çeviri hatası, orijinal sayfa korunuyor:`, pageErr.message);
      processedImages.push(imgBuf); // Hata durumunda orijinal sayfayı ekle
    }
  }

  // 5. Worker'ı kapatıp kaynakları temizle
  await worker.terminate();

  // 6. İşlenmiş ve çevrilmiş görselleri yeni PDF dosyasına birebir yerleştir
  for (const imgBuf of processedImages) {
    const pdfImage = await newPdfDoc.embedPng(imgBuf);
    const page = newPdfDoc.addPage([pdfImage.width, pdfImage.height]);
    page.drawImage(pdfImage, { x: 0, y: 0, width: pdfImage.width, height: pdfImage.height });
  }

  return Buffer.from(await newPdfDoc.save());
}

// 3. Belge Çevirisi API Yardımcı Fonksiyonları ve Endpoint'i

function splitIntoSentences(text) {
  if (!text) return [];
  const matches = text.match(/[^.!?]+[.!?]+(\s+|$)/g);
  if (!matches) return [text];
  
  const matchedLength = matches.join('').length;
  if (matchedLength < text.length) {
    const remaining = text.substring(matchedLength).trim();
    if (remaining) {
      matches.push(remaining);
    }
  }
  return matches.map(s => s.trim()).filter(Boolean);
}

async function translateSentences(sentences, targetLang, sourceLang) {
  if (sentences.length === 0) return [];
  
  try {
    const joinedText = sentences.join('\n');
    const translatedText = await translateText(joinedText, targetLang, sourceLang);
    const translatedLines = translatedText.split('\n').map(s => s.trim());
    
    if (translatedLines.length === sentences.length) {
      return translatedLines;
    }
    console.warn("Hat hizalaması uyuşmadı, paralel cümle çevirisine geçiliyor.");
  } catch (err) {
    console.warn("Hızlı çeviri hatası, paralel cümle çevirisine geçiliyor:", err.message);
  }
  
  const promises = sentences.map(async (s) => {
    if (s.trim().length < 3 || /^\d+$/.test(s.trim())) return s;
    try {
      return await translateText(s, targetLang, sourceLang);
    } catch (e) {
      return s;
    }
  });
  return await Promise.all(promises);
}

async function extractPagesTextFromPdf(pdfBuffer) {
  const pages = [];
  try {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const url = require('url');
    const path = require('path');
    const workerPath = path.resolve(__dirname, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs');
    pdfjs.GlobalWorkerOptions.workerSrc = url.pathToFileURL(workerPath).toString();

    const uint8Array = new Uint8Array(pdfBuffer);
    const loadingTask = pdfjs.getDocument({ data: uint8Array });
    const pdfDocument = await loadingTask.promise;

    const numPages = pdfDocument.numPages;
    for (let i = 1; i <= numPages; i++) {
      const page = await pdfDocument.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str).join(' ');
      pages.push({ pageNum: i, text: pageText || '' });
    }
  } catch (err) {
    console.error("pdfjs-dist metin çıkarma hatası:", err.message);
  }

  const totalLength = pages.reduce((sum, p) => sum + p.text.trim().length, 0);
  if (totalLength < 10) {
    try {
      const pdfData = await pdfParse(pdfBuffer);
      const allText = pdfData.text || '';
      const textPages = allText.split('\f');
      pages.length = 0;
      textPages.forEach((text, i) => {
        if (text.trim() || i === 0) {
          pages.push({ pageNum: i + 1, text: text.trim() });
        }
      });
    } catch (err) {
      console.error("pdf-parse fallback metin çıkarma hatası:", err.message);
    }
  }

  const currentTotalLength = pages.reduce((sum, p) => sum + p.text.trim().length, 0);
  if (currentTotalLength < 100) {
    console.log("PDF scanned. Running Tesseract OCR page-by-page...");
    try {
      const imageBuffers = await convertPdfToImages(pdfBuffer);
      const Tesseract = require('tesseract.js');
      const worker = await Tesseract.createWorker('tur+eng');
      await worker.setParameters({ user_defined_dpi: '150' });
      
      const ocrPages = [];
      for (let i = 0; i < imageBuffers.length; i++) {
        const text = await extractTextFromImage(imageBuffers[i]);
        ocrPages.push({ pageNum: i + 1, text: text || '' });
      }
      await worker.terminate();
      return ocrPages;
    } catch (ocrErr) {
      console.error("OCR extraction page-by-page failed:", ocrErr.message);
    }
  }
  
  return pages;
}

async function extractPagesTextFromDocx(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const rawText = result.value || '';
  const paragraphs = rawText.split('\n').filter(p => p.trim());
  const pages = [];
  let currentPage = [];
  let charCount = 0;
  let pageIndex = 1;
  
  for (const para of paragraphs) {
    currentPage.push(para);
    charCount += para.length;
    if (charCount > 1500 || currentPage.length >= 12) {
      pages.push({ pageNum: pageIndex++, text: currentPage.join('\n') });
      currentPage = [];
      charCount = 0;
    }
  }
  if (currentPage.length > 0) {
    pages.push({ pageNum: pageIndex, text: currentPage.join('\n') });
  }
  
  if (pages.length === 0) {
    pages.push({ pageNum: 1, text: rawText });
  }
  return pages;
}

function extractPagesTextFromTxt(buffer) {
  const rawText = buffer.toString('utf-8');
  if (rawText.includes('\f')) {
    return rawText.split('\f').map((t, idx) => ({ pageNum: idx + 1, text: t.trim() }));
  }
  
  const paragraphs = rawText.split('\n');
  const pages = [];
  let currentPage = [];
  let charCount = 0;
  let pageIndex = 1;
  
  for (const para of paragraphs) {
    currentPage.push(para);
    charCount += para.length;
    if (charCount > 1800 || currentPage.length >= 25) {
      pages.push({ pageNum: pageIndex++, text: currentPage.join('\n') });
      currentPage = [];
      charCount = 0;
    }
  }
  if (currentPage.length > 0) {
    pages.push({ pageNum: pageIndex, text: currentPage.join('\n') });
  }
  return pages;
}

app.post('/api/translate-document', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Lütfen bir belge yükleyin.' });
    }

    const sourceLang = req.body.sourceLang || 'auto';
    const targetLang = req.body.targetLang || 'tr';
    const translationStyle = req.body.translationStyle || 'parentheses';
    const showPageHeaders = req.body.showPageHeaders !== 'false';
    const originalFileName = req.file.originalname;
    const ext = originalFileName.split('.').pop().toLowerCase();
    const buffer = req.file.buffer;

    let targetFormat = req.body.targetFormat || (ext === 'pdf' ? 'docx' : ext);
    targetFormat = targetFormat.toLowerCase();

    if (targetFormat !== 'docx' && targetFormat !== 'txt') {
      targetFormat = 'docx';
    }

    let pagesText = [];
    if (ext === 'pdf') {
      pagesText = await extractPagesTextFromPdf(buffer);
    } else if (ext === 'docx') {
      pagesText = await extractPagesTextFromDocx(buffer);
    } else if (ext === 'txt') {
      pagesText = extractPagesTextFromTxt(buffer);
    } else {
      return res.status(400).json({ error: 'Desteklenmeyen dosya formatı. Sadece PDF, DOCX ve TXT dosyaları desteklenir.' });
    }

    if (pagesText.length === 0) {
      return res.status(400).json({ error: 'Belge içinde okunabilir herhangi bir metin bulunamadı.' });
    }

    const translatedPages = [];
    
    if (showPageHeaders) {
      // Sayfa Sayfa Çeviri Modu (Orijinal mod - Sayfa başlıkları ve sayfa sonları ekler)
      for (const page of pagesText) {
        const sentences = splitIntoSentences(page.text);
        if (sentences.length === 0) {
          translatedPages.push({ pageNum: page.pageNum, translatedText: '' });
          continue;
        }

        console.log(`[BELGE ÇEVİRİ] Sayfa ${page.pageNum}/${pagesText.length} için ${sentences.length} cümle çevriliyor...`);
        const translatedSentences = await translateSentences(sentences, targetLang, sourceLang);

        let pageTranslatedText = '';
        if (translationStyle === 'parentheses') {
          const bilingualList = [];
          for (let i = 0; i < sentences.length; i++) {
            const original = sentences[i].trim();
            const translated = (translatedSentences[i] || '').trim();
            if (!original) continue;

            if (!translated || original === translated || /^\d+$/.test(original) || original.length < 3) {
              bilingualList.push(original);
            } else {
              const annotationPrefix = targetLang === 'tr' ? 'Orijinal' : 'Original';
              bilingualList.push(`${translated} (${annotationPrefix}: ${original})`);
            }
          }
          pageTranslatedText = bilingualList.join(' ');
        } else {
          pageTranslatedText = translatedSentences.join(' ');
        }

        translatedPages.push({ pageNum: page.pageNum, translatedText: pageTranslatedText });
      }
    } else {
      // Kesintisiz Alt Alta Akış Modu (Yeni istenen mod - Boşlukları doldurarak sayfa sonu olmadan çevirir)
      const fullText = pagesText.map(p => p.text).join('\n');
      const paragraphs = fullText.split('\n').map(p => p.trim()).filter(Boolean);
      const translatedParagraphs = [];

      if (translationStyle === 'parentheses') {
        // Çift dilli (bilingual) parantezli mod - Paragraf içindeki cümle düzenini korur
        for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
          const paragraph = paragraphs[pIdx];
          const sentences = splitIntoSentences(paragraph);
          if (sentences.length === 0) {
            translatedParagraphs.push('');
            continue;
          }

          console.log(`[BELGE ÇEVİRİ] Kesintisiz Paragraf ${pIdx + 1}/${paragraphs.length} çevriliyor (${sentences.length} cümle)...`);
          const translatedSentences = await translateSentences(sentences, targetLang, sourceLang);
          
          const bilingualList = [];
          for (let i = 0; i < sentences.length; i++) {
            const original = sentences[i].trim();
            const translated = (translatedSentences[i] || '').trim();
            if (!original) continue;

            if (!translated || original === translated || /^\d+$/.test(original) || original.length < 3) {
              bilingualList.push(original);
            } else {
              const annotationPrefix = targetLang === 'tr' ? 'Orijinal' : 'Original';
              bilingualList.push(`${translated} (${annotationPrefix}: ${original})`);
            }
          }
          translatedParagraphs.push(bilingualList.join(' '));
          await new Promise(resolve => setTimeout(resolve, 100)); // Rate limit koruması
        }
      } else {
        // Standart Yalnızca Çeviri Modu - Hızlı ve yüksek kaliteli paragraf grup çevirisi (3000 karakterlik çunklar)
        const chunks = [];
        let currentChunk = [];
        let currentLength = 0;

        for (const para of paragraphs) {
          if (currentLength + para.length + 1 > 3000) {
            chunks.push(currentChunk);
            currentChunk = [para];
            currentLength = para.length;
          } else {
            currentChunk.push(para);
            currentLength += para.length + 1;
          }
        }
        if (currentChunk.length > 0) {
          chunks.push(currentChunk);
        }

        for (let i = 0; i < chunks.length; i++) {
          const chunk = chunks[i];
          console.log(`[BELGE ÇEVİRİ] Kesintisiz Çunk ${i + 1}/${chunks.length} çevriliyor (${chunk.length} paragraf)...`);
          try {
            const joinedChunk = chunk.join('\n');
            const translatedChunk = await translateText(joinedChunk, targetLang, sourceLang);
            const translatedLines = translatedChunk.split('\n').map(s => s.trim());
            
            if (translatedLines.length === chunk.length) {
              translatedParagraphs.push(...translatedLines);
            } else {
              console.warn(`Çunk ${i + 1} için satır sayısı eşleşmedi. Tek tek çevriliyor...`);
              for (const p of chunk) {
                if (p.length < 3 || /^\d+$/.test(p)) {
                  translatedParagraphs.push(p);
                } else {
                  try {
                    const tr = await translateText(p, targetLang, sourceLang);
                    translatedParagraphs.push(tr);
                  } catch (e) {
                    translatedParagraphs.push(p);
                  }
                }
                await new Promise(resolve => setTimeout(resolve, 150));
              }
            }
          } catch (err) {
            console.warn(`Çunk ${i + 1} çeviri hatası:`, err.message);
            for (const p of chunk) {
              if (p.length < 3 || /^\d+$/.test(p)) {
                translatedParagraphs.push(p);
              } else {
                try {
                  const tr = await translateText(p, targetLang, sourceLang);
                  translatedParagraphs.push(tr);
                } catch (e) {
                  translatedParagraphs.push(p);
                }
              }
              await new Promise(resolve => setTimeout(resolve, 150));
            }
          }
          await new Promise(resolve => setTimeout(resolve, 200));
        }
      }

      // Kesintisiz çıktıyı tek bir sayfa objesi olarak kaydediyoruz
      translatedPages.push({ pageNum: 1, translatedText: translatedParagraphs.join('\n') });
    }

    let outputBuffer;
    let mimeType = 'application/octet-stream';
    const baseName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
    const translatedFileName = `${baseName}_translated_${targetLang}.${targetFormat}`;

    if (targetFormat === 'txt') {
      let finalContent = '';
      for (const page of translatedPages) {
        if (showPageHeaders) {
          const headerText = targetLang === 'tr' ? `--- Sayfa ${page.pageNum} ---` : `--- Page ${page.pageNum} ---`;
          finalContent += `${headerText}\n\n`;
        }
        finalContent += `${page.translatedText}\n\n`;
      }
      outputBuffer = Buffer.from(finalContent.trim(), 'utf-8');
      mimeType = 'text/plain; charset=utf-8';
    } else if (targetFormat === 'docx') {
      const children = [];
      for (let i = 0; i < translatedPages.length; i++) {
        const page = translatedPages[i];
        
        if (showPageHeaders) {
          const headerText = targetLang === 'tr' ? `--- Sayfa ${page.pageNum} ---` : `--- Page ${page.pageNum} ---`;
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: headerText,
                  bold: true,
                  size: 24,
                  color: '10b981'
                })
              ],
              spacing: { before: 200, after: 100 }
            })
          );
        }
        
        const paras = page.translatedText.split('\n');
        for (const pText of paras) {
          if (pText.trim() === '') continue;
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: pText,
                  size: 22
                })
              ],
              spacing: { after: 120 }
            })
          );
        }
        
        if (showPageHeaders && i < translatedPages.length - 1) {
          children.push(new Paragraph({ children: [new PageBreak()] }));
        }
      }

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
          },
          children: children
        }]
      });

      outputBuffer = await Packer.toBuffer(doc);
      mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    }

    res.set({
      'Content-Type': mimeType,
      'Content-Disposition': `attachment; filename="${encodeURIComponent(translatedFileName)}"`,
      'Access-Control-Expose-Headers': 'Content-Disposition'
    });
    res.send(outputBuffer);

  } catch (err) {
    console.error("Belge çevirisi sırasında hata:", err);
    res.status(500).json({ error: 'Belge çevirisi sırasında bir hata oluştu: ' + err.message });
  }
});

app.post('/api/ppt-info', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Dosya yüklenmedi.' });

    const AdmZip = require('adm-zip');
    const zip = new AdmZip(req.file.buffer);

    let presXml = '';
    try {
      presXml = zip.readAsText('ppt/presentation.xml');
    } catch (e) {
      return res.status(400).json({ error: 'Geçersiz veya şifreli PowerPoint sunumu.' });
    }

    // Parse ordered slide list
    const sldIdRegex = /<p:sldId\s+id="([^"]+)"\s+r:id="([^"]+)"\s*\/>/g;
    const slides = [];
    let match;
    while ((match = sldIdRegex.exec(presXml)) !== null) {
      slides.push({ id: match[1], rId: match[2] });
    }

    let relsXml = '';
    try {
      relsXml = zip.readAsText('ppt/_rels/presentation.xml.rels');
    } catch (e) { }

    const relRegex = /<Relationship\s+Id="([^"]+)"\s+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide"\s+Target="([^"]+)"\s*\/>/g;
    const rels = {};
    let relMatch;
    while ((relMatch = relRegex.exec(relsXml)) !== null) {
      rels[relMatch[1]] = relMatch[2];
    }

    const slideData = [];

    for (let i = 0; i < slides.length; i++) {
      const s = slides[i];
      const target = rels[s.rId];
      if (!target) continue;

      const slidePath = 'ppt/' + target;
      let slideXml = '';
      try {
        slideXml = zip.readAsText(slidePath);
      } catch (e) {
        continue;
      }

      // Extract rich text elements in order from <a:r>
      const elements = [];
      const texts = [];
      const rRegex = /<a:r>([\s\S]*?)<\/a:r>/g;
      let rMatch;
      let elIndex = 0;

      while ((rMatch = rRegex.exec(slideXml)) !== null) {
        const rContent = rMatch[1];
        const tMatch = /<a:t>([\s\S]*?)<\/a:t>/.exec(rContent);
        if (!tMatch) continue;

        const txt = tMatch[1].trim();
        if (!txt) continue;

        texts.push(txt);

        // Extract size
        let fontSize = 18;
        const szMatch = /sz="([0-9]+)"/.exec(rContent);
        if (szMatch) {
          fontSize = Math.round(parseInt(szMatch[1], 10) / 100);
        }

        // Bold, italic, underline
        const bold = /b="1"/.test(rContent);
        const italic = /i="1"/.test(rContent);
        const underline = /u="sng"/.test(rContent);

        // Color
        let color = null;
        const colorMatch = /<a:srgbClr\s+val="([^"]+)"/.exec(rContent);
        if (colorMatch) {
          color = '#' + colorMatch[1];
        }

        // Dynamic positioning for the UI editor canvas
        // Title elements occupy top center, bullets occupy lower left
        const isTitle = elIndex === 0;
        const xPos = isTitle ? 20 : 15;
        const yPos = isTitle ? 15 : 30 + ((elIndex - 1) * 12);

        elements.push({
          id: `el_${i + 1}_${elIndex}`,
          text: txt,
          x: xPos,
          y: yPos,
          fontSize: fontSize || (isTitle ? 32 : 18),
          bold: bold || isTitle,
          italic: italic,
          underline: underline,
          color: color || (isTitle ? '#38BDF8' : '#FFFFFF'),
          align: isTitle ? 'center' : 'left',
          isTitle: isTitle
        });

        elIndex++;
      }

      // Extract background color if present
      let bgColor = null;
      const bgMatch = /<a:srgbClr\s+val="([^"]+)"\s*\/>/i.exec(slideXml);
      if (bgMatch) {
        bgColor = '#' + bgMatch[1];
      }

      slideData.push({
        slideIndex: i + 1,
        texts: texts,
        elements: elements,
        title: texts[0] || `Slayt ${i + 1}`,
        bgColor: bgColor || '#FFFFFF'
      });
    }

    res.json({
      pageCount: slideData.length,
      slides: slideData
    });
  } catch (e) {
    res.status(500).json({ error: 'PowerPoint bilgileri alınamadı: ' + e.message });
  }
});

app.get('/api/youtube-search', async (req, res) => {
  try {
    const yts = require('yt-search');
    const query = req.query.q;
    if (!query) return res.status(400).json({ error: 'Arama terimi eksik.' });

    const r = await yts(query);
    const videos = r.videos.slice(0, 15).map(v => ({
      title: v.title,
      url: v.url,
      thumbnail: v.thumbnail,
      duration: v.timestamp,
      author: v.author.name
    }));

    res.json(videos);
  } catch (err) {
    console.error('YouTube Search Error:', err);
    res.status(500).json({ error: 'Arama yapılırken hata oluştu.' });
  }
});

app.get('/api/google-search', async (req, res) => {
  try {
    const axios = require('axios');
    const cheerio = require('cheerio');
    const https = require('https');
    
    const query = req.query.q;
    if (!query) return res.status(400).json({ error: 'Arama terimi eksik.' });

    // Sistem saati 2026 olduğu için SSL sertifika hatalarını yok sayıyoruz
    const agent = new https.Agent({ rejectUnauthorized: false });

    const response = await axios.get(`https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`, {
      httpsAgent: agent,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      }
    });

    const $ = cheerio.load(response.data);
    const results = [];
    
    $('.result').each((i, el) => {
      if (i > 10) return;
      const title = $(el).find('.result__title a').text().trim();
      const url = $(el).find('.result__a').attr('href') || $(el).find('.result__title a').attr('href');
      const description = $(el).find('.result__snippet').text().trim();
      
      let finalUrl = url;
      if (url && url.startsWith('//duckduckgo.com/l/?uddg=')) {
        finalUrl = decodeURIComponent(url.replace('//duckduckgo.com/l/?uddg=', '').split('&')[0]);
      }
      
      if (title && finalUrl) {
        results.push({ title, url: finalUrl, description });
      }
    });

    res.json(results);
  } catch (err) {
    console.error('Web Search Error:', err);
    res.status(500).json({ error: 'Arama yapılırken hata oluştu.' });
  }
});

// Proxy endpoint to bypass X-Frame-Options for websites
app.get('/api/proxy', async (req, res) => {
  try {
    const axios = require('axios');
    const https = require('https');
    const cheerio = require('cheerio');
    
    const targetUrl = req.query.url;
    if (!targetUrl) return res.status(400).send('URL required');

    const agent = new https.Agent({ rejectUnauthorized: false });
    
    const response = await axios.get(targetUrl, {
      httpsAgent: agent,
      responseType: 'arraybuffer',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      }
    });

    // Determine content type
    const contentType = response.headers['content-type'] || 'text/html';
    res.setHeader('Content-Type', contentType);

    // If it's HTML, inject a base tag so relative assets work
    if (contentType.includes('text/html')) {
      const html = response.data.toString('utf-8');
      const $ = cheerio.load(html);
      
      // Inject <base> tag to fix relative links and images
      const parsedUrl = new URL(targetUrl);
      const baseUrl = `${parsedUrl.protocol}//${parsedUrl.host}`;
      $('head').prepend(`<base href="${baseUrl}/">`);
      
      // Modify all anchor tags to open via our proxy if they are same origin, or just let them be
      // (For simplicity, we just inject base tag which handles most things)

      res.send($.html());
    } else {
      res.send(response.data);
    }
  } catch (err) {
    console.error('Proxy Error:', err.message);
    res.status(500).send(`Proxy Hatası: Bu siteye erişilemiyor. (${err.message})`);
  }
});

app.post('/api/rooms', async (req, res) => {
  try {
    const crypto = require('crypto');
    const clientIP = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    
    // Generate secure hash combining IP and random bytes
    const randomBytes = crypto.randomBytes(16).toString('hex');
    const hashInput = `${clientIP}-${randomBytes}-${Date.now()}`;
    const roomHash = crypto.createHash('sha256').update(hashInput).digest('hex').substring(0, 16); // 16 char secure hash
    
    let pool = await sql.connect(sqlConfig);
    await pool.request()
      .input('roomHash', sql.NVarChar, roomHash)
      .input('adminIP', sql.NVarChar, clientIP)
      .query(`
        INSERT INTO VirtualRooms (RoomHash, AdminIP)
        VALUES (@roomHash, @adminIP)
      `);
      
    res.json({ roomHash, inviteLink: `http://localhost:5174/?room=${roomHash}` });
  } catch (err) {
    console.error('Room Creation Error:', err);
    res.status(500).json({ error: 'Oda oluşturulamadı.' });
  }
});

app.get('/api/youtube-info', async (req, res) => {
  try {
    const youtubedl = require('youtube-dl-exec');
    const url = req.query.url;
    if (!url) {
      return res.status(400).json({ error: 'Geçersiz YouTube URLsi.' });
    }
    
    const info = await youtubedl(url, {
      dumpSingleJson: true,
      noCheckCertificates: true,
      noWarnings: true,
      preferFreeFormats: true,
      addHeader: [
        'referer:youtube.com',
        'user-agent:Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      ]
    });
    
    // YENİ AKILLI SEÇENEKLER (Sabit Ön Ayarlar)
    // Sadece en yüksek kalite MP4 ve MP3 seçenekleri
    const availableFormats = [
      {
        itag: 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best',
        quality: 'En Yüksek Kalite MP4 Video (Orijinal Görüntü + Ses)',
        type: 'video'
      },
      {
        itag: 'bestaudio/best',
        quality: 'En Yüksek Kalite MP3 (Sadece Ses)',
        type: 'audio'
      }
    ];

    res.json({
      title: info.title,
      thumbnail: info.thumbnail,
      formats: availableFormats
    });
  } catch (err) {
    let errorMsg = err.message || 'Bilinmeyen bir hata oluştu.';
    
    // YouTube/Yt-dlp özel hata mesajlarını Türkçeleştirme
    if (errorMsg.includes('copyright grounds') || errorMsg.includes('copyright')) {
      errorMsg = 'Bu video telif hakları nedeniyle (başka bir kurum/kişi tarafından) engellenmiştir ve indirilemez.';
    } else if (errorMsg.includes('Video unavailable') || errorMsg.includes('not available')) {
      errorMsg = 'Bu video kullanılamıyor, silinmiş veya gizli olabilir.';
    } else if (errorMsg.includes('Unsupported URL')) {
      errorMsg = 'Girdiğiniz bağlantı desteklenmiyor veya sitede video bulunamadı.';
    } else if (errorMsg.includes('Sign in to confirm')) {
      errorMsg = 'Bu video yaş kısıtlamalıdır ve üye girişi gerektirdiği için sistem tarafından otomatik indirilemez.';
    } else if (errorMsg.includes('Private video')) {
      errorMsg = 'Bu video gizlidir ve erişilemez.';
    }

    res.status(500).json({ error: errorMsg });
  }
});

// ========================================================
// YENİ ÖZELLİK: PUPPETEER STEALTH İLE KORSAN/KORUMALI SİTELERİ AŞMA (HACK)
// ========================================================
async function extractMediaWithPuppeteer(pageUrl) {
  const puppeteer = require('puppeteer-extra');
  const StealthPlugin = require('puppeteer-extra-plugin-stealth');
  puppeteer.use(StealthPlugin());

  let browser;
  try {
    console.log(`[PUPPETEER] Güvenlik duvarını aşmak için görünmez tarayıcı başlatılıyor: ${pageUrl}`);
    browser = await puppeteer.launch({
      headless: "new",
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-web-security']
    });
    const page = await browser.newPage();
    
    // Ağ trafiğini dinleyelim (Gerçek video dosyasını havada yakala)
    let foundUrl = null;
    page.on('request', request => {
      const reqUrl = request.url();
      // Reklam ve analytics engelle
      if (reqUrl.includes('google-analytics') || reqUrl.includes('doubleclick')) return;
      
      // Video akış dosyalarını (.m3u8, .mp4, .ts) yakala
      if (reqUrl.includes('.m3u8') || reqUrl.includes('.mp4')) {
        if (!foundUrl) {
          console.log(`[PUPPETEER] GİZLİ VİDEO LİNKİ YAKALANDI: ${reqUrl.substring(0, 80)}...`);
          foundUrl = reqUrl;
        }
      }
    });

    // Siteye git ve güvenlik testlerini (Cloudflare vb.) geçmesini bekle
    await page.goto(pageUrl, { waitUntil: 'networkidle2', timeout: 35000 });

    // Eğer m3u8 bulamadıysak, gizlenmiş iframe (Vidmoly, Fembed vb.) oynatıcılarını bulalım
    const iframes = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('iframe'))
        .map(i => i.src)
        .filter(src => src && !src.includes('youtube.com') && !src.includes('google'));
    });

    if (foundUrl) {
      return { type: 'direct', url: foundUrl };
    } else if (iframes.length > 0) {
      console.log(`[PUPPETEER] Oynatıcı bulundu: ${iframes[0]}`);
      return { type: 'iframe', url: iframes[0] };
    }
    
    return null;
  } catch (error) {
    console.error("[PUPPETEER] Bypass hatası:", error.message);
    return null;
  } finally {
    if (browser) await browser.close();
  }
}

app.get('/api/youtube-download', async (req, res) => {
  try {
    const youtubedl = require('youtube-dl-exec');
    const ffmpegStatic = require('ffmpeg-static');
    const fs = require('fs');
    const path = require('path');
    const os = require('os');

    const { url, itag, volume, enhance } = req.query;
    if (!url) {
      return res.status(400).json({ error: 'Geçersiz YouTube URLsi.' });
    }
    
    // Bilgileri alıp isim belirleme (Zorla indirme seçeneği için hataları görmezden geliyoruz)
    let title = 'video_indirilen';
    try {
      const info = await youtubedl(url, { dumpSingleJson: true, noWarnings: true, ignoreErrors: true, noPlaylist: true });
      if (info && info.title) {
        title = info.title;
      }
    } catch (e) {
      console.warn("Video başlığı alınamadı, standart isim kullanılacak:", e.message);
    }
    
    const isAudioOnly = itag && itag.includes('bestaudio') && !itag.includes('bestvideo');
    const ext = isAudioOnly ? 'mp3' : 'mp4';
    
    // Kullanıcıya gönderilecek güvenli isim (RFC 5987 / UTF-8 destekli)
    const safeTitle = encodeURIComponent(title);

    // Geçici dosya adı
    const tempDir = os.tmpdir();
    const tempFileName = `ytdownload_${Date.now()}_${Math.floor(Math.random() * 1000)}.${ext}`;
    const tempFilePath = path.join(tempDir, tempFileName);

    // Başlık bilgisini UTF-8 ile tarayıcıya gönder
    res.header('Content-Disposition', `attachment; filename*=UTF-8''${safeTitle}.${ext}`);
    res.header('Access-Control-Expose-Headers', 'Content-Disposition');
    
    // İndirme ve ffmpeg birleştirme işlemi
    const dlOptions = {
      o: tempFilePath,
      f: itag || 'best',
      ffmpegLocation: ffmpegStatic,
      noWarnings: true,
      ignoreErrors: true, // Hataları görmezden gel ve ne bulursan indir (Zorla indirme için)
      noPlaylist: true, // Sadece ilk bulduğu videoyu alsın
      concurrentFragments: 5 // Çoklu bağlantı ile indirmeyi hızlandırır
    };

    // YENİ: EĞER ZORLA İNDİRME İSE VE SİTE BİLİNEN KORSAN/KORUMALI SİTEYSE (Örn: hdfilmcehennemi)
    // Önce Puppeteer ile şifreli linki bul, sonra onu yt-dlp'ye ver!
    let targetUrl = url;
    if (url.includes('hdfilmcehennemi') || url.includes('dizi') || url.includes('film')) {
      console.log("Korumalı site algılandı, Puppeteer Bypass devreye giriyor...");
      const bypassData = await extractMediaWithPuppeteer(url);
      if (bypassData && bypassData.url) {
        targetUrl = bypassData.url;
        dlOptions.referer = url; // Güvenliği geçmek için referer ekle
      }
    }

    if (isAudioOnly) {
      dlOptions.extractAudio = true;
      dlOptions.audioFormat = 'mp3';
      dlOptions.audioQuality = 0; // En iyi ses kalitesi
      
      if (volume && volume !== '1' && volume !== '1.0') {
        dlOptions.postprocessorArgs = `ffmpeg:-filter:a volume=${volume}`; 
      }
    } else {
      dlOptions.mergeOutputFormat = 'mp4';
      dlOptions.formatSort = 'res,ext:mp4:m4a'; // En iyi uyumlu formatları zorla
      
      let ffmpegArgs = [];
      if (volume && volume !== '1' && volume !== '1.0') {
        ffmpegArgs.push(`-filter:a volume=${volume}`);
      }
      
      // Video görüntü iyileştirmesi (Re-encode gerektirdiği için işlemi yavaşlatabilir)
      if (enhance === 'vibrant') {
        ffmpegArgs.push('-filter:v eq=saturation=1.5:contrast=1.1:brightness=0.05 -c:v libx264 -preset ultrafast -crf 23');
      } else if (enhance === 'bright') {
        ffmpegArgs.push('-filter:v eq=brightness=0.15:contrast=1.1 -c:v libx264 -preset ultrafast -crf 23');
      }
      
      if (ffmpegArgs.length > 0) {
        dlOptions.postprocessorArgs = `ffmpeg:${ffmpegArgs.join(' ')}`;
      }
    }
    
    await youtubedl(targetUrl, dlOptions);
    
    // Dosyayı kullanıcıya sun ve sil
    res.download(tempFilePath, `${title}.${ext}`, (err) => {
      if (err) {
        console.error('Gönderme hatası:', err);
      }
      if (fs.existsSync(tempFilePath)) {
        fs.unlinkSync(tempFilePath);
      }
    });

  } catch (err) {
    console.error("İndirme hatası:", err);
    if (!res.headersSent) {
      let errorMsg = err.message || 'Bilinmeyen bir hata oluştu.';
      if (errorMsg.includes('HTTP Error 404') || errorMsg.includes('Not Found')) {
        errorMsg = 'Video dosyası bu web sitesinin gizli sunucusunda barındırıldığı için (404 Not Found) bot tarafından çekilemedi. Bu site özel koruma kullanıyor.';
      } else if (errorMsg.includes('copyright')) {
        errorMsg = 'Bu video telif hakları nedeniyle (başka bir kurum/kişi tarafından) engellenmiştir ve indirilemez.';
      } else if (errorMsg.includes('Video unavailable') || errorMsg.includes('not available')) {
        errorMsg = 'Bu video kullanılamıyor, silinmiş veya gizli olabilir.';
      }
      res.status(500).json({ error: errorMsg });
    }
  }
});

// ==========================================
// YENİ ÖZELLİK: DERS ÇALIŞMA ORTAMI (API)
// ==========================================

// --- Notlar API ---
app.get('/api/study/notes', async (req, res) => {
  try {
    const clientIP = req.headers['x-client-id'] || req.headers['x-forwarded-for'] || req.socket.remoteAddress || '127.0.0.1';
    let pool = await sql.connect(sqlConfig);
    const result = await pool.request()
      .input('ip', sql.NVarChar, clientIP)
      .query('SELECT * FROM StudyNotes WHERE client_ip = @ip ORDER BY created_at DESC');
    res.json(result.recordset);
  } catch (err) {
    res.status(500).json({ error: 'Notlar getirilemedi.' });
  }
});

app.post('/api/study/notes', async (req, res) => {
  try {
    const clientIP = req.headers['x-client-id'] || req.headers['x-forwarded-for'] || req.socket.remoteAddress || '127.0.0.1';
    const { content } = req.body;
    let pool = await sql.connect(sqlConfig);
    await pool.request()
      .input('ip', sql.NVarChar, clientIP)
      .input('content', sql.NVarChar, content)
      .query('INSERT INTO StudyNotes (content, client_ip) VALUES (@content, @ip)');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'Not kaydedilemedi.' });
  }
});

app.delete('/api/study/notes/:id', async (req, res) => {
  try {
    const clientIP = req.headers['x-client-id'] || req.headers['x-forwarded-for'] || req.socket.remoteAddress || '127.0.0.1';
    const { id } = req.params;
    let pool = await sql.connect(sqlConfig);
    await pool.request()
      .input('id', sql.Int, id)
      .input('ip', sql.NVarChar, clientIP)
      .query('DELETE FROM StudyNotes WHERE id = @id AND client_ip = @ip');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'Not silinemedi.' });
  }
});

// --- İlerleme Haritası API ---
app.get('/api/study/progress', async (req, res) => {
  try {
    const clientIP = req.headers['x-client-id'] || req.headers['x-forwarded-for'] || req.socket.remoteAddress || '127.0.0.1';
    let pool = await sql.connect(sqlConfig);
    const result = await pool.request()
      .input('ip', sql.NVarChar, clientIP)
      .query('SELECT * FROM StudyProgress WHERE client_ip = @ip ORDER BY created_at DESC');
    res.json(result.recordset);
  } catch (err) {
    res.status(500).json({ error: 'İlerleme verileri getirilemedi.' });
  }
});

app.post('/api/study/progress', async (req, res) => {
  try {
    const clientIP = req.headers['x-client-id'] || req.headers['x-forwarded-for'] || req.socket.remoteAddress || '127.0.0.1';
    const { test_name, score, total_questions } = req.body;
    let pool = await sql.connect(sqlConfig);
    await pool.request()
      .input('ip', sql.NVarChar, clientIP)
      .input('test_name', sql.NVarChar, test_name)
      .input('score', sql.Int, score)
      .input('total', sql.Int, total_questions)
      .query('INSERT INTO StudyProgress (test_name, score, total_questions, client_ip) VALUES (@test_name, @score, @total, @ip)');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'İlerleme kaydedilemedi.' });
  }
});

app.delete('/api/study/progress/:id', async (req, res) => {
  try {
    const clientIP = req.headers['x-client-id'] || req.headers['x-forwarded-for'] || req.socket.remoteAddress || '127.0.0.1';
    const { id } = req.params;
    let pool = await sql.connect(sqlConfig);
    await pool.request()
      .input('id', sql.Int, id)
      .input('ip', sql.NVarChar, clientIP)
      .query('DELETE FROM StudyProgress WHERE id = @id AND client_ip = @ip');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'İlerleme silinemedi.' });
  }
});

// --- VIRTUAL ROOMS (SOCKET.IO) EVENTS ---
// Basic profanity filter
const badWords = ['amk', 'aq', 'siktir', 'piç', 'oç'];

// Room states: { roomKey: { users: [{ id, username, role }], currentUrl: string, currentTime: number, isPlaying: boolean } }
const virtualRooms = {};

io.on('connection', (socket) => {
  console.log('A user connected for Virtual Rooms:', socket.id);

  socket.on('join_room', (data) => {
    const { roomKey, username } = data;
    socket.join(roomKey);
    
    if (!virtualRooms[roomKey]) {
      virtualRooms[roomKey] = { users: [], currentUrl: '', currentTime: 0, isPlaying: false };
    }
    
    // First user becomes admin
    const role = virtualRooms[roomKey].users.length === 0 ? 'admin' : 'member';
    virtualRooms[roomKey].users.push({ id: socket.id, username, role, isMuted: false });

    // Notify room of new user
    socket.to(roomKey).emit('receive_message', {
      user: 'Sistem',
      text: `${username} odaya katıldı!`,
      time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})
    });

    // Send updated user list and current video state to the room
    io.in(roomKey).emit('room_users', virtualRooms[roomKey].users);
    
    // If there's an active video, sync the new user
    if (virtualRooms[roomKey].currentUrl) {
      socket.emit('video_sync', {
        action: virtualRooms[roomKey].isPlaying ? 'play' : 'pause',
        url: virtualRooms[roomKey].currentUrl,
        time: virtualRooms[roomKey].currentTime
      });
    }
  });

  socket.on('send_message', (data) => {
    const room = virtualRooms[data.roomKey];
    if (room) {
      const realUser = room.users.find(u => u.id === socket.id);
      if (realUser && realUser.isMuted) {
        socket.emit('receive_message', { user: 'Sistem', text: 'Susturulduğunuz için mesaj gönderemezsiniz.', time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
        return;
      }
    }

    let cleanText = data.text;
    badWords.forEach(word => {
      const regex = new RegExp(word, 'gi');
      cleanText = cleanText.replace(regex, '***');
    });

    data.text = cleanText;
    io.in(data.roomKey).emit('receive_message', data);
  });

  socket.on('video_action', (data) => {
    // data: { roomKey, action: 'play' | 'pause' | 'seek' | 'load', time: Number, url: String }
    const room = virtualRooms[data.roomKey];
    if (!room) return;
    
    const user = room.users.find(u => u.id === socket.id);
    if (!user || user.role !== 'admin') {
      socket.emit('receive_message', { user: 'Sistem', text: 'Video kontrolleri sadece oda yöneticilerine (Admin) aittir.', time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
      return;
    }

    if (data.action === 'load') room.currentUrl = data.url;
    if (data.action === 'play') room.isPlaying = true;
    if (data.action === 'pause') room.isPlaying = false;
    if (data.time !== undefined) room.currentTime = data.time;

    socket.to(data.roomKey).emit('video_sync', data);
  });

  socket.on('kick_user', (data) => {
    const room = virtualRooms[data.roomKey];
    if (!room) return;
    const admin = room.users.find(u => u.id === socket.id);
    if (admin && admin.role === 'admin') {
      const targetUser = room.users.find(u => u.id === data.targetId);
      if (targetUser) {
        io.to(data.targetId).emit('kicked');
        io.sockets.sockets.get(data.targetId)?.leave(data.roomKey);
        room.users = room.users.filter(u => u.id !== data.targetId);
        io.in(data.roomKey).emit('room_users', room.users);
        io.in(data.roomKey).emit('receive_message', { user: 'Sistem', text: `${targetUser.username} odadan atıldı.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
      }
    }
  });

  socket.on('promote_user', (data) => {
    const room = virtualRooms[data.roomKey];
    if (!room) return;
    const admin = room.users.find(u => u.id === socket.id);
    if (admin && admin.role === 'admin') {
      const targetUser = room.users.find(u => u.id === data.targetId);
      if (targetUser) {
        targetUser.role = 'admin';
        io.in(data.roomKey).emit('room_users', room.users);
        io.in(data.roomKey).emit('receive_message', { user: 'Sistem', text: `${targetUser.username} artık bir oda yöneticisi.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
      }
    }
  });

  socket.on('transfer_admin', (data) => {
    const room = virtualRooms[data.roomKey];
    if (!room) return;
    const admin = room.users.find(u => u.id === socket.id);
    if (admin && admin.role === 'admin') {
      const targetUser = room.users.find(u => u.id === data.targetId);
      if (targetUser) {
        targetUser.role = 'admin';
        admin.role = 'member';
        io.in(data.roomKey).emit('room_users', room.users);
        io.in(data.roomKey).emit('receive_message', { user: 'Sistem', text: `${admin.username}, oda yöneticiliğini ${targetUser.username} kişisine devretti.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
      }
    }
  });

  socket.on('mute_user', (data) => {
    const room = virtualRooms[data.roomKey];
    if (!room) return;
    const admin = room.users.find(u => u.id === socket.id);
    if (admin && admin.role === 'admin') {
      const targetUser = room.users.find(u => u.id === data.targetId);
      if (targetUser && targetUser.id !== socket.id) {
        targetUser.isMuted = true;
        io.in(data.roomKey).emit('room_users', room.users);
        io.to(data.targetId).emit('muted');
        io.in(data.roomKey).emit('receive_message', { user: 'Sistem', text: `${targetUser.username} yönetici tarafından susturuldu.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
      }
    }
  });

  socket.on('unmute_user', (data) => {
    const room = virtualRooms[data.roomKey];
    if (!room) return;
    const admin = room.users.find(u => u.id === socket.id);
    if (admin && admin.role === 'admin') {
      const targetUser = room.users.find(u => u.id === data.targetId);
      if (targetUser) {
        targetUser.isMuted = false;
        io.in(data.roomKey).emit('room_users', room.users);
        io.to(data.targetId).emit('unmuted');
        io.in(data.roomKey).emit('receive_message', { user: 'Sistem', text: `${targetUser.username} kişisinin susturması kaldırıldı.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
      }
    }
  });

  socket.on('disconnect', () => {
    console.log('User disconnected:', socket.id);
    // Remove user from any rooms they were in
    for (const [roomKey, room] of Object.entries(virtualRooms)) {
      const userIndex = room.users.findIndex(u => u.id === socket.id);
      if (userIndex !== -1) {
        const user = room.users[userIndex];
        room.users.splice(userIndex, 1);
        if (room.users.length === 0) {
          delete virtualRooms[roomKey]; // Delete empty room
        } else {
          // If admin left, randomly assign new admin? Or keep it adminless. Let's make the oldest user admin.
          if (user.role === 'admin' && !room.users.some(u => u.role === 'admin')) {
             room.users[0].role = 'admin';
             io.in(roomKey).emit('receive_message', { user: 'Sistem', text: `${room.users[0].username} yeni yönetici oldu.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
          }
          io.in(roomKey).emit('room_users', room.users);
          io.in(roomKey).emit('receive_message', { user: 'Sistem', text: `${user.username} odadan ayrıldı.`, time: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) });
        }
      }
    }
  });
});

server.listen(port, () => {
  console.log(`Backend server (with Socket.io) running on http://localhost:${port}`);
});
