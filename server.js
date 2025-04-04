require('dotenv').config();

const express = require('express');
const cors = require('cors');
const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const PDFDocument = require('pdfkit');
const PptxGenJS = require('pptxgenjs');
const AdmZip = require('adm-zip');

const app = express();

// 1) Environment variables
const PORT = process.env.PORT || 4000;
const baseUrl = process.env.BASE_URL || `http://localhost:${PORT}`;

// 2) Basic Express config & CORS
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type']
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 3) Directory for final downloadable files
const downloadsDir = path.join(__dirname, 'downloads');
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir);
}

// 4) Serve files from /downloads/:filename
app.get('/downloads/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(downloadsDir, filename);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }

  // Send correct headers
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Referrer-Policy', 'no-referrer');

  const ext = path.extname(filename).toLowerCase();
  switch (ext) {
    case '.pdf':
      res.setHeader('Content-Type', 'application/pdf');
      break;
    case '.pptx':
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      break;
    case '.zip':
      res.setHeader('Content-Type', 'application/zip');
      break;
    default:
      res.setHeader('Content-Type', 'application/octet-stream');
  }

  // Stream the file
  const fileStream = fs.createReadStream(filePath);
  fileStream.pipe(res);

  // Cleanup after 3 minutes
  setTimeout(() => {
    if (fs.existsSync(filePath)) {
      fs.unlink(filePath, (err) => {
        if (err) console.error(`Failed to delete ${filePath}:`, err);
      });
    }
  }, 3 * 60 * 1000);
});

// 5) Image sizes dictionary
const imageSizesDict = {
  '320': { quality: 85, width: 320 },
  '638': { quality: 85, width: 638 },
  '2048': { quality: 75, width: 2048 },
};

/**
 * 6) /api/get-slides
 *    Scrapes the SlideShare page for the preview slides
 */
app.post('/api/get-slides', async (req, res) => {
  try {
    const { slideshareUrl } = req.body;
    if (!slideshareUrl) {
      return res.status(400).json({ error: 'No SlideShare URL provided.' });
    }

    // Load SlideShare HTML
    const response = await axios.get(slideshareUrl);
    const html = response.data;
    const $ = cheerio.load(html);

    // Find __NEXT_DATA__ JSON
    let nextDataScript = null;
    $('script').each((i, script) => {
      if ($(script).attr('id') === '__NEXT_DATA__') {
        nextDataScript = $(script).html();
      }
    });
    if (!nextDataScript) {
      return res.status(404).json({ error: 'Could not find __NEXT_DATA__ script on that page.' });
    }

    const jsonData = JSON.parse(nextDataScript);
    const props = jsonData?.props?.pageProps || {};
    const slideshow = props?.slideshow || {};
    const slides = slideshow?.slides || {};

    const totalSlides = slideshow.totalSlides || 0;
    const imageLocation = slides.imageLocation || '';
    const imageTitle = slides.title || '';
    const host = slides.host || '';

    // Build low-res preview URLs
    const { width: previewWidth, quality: previewQuality } = imageSizesDict['320'];
    const slideImagesPreview = [];
    for (let i = 1; i <= totalSlides; i++) {
      const url = `${host}/${imageLocation}/${previewQuality}/${imageTitle}-${i}-${previewWidth}.jpg`;
      slideImagesPreview.push(url);
    }

    return res.json({
      totalSlides,
      slideImagesPreview,
      slideshowInfo: { host, imageLocation, imageTitle },
    });
  } catch (err) {
    console.error('Error in /api/get-slides:', err);
    return res.status(500).json({ error: err.message });
  }
});

/**
 * 7) /api/generate-file
 *    Downloads selected slides in chosen resolution, then
 *    builds final PDF/PPTX/ZIP. Download is done in parallel.
 */
app.post('/api/generate-file', async (req, res) => {
  let downloadedFiles = []; // keep track for cleanup on error

  try {
    const { slideshowInfo, resolution, outputFormat, selectedIndices } = req.body;

    if (!slideshowInfo || !resolution || !outputFormat || !selectedIndices) {
      return res.status(400).json({ error: 'Missing required data.' });
    }
    if (!Array.isArray(selectedIndices) || selectedIndices.length === 0) {
      return res.status(400).json({ error: 'No slides selected.' });
    }

    const { host, imageLocation, imageTitle } = slideshowInfo;
    const sizeEntry = imageSizesDict[resolution];
    if (!sizeEntry) {
      return res.status(400).json({ error: `Invalid resolution: ${resolution}` });
    }
    const { width, quality } = sizeEntry;

    // Build final array of slide URLs
    const slideImagesFinal = selectedIndices.map((idx) => {
      const realSlideNum = idx + 1; // user slides are 0-based
      return `${host}/${imageLocation}/${quality}/${imageTitle}-${realSlideNum}-${width}.jpg`;
    });

    // Folder for temp images
    const tempFolder = path.join(__dirname, 'temp_slides');
    if (!fs.existsSync(tempFolder)) {
      fs.mkdirSync(tempFolder, { recursive: true });
    }

    // Decide final file format for images
    // If user asked "png", we store slides as .png. Otherwise .jpg
    const wantsPng = (outputFormat.toLowerCase() === 'png');

    // 8) Parallel download + convert
    await Promise.all(
      slideImagesFinal.map(async (imgUrl, i) => {
        const ext = wantsPng ? 'png' : 'jpg';
        const filename = `slide_${i}.${ext}`;
        const filepath = path.join(tempFolder, filename);

        // Download raw arraybuffer
        const response = await axios.get(imgUrl, { responseType: 'arraybuffer' });
        if (response.status !== 200) {
          throw new Error(`Failed to download image #${i}: HTTP ${response.status}`);
        }
        const originalBuffer = Buffer.from(response.data);

        // Convert to desired format using Sharp
        let finalBuffer;
        if (wantsPng) {
          finalBuffer = await sharp(originalBuffer).png().toBuffer();
        } else {
          finalBuffer = await sharp(originalBuffer).jpeg().toBuffer();
        }

        fs.writeFileSync(filepath, finalBuffer);
        downloadedFiles.push(filepath);
      })
    );

    // 9) Build final file
    let finalExt = outputFormat.toLowerCase();
    if (finalExt === 'jpg' || finalExt === 'png') {
      // We'll do a .zip so there's a single download link
      finalExt = 'zip';
    }

    const finalFilename = `slides_${Date.now()}.${finalExt}`;
    const finalFilePath = path.join(downloadsDir, finalFilename);

    switch (outputFormat.toLowerCase()) {
      case 'jpg':
      case 'png':
      case 'zip': {
        // Zip up all the images
        const zip = new AdmZip();
        for (const file of downloadedFiles) {
          zip.addLocalFile(file);
        }
        fs.writeFileSync(finalFilePath, zip.toBuffer());
        break;
      }

      case 'pdf': {
        // Build a PDF: each image -> new page
        const pdfDoc = new PDFDocument({ autoFirstPage: false });
        const writeStream = fs.createWriteStream(finalFilePath);
        pdfDoc.pipe(writeStream);

        for (const file of downloadedFiles) {
          const image = pdfDoc.openImage(file);
          pdfDoc.addPage({ size: [image.width, image.height] });
          pdfDoc.image(image, 0, 0);
        }
        pdfDoc.end();

        await new Promise((resolve) => writeStream.on('finish', resolve));
        break;
      }

      case 'pptx': {
        // Build a PPTX: each slide -> one full-image slide
        const pptx = new PptxGenJS();
        for (const file of downloadedFiles) {
          const slide = pptx.addSlide();
          const fileBuffer = fs.readFileSync(file);
          const b64 = fileBuffer.toString('base64');

          slide.addImage({
            data: `data:image/${wantsPng ? 'png' : 'jpeg'};base64,${b64}`,
            x: 0, y: 0, w: '100%', h: '100%',
          });
        }
        const b64pptx = await pptx.write('base64');
        const pptxBuffer = Buffer.from(b64pptx, 'base64');
        fs.writeFileSync(finalFilePath, pptxBuffer);
        break;
      }

      default:
        cleanupTempFiles(downloadedFiles);
        return res.status(400).json({ error: `Invalid output format: ${outputFormat}` });
    }

    // 10) Remove temp images
    cleanupTempFiles(downloadedFiles);

    // 11) Return a direct download URL
    const downloadUrl = `${baseUrl}/downloads/${finalFilename}`;
    return res.json({ downloadUrl });
  } catch (err) {
    console.error('Error in /api/generate-file:', err);
    cleanupTempFiles(downloadedFiles);
    return res.status(500).json({ error: err.message });
  }
});

// Utility function to remove temp files
function cleanupTempFiles(files) {
  for (const filePath of files) {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  }
}

// Start the server
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
