require('dotenv').config(); // <-- Make sure you install dotenv and load it at the top

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

const PORT = 4000;
const baseUrl = process.env.BASE_URL || `http://localhost:${PORT}`;

// CORS config
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type']
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Set up 'downloads' folder
const downloadsDir = path.join(__dirname, 'downloads');
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir);
}

/**
 * Serve generated files from /downloads/:filename
 */
app.get('/downloads/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(downloadsDir, filename);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }

  // Set relevant headers
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

  const fileStream = fs.createReadStream(filePath);
  fileStream.pipe(res);

  // Optional cleanup after 5 minutes
  setTimeout(() => {
    if (fs.existsSync(filePath)) {
      fs.unlink(filePath, (err) => {
        if (err) console.error(`Failed to delete ${filePath}:`, err);
      });
    }
  }, 5 * 60 * 1000);
});

// Helper: image sizes
const imageSizesDict = {
  '320': { quality: 85, width: 320 },
  '638': { quality: 85, width: 638 },
  '2048': { quality: 75, width: 2048 },
};

/**
 * (1) /api/get-slides
 *     Scrape SlideShare for the slides preview
 */
app.post('/api/get-slides', async (req, res) => {
  try {
    const { slideshareUrl } = req.body;
    if (!slideshareUrl) {
      return res.status(400).json({ error: 'No SlideShare URL provided.' });
    }

    const response = await axios.get(slideshareUrl);
    const html = response.data;
    const $ = cheerio.load(html);

    let nextDataScript = null;
    $('script').each((i, script) => {
      if ($(script).attr('id') === '__NEXT_DATA__') {
        nextDataScript = $(script).html();
      }
    });
    if (!nextDataScript) {
      return res.status(404).json({
        error: 'Could not find __NEXT_DATA__ script tag in the SlideShare page.',
      });
    }

    const jsonData = JSON.parse(nextDataScript);
    const props = jsonData?.props?.pageProps || {};
    const slideshow = props?.slideshow || {};
    const slides = slideshow?.slides || {};

    const totalSlides = slideshow.totalSlides || 0;
    const imageLocation = slides.imageLocation || '';
    const imageTitle = slides.title || '';
    const host = slides.host || '';

    const previewSize = imageSizesDict['320'];
    const { width: previewWidth, quality: previewQuality } = previewSize;

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
  } catch (error) {
    console.error('Error in /api/get-slides:', error.message);
    return res.status(500).json({ error: error.message });
  }
});

/**
 * (2) /api/generate-file
 *     Generate the requested file (PDF, PPT, ZIP, etc) from the selected slides
 */
app.post('/api/generate-file', async (req, res) => {
  let downloadedFiles = []; // We'll keep track in case we need cleanup
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
      return res.status(400).json({ error: 'Invalid resolution chosen.' });
    }

    const { width, quality } = sizeEntry;

    // Build final slide image URLs
    const slideImagesFinal = [];
    for (const index of selectedIndices) {
      const realNum = index + 1;
      const url = `${host}/${imageLocation}/${quality}/${imageTitle}-${realNum}-${width}.jpg`;
      slideImagesFinal.push(url);
    }

    // Download each image to a temp folder
    const tempFolder = path.join(__dirname, 'temp_slides');
    if (!fs.existsSync(tempFolder)) {
      fs.mkdirSync(tempFolder, { recursive: true });
    }

    const isPng = outputFormat === 'png';
    for (let i = 0; i < slideImagesFinal.length; i++) {
      const imgUrl = slideImagesFinal[i];
      const filename = `slide_${i}.${isPng ? 'png' : 'jpg'}`;
      const filepath = path.join(tempFolder, filename);

      const response = await axios.get(imgUrl, { responseType: 'arraybuffer' });
      const buffer = Buffer.from(response.data, 'binary');

      if (isPng) {
        // Convert to PNG with sharp
        await sharp(buffer).png().toFile(filepath);
      } else {
        // Just save as JPEG
        fs.writeFileSync(filepath, buffer);
      }
      downloadedFiles.push(filepath);
    }

    // Prepare final output filename
    let extension = outputFormat;
    if (outputFormat === 'jpg' || outputFormat === 'png') {
      // user asked for jpg/png, we zip them up
      extension = 'zip';
    }
    const finalFilename = `slides_${Date.now()}.${extension}`;
    const finalFilePath = path.join(downloadsDir, finalFilename);

    // Build final file
    switch (outputFormat) {
      case 'jpg':
      case 'png':
      case 'zip': {
        // Zip up the images
        const zip = new AdmZip();
        for (const file of downloadedFiles) {
          zip.addLocalFile(file);
        }
        fs.writeFileSync(finalFilePath, zip.toBuffer());
        break;
      }
      case 'pdf': {
        const pdfDoc = new PDFDocument({ autoFirstPage: false });
        const stream = fs.createWriteStream(finalFilePath);
        pdfDoc.pipe(stream);

        for (const file of downloadedFiles) {
          const image = pdfDoc.openImage(file);
          pdfDoc.addPage({ size: [image.width, image.height] });
          pdfDoc.image(image, 0, 0);
        }
        pdfDoc.end();
        await new Promise((resolve) => stream.on('finish', resolve));
        break;
      }
      case 'pptx': {
        const pptx = new PptxGenJS();
        for (const file of downloadedFiles) {
          const slide = pptx.addSlide();
          const fileBuffer = fs.readFileSync(file);
          const base64data = fileBuffer.toString('base64');
          slide.addImage({
            data: `data:image/${isPng ? 'png' : 'jpeg'};base64,${base64data}`,
            x: 0, y: 0, w: '100%', h: '100%',
          });
        }
        const b64 = await pptx.write('base64');
        const pptxBuffer = Buffer.from(b64, 'base64');
        fs.writeFileSync(finalFilePath, pptxBuffer);
        break;
      }
      default: {
        cleanupTempFiles(downloadedFiles);
        return res.status(400).json({ error: 'Invalid output format.' });
      }
    }

    // Cleanup downloaded images
    cleanupTempFiles(downloadedFiles);

    // Build the final download URL using our BASE_URL
    const downloadUrl = `${baseUrl}/downloads/${finalFilename}`;
    return res.json({ downloadUrl });
  } catch (error) {
    console.error('Error in /api/generate-file:', error.message);
    cleanupTempFiles(downloadedFiles);
    return res.status(500).json({ error: error.message });
  }
});

// Utility to delete temp files
function cleanupTempFiles(files) {
  for (const filePath of files) {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  }
}

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
