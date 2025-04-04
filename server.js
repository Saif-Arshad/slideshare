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

// 1) Load environment variables
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

// 3) A folder to hold final downloadable files
const downloadsDir = path.join(__dirname, 'downloads');
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir);
}

// 4) Serve the final file from GET /downloads/:filename
app.get('/downloads/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(downloadsDir, filename);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }

  // Set appropriate headers
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

  // Pipe the file
  const fileStream = fs.createReadStream(filePath);
  fileStream.pipe(res);

  // Clean up the file after 3 minutes
  setTimeout(() => {
    if (fs.existsSync(filePath)) {
      fs.unlink(filePath, (err) => {
        if (err) console.error(`Failed to delete ${filePath}:`, err);
      });
    }
  }, 3 * 60 * 1000);
});

// 5) Dictionary of possible image sizes for SlideShare
const imageSizesDict = {
  '320': { quality: 85, width: 320 },
  '638': { quality: 85, width: 638 },
  '2048': { quality: 75, width: 2048 },
};

/**
 * 6) /api/get-slides
 *    Scrapes the SlideShare page and finds the array of preview images
 */
app.post('/api/get-slides', async (req, res) => {
  try {
    const { slideshareUrl } = req.body;
    if (!slideshareUrl) {
      return res.status(400).json({ error: 'No SlideShare URL provided.' });
    }

    // Fetch the SlideShare HTML
    const response = await axios.get(slideshareUrl);
    const html = response.data;
    const $ = cheerio.load(html);

    // Find the __NEXT_DATA__ script
    let nextDataScript = null;
    $('script').each((i, script) => {
      if ($(script).attr('id') === '__NEXT_DATA__') {
        nextDataScript = $(script).html();
      }
    });
    if (!nextDataScript) {
      return res.status(404).json({
        error: 'Could not find __NEXT_DATA__ script tag on that SlideShare page.'
      });
    }

    const jsonData = JSON.parse(nextDataScript);
    const props = jsonData?.props?.pageProps || {};
    const slideshow = props?.slideshow || {};
    const slides = slideshow?.slides || {};

    // Basic info
    const totalSlides = slideshow.totalSlides || 0;
    const imageLocation = slides.imageLocation || '';
    const imageTitle = slides.title || '';
    const host = slides.host || '';

    // Build preview URLs using the low-res "320" preset
    const { width: previewWidth, quality: previewQuality } = imageSizesDict['320'];
    const slideImagesPreview = [];
    for (let i = 1; i <= totalSlides; i++) {
      const url = `${host}/${imageLocation}/${previewQuality}/${imageTitle}-${i}-${previewWidth}.jpg`;
      slideImagesPreview.push(url);
    }

    // Return them to the client
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
 *    Download each selected slide in the chosen resolution,
 *    forcibly convert to a known format (JPEG or PNG),
 *    then build final PDF / PPTX / ZIP output.
 */
app.post('/api/generate-file', async (req, res) => {
  let downloadedFiles = []; // For cleanup if something goes wrong

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

    // Build final image URLs
    const slideImagesFinal = selectedIndices.map((idx) => {
      const realSlideNum = idx + 1; // user slides are 0-based in array
      return `${host}/${imageLocation}/${quality}/${imageTitle}-${realSlideNum}-${width}.jpg`;
    });

    // We'll store them in a temp folder
    const tempFolder = path.join(__dirname, 'temp_slides');
    if (!fs.existsSync(tempFolder)) {
      fs.mkdirSync(tempFolder, { recursive: true });
    }

    // For PDFKit or PPTX, we typically want .jpg or .png
    // We'll convert everything to the user's desired format:
    // - If user wants "png", then we'll store as PNG.
    // - Otherwise, we'll store as JPEG (pdf, pptx, zip, jpg => all become .jpg).
    const wantsPng = (outputFormat.toLowerCase() === 'png');

    // 8) Download each slide, forcibly convert format with Sharp
    for (let i = 0; i < slideImagesFinal.length; i++) {
      const imgUrl = slideImagesFinal[i];

      // Decide on extension for the local file
      const ext = wantsPng ? 'png' : 'jpg';
      const filename = `slide_${i}.${ext}`;
      const filepath = path.join(tempFolder, filename);

      // Download raw arraybuffer
      const response = await axios.get(imgUrl, { responseType: 'arraybuffer' });
      if (response.status !== 200) {
        throw new Error(`Failed to download image: HTTP ${response.status}`);
      }
      const originalBuffer = Buffer.from(response.data);

      // Convert buffer to the correct format with Sharp
      let finalBuffer;
      if (wantsPng) {
        finalBuffer = await sharp(originalBuffer).png().toBuffer();
      } else {
        finalBuffer = await sharp(originalBuffer).jpeg().toBuffer();
      }

      // Save that standardized image
      fs.writeFileSync(filepath, finalBuffer);
      downloadedFiles.push(filepath);
    }

    // 9) Build the final file
    //    If user asked for "jpg" or "png" specifically, we actually zip them up anyway.
    //    Because we only return a single link to download everything.
    let finalExt = outputFormat.toLowerCase();
    if (finalExt === 'jpg' || finalExt === 'png') {
      // We'll do a .zip so all images are in one file
      finalExt = 'zip';
    }

    const finalFilename = `slides_${Date.now()}.${finalExt}`;
    const finalFilePath = path.join(downloadsDir, finalFilename);

    // Switch on the chosen output
    switch (outputFormat.toLowerCase()) {
      case 'jpg':
      case 'png':
      case 'zip': {
        // Just zip up the images (which are all either .jpg or .png)
        const zip = new AdmZip();
        for (const file of downloadedFiles) {
          zip.addLocalFile(file);
        }
        fs.writeFileSync(finalFilePath, zip.toBuffer());
        break;
      }

      case 'pdf': {
        // Create a PDF with each image as a page
        const pdfDoc = new PDFDocument({ autoFirstPage: false });
        const writeStream = fs.createWriteStream(finalFilePath);
        pdfDoc.pipe(writeStream);

        for (const file of downloadedFiles) {
          const image = pdfDoc.openImage(file);
          pdfDoc.addPage({ size: [image.width, image.height] });
          pdfDoc.image(image, 0, 0);
        }
        pdfDoc.end();

        // Wait for the PDF to be fully written
        await new Promise((resolve) => writeStream.on('finish', resolve));
        break;
      }

      case 'pptx': {
        // Create a PPTX, each slide is one image
        const pptx = new PptxGenJS();
        for (const file of downloadedFiles) {
          const slide = pptx.addSlide();
          const fileBuffer = fs.readFileSync(file);
          const b64 = fileBuffer.toString('base64');

          // If we used PNG, the data type is image/png; else image/jpeg
          slide.addImage({
            data: `data:image/${wantsPng ? 'png' : 'jpeg'};base64,${b64}`,
            x: 0, y: 0, w: '100%', h: '100%'
          });
        }
        const b64pptx = await pptx.write('base64');
        const pptxBuffer = Buffer.from(b64pptx, 'base64');
        fs.writeFileSync(finalFilePath, pptxBuffer);
        break;
      }

      default:
        // Unknown format, cleanup and bail
        cleanupTempFiles(downloadedFiles);
        return res.status(400).json({ error: `Invalid output format: ${outputFormat}` });
    }

    // 10) Cleanup the temp images
    cleanupTempFiles(downloadedFiles);

    // 11) Respond with a direct download URL
    const downloadUrl = `${baseUrl}/downloads/${finalFilename}`;
    return res.json({ downloadUrl });
  } catch (err) {
    console.error('Error in /api/generate-file:', err);
    // Cleanup any partial downloads if we crashed
    cleanupTempFiles(downloadedFiles);
    return res.status(500).json({ error: err.message });
  }
});

/**
 * Utility to remove the downloaded temp images
 */
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
