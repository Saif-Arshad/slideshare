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

// Make sure this is BEFORE the routes:
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const downloadsDir = path.join(__dirname, 'downloads');
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir);
}
app.use(
  '/downloads',
  cors(), 
  express.static(downloadsDir)
);

// Available sizes for "resolution"
const imageSizesDict = {
  '320': { quality: 85, width: 320 },
  '638': { quality: 85, width: 638 },
  '2048': { quality: 75, width: 2048 },
};

/**
 * 1) /api/get-slides: fetch slideshow data + preview URLs
 *    from SlideShare page HTML.
 */
app.post('/api/get-slides', async (req, res) => {
  try {
    const { slideshareUrl } = req.body;
    if (!slideshareUrl) {
      return res.status(400).json({ error: 'No SlideShare URL provided.' });
    }

    // 1) Grab HTML of the SlideShare page
    const response = await axios.get(slideshareUrl);
    const html = response.data;

    // 2) Use cheerio to find the __NEXT_DATA__ script tag
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

    // 3) Parse the JSON to extract the slideshow info
    const jsonData = JSON.parse(nextDataScript);
    const props = jsonData?.props?.pageProps || {};
    const slideshow = props?.slideshow || {};
    const slides = slideshow?.slides || {};

    const totalSlides = slideshow.totalSlides || 0;
    const imageLocation = slides.imageLocation || '';
    const imageTitle = slides.title || '';
    const host = slides.host || '';

    // We'll provide "preview" images at 320px
    const previewSize = imageSizesDict['320']; // Low res
    const { width: previewWidth, quality: previewQuality } = previewSize;

    // 4) Build array of preview image URLs
    const slideImagesPreview = [];
    for (let i = 1; i <= totalSlides; i++) {
      const url = `${host}/${imageLocation}/${previewQuality}/${imageTitle}-${i}-${previewWidth}.jpg`;
      slideImagesPreview.push(url);
    }

    // 5) Return the data to the frontend
    return res.json({
      totalSlides,
      slideImagesPreview,
      slideshowInfo: {
        host,
        imageLocation,
        imageTitle,
      },
    });
  } catch (error) {
    console.error('Error in /api/get-slides:', error.message);
    return res.status(500).json({ error: error.message });
  }
});

/**
 * 2) /api/generate-file:
 *    - Receives slideshow info, resolution, format, and selectedIndices
 *    - Builds the final image URLs
 *    - Downloads images
 *    - Converts to PDF, PPTX, or zipped images
 *    - Saves final file in /downloads
 *    - Returns a download URL
 */
app.post('/api/generate-file', async (req, res) => {
  try {
    // Debug: log out what the server sees in req.body
    console.log('DEBUG /api/generate-file body:', req.body);

    const { slideshowInfo, resolution, outputFormat, selectedIndices } = req.body;

    if (!slideshowInfo || !resolution || !outputFormat || !selectedIndices) {
      return res.status(400).json({ error: 'Missing required data.' });
    }
    if (!Array.isArray(selectedIndices) || selectedIndices.length === 0) {
      return res.status(400).json({ error: 'No slides selected.' });
    }

    // Extract info for building final URLs
    const { host, imageLocation, imageTitle } = slideshowInfo;
    const sizeEntry = imageSizesDict[resolution];
    if (!sizeEntry) {
      return res.status(400).json({ error: 'Invalid resolution chosen.' });
    }

    const { width, quality } = sizeEntry;

    // Reconstruct the final image URLs (JPG by default, unless converting to PNG)
    const slideImagesFinal = [];
    for (const index of selectedIndices) {
      // For example, if index=0 => slide #1
      const realNum = index + 1;
      const url = `${host}/${imageLocation}/${quality}/${imageTitle}-${realNum}-${width}.jpg`;
      slideImagesFinal.push(url);
    }

    // Create a temp folder for downloads
    const tempFolder = path.join(__dirname, 'temp_slides');
    if (!fs.existsSync(tempFolder)) {
      fs.mkdirSync(tempFolder, { recursive: true });
    }

    // Download each slide image to temp folder
    const downloadedFiles = [];
    // If outputFormat === 'png', we convert images to PNG. Otherwise keep as JPG.
    const isPng = outputFormat === 'png';

    for (let i = 0; i < slideImagesFinal.length; i++) {
      const imgUrl = slideImagesFinal[i];
      // Name the file in local temp
      const filename = `slide_${i}.${isPng ? 'png' : 'jpg'}`;
      const filepath = path.join(tempFolder, filename);

      // Download image (as a buffer)
      const response = await axios.get(imgUrl, { responseType: 'arraybuffer' });
      const buffer = Buffer.from(response.data, 'binary');

      // Possibly convert to PNG
      if (isPng) {
        await sharp(buffer).png().toFile(filepath);
      } else {
        fs.writeFileSync(filepath, buffer);
      }

      downloadedFiles.push(filepath);
    }

    // Build final output file
    let finalFilename = '';
    let extension = outputFormat;
    if (outputFormat === 'jpg' || outputFormat === 'png') {
      // We'll zip them if user asked for 'jpg' or 'png'
      extension = 'zip';
    }
    finalFilename = `slides_${Date.now()}.${extension}`;
    const finalFilePath = path.join(downloadsDir, finalFilename);

    // Generate final file based on format
    switch (outputFormat) {
      // If user wants JPG or PNG, we make a ZIP of images
      case 'jpg':
      case 'png':
      case 'zip': {
        const zip = new AdmZip();
        for (const file of downloadedFiles) {
          zip.addLocalFile(file);
        }
        fs.writeFileSync(finalFilePath, zip.toBuffer());
        break;
      }

      // PDF
      case 'pdf': {
        const pdfDoc = new PDFDocument({ autoFirstPage: false });
        const stream = fs.createWriteStream(finalFilePath);
        pdfDoc.pipe(stream);

        for (const file of downloadedFiles) {
          const image = pdfDoc.openImage(file);
          // Make a new page matching the image dimensions
          pdfDoc.addPage({ size: [image.width, image.height] });
          pdfDoc.image(image, 0, 0);
        }
        pdfDoc.end();

        // Wait for PDF to finish writing
        await new Promise((resolve) => stream.on('finish', resolve));
        break;
      }

      // PPTX
      case 'pptx': {
        const pptx = new PptxGenJS();
        for (const file of downloadedFiles) {
          const slide = pptx.addSlide();
          const fileBuffer = fs.readFileSync(file);
          const base64data = fileBuffer.toString('base64');
          slide.addImage({
            data: `data:image/${isPng ? 'png' : 'jpeg'};base64,${base64data}`,
            x: 0,
            y: 0,
            w: '100%',
            h: '100%',
          });
        }
        const b64 = await pptx.write('base64');
        const pptxBuffer = Buffer.from(b64, 'base64');
        fs.writeFileSync(finalFilePath, pptxBuffer);
        break;
      }

      default:
        // Cleanup & return error
        cleanupTempFiles(downloadedFiles);
        return res.status(400).json({ error: 'Invalid output format.' });
    }

    // Cleanup downloaded images from temp folder
    cleanupTempFiles(downloadedFiles);

    // Build the final download link
    const downloadUrl = `${req.protocol}://${req.get('host')}/downloads/${finalFilename}`;
    return res.json({ downloadUrl });
  } catch (error) {
    console.error('Error in /api/generate-file:', error.message);
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

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
