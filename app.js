require('dotenv').config(); // Load environment variables from .env file

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs-extra');
const { PDFDocument } = require('pdf-lib');
const sharp = require('sharp');
const mammoth = require('mammoth');
const puppeteer = require('puppeteer');
const htmlPdf = require('html-pdf-node');
const Jimp = require('jimp');
const heicConvert = require('heic-convert');
const { Document, Packer, Paragraph, ImageRun } = require('docx');
const PptxGenJS = require('pptxgenjs');
const imageJs = require('image-js'); // Add image-js for DDS support
const { exec } = require('child_process'); // Add child_process for running shell commands
const cors = require('cors');

const app = express();
const port = 8000;

const isProduction = process.env.NODE_ENV === 'production';
const liveStoragePath = process.env.LIVE_STORAGE_PATH || '/tmp';
const localStoragePath = process.env.LOCAL_STORAGE_PATH || './upload';
const storagePath = isProduction ? liveStoragePath : localStoragePath;

// Ensure storage path exists
fs.ensureDirSync(storagePath);

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, storagePath);
  },
  filename: function (req, file, cb) {
    cb(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname));
  },
});

const upload = multer({ storage: storage }).single('image');

app.use(cors());
app.use('/download', express.static(storagePath));

app.post('/convert', upload, async (req, res) => {
  try {
    let toFormat = req.body.to;

    console.log('🚀 ~ app.post ~ toFormat:', toFormat);
    console.log('🚀 ~ app.post ~ req.file.filename:', req.file.filename);

    if (!req.file) {
      return res.status(400).send({ error: 'No file uploaded' });
    }

    const inputPath = path.join(storagePath, req.file.filename);
    const outputFilename = `${req.file.originalname.split('.')[0]}-${Date.now()}.${toFormat}`;
    const outputPath = path.join(storagePath, outputFilename);
    const downloadLink = `${req.protocol}://${req.get('host')}/download/${outputFilename}`;

    if (req.file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      if (toFormat === 'html') {
        const result = await mammoth.convertToHtml({ path: inputPath });
        fs.writeFileSync(outputPath, result.value);
      } else if (toFormat === 'pdf') {
        const result = await mammoth.convertToHtml({ path: inputPath });
        const pdfBuffer = await convertHtmlToPdf(result.value);
        fs.writeFileSync(outputPath, pdfBuffer);
      } else if (['jpeg', 'jpg', 'png', 'gif', 'tiff', 'webp', 'hdr', 'avif'].includes(toFormat)) {
        const result = await mammoth.convertToHtml({ path: inputPath });
        const imageBuffer = await convertHtmlToImage(result.value, toFormat);
        fs.writeFileSync(outputPath, imageBuffer);
      } else {
        throw new Error(`Unsupported conversion format: ${toFormat}`);
      }
    } else {
      if (toFormat === 'pdf') {
        const pdfDoc = await PDFDocument.create();
        const image = sharp(inputPath);
        const metadata = await image.metadata();
        const imageBuffer = await image.toBuffer();
        let embeddedImage;

        switch (metadata.format) {
          case 'jpeg':
          case 'jpg':
            embeddedImage = await pdfDoc.embedJpg(imageBuffer);
            break;
          case 'png':
          case 'gif':
          case 'tiff':
          case 'webp':
          case 'avif':
            embeddedImage = await pdfDoc.embedPng(imageBuffer);
            break;
          case 'hdr':
            const hdrImage = await Jimp.read(inputPath);
            const hdrBuffer = await hdrImage.getBufferAsync(Jimp.MIME_PNG);
            embeddedImage = await pdfDoc.embedPng(hdrBuffer);
            break;
          case 'dds':
            const ddsImage = await Jimp.read(inputPath);
            const ddsBuffer = await ddsImage.getBufferAsync(Jimp.MIME_PNG); // Convert to PNG buffer
            const { data, width, height } = await imageJs.Image.load(ddsBuffer); // Load image with image-js
            const encodedDdsBuffer = imageJs.encode(imageJs.getFormat('DDS'), {
              width,
              height,
              data,
            }); // Encode to DDS format
            fs.writeFileSync(outputPath, Buffer.from(encodedDdsBuffer));
            break;
          case 'HEIC':
          case 'HEIF':
            const heicImageBuffer = await fs.readFile(inputPath);
            const heicImage = await heicConvert({
              buffer: heicImageBuffer,
              format: 'HEIF', // Convert HEIC/HEIF to JPEG for embedding in PDF
            });
            embeddedImage = await pdfDoc.embedJpg(heicImage);
            break;
          default:
            throw new Error(`Unsupported image format: ${metadata.format}`);
        }
        const page = pdfDoc.addPage([embeddedImage.width, embeddedImage.height]);
        page.drawImage(embeddedImage, {
          x: 0,
          y: 0,
          width: embeddedImage.width,
          height: embeddedImage.height,
        });
        const pdfBytes = await pdfDoc.save();
        fs.writeFileSync(outputPath, pdfBytes);
      } else if (toFormat === 'docx') {
        const imageBuffer = await fs.readFile(inputPath);
        const doc = new Document({
          sections: [
            {
              properties: {},
              children: [
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: imageBuffer,
                      transformation: {
                        width: 600,
                        height: 400,
                      },
                    }),
                  ],
                }),
              ],
            },
          ],
        });
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(outputPath, buffer);
      } else if (toFormat === 'pptx') {
        const imageBuffer = await fs.readFile(inputPath);
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addImage({
          data: `data:image/${req.file.mimetype.split('/')[1]};base64,${imageBuffer.toString('base64')}`,
          x: 0.5,
          y: 0.5,
          w: 8.5,
          h: 6,
        });
        await pptx.writeFile({ fileName: outputPath });
      } else if (toFormat === 'odp') {
        const command = `libreoffice --headless --convert-to odp "${inputPath}" --outdir "${path.dirname(outputPath)}"`;
        exec(command, (err, stdout, stderr) => {
          if (err) {
            throw new Error(`Error converting to ODP: ${err.message}`);
          }
          console.log(stdout);
        });
      } else if (['HEIC', 'HEIF'].includes(toFormat)) {
        const image = await Jimp.read(inputPath);
        const buffer = await image.getBufferAsync(Jimp.MIME_PNG);
        const heicImage = await heicConvert({
          buffer,
          format: 'HEIC',
        });
        fs.writeFileSync(outputPath, heicImage);
      } else if (toFormat === 'rgb') {
        const image = await Jimp.read(inputPath);
        await image.writeAsync(outputPath); // Writes the image as RGB format
      } else if (toFormat === 'jp2') {
        await sharp(inputPath).toFormat('jp2').toFile(outputPath);
      } else if (toFormat === 'jfif') {
        await sharp(inputPath).toFormat('jpeg').toFile(outputPath);
      } else {
        await Jimp.read(inputPath)
          .then((image) => {
            return image.writeAsync(outputPath); // Writes the image in the requested format
          })
          .catch((err) => {
            throw new Error(`Error processing image: ${err.message}`);
          });
      }
    }

    setTimeout(() => {
      fs.remove(inputPath, (err) => {
        if (err) {
          console.error('Error deleting original file:', err);
        } else {
          console.log('Original file deleted successfully');
        }
      });
    }, 1000);

    return res.status(200).send({ message: 'Conversion successful', downloadLink: downloadLink });
  } catch (err) {
    console.log('***************************');
    console.error('Error during conversion:', err);
    res.status(500).send({ error: 'Conversion failed' });
  }
});

async function convertHtmlToPdf(html) {
  let file = { content: html };
  let options = { format: 'A4' };
  let pdfBuffer = await htmlPdf.generatePdf(file, options);
  return pdfBuffer;
}

async function convertHtmlToImage(html, format) {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.setContent(html);
  const imageBuffer = await page.screenshot({ fullPage: true, type: format });
  await browser.close();
  return imageBuffer;
}

app.listen(port, () => {
  console.log('Server listening successfully on port', port);
});
