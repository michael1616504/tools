/**
 * Azure Function: /api/convert-image
 *
 * Converts an image (HEIC, HEIF, JPG, PNG, WebP, TIFF) to a single-page PDF.
 * Uses `sharp` for image decoding (handles HEIC via libvips on Linux)
 * and `pdf-lib` to embed the image in a properly-sized PDF page.
 *
 * Request:  POST { file_base64: string, filename?: string, page_size?: "letter"|"a4"|"fit" }
 * Response: { file_base64: string, filename: string, width: number, height: number, message: string }
 */

const { PDFDocument } = require('pdf-lib');
let sharp;

module.exports = async function (context, req) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  if (req.method === 'OPTIONS') { context.res = { status: 204, headers }; return; }
  if (req.method !== 'POST') { context.res = { status: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) }; return; }

  try {
    // Lazy-load sharp (may not be available in all environments)
    if (!sharp) {
      try { sharp = require('sharp'); } catch (e) {
        context.res = {
          status: 501, headers,
          body: JSON.stringify({ error: 'Image conversion not available — sharp library not installed on this server.' })
        };
        return;
      }
    }

    const { file_base64, filename, page_size } = req.body || {};
    if (!file_base64) {
      context.res = { status: 400, headers, body: JSON.stringify({ error: 'Missing file_base64' }) };
      return;
    }

    const srcBuf = Buffer.from(file_base64, 'base64');
    const name = (filename || 'image').replace(/\.[^.]+$/, '') + '.pdf';
    const size = page_size || 'letter';

    // ── Convert image to JPEG using sharp ──
    // sharp handles HEIC, HEIF, WebP, TIFF, PNG, JPEG, AVIF, etc.
    const sharpImg = sharp(srcBuf);
    const metadata = await sharpImg.metadata();

    const imgW = metadata.width;
    const imgH = metadata.height;

    // Convert to JPEG for embedding (pdf-lib handles JPG and PNG)
    // Use PNG if source has alpha channel, otherwise JPEG for smaller size
    let imgBytes, imgFormat;
    if (metadata.hasAlpha) {
      imgBytes = await sharpImg.png().toBuffer();
      imgFormat = 'png';
    } else {
      imgBytes = await sharpImg.jpeg({ quality: 92 }).toBuffer();
      imgFormat = 'jpg';
    }

    // ── Build PDF with pdf-lib ──
    const pdfDoc = await PDFDocument.create();

    const img = imgFormat === 'png'
      ? await pdfDoc.embedPng(imgBytes)
      : await pdfDoc.embedJpg(imgBytes);

    // Page dimensions
    let pw, ph;
    if (size === 'fit') {
      pw = imgW;
      ph = imgH;
    } else if (size === 'a4') {
      pw = 595.28;
      ph = 841.89;
    } else {
      // letter
      pw = 612;
      ph = 792;
    }

    const page = pdfDoc.addPage([pw, ph]);

    if (size === 'fit') {
      page.drawImage(img, { x: 0, y: 0, width: pw, height: ph });
    } else {
      const margin = 18;
      const maxW = pw - margin * 2;
      const maxH = ph - margin * 2;
      const scale = Math.min(maxW / imgW, maxH / imgH, 1);
      const drawW = imgW * scale;
      const drawH = imgH * scale;
      const x = (pw - drawW) / 2;
      const y = (ph - drawH) / 2;
      page.drawImage(img, { x, y, width: drawW, height: drawH });
    }

    const pdfBytes = await pdfDoc.save({ useObjectStreams: false });

    context.res = {
      status: 200, headers,
      body: JSON.stringify({
        file_base64: Buffer.from(pdfBytes).toString('base64'),
        filename: name,
        width: imgW,
        height: imgH,
        message: `Converted ${imgW}×${imgH} ${metadata.format || 'image'} to PDF (${size} page).`
      })
    };

  } catch (e) {
    context.log.error('convert-image error:', e);

    // Provide a helpful message for HEIC if libvips doesn't support it
    const isHeicError = e.message && (e.message.includes('heif') || e.message.includes('heic'));
    const msg = isHeicError
      ? 'HEIC conversion failed — the server may need libheif support. Try converting in the browser instead.'
      : 'Error converting image.';

    context.res = {
      status: 500, headers,
      body: JSON.stringify({ error: msg, detail: e.message })
    };
  }
};
