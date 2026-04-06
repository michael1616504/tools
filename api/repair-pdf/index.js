/**
 * Azure Function: /api/repair-pdf
 *
 * Accepts a base64-encoded PDF and returns a repaired version.
 *
 * Repair strategies (applied in order):
 *  1. Strip corrupt data appended after %%EOF (iLovePDF, Canon scanner merges)
 *  2. Detect cross-reference scrambling — if truncation loses pages, attempt
 *     page-level extraction from each valid sub-document embedded in the file
 *  3. Re-linearise using pdf-lib to produce a clean single-revision file
 *
 * Request:  POST { file_base64: string, filename?: string }
 * Response: { repaired: bool, pages: number, original_pages: number|null,
 *             file_base64: string, filename: string, message: string }
 */

const { PDFDocument } = require('pdf-lib');

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
    const { file_base64, filename } = req.body || {};
    if (!file_base64) {
      context.res = { status: 400, headers, body: JSON.stringify({ error: 'Missing file_base64' }) };
      return;
    }

    const srcBuf = Buffer.from(file_base64, 'base64');
    const name = filename || 'repaired.pdf';

    // ── Step 1: Find all %%EOF positions ──
    const eofs = findEOFs(srcBuf);
    const isLin = isLinearized(srcBuf);

    // Try opening the full file first
    let fullDoc = null;
    let fullPages = 0;
    try {
      fullDoc = await PDFDocument.load(srcBuf, { ignoreEncryption: true });
      fullPages = fullDoc.getPageCount();
    } catch (e) {
      // Can't open full file — will try sub-documents
    }

    // If only 0 or 1 %%EOF, nothing to repair
    if (eofs.length <= 1 && fullDoc) {
      const outBytes = await fullDoc.save({ useObjectStreams: false });
      context.res = {
        status: 200, headers,
        body: JSON.stringify({
          repaired: false,
          pages: fullPages,
          original_pages: fullPages,
          file_base64: Buffer.from(outBytes).toString('base64'),
          filename: name,
          message: 'File is clean — no repair needed.'
        })
      };
      return;
    }

    // ── Step 2: Try truncation at the "real" %%EOF ──
    const realIdx = (isLin && eofs.length >= 2) ? 1 : 0;
    const realEnd = eofs[realIdx] + 5;

    // Skip whitespace
    let a = realEnd;
    while (a < srcBuf.length && (srcBuf[a] === 0x0A || srcBuf[a] === 0x0D || srcBuf[a] === 0x20)) a++;

    let needsTruncation = false;
    if (realIdx < eofs.length - 1 && a < srcBuf.length - 10) {
      const afterStr = srcBuf.slice(a, Math.min(a + 40, srcBuf.length)).toString('latin1');
      const validUpdate = /^\d+\s+\d+\s+obj/.test(afterStr) || afterStr.startsWith('xref');
      if (!validUpdate) needsTruncation = true;
      // Also check for known corruption markers
      const sample = srcBuf.slice(realEnd, Math.min(realEnd + 50000, srcBuf.length)).toString('latin1');
      if (sample.includes('iLovePDF')) needsTruncation = true;
    }

    if (needsTruncation) {
      // Truncate to the real %%EOF
      let end = realEnd;
      while (end < srcBuf.length && (srcBuf[end] === 0x0A || srcBuf[end] === 0x0D)) {
        end++;
        if (end - realEnd > 2) break;
      }
      const truncated = srcBuf.slice(0, end);

      try {
        const truncDoc = await PDFDocument.load(truncated, { ignoreEncryption: true });
        const truncPages = truncDoc.getPageCount();

        // ── Step 3: Check if truncation lost pages compared to full file ──
        // If so, try extracting from each sub-document
        if (fullDoc && fullPages > truncPages) {
          // Try to extract valid pages from sub-documents
          const extracted = await extractSubDocPages(srcBuf, eofs, isLin);
          if (extracted && extracted.pages > truncPages) {
            // Sub-document extraction recovered more pages
            context.res = {
              status: 200, headers,
              body: JSON.stringify({
                repaired: true,
                pages: extracted.pages,
                original_pages: fullPages,
                file_base64: Buffer.from(extracted.bytes).toString('base64'),
                filename: name,
                message: `Deep repair: recovered ${extracted.pages} pages from ${eofs.length} sub-documents (full file had ${fullPages} pages with corrupt cross-refs).`
              })
            };
            return;
          }
        }

        // Truncation is the best we can do
        const outBytes = await truncDoc.save({ useObjectStreams: false });
        context.res = {
          status: 200, headers,
          body: JSON.stringify({
            repaired: true,
            pages: truncPages,
            original_pages: fullPages || null,
            file_base64: Buffer.from(outBytes).toString('base64'),
            filename: name,
            message: `Repaired: stripped ${(srcBuf.length - end).toLocaleString()} bytes of corrupt appended data. ${truncPages} page(s) recovered.`
          })
        };
        return;

      } catch (truncErr) {
        // Truncated version won't open either — try sub-document extraction
        const extracted = await extractSubDocPages(srcBuf, eofs, isLin);
        if (extracted && extracted.pages > 0) {
          context.res = {
            status: 200, headers,
            body: JSON.stringify({
              repaired: true,
              pages: extracted.pages,
              original_pages: fullPages || null,
              file_base64: Buffer.from(extracted.bytes).toString('base64'),
              filename: name,
              message: `Deep repair: recovered ${extracted.pages} pages from embedded sub-documents.`
            })
          };
          return;
        }
      }
    }

    // File opens fine and doesn't need truncation — re-save clean
    if (fullDoc) {
      const outBytes = await fullDoc.save({ useObjectStreams: false });
      context.res = {
        status: 200, headers,
        body: JSON.stringify({
          repaired: false,
          pages: fullPages,
          original_pages: fullPages,
          file_base64: Buffer.from(outBytes).toString('base64'),
          filename: name,
          message: 'File is clean — no repair needed.'
        })
      };
      return;
    }

    // Nothing worked
    context.res = {
      status: 422, headers,
      body: JSON.stringify({ error: 'Could not repair this PDF. The file may be too severely damaged.' })
    };

  } catch (e) {
    context.log.error('repair-pdf error:', e);
    context.res = {
      status: 500, headers,
      body: JSON.stringify({ error: 'Internal error during repair.', detail: e.message })
    };
  }
};

/* ── Helpers ── */

function findEOFs(buf) {
  const eofs = [];
  for (let i = 0; i <= buf.length - 5; i++) {
    if (buf[i] === 0x25 && buf[i+1] === 0x25 && buf[i+2] === 0x45 && buf[i+3] === 0x4F && buf[i+4] === 0x46)
      eofs.push(i);
  }
  return eofs;
}

function isLinearized(buf) {
  const head = buf.slice(0, Math.min(1024, buf.length)).toString('latin1');
  return /\/Linearized\s/.test(head);
}

/**
 * Extract valid pages from sub-documents embedded in a corrupt file.
 * Each %%EOF boundary is tested as a potential valid sub-document.
 * Pages are de-duplicated by taking only the LAST (most complete) valid version.
 */
async function extractSubDocPages(buf, eofs, isLin) {
  // Try each %%EOF from last to first, find the best valid sub-document
  let bestDoc = null;
  let bestPages = 0;
  let bestBytes = null;

  for (let idx = eofs.length - 1; idx >= 0; idx--) {
    let end = eofs[idx] + 5;
    while (end < buf.length && (buf[end] === 0x0A || buf[end] === 0x0D)) {
      end++;
      if (end - (eofs[idx] + 5) > 2) break;
    }
    const chunk = buf.slice(0, end);
    try {
      const doc = await PDFDocument.load(chunk, { ignoreEncryption: true });
      const pages = doc.getPageCount();
      if (pages > bestPages) {
        bestPages = pages;
        bestDoc = doc;
        bestBytes = chunk;
      }
    } catch (e) {
      // This sub-document isn't valid, skip
    }
  }

  if (!bestDoc || bestPages === 0) return null;

  // Re-save as a clean PDF
  const outDoc = await PDFDocument.create();
  const copied = await outDoc.copyPages(bestDoc, bestDoc.getPageIndices());
  copied.forEach(p => outDoc.addPage(p));
  const outBytes = await outDoc.save({ useObjectStreams: false });

  return { pages: bestPages, bytes: outBytes };
}
