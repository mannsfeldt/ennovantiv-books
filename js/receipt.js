// Receipt capture: shrink the photo, OCR it in the browser (Tesseract.js),
// and guess vendor / date / total. The image never leaves your browser
// except when it is saved to your own OneDrive/SharePoint receipts folder.

const TESSERACT_URL = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';

export async function compressImage(file, maxSide = 1800, quality = 0.82) {
  if (!file.type.startsWith('image/')) return file; // PDFs etc. upload as-is
  const img = await new Promise((res, rej) => {
    const i = new Image();
    i.onload = () => res(i);
    i.onerror = rej;
    i.src = URL.createObjectURL(file);
  });
  const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
  const c = document.createElement('canvas');
  c.width = Math.round(img.width * scale);
  c.height = Math.round(img.height * scale);
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
  return new Promise(res => c.toBlob(res, 'image/jpeg', quality));
}

function loadScript(src) {
  return new Promise((res, rej) => {
    if (window.Tesseract) return res();
    const s = document.createElement('script');
    s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not load OCR library'));
    document.head.appendChild(s);
  });
}

export async function readReceipt(blob, onProgress) {
  await loadScript(TESSERACT_URL);
  const { data } = await Tesseract.recognize(blob, 'eng', {
    logger: m => m.status === 'recognizing text' && onProgress?.(Math.round(m.progress * 100)),
  });
  return { text: data.text, ...parseReceipt(data.text) };
}

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const pad = n => String(n).padStart(2, '0');

export function parseReceipt(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const money = l => [...l.matchAll(/(?:\$\s*)?(\d{1,5}(?:,\d{3})*[.,]\d{2})\b/g)].map(m => Number(m[1].replace(/,(?=\d{3})/g, '').replace(',', '.')));

  // Total: prefer a "total" line that is not subtotal; else the largest amount
  let total = null;
  for (const l of lines) {
    if (/total|amount due|balance due|grand/i.test(l) && !/sub\s*-?total/i.test(l)) {
      const a = money(l);
      if (a.length) total = Math.max(total ?? 0, ...a);
    }
  }
  if (total === null) {
    const all = lines.flatMap(money);
    if (all.length) total = Math.max(...all);
  }

  // Date
  let date = null;
  const t = text.replace(/\s+/g, ' ');
  let m;
  if ((m = t.match(/\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/))) date = `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  else if ((m = t.match(/\b(\d{1,2})[-/.](\d{1,2})[-/.](20\d{2}|\d{2})\b/))) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    let a = Number(m[1]), b = Number(m[2]);
    // Canadian receipts are mixed; assume DD/MM if first > 12, else MM/DD
    const [mo, d] = a > 12 ? [b, a] : [a, b];
    date = `${y}-${pad(mo)}-${pad(d)}`;
  } else if ((m = t.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(20\d{2})/i))) {
    date = `${m[3]}-${pad(MONTHS[m[1].toLowerCase()])}-${pad(m[2])}`;
  }

  // Vendor: first line with real letters that isn't an address/phone
  const vendor = lines.find(l => /[a-z]{3,}/i.test(l) && !/\d{3}[- ]\d{4}|street|st\.|ave|road|rd\.|www|http|receipt|invoice/i.test(l)) || '';

  const hst = /hst|gst|tax/i.test(text);
  return { total, date, vendor: vendor.replace(/[^\w &'.-]/g, '').trim().slice(0, 40), hst };
}
