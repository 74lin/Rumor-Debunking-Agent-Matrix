#!/usr/bin/env node
import { chromium } from 'playwright';
import { PDFDocument } from 'pdf-lib';
import fs from 'fs';
import path from 'path';

const args = { width: 1920, height: 1080 };
const a = process.argv.slice(2);
for (let i = 0; i < a.length; i += 2) {
  const k = a[i].replace(/^--/, '');
  args[k] = a[i + 1];
}
if (!args.slides || !args.out) {
  console.error('用法: node export_deck_pdf.mjs --slides <dir> --out <file.pdf>');
  process.exit(1);
}

// Resolve paths
const slidesDir = path.resolve(args.slides);
const outFile = path.resolve(args.out);

console.log('Slides dir:', slidesDir);
console.log('Output file:', outFile);

const files = fs.readdirSync(slidesDir)
  .filter(f => f.endsWith('.html'))
  .sort();

console.log(`Found ${files.length} slides`);

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.setViewportSize({ width: 1920, height: 1080 });

  const pdfDoc = await PDFDocument.create();

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const fullPath = path.join(slidesDir, f);
    console.log(`  [${i + 1}/${files.length}] ${f}`);
    await page.goto('file://' + fullPath);
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(500);
    const pdfBytes = await page.pdf({ width: 1920, height: 1080, printBackground: true });
    const slidePdf = await PDFDocument.load(pdfBytes);
    const [copied] = await pdfDoc.copyPages(slidePdf, [0]);
    pdfDoc.addPage(copied);
  }

  await browser.close();

  // Save with explicit file write
  const pdfBytesOut = await pdfDoc.save();
  fs.writeFileSync(outFile, pdfBytesOut);

  console.log(`\nDone: ${outFile} (${pdfBytesOut.length} bytes)`);
  console.log('File exists:', fs.existsSync(outFile));
})().catch(e => {
  console.error('Error:', e);
  process.exit(1);
});