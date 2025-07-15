import PPTXGenJS from 'pptxgenjs';
import { put } from '@vercel/blob';
import fs from 'fs/promises';
import path from 'path';

// Change this if your logo is in a different folder:
const LOGO_PATH = './sp_global_logo.png';

function setA4Landscape(pptx) {
  pptx.layout = {
    name: 'A4',
    width: 29.7 / 2.54, // cm to inches
    height: 21 / 2.54
  };
}

async function addFooterWithLogo(pptx, slide, pageNum) {
  const logoWidth = 6.0 / 2.54; // cm to inches
  const logoHeight = 2.4 / 2.54;
  const footerY = 18.03 / 2.54;
  let leftTextX;
  try {
    await fs.access(LOGO_PATH);
    slide.addImage({ path: LOGO_PATH, x: 1/2.54, y: footerY, w: logoWidth, h: logoHeight });
    leftTextX = 7.2 / 2.54;
  } catch {
    // Logo not present
    leftTextX = 1 / 2.54;
  }
  slide.addText("Permission to reprint or distribute any content from this presentation requires the prior written approval of S&P Global Market Intelligence.",
    { x: leftTextX, y: footerY, w: 10/2.54, h: 1.5/2.54, fontSize: 10, color: '808080', align: 'left' }
  );
  slide.addText(pageNum.toString(),
    { x: (pptx.layout.width - 3/2.54), y: footerY, w: 2.5/2.54, h: 1.5/2.54, fontSize: 14, color: '808080', align: 'right' }
  );
}

function addDatesAvailableBox(slide, left, top, width, height, datesText) {
  slide.addShape(pptxgen.ShapeType.rect, { x:left, y:top, w:width, h:height, fill: { color: "e0eaee" }, line: { color: "e0eaee" } });
  slide.addText([
    { text: "Proposed Dates: ", options: { fontSize: 14, bold: true, color: "CC0000" } },
    { text: datesText, options: { fontSize: 14, color: "000000" } }
  ], { x: left, y: top, w: width, h: height, align: 'left', fontSize: 14, valign: 'middle' });
}

function createFrontPage(pptx, heading, dateToPresent) {
  const slide = pptx.addSlide();
  slide.addShape(pptxgen.ShapeType.rect,
    { x: 0, y: 0, w: pptx.layout.width, h: pptx.layout.height, fill: { color: '999999' } }
  );
  slide.addText("S&P Global", { x: 1/2.54, y: 1/2.54, w: 8/2.54, h: 2/2.54, fontSize: 20, bold: true, color: "FFFFFF" });
  slide.addText("Market Intelligence", { x: 1/2.54, y: 2.1/2.54, w: 8/2.54, h: 2/2.54, fontSize: 20, color: "FFFFFF" });
  slide.addText(heading, { x: 1/2.54, y: 6/2.54, w: 22/2.54, h: 4/2.54, fontSize: 54, bold: true, color: "FFFFFF" });
  slide.addText(dateToPresent, { x: 1/2.54, y: 17/2.54, w: 8/2.54, h: 2/2.54, fontSize: 32, color: "FFFFFF" });
  slide.addText("S&P Global Market Intelligence", 
    { x: (pptx.layout.width - 9/2.54), y: (pptx.layout.height - 2/2.54), w: 8/2.54, h: 1/2.54, fontSize: 14, color: "FFFFFF", align: 'right' });
}

// Helper to create content slide
async function createContentSlide(pptx, slide, idx, venue, pageNum) {
  const venueName = venue.venue_name || "";
  const venueCity = venue.venue_city || "";
  const venueGuestRooms = venue.venue_guest_rooms || "";
  const proposedDates = venue.proposed_dates || "";
  const averageDailyRate = venue.average_daily_rate || "";
  const totalFandB = venue.total_FandB || "";
  const additionalFees = venue.additional_fees || "";

  slide.addText(
    `Recommendation #${idx + 1} – ${venueName} : ${venueGuestRooms} rooms`,
    { x: 1/2.54, y: 1/2.54, w: 21/2.54, h: 2/2.54, fontSize: 32, bold: true }
  );

  addDatesAvailableBox(slide, 1/2.54, 5.54/2.54, 10/2.54, 1.2/2.54, proposedDates);

  const overviewText =
    `• City: ${venueCity}\n• Guest Rooms: ${venueGuestRooms}\n• Average Daily Rate: ${averageDailyRate}\n• Total Food & Beverages: ${totalFandB}\n• Additional Fees: ${additionalFees}`;
  const overviewLeft = 1/2.54, overviewTop = 8/2.54, overviewWidth = 12/2.54, overviewHeight = 7/2.54;

  slide.addShape(pptxgen.ShapeType.rect, { x: overviewLeft, y: overviewTop, w: overviewWidth, h: overviewHeight, line: { color: "000000", width: 2 }, fill: { color: "FFFFFF" } });

  slide.addText("Hotel Overview",
    { x: overviewLeft, y: overviewTop, w: overviewWidth, h: 1/2.54, fontSize: 16, bold: true, color: "FF0000" });
  slide.addText(overviewText,
    { x: overviewLeft, y: overviewTop + 1/2.54, w: overviewWidth, h: overviewHeight - 1/2.54, fontSize: 14, color: "000000" });

  // Add images and labels
  const imgW = 7/2.54, imgH = 4/2.54, gapH = 0.5/2.54, gapV = 2.54/2.54, imgL = 14.5/2.54, imgT = 5.54/2.54;
  const labels = ["Main Ballroom", "Bedroom", "Breakout room", "Outdoor space"];
  const positions = [
    [imgL, imgT],
    [imgL + imgW + gapH, imgT],
    [imgL, imgT + imgH + gapV],
    [imgL + imgW + gapH, imgT + imgH + gapV]
  ];
  for (let i = 0; i < labels.length; ++i) {
    const [left, top] = positions[i];
    // Gray box as placeholder
    slide.addShape(pptxgen.ShapeType.rect, { x:left, y:top, w:imgW, h:imgH, fill:{ color: "e6e6e6" }, line:{ color: "c8c8c8" } });
    // Label
    slide.addText(labels[i], { x:left, y:top + imgH + 0.2/2.54, w:imgW, h:1.0/2.54, fontSize: 12, align:'center' });
  }
  await addFooterWithLogo(pptx, slide, pageNum);
}

function parseInput(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const heading = lines[0];
  const dateToPresent = lines[1];
  const num = parseInt(lines[2]);
  const recs = Array(num).fill().map(() => ({}));
  for (const arg of lines.slice(3)) {
    const [key, ...rest] = arg.split('=');
    const value = rest.join('=');
    const match = /^R(\d+)\.(.*)/.exec(key);
    if (match) {
      const idx = parseInt(match[1], 10) - 1;
      if (idx >= 0 && idx < num) {
        recs[idx][match[2]] = value;
      }
    }
  }
  return { heading, dateToPresent, num, recs };
}

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST')
      return res.status(405).json({ error: 'Method not allowed' });

    let inputText;
    if (req.body && typeof req.body === 'object' && 'text' in req.body) {
      inputText = req.body.text;
    } else if (typeof req.body === 'string') {
      inputText = req.body;
    } else {
      inputText = await new Promise((resolve, reject) => {
        let data = '';
        req.on('data', chunk => (data += chunk));
        req.on('end', () => resolve(data));
        req.on('error', reject);
      });
    }
    if (!inputText) return res.status(400).json({ error: 'Missing input text' });

    const pptx = new PPTXGenJS();
    setA4Landscape(pptx);
    const { heading, dateToPresent, num, recs } = parseInput(inputText);

    createFrontPage(pptx, heading, dateToPresent);

    let slideNum = 2;
    for (let i = 0; i < recs.length; ++i) {
      const slide = pptx.addSlide();
      await createContentSlide(pptx, slide, i, recs[i], slideNum);
      slideNum += 1;
    }

    const filename = (heading ? heading.replace(/\s+/g, "_") : "presentation") + ".pptx";
    // Return as Buffer, then upload to Blob
    const buffer = await pptx.write('nodebuffer');

    const { url } = await put(filename, buffer, { access: 'public' });
    return res.status(200).json({ url });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: String(e) });
  }
}
