import PPTXGenJS from "pptxgenjs";
import { put } from "@vercel/blob";
import fs from "fs/promises";

// Logo location: place `sp_global_logo.png` in your repo root or update this path.
const LOGO_PATH = "./sp_global_logo.png";

function cmToInch(cm) {
  return cm / 2.54;
}

async function addFooterWithLogo(pptx, slide, pageNum) {
  const logoWidth = cmToInch(6.0);
  const logoHeight = cmToInch(2.4);
  const footerY = cmToInch(18.03);
  let leftTextX;
  try {
    await fs.access(LOGO_PATH);
    slide.addImage({ path: LOGO_PATH, x: cmToInch(1), y: footerY, w: logoWidth, h: logoHeight });
    leftTextX = cmToInch(7.2);
  } catch {
    leftTextX = cmToInch(1);
  }
  slide.addText(
    "Permission to reprint or distribute any content from this presentation requires the prior written approval of S&P Global Market Intelligence.",
    { x: leftTextX, y: footerY, w: cmToInch(10), h: cmToInch(1.5), fontSize: 10, color: "808080", align: "left" }
  );
  slide.addText(
    String(pageNum),
    { x: pptx.layout.width - cmToInch(3), y: footerY, w: cmToInch(2.5), h: cmToInch(1.5), fontSize: 14, color: "808080", align: "right" }
  );
}

function addDatesAvailableBox(slide, left, top, width, height, datesText) {
  slide.addShape("rect", { x: left, y: top, w: width, h: height, fill: { color: "e0eaee" }, line: { color: "e0eaee" } });
  slide.addText(
    [
      { text: "Proposed Dates: ", options: { fontSize: 14, bold: true, color: "CC0000" } },
      { text: datesText, options: { fontSize: 14, color: "000000" } }
    ],
    { x: left, y: top, w: width, h: height, align: "left", valign: "middle" }
  );
}

function createFrontPage(pptx, heading, dateToPresent) {
  const slide = pptx.addSlide();
  slide.addShape("rect", { x: 0, y: 0, w: pptx.layout.width, h: pptx.layout.height, fill: { color: "999999" } });
  slide.addText("S&P Global", { x: cmToInch(1), y: cmToInch(1), w: cmToInch(8), h: cmToInch(2), fontSize: 20, bold: true, color: "FFFFFF" });
  slide.addText("Market Intelligence", { x: cmToInch(1), y: cmToInch(2.1), w: cmToInch(8), h: cmToInch(2), fontSize: 20, color: "FFFFFF" });
  slide.addText(heading, { x: cmToInch(1), y: cmToInch(6), w: cmToInch(22), h: cmToInch(4), fontSize: 54, bold: true, color: "FFFFFF" });
  slide.addText(dateToPresent, { x: cmToInch(1), y: cmToInch(17), w: cmToInch(8), h: cmToInch(2), fontSize: 32, color: "FFFFFF" });
  slide.addText(
    "S&P Global Market Intelligence",
    { x: pptx.layout.width - cmToInch(9), y: pptx.layout.height - cmToInch(2), w: cmToInch(8), h: cmToInch(1), fontSize: 14, color: "FFFFFF", align: "right" }
  );
}

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
    { x: cmToInch(1), y: cmToInch(1), w: cmToInch(21), h: cmToInch(2), fontSize: 32, bold: true }
  );
  addDatesAvailableBox(slide, cmToInch(1), cmToInch(5.54), cmToInch(10), cmToInch(1.2), proposedDates);

  const overviewText =
    `• City: ${venueCity}\n• Guest Rooms: ${venueGuestRooms}\n• Average Daily Rate: ${averageDailyRate}\n• Total Food & Beverages: ${totalFandB}\n• Additional Fees: ${additionalFees}`;

  const overviewLeft = cmToInch(1), overviewTop = cmToInch(8), overviewWidth = cmToInch(12), overviewHeight = cmToInch(7);

  slide.addShape("rect", {
    x: overviewLeft, y: overviewTop, w: overviewWidth, h: overviewHeight, line: { color: "000000", width: 2 }, fill: { color: "FFFFFF" }
  });

  slide.addText("Hotel Overview", {
    x: overviewLeft, y: overviewTop, w: overviewWidth, h: cmToInch(1), fontSize: 16, bold: true, color: "FF0000"
  });
  slide.addText(overviewText, {
    x: overviewLeft, y: overviewTop + cmToInch(1), w: overviewWidth, h: overviewHeight - cmToInch(1), fontSize: 14, color: "000000"
  });

  // Placeholder gray boxes and labels
  const imgW = cmToInch(7), imgH = cmToInch(4), gapH = cmToInch(0.5), gapV = cmToInch(2.54), imgL = cmToInch(14.5), imgT = cmToInch(5.54);
  const labels = ["Main Ballroom", "Bedroom", "Breakout room", "Outdoor space"];
  const positions = [
    [imgL, imgT],
    [imgL + imgW + gapH, imgT],
    [imgL, imgT + imgH + gapV],
    [imgL + imgW + gapH, imgT + imgH + gapV]
  ];
  for (let i = 0; i < labels.length; ++i) {
    const [left, top] = positions[i];
    slide.addShape("rect", {
      x: left, y: top, w: imgW, h: imgH, fill: { color: "e6e6e6" }, line: { color: "c8c8c8" }
    });
    slide.addText(labels[i], {
      x: left, y: top + imgH + cmToInch(0.2), w: imgW, h: cmToInch(1.0), fontSize: 12, align: "center"
    });
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
    pptx.defineLayout({ name: "A4", width: 11.6929, height: 8.2677 }); // in inches
    pptx.layout = "A4";
    const { heading, dateToPresent, num, recs } = parseInput(inputText);
    createFrontPage(pptx, heading, dateToPresent);

    let slideNum = 2;
    for (let i = 0; i < recs.length; ++i) {
      const slide = pptx.addSlide();
      await createContentSlide(pptx, slide, i, recs[i], slideNum);
      slideNum += 1;
    }

    const filename = (heading ? heading.replace(/\s+/g, "_") : "presentation") + ".pptx";
    const buffer = await pptx.write("nodebuffer");

    const { url } = await put(filename, buffer, { access: "public" });
    return res.status(200).json({ url });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: String(e) });
  }
}
