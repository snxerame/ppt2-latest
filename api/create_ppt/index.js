// api/create_ppt/index.js

import PPTXGenJS from "pptxgenjs";
import { put } from "@vercel/blob";

function cmToInch(cm) {
  return cm / 2.54;
}

// --- Custom-drawn S&P Global Market Intelligence "logo" as elements ---
function addSPGlobalLogo(slide, x, y, width, height) {
  // Adjust size and positions proportionally
  const barHeight = height * 0.18;
  const textHeight = height * 0.4;
  const gap = height * 0.12;

  // Black underline
  slide.addShape("rect", {
    x: x,
    y: y,
    w: width * 0.55,
    h: barHeight,
    fill: { color: "000000" },
    line: { color: "000000" }
  });

  // "S&P Global" Red
  slide.addText("S&P Global", {
    x: x,
    y: y + barHeight + gap / 2,
    w: width * 0.55,
    h: textHeight,
    fontSize: Math.round(textHeight * 0.8 * 21), // scale for pptx font size
    bold: true,
    color: "CC0A1E",
    fontFace: "Arial"
  });

  // "Market Intelligence" Black
  slide.addText("Market Intelligence", {
    x: x,
    y: y + barHeight + gap / 2 + textHeight,
    w: width * 1.15,
    h: textHeight,
    fontSize: Math.round(textHeight * 0.8 * 21),
    bold: false,
    color: "222222",
    fontFace: "Arial"
  });
}

function addFooter(slide, pptx, pageNum) {
  const footerY = pptx.layout.height - cmToInch(2.0);
  // Footer disclaimer at left
  slide.addText(
    "Permission to reprint or distribute any content from this presentation requires the prior written approval of S&P Global Market Intelligence.",
    {
      x: cmToInch(1),
      y: footerY,
      w: cmToInch(18),
      h: cmToInch(1.3),
      fontSize: 10,
      color: "808080",
      align: "left"
    }
  );
  // Page number at bottom right
  slide.addText(
    String(pageNum),
    {
      x: pptx.layout.width - cmToInch(2.5),
      y: footerY,
      w: cmToInch(2.0),
      h: cmToInch(1.3),
      fontSize: 14,
      color: "808080",
      align: "right"
    }
  );
}

function addDatesAvailableBox(slide, left, top, width, height, datesText) {
  slide.addShape("rect", { x: left, y: top, w: width, h: height, fill: { color: "e0eaee" }, line: { color: "e0eaee" } });
  slide.addText(
    [
      { text: "Proposed Dates: ", options: { fontSize: 14, bold: true, color: "CC0A1E" } },
      { text: datesText, options: { fontSize: 14, color: "000000" } }
    ],
    { x: left, y: top, w: width, h: height, align: "left", valign: "middle" }
  );
}

function createFrontPage(pptx, heading, dateToPresent) {
  const slide = pptx.addSlide();
  // Full-page grey background
  slide.addShape("rect", {
    x: 0, y: 0, w: pptx.layout.width, h: pptx.layout.height,
    fill: { color: "444444" }
  });

  // Custom S&P Global Market Intelligence "logo" at top left
  addSPGlobalLogo(
    slide,
    cmToInch(1),
    cmToInch(1.2),
    cmToInch(10),
    cmToInch(2.0)
  );

  // Title (centered, white)
  slide.addText(heading, {
    x: 0,
    y: cmToInch(6),
    w: pptx.layout.width,
    h: cmToInch(2.6),
    fontSize: 52,
    bold: true,
    color: "FFFFFF",
    align: "center"
  });

  // Date (centered, white, below title)
  slide.addText(dateToPresent, {
    x: 0,
    y: cmToInch(9),
    w: pptx.layout.width,
    h: cmToInch(2),
    fontSize: 32,
    color: "FFFFFF",
    align: "center"
  });

  // “S&P Market Analysis” as subtitle (centered, white, small, below date)
  slide.addText(
    "S&P Market Analysis",
    {
      x: 0,
      y: cmToInch(11.3),
      w: pptx.layout.width,
      h: cmToInch(1.5),
      fontSize: 20,
      color: "FFFFFF",
      align: "center"
    }
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
    x: overviewLeft, y: overviewTop, w: overviewWidth, h: cmToInch(1), fontSize: 16, bold: true, color: "CC0A1E"
  });
  slide.addText(overviewText, {
    x: overviewLeft, y: overviewTop + cmToInch(1), w: overviewWidth, h: overviewHeight - cmToInch(1), fontSize: 14, color: "000000"
  });

  // Placeholder boxes and labels
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
  await addFooter(slide, pptx, pageNum);
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
