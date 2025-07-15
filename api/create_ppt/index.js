import PPTXGenJS from "pptxgenjs";
import { put, del } from "@vercel/blob";

// --------- Helpers ----------
function cmToInch(cm) {
  return cm / 2.54;
}



// Adds disclaimer (bottom left) and page number (bottom right) on all slides
function addFooter(slide, pageNum) {
  // Hardcoded for A4 landscape: 8.2677 inch height, minus ~2.0cm bottom margin for footer.
  const footerY = 7.48; // 8.2677 - 0.7874 (2.0 cm in inches)
  // Disclaimer (left)
  slide.addText(
    "",
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
  // Page number (right)
  slide.addText(
    String(pageNum),
    {
      x: 11.7 - cmToInch(2.5), // A4 width minus 2.5cm
      y: footerY,
      w: cmToInch(2.0),
      h: cmToInch(1.3),
      fontSize: 14,
      color: "808080",
      align: "right"
    }
  );
}

// Dates textbox for proposed dates
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

// ------ Front Page ------
function createFrontPage(pptx, heading, dateToPresent) {
  const slide = pptx.addSlide();
  slide.background = { fill: "444444" };


  // Title (centered, white)
  slide.addText(heading, {
    x: 0, y: cmToInch(6), w: pptx.layout.width, h: cmToInch(2.6),
    fontSize: 52, bold: true, color: "FFFFFF", align: "center"
  });

  // Date (centered, white, below title)
  slide.addText(dateToPresent, {
    x: 0, y: cmToInch(9), w: pptx.layout.width, h: cmToInch(2),
    fontSize: 32, color: "FFFFFF", align: "center"
  });

  // Subtitle (centered, white, below date)
  slide.addText("S&P Market Analysis", {
    x: 0, y: cmToInch(11.3), w: pptx.layout.width, h: cmToInch(1.5),
    fontSize: 20, color: "FFFFFF", align: "center"
  });

  // Footer in white, bottom right
  slide.addText(
    "S&P Global Market Intelligence",
    {
      x: 11.7 - cmToInch(10), // A4 width minus 10cm
      y: 7.48, // ~2cm from bottom
      w: cmToInch(9.1),
      h: cmToInch(1),
      fontSize: 15,
      color: "FFFFFF",
      align: "right",
      fontFace: "Arial"
    }
  );
}

// ------ Content Slide ------
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

  addFooter(slide, pageNum);

  // Branded footer, bottom left: "S&P Global" (red, bold), "Market Intelligence" (black)
  const sx = cmToInch(1);
  const sy = 7.48; // A4 height - 2cm
  const sLineH = cmToInch(0.7);

  slide.addText("S&P Global", {
    x: sx,
    y: sy,
    w: cmToInch(3.4),
    h: sLineH,
    fontSize: 15,
    bold: true,
    color: "CC0A1E",
    fontFace: "Arial",
    align: "left"
  });
  slide.addText("Market Intelligence", {
    x: sx + cmToInch(3.2),
    y: sy,
    w: cmToInch(7),
    h: sLineH,
    fontSize: 15,
    color: "222222",
    fontFace: "Arial",
    align: "left"
  });
}

// --- Input Parser ---
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

// --- API Handler ---
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
    pptx.defineLayout({ name: "A4", width: 11.6929, height: 8.2677 });
    pptx.layout = "A4";
    const { heading, dateToPresent, num, recs } = parseInput(inputText);

    createFrontPage(pptx, heading, dateToPresent);

    let slideNum = 2;
    for (let i = 0; i < recs.length; ++i) {
      const slide = pptx.addSlide();
      await createContentSlide(pptx, slide, i, recs[i], slideNum);
      slideNum += 1;
    }

    const filename = `${new Date().toISOString().replace(/[:T]/g, '-').slice(0,19)}.pptx`;


    const buffer = await pptx.write("nodebuffer");

    // Delete possible previous file before upload!
    try {
      await del(filename);
    } catch (err) {
      // It's fine for the file not to exist yet; ignore error.
    }

    const { url } = await put(filename, buffer, {
      access: "public"
      // allowOverwrite: true not needed since we're deleting first
    });

    return res.status(200).json({ url });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: String(e) });
  }
}
