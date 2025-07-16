import sys
import os
import json
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

LOGO_PATH = "sp_global_logo.png"

def set_a4_landscape(prs):
    prs.slide_width = Cm(29.7)
    prs.slide_height = Cm(21)

def add_footer_with_logo(prs, slide, page_num):
    logo_width = Cm(6.0)
    logo_height = Cm(2.4)
    footer_y = Cm(18.03)
    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, Cm(1), footer_y, width=logo_width, height=logo_height)
        left_text_x = Cm(7.2)
    else:
        left_text_x = Cm(1)
    left_box = slide.shapes.add_textbox(left_text_x, footer_y, Cm(14), Cm(1.5))
    left_frame = left_box.text_frame
    left_frame.clear()
    p_left = left_frame.add_paragraph()
    p_left.text = "Permission to reprint or distribute any content from this presentation requires the prior written approval of S&P Global Market Intelligence. "
    p_left.font.size = Pt(10)
    p_left.font.color.rgb = RGBColor(128, 128, 128)
    p_left.alignment = PP_ALIGN.LEFT
    right_box = slide.shapes.add_textbox(prs.slide_width - Cm(3), footer_y, Cm(2.5), Cm(1.5))
    right_frame = right_box.text_frame
    right_frame.clear()
    p_right = right_frame.add_paragraph()
    p_right.text = str(page_num)
    p_right.font.size = Pt(14)
    p_right.font.color.rgb = RGBColor(128, 128, 128)
    p_right.alignment = PP_ALIGN.RIGHT

def add_dates_available_box(slide, left, top, width, height, dates_text):
    text_box = slide.shapes.add_textbox(left, top, width, height)
    text_box.fill.solid()
    text_box.fill.fore_color.rgb = RGBColor(224, 234, 238)
    text_box.line.color.rgb = RGBColor(224, 234, 238)
    tf = text_box.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run1 = p.add_run()
    run1.text = "Proposed Dates: "
    run1.font.size = Pt(14)
    run1.font.bold = True
    run1.font.color.rgb = RGBColor(204, 0, 0)
    run2 = p.add_run()
    run2.text = str(dates_text)
    run2.font.size = Pt(14)
    run2.font.bold = False
    run2.font.color.rgb = RGBColor(0, 0, 0)
    return text_box

def create_front_page(prs, heading, date_to_present):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    # Full background block
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(153, 153, 153)
    shape.line.fill.background()
    # S&P Global
    brand_box = slide.shapes.add_textbox(Cm(1), Cm(1), Cm(8), Cm(2))
    brand_frame = brand_box.text_frame
    brand_frame.clear()
    p1 = brand_frame.add_paragraph()
    p1.text = "S&P Global"
    p1.font.size = Pt(20)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)
    p2 = brand_frame.add_paragraph()
    p2.text = "Market Intelligence"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(255, 255, 255)
    # Title
    title_box = slide.shapes.add_textbox(Cm(1), Cm(6), Cm(22), Cm(4))
    title_frame = title_box.text_frame
    title_frame.clear()
    p = title_frame.add_paragraph()
    p.text = heading
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    # Date
    date_box = slide.shapes.add_textbox(Cm(1), Cm(17), Cm(8), Cm(2))
    date_frame = date_box.text_frame
    date_frame.clear()
    p_date = date_frame.add_paragraph()
    p_date.text = date_to_present
    p_date.font.size = Pt(32)
    p_date.font.color.rgb = RGBColor(255, 255, 255)
    # Bottom right
    bottom_right_box = slide.shapes.add_textbox(prs.slide_width - Cm(9), prs.slide_height - Cm(2), Cm(8), Cm(1))
    bottom_right_frame = bottom_right_box.text_frame
    bottom_right_frame.clear()
    p_br = bottom_right_frame.add_paragraph()
    p_br.text = "S&P Global Market Intelligence"
    p_br.font.size = Pt(14)
    p_br.font.color.rgb = RGBColor(255, 255, 255)
    p_br.alignment = PP_ALIGN.RIGHT

def make_overview_text(venue_info):
    lines = []
    def line(label, key):
        value = venue_info.get(key, "NA")
        return f"• {label}: {value}" if value else f"• {label}: NA"
    # Core details (always include even if missing)
    fields = [
        ("Venue City", "Venue City"),
        ("Guest Rooms", "Venue Guest Rooms"),
        ("Daily Average Rate", "Daily Average Rate"),
        ("Total F&B", "Total F&B"),
        ("Additional Fees", "Additional fees"),
        ("Distance from Chicago ORD", "Distance from Chicago ORD"),
        ("Distance from Chicago MDW", "Distance from Chicago MDW"),
        ("Distance from DALAS DFW", "Distance from DALAS DFW"),
        ("Distance from DALAS Love Field", "Distance from DALAS Love Field"),
        ("Distance from Nashville BNA", "Distance from Nashville BNA"),
        ("Availability", "Availability"),
        ("Status", "Status"),
        ("Website Link", "Website Link"),
    ]
    for label, key in fields:
        lines.append(line(label, key))

    # Handle course list specially
    courses = venue_info.get("Nearby Golf Courts")
    if courses and isinstance(courses, list):
        lines.append("• Nearby Golf Courts:")
        for course in courses:
            cname = course.get("Name", "N/A")
            dist = course.get("distance", "N/A")
            lines.append(f"     - {cname} ({dist})")
    return "\n".join(lines)

def create_content_slide(prs, idx, venue_info, page_num, venue_title=""):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    venue_name = venue_info.get("Venue Name", "")
    venue_guest_rooms = venue_info.get("Venue Guest Rooms", "")

    heading_val = f"Recommendation #{idx+1} – {venue_name} : {venue_guest_rooms} rooms"
    heading_box = slide.shapes.add_textbox(Cm(1), Cm(1), Cm(21), Cm(2))
    heading_frame = heading_box.text_frame
    heading_frame.clear()
    p_heading = heading_frame.add_paragraph()
    p_heading.text = heading_val
    p_heading.font.size = Pt(32)
    p_heading.font.bold = True

    add_dates_available_box(
        slide, Cm(1), Cm(5.54), Cm(14.4), Cm(1.2), venue_info.get("Proposed dates", "")
    )

    overview_text = make_overview_text(venue_info)
    overview_left = Cm(1)
    overview_top = Cm(7)
    overview_width = Cm(14.4)
    overview_height = Cm(8)
    box_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        overview_left,
        overview_top,
        overview_width,
        overview_height
    )
    box_shape.line.color.rgb = RGBColor(0, 0, 0)
    box_shape.line.width = Pt(2)
    box_shape.fill.background()
    overview_box = slide.shapes.add_textbox(
        overview_left,
        overview_top,
        overview_width,
        overview_height
    )
    overview_frame = overview_box.text_frame
    overview_frame.clear()
    p_overview = overview_frame.add_paragraph()
    p_overview.text = "Hotel Overview"
    p_overview.font.size = Pt(16)
    p_overview.font.bold = True
    p_overview.font.color.rgb = RGBColor(255, 0, 0)
    p_overview.space_after = Pt(8)
    p_overview2 = overview_frame.add_paragraph()
    p_overview2.text = overview_text
    p_overview2.font.size = Pt(12)
    p_overview2.font.color.rgb = RGBColor(0, 0, 0)

    # Hotel Link (optional, as a separate box)
    weblink = venue_info.get("Website Link")
    if weblink:
        link_box = slide.shapes.add_textbox(Cm(1), Cm(15.6), Cm(12), Cm(1))
        link_frame = link_box.text_frame
        link_frame.clear()
        p_link = link_frame.add_paragraph()
        p_link.text = "Website: " + weblink
        p_link.font.size = Pt(10)
        p_link.font.color.rgb = RGBColor(0, 0, 160)
        p_link.alignment = PP_ALIGN.LEFT

    # Placeholder shapes and picture frames - here, just as gray blocks and labels
    img_w = Cm(7)
    img_h = Cm(4)
    gap_horizontal = Cm(0.5)
    gap_vertical = Cm(2.54)
    img_left = Cm(17)
    img_top = Cm(5.54)

    labels = ["Main Ballroom", "Bedroom", "Breakout room", "Outdoor space"]
    positions = [
        (img_left, img_top),
        (img_left + img_w + gap_horizontal, img_top),
        (img_left, img_top + img_h + gap_vertical),
        (img_left + img_w + gap_horizontal, img_top + img_h + gap_vertical)
    ]
    for label, (left, top) in zip(labels, positions):
        label_top = top + img_h + Cm(0.2)
        label_height = Cm(1.0)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, img_w, img_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(230, 230, 230)
        shape.line.color.rgb = RGBColor(200, 200, 200)
        label_box = slide.shapes.add_textbox(left, label_top, img_w, label_height)
        label_frame = label_box.text_frame
        label_frame.clear()
        p_label = label_frame.add_paragraph()
        p_label.text = label
        p_label.font.size = Pt(12)
        p_label.alignment = PP_ALIGN.CENTER

    add_footer_with_logo(prs, slide, page_num)

def main():
    if len(sys.argv) < 4:
        print("Usage: python script.py <input.json> <Heading> <Date_to_be_presented>")
        sys.exit(1)
    input_json = sys.argv[1]
    heading = sys.argv[2]
    date_to_present = sys.argv[3]

    with open(input_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    hotel_refs = data.get("Hotels", {})
    number_of_hotels = int(data.get("Number_of_hotels", 0))
    prs = Presentation()
    set_a4_landscape(prs)
    create_front_page(prs, heading, date_to_present)
    slide_num = 2

    # Create slide for every RX that exists in the input.
    for i in range(1, number_of_hotels+1):
        ref = f"R{i}"
        venue_info = data.get(ref)
        if venue_info is not None:
            create_content_slide(prs, i-1, venue_info, slide_num, venue_title=hotel_refs.get(ref,""))
            slide_num += 1
        else:
            # Optionally alert if RX missing
            print(f"Warning: No data for {ref} in input.")

    # Output file
    output_pptx = os.path.splitext(os.path.basename(input_json))[0] + ".pptx"
    prs.save(output_pptx)
    print(f"Presentation saved as {output_pptx}")

if __name__ == "__main__":
    main()
