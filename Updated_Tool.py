import streamlit as st
import re
import math
import csv
import io
from pathlib import Path
from fpdf import FPDF # New dependency: pip install fpdf2

# --- 1. CONFIGURATION DATA ---
XML_TIMEBASE_MAP = {
    "10.00 fps": "10", "12.00 fps": "12", "15.00 fps": "15",
    "23.976 fps": "24", "24.00 fps": "24", "25.00 fps": "25", "29.97 fps": "30",
    "30.00 fps": "30", "50.00 fps": "50", "59.94 fps": "60", "60.00 fps": "60"
}

RES_MAP = {
    "1080x1920 (Vertical HD)": (1080, 1920),
    "1920x1080 (Landscape HD)": (1920, 1080),
    "2160x3840 (Vertical 4K)": (2160, 3840),
    "3840x2160 (Landscape 4K)": (3840, 2160),
    "1080x1080 (Square)": (1080, 1080)
}

# FPS Calculation Logic (Untouched as requested)
def tc_to_frames(tc, fps_choice):
    try:
        clean_tc = tc.replace(';', ':')
        parts = list(map(int, clean_tc.split(':')))
        h, m, s, f = parts
        total_minutes = (h * 60) + m
        if "29.97" in fps_choice:
            frame_number = ((total_minutes * 60) + s) * 30 + f
            drop_frames = 2 * (total_minutes - (total_minutes // 10))
            return frame_number - drop_frames
        elif "59.94" in fps_choice:
            frame_number = ((total_minutes * 60) + s) * 60 + f
            drop_frames = 4 * (total_minutes - (total_minutes // 10))
            return frame_number - drop_frames
        else:
            base = 24 if "23.976" in fps_choice else float(fps_choice.split(' ')[0])
            return math.floor((h * 3600 * base) + (m * 60 * base) + (s * base) + f)
    except: return 0

# --- 2. UI SETUP ---
st.set_page_config(page_title="QOMY Feedback Tool", page_icon="🎬", layout="centered")

st.title("🎬 QOMY Feedback Tool")
st.markdown("Upload your Premiere CSV to generate formatted **PDF Feedback Docs** and **XML Markers**.")

st.sidebar.header("GLOBAL SETTINGS")

st.sidebar.write("Select Premiere Sequence FPS:")
fps_choice = st.sidebar.selectbox("FPS Dropdown:", list(XML_TIMEBASE_MAP.keys()), index=6, label_visibility="collapsed")

st.sidebar.write("Select Sequence Resolution:")
res_choice = st.sidebar.selectbox("Resolution Dropdown:", list(RES_MAP.keys()), index=0, label_visibility="collapsed")
width, height = RES_MAP[res_choice]

# --- 3. WORKFLOW ---
csv_file = st.file_uploader("Select Premiere CSV", type="csv")

if csv_file:
    try:
        raw_data = csv_file.read()
        content = ""
        for enc in ['utf-8-sig', 'utf-16', 'cp1252']:
            try:
                content = raw_data.decode(enc)
                if "Marker Name" in content: break
            except: continue
        
        lines = content.splitlines()
        if len(lines) > 0:
            first_line = lines[0]
            delim = '\t' if '\t' in first_line else ','
            reader = csv.DictReader(lines, delimiter=delim)
        else:
            raise ValueError("The uploaded file is empty.")
        
        base_name = Path(csv_file.name).stem
        final_filename = f"{base_name}_feedback"

        # --- PDF GENERATION ---
        pdf = FPDF()
        pdf.add_page()
        
        # Load Custom Fonts
        # Ensure these .ttf files exist in your root directory!
        try:
            pdf.add_font("Oswald", "M", "Oswald-Medium.ttf")
            pdf.add_font("Satoshi", "", "Satoshi-Regular.ttf")
            header_font = "Oswald"
            body_font = "Satoshi"
        except:
            st.warning("Font files not found. Falling back to Arial.")
            header_font = "Arial"
            body_font = "Arial"

        # Title
        pdf.set_font(header_font, "M" if header_font != "Arial" else "B", 18)
        pdf.cell(0, 15, base_name, ln=True, align='C')
        pdf.ln(5)

        xml_markers = ""
        
        for row in reader:
            name = row.get('Marker Name', '').strip()
            desc = row.get('Description', '').strip()
            comment = name if len(name) >= len(desc) else desc
            
            in_tc = row.get('In', '00:00:00:00')
            out_tc = row.get('Out', in_tc)
            ts_display = in_tc if in_tc == out_tc else f"{in_tc} - {out_tc}"
            
            # PDF Content
            # Timestamp (Oswald Medium)
            pdf.set_font(header_font, "M" if header_font != "Arial" else "B", 11)
            pdf.cell(0, 7, ts_display, ln=True)
            
            # Comment (Satoshi Regular)
            pdf.set_font(body_font, "", 11)
            pdf.multi_cell(0, 6, comment)
            pdf.ln(4)
            
            # XML Logic (Added Square Pixel Aspect Ratio)
            start_f = tc_to_frames(in_tc, fps_choice)
            end_f = tc_to_frames(out_tc, fps_choice)
            clean_cmt = comment.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            xml_markers += f"<marker><name>NOTE</name><comment>{clean_cmt}</comment><in>{int(start_f)}</in><out>{int(end_f)}</out></marker>"

        # Prepare PDF buffer
        pdf_output = pdf.output(dest='S')
        
        # Prepare XML parameters
        timebase = XML_TIMEBASE_MAP.get(fps_choice, "30")
        ntsc = "TRUE" if (".97" in fps_choice or ".94" in fps_choice) else "FALSE"
        
        # Updated XML with square pixel aspect ratio tag
        full_xml = (
            f'<?xml version="1.0" encoding="UTF-8"?>'
            f'<xmeml version="4"><project><children><sequence>'
            f'<name>{final_filename}</name>'
            f'<rate><timebase>{timebase}</timebase><ntsc>{ntsc}</ntsc></rate>'
            f'<media><video><format><samplecharacteristics>'
            f'<width>{width}</width><height>{height}</height>'
            f'<pixelaspectratio>square</pixelaspectratio>' # Forced Square Pixels
            f'</samplecharacteristics></format></video></media>'
            f'{xml_markers}'
            f'</sequence></children></project></xmeml>'
        )

        st.divider()
        st.success(f"Processed: {base_name}")
        
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇️ Download PDF", data=pdf_output, file_name=f"{final_filename}.pdf", mime="application/pdf")
        with c2:
            st.download_button("⬇️ Download XML", data=full_xml, file_name=f"{final_filename}.xml")

    except Exception as e:
        st.error(f"Error: {e}")
