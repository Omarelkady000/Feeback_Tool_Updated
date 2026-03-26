import streamlit as st
import math
import csv
import io
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# --- 1. CONFIGURATION ---
XML_TIMEBASE_MAP = {
    "23.976 fps": 23.976, "24.00 fps": 24, "25.00 fps": 25, 
    "29.97 fps": 29.97, "30.00 fps": 30, "50.00 fps": 50, 
    "59.94 fps": 59.94, "60.00 fps": 60
}

RES_MAP = {
    "1080x1920 (Vertical HD)": (1080, 1920),
    "1920x1080 (Landscape HD)": (1920, 1080),
    "2160x3840 (Vertical 4K)": (2160, 3840),
    "3840x2160 (Landscape 4K)": (3840, 2160)
}

def tc_to_frames(tc, source_fps, target_fps):
    try:
        clean_tc = tc.replace(';', ':')
        h, m, s, f = list(map(int, clean_tc.split(':')))
        
        # 1. Convert CSV Timecode to "Real Time" (Seconds)
        # We divide 'f' by the 'source_fps' because that is the ruler it was measured on.
        total_seconds = (h * 3600) + (m * 60) + s + (f / source_fps)
        
        # 2. Convert "Real Time" to the new "Target Ruler" (Frames)
        return math.floor(total_seconds * target_fps)
    except:
        return 0

def set_font(run, size=11, bold=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:ascii'), 'Arial')
    run.font.size = Pt(size)
    run.bold = bold

# --- 2. UI ---
st.set_page_config(page_title="QOMY Feedback Tool", page_icon="🎬")
st.title("🎬 QOMY Feedback Tool")

st.sidebar.header("PROJECT SETTINGS")

# Target: The sequence you are building now
st.sidebar.subheader("1. New Sequence (Target)")
target_fps_choice = st.sidebar.selectbox("Target FPS:", list(XML_TIMEBASE_MAP.keys()), index=4)
target_fps_val = XML_TIMEBASE_MAP[target_fps_choice]

res_choice = st.sidebar.selectbox("Resolution:", list(RES_MAP.keys()), index=0)
width, height = RES_MAP[res_choice]

# Source: The project the CSV came from
st.sidebar.subheader("2. Original CSV (Source)")
st.sidebar.info("Match this to the FPS of the project that exported the CSV.")
source_fps_choice = st.sidebar.selectbox("Source FPS:", list(XML_TIMEBASE_MAP.keys()), index=4)
source_fps_val = XML_TIMEBASE_MAP[source_fps_choice]

# --- 3. PROCESSING ---
csv_file = st.file_uploader("Upload Premiere CSV", type="csv")
logo_file = st.file_uploader("Upload Logo (Optional)", type=["png", "jpg"])

if csv_file:
    try:
        raw_data = csv_file.read()
        content = ""
        for enc in ['utf-8-sig', 'utf-16', 'cp1252']:
            try:
                content = raw_data.decode(enc)
                if "Marker Name" in content: break
            except: continue
        
        dialect = csv.Sniffer().sniff(content[:2000])
        reader = csv.DictReader(content.splitlines(), dialect=dialect)
        
        doc = Document()
        base_name = Path(csv_file.name).stem
        final_filename = f"{base_name}feedback"
        
        # Word Doc Title
        title_para = doc.add_heading('', 0)
        title_run = title_para.add_run(base_name)
        set_font(title_run, size=18, bold=True)
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        xml_markers = ""
        for row in reader:
            in_tc = row.get('In', '00:00:00:00')
            out_tc = row.get('Out', in_tc)
            name = row.get('Marker Name', '').strip()
            desc = row.get('Description', '').strip()
            comment = name if len(name) >= len(desc) else desc
            
            # Write to Word
            p = doc.add_paragraph()
            run_ts = p.add_run(f"{in_tc} - {out_tc}" if in_tc != out_tc else in_tc)
            set_font(run_ts, bold=True)
            p_cmt = doc.add_paragraph()
            run_cmt = p_cmt.add_run(comment)
            set_font(run_cmt)
            
            # Calculate XML Frames
            start_f = tc_to_frames(in_tc, source_fps_val, target_fps_val)
            end_f = tc_to_frames(out_tc, source_fps_val, target_fps_val)
            
            clean_cmt = comment.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            xml_markers += f"<marker><name>NOTE</name><comment>{clean_cmt}</comment><in>{int(start_f)}</in><out>{int(end_f)}</out></marker>"

        # Export
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        
        timebase = "24" if target_fps_val == 23.976 else str(int(target_fps_val))
        ntsc = "TRUE" if target_fps_val in [23.976, 29.97, 59.94] else "FALSE"
        
        full_xml = f'<?xml version="1.0" encoding="UTF-8"?><xmeml version="4"><project><children><sequence><name>{final_filename}</name><rate><timebase>{timebase}</timebase><ntsc>{ntsc}</ntsc></rate><media><video><format><samplecharacteristics><width>{width}</width><height>{height}</height></samplecharacteristics></format></video></media>{xml_markers}</sequence></children></project></xmeml>'

        st.divider()
        st.success(f"Generated files for: {base_name}")
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Download Docx", doc_io, f"{final_filename}.docx")
        c2.download_button("⬇️ Download XML", full_xml, f"{final_filename}.xml")

    except Exception as e:
        st.error(f"Error: {e}")
