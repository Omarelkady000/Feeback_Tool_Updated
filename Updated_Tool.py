import streamlit as st
import re
import requests
import math
import csv
import io
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# --- 1. CORE MATH ---
XML_TIMEBASE_MAP = {
    "10.00 fps": "10", "12.00 fps": "12", "15.00 fps": "15",
    "23.976 fps": "24", "24.00 fps": "24", "25.00 fps": "25", "29.97 fps": "30",
    "30.00 fps": "30", "50.00 fps": "50", "59.94 fps": "60", "60.00 fps": "60"
}

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

def apply_style(run, size=11, bold=False, color=None):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:ascii'), 'Arial')
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color

# --- 2. UI SETUP ---
st.set_page_config(page_title="QOMY Feedback Tool", page_icon="🎬", layout="centered")

st.title("🎬 QOMY Feedback Tool")
st.markdown("Upload your Premiere CSV to generate formatted Feedback Docs and XML Markers.")

# Global Settings in Sidebar
st.sidebar.header("Settings")
fps_choice = st.sidebar.selectbox("Sequence FPS:", list(XML_TIMEBASE_MAP.keys()), index=6)

# --- SINGLE MODE WORKFLOW ---
csv_file = st.file_uploader("Upload Premiere Markers CSV", type="csv")
logo_file = st.file_uploader("Upload Brand Logo (Optional)", type=["png", "jpg"])

if csv_file:
    try:
        raw_data = csv_file.read()
        content = ""
        # Handle various CSV encodings
        for enc in ['utf-8-sig', 'utf-16', 'cp1252']:
            try:
                content = raw_data.decode(enc)
                if "Marker Name" in content: break
            except: continue
        
        if "Marker Name" not in content:
            st.error("Invalid CSV format. Please ensure you exported 'Markers' from Premiere.")
        else:
            dialect = csv.Sniffer().sniff(content[:2000])
            reader = csv.DictReader(content.splitlines(), dialect=dialect)
            
            doc = Document()
            
            # 1. Logo Handling
            if logo_file:
                doc.add_picture(io.BytesIO(logo_file.read()), width=Inches(1.5))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 2. Header (Filename)
            title_name = Path(csv_file.name).stem
            header = doc.add_paragraph()
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_h = header.add_run(title_name)
            apply_style(run_h, size=14, bold=True)

            xml_markers = ""
            
            for row in reader:
                name = row.get('Marker Name', '').strip()
                desc = row.get('Description', '').strip()
                # Premiere sometimes puts the text in Name or Description
                comment = name if len(name) >= len(desc) else desc
                
                in_tc = row.get('In', '00:00:00:00')
                out_tc = row.get('Out', in_tc)
                
                # Timestamp Formatting (Exact match to your reference)
                ts_display = in_tc if in_tc == out_tc else f"{in_tc} - {out_tc}"
                
                # Color Logic (Red for remove/cut)
                is_negative = bool(re.search(r'\b(remove|cut|delete)\b', comment, re.IGNORECASE))
                text_color = RGBColor(255, 0, 0) if is_negative else None
                
                # Add Line: Timecode + Double Space + Comment (Arial)
                p = doc.add_paragraph()
                run_ts = p.add_run(f"{ts_display}  ")
                apply_style(run_ts, bold=True, color=text_color)
                
                run_cmt = p.add_run(comment)
                apply_style(run_cmt, color=text_color)
                
                # XML logic
                start_f = tc_to_frames(in_tc, fps_choice)
                end_f = tc_to_frames(out_tc, fps_choice)
                clean_cmt = comment.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                xml_markers += f"<marker><name>NOTE</name><comment>{clean_cmt}</comment><in>{int(start_f)}</in><out>{int(end_f)}</out></marker>"

            # XML Finalizing
            timebase = XML_TIMEBASE_MAP.get(fps_choice, "30")
            ntsc = "TRUE" if (".97" in fps_choice or ".94" in fps_choice) else "FALSE"
            full_xml = f'<?xml version="1.0" encoding="UTF-8"?><xmeml version="4"><project><children><sequence><name>QOMY_IMPORT</name><rate><timebase>{timebase}</timebase><ntsc>{ntsc}</ntsc></rate><media><video><format><samplecharacteristics><width>1920</width><height>1080</height></samplecharacteristics></format></video></media>{xml_markers}</sequence></children></project></xmeml>'

            # Create Downloadable Buffers
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            
            st.divider()
            st.success(f"Successfully processed: {title_name}")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="⬇️ Download Word Doc",
                    data=doc_io,
                    file_name=f"{title_name}_Feedback.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with col2:
                st.download_button(
                    label="⬇️ Download Premiere XML",
                    data=full_xml,
                    file_name=f"{title_name}_Markers.xml",
                    mime="application/xml"
                )

    except Exception as e:
        st.error(f"Processing failed: {e}")
