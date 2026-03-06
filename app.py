import io
import re

import streamlit as st
from docx import Document
from deep_translator import GoogleTranslator
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Language options ──────────────────────────────────────────────────────────

LANGUAGES = {
    "English": "en",
    "Chinese (Simplified)": "zh-CN",
    "Chinese (Traditional)": "zh-TW",
    "Spanish": "es",
    "French": "fr",
    "Korean": "ko",
    "Japanese": "ja",
    "Portuguese": "pt",
    "Arabic": "ar",
    "Hindi": "hi",
    "German": "de",
    "Italian": "it",
    "Russian": "ru",
    "Vietnamese": "vi",
    "Thai": "th",
    "Indonesian": "id",
    "Malay": "ms",
    "Tagalog": "tl",
    "Ukrainian": "uk",
    "Dutch": "nl",
}

# ── Theme palette ─────────────────────────────────────────────────────────────

THEMES = {
    "Dark (Navy)": {
        "bg":      RGBColor(0x1A, 0x1A, 0x2E),
        "orig":    RGBColor(0xFF, 0xFF, 0xFF),
        "trans":   RGBColor(0xFF, 0xD7, 0x00),
        "divider": RGBColor(0x44, 0x44, 0x66),
        "label":   RGBColor(0x88, 0x88, 0xAA),
    },
    "Dark (Blue)": {
        "bg":      RGBColor(0x0D, 0x2B, 0x55),
        "orig":    RGBColor(0xFF, 0xFF, 0xFF),
        "trans":   RGBColor(0x7E, 0xC8, 0xE3),
        "divider": RGBColor(0x1E, 0x4D, 0x8C),
        "label":   RGBColor(0x7E, 0xC8, 0xE3),
    },
    "Light": {
        "bg":      RGBColor(0xF8, 0xF8, 0xF8),
        "orig":    RGBColor(0x1A, 0x1A, 0x2E),
        "trans":   RGBColor(0xB0, 0x30, 0x20),
        "divider": RGBColor(0xCC, 0xCC, 0xCC),
        "label":   RGBColor(0x88, 0x88, 0x88),
    },
    "Black": {
        "bg":      RGBColor(0x00, 0x00, 0x00),
        "orig":    RGBColor(0xFF, 0xFF, 0xFF),
        "trans":   RGBColor(0x00, 0xFF, 0xCC),
        "divider": RGBColor(0x33, 0x33, 0x33),
        "label":   RGBColor(0x66, 0x66, 0x66),
    },
}

# ── Text processing ───────────────────────────────────────────────────────────

def extract_paragraphs(docx_file) -> list[str]:
    """Return non-empty paragraph strings from a .docx file."""
    doc = Document(docx_file)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def split_sentences(text: str) -> list[str]:
    """Split text into sentences on common terminal punctuation."""
    parts = re.split(r'(?<=[.!?。！？])\s+', text)
    return [s.strip() for s in parts if s.strip()]


def max_chars_for_font(font_pt: int) -> int:
    """
    Estimate the maximum characters that safely fit in one text box half
    (half the slide height at 16:9, with 0.2" top/bottom padding).
    Uses a 0.85 safety factor to allow more text per slide while still
    leaving headroom for translated text being slightly longer.
    """
    BOX_H_PT   = (7.5 - 0.4) / 2 * 72          # ~248 pt per half
    lines      = int(BOX_H_PT / (font_pt * 1.2))
    usable_w   = (13.33 - 0.9) * 72             # ~895 pt
    chars_line = int(usable_w / (font_pt * 0.55))
    return max(20, int(lines * chars_line * 0.85))


def _split_to_fit(text: str, max_chars: int) -> list[str]:
    """Break text into pieces no longer than max_chars at word boundaries."""
    if len(text) <= max_chars:
        return [text]
    words = text.split()
    pieces, current = [], ""
    for word in words:
        candidate = (current + " " + word).strip()
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                pieces.append(current)
            current = word
    if current:
        pieces.append(current)
    return pieces


def build_chunks(paragraphs: list[str], mode: str, n: int, max_chars: int = 9999) -> list[str]:
    """
    Greedy packing: fill each slide with as much text as possible before
    starting a new one, so content is consolidated rather than fragmented.

    mode='paragraph' → units are whole paragraphs; short ones are merged
                        together until max_chars is reached.
    mode='sentence'  → units are sentences; pack up to n sentences per slide
                        (or fewer if max_chars would be exceeded).
    max_chars        → hard upper bound on characters per slide chunk.
    """
    if mode == "paragraph":
        units     = paragraphs
        max_units = 999_999   # no count cap — only max_chars limits merging
    else:
        units = []
        for para in paragraphs:
            units.extend(split_sentences(para))
        max_units = n         # respect the "sentences per slide" setting

    chunks: list[str] = []
    current      = ""
    current_count = 0

    for unit in units:
        # If the unit itself is longer than the limit, word-split it first
        pieces = _split_to_fit(unit, max_chars)
        for piece in pieces:
            candidate = (current + " " + piece).strip()
            if current and (len(candidate) > max_chars or current_count >= max_units):
                chunks.append(current)
                current       = piece
                current_count = 1
            else:
                current       = candidate
                current_count += 1

    if current:
        chunks.append(current)

    return chunks

def chars_per_line_for_font(font_pt: int) -> int:
    """Estimated characters that fit on one line at the given font size."""
    usable_w_pt = (13.33 - 0.9) * 72   # ~895 pt
    return max(10, int(usable_w_pt / (font_pt * 0.55)))


def fix_widow(text: str, chars_per_line: int) -> str:
    """
    Simulate word-wrap. If the last line would contain only one word (widow),
    steal the last word of the previous line and pull it down, then return the
    text with explicit newlines so python-pptx honours the adjusted breaks.
    """
    words = text.split()
    if len(words) < 3:
        return text

    # Simulate word-wrap into lines
    lines: list[list[str]] = [[]]
    for word in words:
        probe = " ".join(lines[-1] + [word])
        if len(probe) <= chars_per_line:
            lines[-1].append(word)
        else:
            lines.append([word])

    if len(lines) < 2 or len(lines[-1]) != 1:
        return text  # no widow — leave untouched (no \n injected)

    if len(lines[-2]) <= 1:
        return text  # previous line too short to donate a word

    # Steal the last word from the previous line
    stolen = lines[-2].pop()
    lines[-1].insert(0, stolen)

    return "\n".join(" ".join(line) for line in lines)


def fix_orphan_chunks(chunks: list[str], max_chars: int, min_words: int = 3) -> list[str]:
    """
    Merge any chunk that is an orphan (fewer than min_words words) into the
    previous chunk when it fits within max_chars; otherwise redistribute words
    evenly between the two chunks.
    """
    if len(chunks) <= 1:
        return chunks

    result = list(chunks)
    i = len(result) - 1
    while i > 0:
        if len(result[i].split()) < min_words:
            merged = (result[i - 1] + " " + result[i]).strip()
            if len(merged) <= max_chars:
                result[i - 1] = merged
                result.pop(i)
            else:
                all_words = result[i - 1].split() + result[i].split()
                mid = len(all_words) // 2
                result[i - 1] = " ".join(all_words[:mid])
                result[i]     = " ".join(all_words[mid:])
        i -= 1

    return result


# ── Translation ───────────────────────────────────────────────────────────────

def translate_chunks(
    chunks: list[str],
    src: str,
    tgt: str,
    progress_bar,
) -> list[str]:
    translator = GoogleTranslator(source=src, target=tgt)
    results = []
    total = len(chunks)
    for idx, text in enumerate(chunks):
        try:
            translated = translator.translate(text) or text
        except Exception:
            translated = text  # fall back to original on any error
        results.append(translated)
        progress_bar.progress((idx + 1) / total)
    return results

# ── PowerPoint builder ────────────────────────────────────────────────────────

def _set_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_textbox(slide, left, top, width, height, text, color, font_pt, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for idx, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_pt)
        run.font.color.rgb = color
        run.font.bold = bold


def build_pptx(
    chunks: list[str],
    translations: list[str],
    font_pt: int,
    theme_name: str,
    title: str,
) -> Presentation:
    palette = THEMES[theme_name]
    bg      = palette["bg"]
    orig_c  = palette["orig"]
    trans_c = palette["trans"]
    div_c   = palette["divider"]

    prs = Presentation()
    prs.slide_width  = Inches(13.33)   # 16:9 widescreen
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # truly blank

    # ── Title slide ──────────────────────────────────────────────────────────
    if title:
        slide = prs.slides.add_slide(blank_layout)
        _set_bg(slide, bg)
        _add_textbox(
            slide,
            Inches(0.5), Inches(2.0),
            Inches(12.33), Inches(3.5),
            title,
            orig_c,
            max(font_pt, 60),
            bold=True,
            align=PP_ALIGN.CENTER,
        )

    # ── Content slides ────────────────────────────────────────────────────────
    # Vertical layout (no labels):
    #   V_PAD  |  orig_box  |  DIV_GAP  |  divider  |  DIV_GAP  |  trans_box  |  V_PAD
    MARGIN  = Inches(0.45)
    TEXT_W  = Inches(13.33) - 2 * MARGIN
    V_PAD   = Inches(0.2)
    DIV_H   = Inches(0.07)
    DIV_GAP = Inches(0.08)

    BOX_H   = (Inches(7.5) - 2 * V_PAD - DIV_H - 2 * DIV_GAP) / 2

    orig_top  = V_PAD
    div_top   = orig_top  + BOX_H + DIV_GAP
    trans_top = div_top   + DIV_H + DIV_GAP

    cpl = chars_per_line_for_font(font_pt)

    for orig_text, trans_text in zip(chunks, translations):
        orig_text  = fix_widow(orig_text,  cpl)
        trans_text = fix_widow(trans_text, cpl)
        slide = prs.slides.add_slide(blank_layout)
        _set_bg(slide, bg)

        # Original text (top half)
        _add_textbox(
            slide, MARGIN, orig_top, TEXT_W, BOX_H,
            orig_text, orig_c, font_pt,
        )

        # Divider bar
        divider = slide.shapes.add_shape(
            1,  # MSO_SHAPE_TYPE.RECTANGLE
            MARGIN, div_top, TEXT_W, DIV_H,
        )
        divider.fill.solid()
        divider.fill.fore_color.rgb = div_c
        divider.line.fill.background()

        # Translated text (bottom half)
        _add_textbox(
            slide, MARGIN, trans_top, TEXT_W, BOX_H,
            trans_text, trans_c, font_pt,
        )

    return prs

# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Bilingual Sermon Slides",
    layout="wide",
)

st.title("Bilingual Sermon Slide Generator")
st.caption(
    "Upload a Word sermon document and generate bilingual PowerPoint slides "
    "your audience can follow while the sermon is preached."
)

# ── Sidebar settings ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")

    src_lang = st.selectbox(
        "Sermon language (source)",
        list(LANGUAGES.keys()),
        index=0,
    )
    tgt_lang = st.selectbox(
        "Audience language (target)",
        list(LANGUAGES.keys()),
        index=1,
    )

    st.divider()

    split_mode = st.radio("Split slides by", ["Sentences", "Paragraphs"])
    sentences_per_slide = 2
    if split_mode == "Sentences":
        sentences_per_slide = st.slider(
            "Sentences per slide", min_value=1, max_value=4, value=2
        )

    st.divider()

    font_size = st.slider("Font size (pt)", min_value=45, max_value=80, value=54)
    theme = st.selectbox("Slide theme", list(THEMES.keys()))
    title_text = st.text_input("Presentation title (optional)")

# ── Main area ─────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload sermon document (.docx)",
    type=["docx"],
    help="Microsoft Word .docx format only.",
)

if uploaded_file:
    st.success(f"Loaded: **{uploaded_file.name}**")

    if st.button("Generate Bilingual Slides", type="primary"):

        with st.spinner("Reading document…"):
            paragraphs = extract_paragraphs(uploaded_file)

        if not paragraphs:
            st.error("No text found in the document. Please check the file.")
            st.stop()

        st.info(f"{len(paragraphs)} paragraph(s) read from document.")

        mode_key  = "paragraph" if split_mode == "Paragraphs" else "sentence"
        max_chars = max_chars_for_font(font_size)
        chunks = build_chunks(paragraphs, mode_key, sentences_per_slide, max_chars)
        chunks = fix_orphan_chunks(chunks, max_chars)
        st.info(f"{len(chunks)} slide(s) will be created.")

        st.write("**Translating…** (this may take a moment for long sermons)")
        progress = st.progress(0)
        translations = translate_chunks(
            chunks,
            LANGUAGES[src_lang],
            LANGUAGES[tgt_lang],
            progress,
        )

        with st.spinner("Building PowerPoint…"):
            prs = build_pptx(
                chunks,
                translations,
                font_pt=font_size,
                theme_name=theme,
                title=title_text,
            )
            buf = io.BytesIO()
            prs.save(buf)
            buf.seek(0)

        st.success(f"Done! {len(chunks)} slides generated.")
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=buf,
            file_name="sermon_bilingual.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".presentationml.presentation"
            ),
        )
