import io
import json
import base64
from pathlib import Path
from typing import Dict, List, Union, Tuple

import streamlit as st
from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError

# YAML for presets
try:
    import yaml  # type: ignore
    _HAVE_YAML = True
except Exception:
    _HAVE_YAML = False

# PyMuPDF for fast page previews
try:
    import fitz  # PyMuPDF
    _HAVE_FITZ = True
except Exception:
    _HAVE_FITZ = False


ICON_URL = "https://raw.githubusercontent.com/baliant/pdftool/main/icon/koru_logo.png"

st.set_page_config(
    page_title="PDF Select, Review & Merge",
    page_icon=ICON_URL,
    layout="wide"
)


def open_reader_from_bytes(data: bytes):
    try:
        bio = io.BytesIO(data)
        reader = PdfReader(bio)
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                st.warning("Encrypted PDF that cannot be opened with an empty password. Skipping.")
                return None
        return reader
    except (PdfReadError, OSError) as e:
        st.warning(f"Cannot open PDF: {e}")
        return None


def parse_one_pagespec(token: str, max_pages: int) -> List[int]:
    t = token.strip().lower()
    if not t:
        return []
    if t == "all":
        return list(range(1, max_pages + 1))
    if "-" in t:
        if t.count("-") > 1:
            raise ValueError(f"Invalid range: {t}")
        start, end = t.split("-", 1)
        if start == "" and end == "":
            raise ValueError(f"Invalid open range: {t}")
        if start == "":
            e = int(end)
            if e < 1:
                raise ValueError(f"Invalid end in range: {t}")
            e = min(e, max_pages)
            return list(range(1, e + 1))
        if end == "":
            s = int(start)
            if s < 1:
                raise ValueError(f"Invalid start in range: {t}")
            s = min(s, max_pages)
            return list(range(s, max_pages + 1))
        s = int(start)
        e = int(end)
        if s < 1 or e < 1 or s > e:
            raise ValueError(f"Invalid range: {t}")
        s = min(s, max_pages)
        e = min(e, max_pages)
        return list(range(s, e + 1))
    n = int(t)
    if n < 1:
        raise ValueError(f"Invalid page number: {t}")
    n = min(n, max_pages)
    return [n]


def parse_pagespec(spec: Union[str, List[str]], max_pages: int) -> List[int]:
    if isinstance(spec, str):
        parts = [p for p in spec.split(",") if p.strip()]
    else:
        parts = list(spec)
    pages: List[int] = []
    for part in parts:
        pages.extend(parse_one_pagespec(part, max_pages))
    # dedup preserve order
    seen = set()
    uniq: List[int] = []
    for p in pages:
        if p not in seen:
            uniq.append(p)
            seen.add(p)
    return uniq


def load_selection_mapping(data: bytes, suffix: str):
    suffix = suffix.lower()
    text = data.decode("utf-8")
    if suffix in (".yaml", ".yml"):
        if not _HAVE_YAML:
            st.error("PyYAML is not installed on the server environment. Install with: pip install pyyaml")
            return {}
        return yaml.safe_load(text) or {}
    return json.loads(text)


def _unique_key(prefix: str, name_or_path: str, data_bytes: bytes | None = None) -> str:
    base = f"{prefix}:{name_or_path}"
    if data_bytes is not None:
        import hashlib
        h = hashlib.md5(data_bytes).hexdigest()[:8]
        base += f":{h}"
    return base


@st.cache_resource(show_spinner=False)
def _load_doc_for_preview(data: bytes):
    if not _HAVE_FITZ:
        return None
    try:
        return fitz.open(stream=data, filetype="pdf")
    except Exception:
        return None


def render_page_image(data: bytes, page_index0: int, zoom: float = 1.5) -> Union[None, bytes]:
    """
    Render one page to PNG. Returns PNG bytes or None.
    """
    if not _HAVE_FITZ:
        return None
    doc = _load_doc_for_preview(data)
    if doc is None:
        return None
    if page_index0 < 0 or page_index0 >= doc.page_count:
        return None
    try:
        page = doc.load_page(page_index0)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        return pix.tobytes("png")
    except Exception:
        return None


def embed_pdf_viewer(data: bytes, height: int = 480):
    """Inline PDF viewer using base64 in an iframe (works for small/medium files)."""
    b64 = base64.b64encode(data).decode("ascii")
    src = f"data:application/pdf;base64,{b64}#view=FitH"
    st.components.v1.html(
        f'<iframe src="{src}" width="100%" height="{height}" style="border:none;"></iframe>',
        height=height, scrolling=False
    )


st.title("📚 PDF Select, Review & Merge")
st.caption("List page counts, visually review pages, pick ranges, merge, and download.")

with st.expander("Page selection syntax help"):
    st.markdown(
        """
**Syntax (1-based):**
- Single page: `7`  
- Range: `3-9`  
- Open start: `-5` (pages 1..5)  
- Open end: `4-` (pages 4..last)  
- Comma list: `1-3,5,10-12`  
- `all` for all pages
"""
    )

tab_upload, tab_folder, tab_review = st.tabs(["📤 Upload PDFs", "📁 Folder path (local run)", "🧐 Review"])

selections: Dict[str, List[int]] = {}
file_entries: List[Tuple[str, bytes]] = []  # list of tuples: (display_name, data_bytes)


# ---------------- Upload Tab ----------------
with tab_upload:
    uploaded = st.file_uploader("Select one or more PDFs", type=["pdf"], accept_multiple_files=True)
    mapping_file = st.file_uploader("Optional selections mapping (YAML or JSON)", type=["yaml","yml","json"], key="map_upload")
    uploaded_mapping = {}
    if mapping_file is not None:
        try:
            uploaded_mapping = load_selection_mapping(mapping_file.getvalue(), Path(mapping_file.name).suffix)
        except Exception as e:
            st.error(f"Failed to load mapping: {e}")

    if uploaded:
        st.subheader("Files")
        for f in uploaded:
            data = f.getvalue()
            reader = open_reader_from_bytes(data)
            if not reader:
                continue
            pages = len(reader.pages)

            left, right = st.columns([2,1])
            with left:
                st.markdown(f"**{f.name}** — {pages} pages")
                default_spec = "all"
                if uploaded_mapping:
                    default_spec = uploaded_mapping.get(f.name, default_spec)
                spec_key = _unique_key("spec", f.name, data)
                spec = st.text_input(f"Pages for {f.name}", value=str(default_spec), key=spec_key)

                # --- Review controls (per file) ---
                with st.expander("Review this PDF"):
                    if not _HAVE_FITZ:
                        st.info("Install **PyMuPDF** (`pip install pymupdf`) to enable page previews.")
                    else:
                        pnum_key = _unique_key("prev", f.name, data)
                        pnum = st.number_input("Preview page", min_value=1, max_value=pages, value=1, step=1, key=pnum_key)
                        png = render_page_image(data, int(pnum)-1, zoom=1.5)
                        if png:
                            st.image(png, caption=f"{f.name} — Page {pnum}", use_column_width=True)
                        show_full_key = _unique_key("viewer", f.name, data)
                        show_full = st.checkbox("Inline full PDF viewer", value=False, key=show_full_key)
                        if show_full:
                            embed_pdf_viewer(data, height=500)

            with right:
                st.caption("Selected pages preview")
                try:
                    pages_list = parse_pagespec(spec, pages)
                    if not pages_list:
                        st.warning("No valid pages; this file will be skipped.")
                    else:
                        st.caption(f"Selected pages: {len(pages_list)}")
                        # Show first 10 indices for quick glance
                        s = ", ".join(map(str, pages_list[:10]))
                        more = "" if len(pages_list) <= 10 else " …"
                        st.text(f"{s}{more}")
                        selections[f.name] = pages_list
                        file_entries.append((f.name, data))
                except Exception as e:
                    st.error(f"Invalid page spec for {f.name}: {e}")


# ---------------- Folder Tab ----------------
with tab_folder:
    st.write("Enter a local folder path to scan for PDFs. (This only works when you run Streamlit locally.)")
    folder = st.text_input("Folder path", value="")
    mapping2 = st.file_uploader("Optional selections mapping (YAML or JSON)", type=["yaml","yml","json"], key="map_folder")
    folder_mapping = {}
    if mapping2 is not None:
        try:
            folder_mapping = load_selection_mapping(mapping2.getvalue(), Path(mapping2.name).suffix)
        except Exception as e:
            st.error(f"Failed to load mapping: {e}")

    if folder:
        p = Path(folder).expanduser()
        if not p.exists():
            st.error("Folder does not exist.")
        else:
            pdfs = sorted(p.rglob("*.pdf"))
            if not pdfs:
                st.warning("No PDFs found in the folder (recursively).")
            else:
                st.subheader("Files")
                for pf in pdfs:
                    try:
                        data = pf.read_bytes()
                    except Exception as e:
                        st.warning(f"Cannot read {pf}: {e}")
                        continue
                    reader = open_reader_from_bytes(data)
                    if not reader:
                        continue
                    pages = len(reader.pages)
                    st.markdown(f"**{pf}** — {pages} pages")
                    default_spec = "all"
                    if folder_mapping:
                        if str(pf) in folder_mapping:
                            default_spec = folder_mapping[str(pf)]
                        elif pf.name in folder_mapping:
                            default_spec = folder_mapping[pf.name]
                    spec_key = _unique_key("spec", str(pf))  # path is unique
                    spec = st.text_input(f"Pages for {pf}", value=str(default_spec), key=spec_key)
                    try:
                        pages_list = parse_pagespec(spec, pages)
                        if not pages_list:
                            st.warning("No valid pages selected; this file will be skipped.")
                        else:
                            selections[str(pf)] = pages_list
                            file_entries.append((str(pf), data))
                    except Exception as e:
                        st.error(f"Invalid page spec for {pf}: {e}")


# ---------------- Global Review Tab ----------------
with tab_review:
    st.write("Preview uploaded or scanned PDFs. Use the controls to flip pages and verify content before merging.")
    if not file_entries:
        st.info("No files loaded yet. Upload or select a folder first.")
    else:
        for name, data in file_entries:
            with st.expander(f"🔍 {Path(name).name}"):
                reader = open_reader_from_bytes(data)
                if not reader:
                    continue
                pages = len(reader.pages)
                if not _HAVE_FITZ:
                    st.info("Install **PyMuPDF** (`pip install pymupdf`) to enable page previews.")
                else:
                    c1, c2 = st.columns([3,2])
                    with c1:
                        slider_key = _unique_key("slider", name, data)
                        pnum = st.slider(f"Page for {Path(name).name}", 1, pages, 1, key=slider_key)
                        png = render_page_image(data, int(pnum)-1, zoom=1.5)
                        if png:
                            st.image(png, caption=f"{Path(name).name} — Page {pnum}", use_column_width=True)
                    with c2:
                        st.write("Quick jump")
                        jump_key = _unique_key("jump", name, data)
                        jump = st.text_input(f"Go to page (1..{pages})", value="", key=jump_key)
                        if jump.strip().isdigit():
                            j = int(jump)
                            if 1 <= j <= pages:
                                st.session_state[slider_key] = j
                        viewer_key = _unique_key("viewer2", name, data)
                        show_full = st.checkbox("Inline full PDF viewer", value=False, key=viewer_key)
                        if show_full:
                            embed_pdf_viewer(data, height=450)


st.markdown("---")
st.subheader("Merge")

colA, colB, colC = st.columns([2,1,1])
with colA:
    out_name = st.text_input("Output file name", value="merged.pdf")
with colB:
    add_bookmarks = st.checkbox("Add bookmarks per source file", value=True)
with colC:
    keep_file_order = st.selectbox("File order", ["Upload/scan order (default)", "Alphabetical by name"], index=0)

if file_entries and selections:
    if keep_file_order == "Alphabetical by name":
        file_entries = sorted(file_entries, key=lambda t: Path(t[0]).name.lower())
    # Build merged PDF
    if st.button("Merge PDFs"):
        writer = PdfWriter()
        for name, data in file_entries:
            reader = open_reader_from_bytes(data)
            if not reader:
                continue
            maxp = len(reader.pages)
            wanted = selections.get(name, list(range(1, maxp + 1)))
            if add_bookmarks and wanted:
                try:
                    # Bookmark to first added page for this file
                    start_idx = len(writer.pages)
                    if hasattr(writer, "add_outline_item"):
                        writer.add_outline_item(Path(name).name, start_idx)
                    elif hasattr(writer, "add_bookmark"):
                        writer.add_bookmark(Path(name).name, start_idx)
                except Exception:
                    pass
            for p1 in wanted:
                idx0 = p1 - 1
                if 0 <= idx0 < maxp:
                    writer.add_page(reader.pages[idx0])

        bio = io.BytesIO()
        writer.write(bio)
        bio.seek(0)

        st.success(f"Merged {len(file_entries)} file(s) into '{out_name}'.")
        st.download_button("Download merged PDF", data=bio, file_name=out_name, mime="application/pdf")
else:
    st.info("Add files and valid page selections to enable merging.")


# Export selection (unchanged but with safe guards)
if file_entries and selections:
    mapping_out = {name: ",".join(map(str, pages)) for name, pages in selections.items()}
    col1, col2 = st.columns(2)
    with col1:
        if _HAVE_YAML:
            if st.button("Export mapping (YAML)"):
                yaml_bytes = yaml.safe_dump(mapping_out, allow_unicode=True).encode("utf-8")
                st.download_button("Download selections.yaml", data=yaml_bytes,
                                   file_name="selections.yaml", mime="text/yaml")
        else:
            st.caption("Install **pyyaml** to export YAML mapping.")
    with col2:
        json_bytes = json.dumps(mapping_out, ensure_ascii=False, indent=2).encode("utf-8")
        st.download_button("Download selections.json", data=json_bytes,
                           file_name="selections.json", mime="application/json")
