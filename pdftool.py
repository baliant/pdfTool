import io
import json
from pathlib import Path
from typing import Dict, List, Union

import streamlit as st
from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError

try:
    import yaml  # type: ignore
    _HAVE_YAML = True
except Exception:
    _HAVE_YAML = False


icon_url = 📚

st.set_page_config(
    page_title="PDF Select & Merge",
    page_icon=icon_url,
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


st.title("📚 PDF Select & Merge")
st.caption("List page counts, pick pages per file, merge, and download. (1-based page numbering)")


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

tab_upload, tab_folder = st.tabs(["📤 Upload PDFs", "📁 Folder path (local run)"])

selections: Dict[str, List[int]] = {}
file_entries = []  # list of tuples: (display_name, data_bytes)

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
            col1, col2 = st.columns([2,1])
            with col1:
                st.markdown(f"**{f.name}** — {pages} pages")
                default_spec = "all"
                # Apply mapping if provided (basename or exact name)
                if uploaded_mapping:
                    if f.name in uploaded_mapping:
                        default_spec = uploaded_mapping[f.name]
                    else:
                        # also try full "name" as key (for symmetry with folder mode)
                        default_spec = uploaded_mapping.get(f.name, default_spec)
                spec = st.text_input(f"Pages for {f.name}", value=str(default_spec), key=f"spec_{f.name}")
            with col2:
                st.text("")  # spacing
                st.text("")  # spacing
                st.caption("Preview (count only)")

            # validate
            try:
                pages_list = parse_pagespec(spec, pages)
                if not pages_list:
                    st.warning("No valid pages selected; this file will be skipped.")
                else:
                    st.caption(f"Selected pages: {len(pages_list)}")
                    selections[f.name] = pages_list
                    file_entries.append((f.name, data))
            except Exception as e:
                st.error(f"Invalid page spec for {f.name}: {e}")

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
                        # keys can be full path or basename
                        if str(pf) in folder_mapping:
                            default_spec = folder_mapping[str(pf)]
                        elif pf.name in folder_mapping:
                            default_spec = folder_mapping[pf.name]
                    spec = st.text_input(f"Pages for {pf}", value=str(default_spec), key=f"spec_{pf}")
                    try:
                        pages_list = parse_pagespec(spec, pages)
                        if not pages_list:
                            st.warning("No valid pages selected; this file will be skipped.")
                        else:
                            selections[str(pf)] = pages_list
                            file_entries.append((str(pf), data))
                    except Exception as e:
                        st.error(f"Invalid page spec for {pf}: {e}")


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
                    # bookmark to first added page for this file
                    # We'll remember the current page count in writer
                    start_idx = len(writer.pages)
                    writer.add_outline_item(Path(name).name, start_idx)
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
