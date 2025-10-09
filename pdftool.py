import io
import json
import base64
from pathlib import Path
from typing import Dict, List, Union, Tuple, Optional

import streamlit as st
from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError, DependencyError  # <-- important

# Optional YAML for presets
try:
    import yaml  # type: ignore
    _HAVE_YAML = True
except Exception:
    _HAVE_YAML = False

# Optional PyMuPDF for fast page previews (thumbnails)
try:
    import fitz  # PyMuPDF
    _HAVE_FITZ = True
except Exception:
    _HAVE_FITZ = False


ICON_URL = "https://raw.githubusercontent.com/baliant/pdftool/main/icon/koru_logo.png"

st.set_page_config(
    page_title="PDF Select, Review & Merge",
    page_icon=ICON_URL,
    layout="wide",
)

CRYPTO_HINT = (
    "Decryption requires a crypto backend. Install one of:\n"
    "`pip install cryptography`  (recommended)\n"
    "or `pip install pycryptodome`"
)

# ---------- Helpers ----------
def try_open_reader(data: bytes) -> Optional[PdfReader]:
    """Open a PdfReader safely (no page access), return None on hard failure."""
    try:
        bio = io.BytesIO(data)
        reader = PdfReader(bio)
        return reader
    except (PdfReadError, OSError) as e:
        st.warning(f"Cannot open PDF: {e}")
        return None


def try_decrypt_reader(reader: PdfReader, password: str) -> bool:
    """Attempt to decrypt. Returns True if unlocked. Surfaces crypto dependency issues nicely."""
    try:
        # pypdf returns int or bool across versions; treat >0 / True as success
        res = reader.decrypt(password)
        return bool(res)
    except DependencyError:
        st.error(f"Cannot decrypt: crypto backend missing.\n\n{CRYPTO_HINT}")
        return False
    except Exception as e:
        # Wrong password or other error
        st.warning(f"Failed to decrypt with the provided password: {e}")
        return False


def get_num_pages_safe(reader: PdfReader) -> int:
    """Return page count, guarding against encryption/crypto issues."""
    try:
        return len(reader.pages)
    except DependencyError:
        # Happens if the file is encrypted and crypto backend is missing
        st.error(f"Cannot read pages (crypto backend missing).\n\n{CRYPTO_HINT}")
        return 0
    except Exception as e:
        st.warning(f"Could not read page count: {e}")
        return 0


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
    """Render one page to PNG. Returns PNG bytes or None."""
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


def page_selector(label: str, pages: int, key: str):
    """Safe page picker: slider for ≥2 pages, number_input for 1 page, skip for 0."""
    try:
        pages = int(pages)
    except Exception:
        pages = 0

    if pages < 1:
        st.warning("This PDF appears to have no pages (cannot preview).")
        return None

    if pages == 1:
        st.caption("This PDF has 1 page.")
        return st.number_input(label, min_value=1, max_value=1, value=1, step=1, key=key)

    # pages >= 2 → slider is fine
    return st.slider(label, min_value=1, max_value=pages, value=1, step=1, key=key)


# ---------- UI ----------
st.title("📚 PDF Select, Review & Merge")
st.caption("Handles encrypted PDFs (with password), page previews, flexible selections, and merging.")

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
file_entries: List[Tuple[str, bytes]] = []  # (display_name, data_bytes)
passwords: Dict[str, str] = {}              # per-file password cache (name/path -> password)


# ----- Upload Tab -----
with tab_upload:
    uploaded = st.file_uploader("Select one or more PDFs", type=["pdf"], accept_multiple_files=True)
    mapping_file = st.file_uploader("Optional selections mapping (YAML or JSON)", type=["yaml", "yml", "json"], key="map_upload")
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
            reader = try_open_reader(data)
            if not reader:
                continue

            # ---- Encrypted handling
            pw_key = _unique_key("pw", f.name, data)
            if reader.is_encrypted:
                st.info(f"🔒 {f.name} is encrypted.")
                # try empty password first (some PDFs use empty user password)
                unlocked = False
                try:
                    unlocked = bool(reader.decrypt(""))
                except DependencyError:
                    st.error(f"{CRYPTO_HINT}")
                except Exception:
                    unlocked = False
                # ask for password if still locked
                if not unlocked:
                    pw = st.text_input(f"Password for {f.name}", type="password", key=pw_key)
                    if pw:
                        unlocked = try_decrypt_reader(reader, pw)
                        if unlocked:
                            passwords[f.name] = pw
                if not unlocked:
                    st.warning("Locked: cannot show page count/preview until decrypted.")
                    # Allow mapping entry but skip preview/count
                    default_spec = uploaded_mapping.get(f.name, "all") if uploaded_mapping else "all"
                    spec_key = _unique_key("spec", f.name, data)
                    spec = st.text_input(f"Pages for {f.name}", value=str(default_spec), key=spec_key)
                    # Don't append file_entries yet because we can't know pages/validate
                    continue  # proceed to next file

            # ---- Now safe to access pages
            pages = get_num_pages_safe(reader)
            left, right = st.columns([2, 1])
            with left:
                st.markdown(f"**{f.name}** — {pages} pages")
                default_spec = uploaded_mapping.get(f.name, "all") if uploaded_mapping else "all"
                spec_key = _unique_key("spec", f.name, data)
                spec = st.text_input(f"Pages for {f.name}", value=str(default_spec), key=spec_key)

                # Review controls
                with st.expander("Review this PDF"):
                    if not _HAVE_FITZ:
                        st.info("Install **PyMuPDF** (`pip install pymupdf`) to enable page previews.")
                    else:
                        pnum_key = _unique_key("prev", f.name, data)
                        pnum = page_selector("Preview page", pages, key=pnum_key)
                        if pnum is not None:
                            png = render_page_image(data, int(pnum) - 1, zoom=1.5)
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
                        s = ", ".join(map(str, pages_list[:10]))
                        more = "" if len(pages_list) <= 10 else " …"
                        st.text(f"{s}{more}")
                        selections[f.name] = pages_list
                        file_entries.append((f.name, data))
                except Exception as e:
                    st.error(f"Invalid page spec for {f.name}: {e}")


# ----- Folder Tab -----
with tab_folder:
    st.write("Enter a local folder path to scan for PDFs. (This only works when you run Streamlit locally.)")
    folder = st.text_input("Folder path", value="")
    mapping2 = st.file_uploader("Optional selections mapping (YAML or JSON)", type=["yaml", "yml", "json"], key="map_folder")
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
                    reader = try_open_reader(data)
                    if not reader:
                        continue

                    # Encrypted handling
                    pw_key = _unique_key("pw", str(pf))
                    if reader.is_encrypted:
                        st.info(f"🔒 {pf.name} is encrypted.")
                        unlocked = False
                        try:
                            unlocked = bool(reader.decrypt(""))
                        except DependencyError:
                            st.error(f"{CRYPTO_HINT}")
                        except Exception:
                            unlocked = False
                        if not unlocked:
                            pw = st.text_input(f"Password for {pf}", type="password", key=pw_key)
                            if pw:
                                unlocked = try_decrypt_reader(reader, pw)
                                if unlocked:
                                    passwords[str(pf)] = pw
                        if not unlocked:
                            st.warning("Locked: cannot show page count/preview until decrypted.")
                            default_spec = folder_mapping.get(str(pf), folder_mapping.get(pf.name, "all")) if folder_mapping else "all"
                            spec_key = _unique_key("spec", str(pf))
                            st.text_input(f"Pages for {pf}", value=str(default_spec), key=spec_key)
                            continue

                    pages = get_num_pages_safe(reader)
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


# ----- Global Review Tab -----
with tab_review:
    st.write("Preview uploaded or scanned PDFs. Use the controls to flip pages and verify content before merging.")
    if not file_entries:
        st.info("No files loaded yet. Upload or select a folder first.")
    else:
        for name, data in file_entries:
            with st.expander(f"🔍 {Path(name).name}"):
                reader = try_open_reader(data)
                if not reader:
                    continue

                # If encrypted, try to unlock using cached password (if any)
                if reader.is_encrypted:
                    pw = passwords.get(name, "")
                    ok = False
                    try:
                        ok = bool(reader.decrypt(pw))
                    except DependencyError:
                        st.error(f"{CRYPTO_HINT}")
                    except Exception:
                        ok = False
                    if not ok:
                        st.warning("Locked: enter password in the Upload/Folder tab to enable preview.")
                        continue

                pages = get_num_pages_safe(reader)
                if not _HAVE_FITZ:
                    st.info("Install **PyMuPDF** (`pip install pymupdf`) to enable page previews.")
                else:
                    c1, c2 = st.columns([3, 2])
                    with c1:
                        slider_key = _unique_key("slider", name, data)
                        pnum = page_selector(f"Page for {Path(name).name}", pages, key=slider_key)
                        if pnum is not None:
                            png = render_page_image(data, int(pnum) - 1, zoom=1.5)
                            if png:
                                st.image(png, caption=f"{Path(name).name} — Page {pnum}", use_column_width=True)
                    with c2:
                        st.write("Inline viewer")
                        viewer_key = _unique_key("viewer2", name, data)
                        show_full = st.checkbox("Show full PDF", value=False, key=viewer_key)
                        if show_full:
                            embed_pdf_viewer(data, height=450)


# ----- Merge -----
st.markdown("---")
st.subheader("Merge")

colA, colB, colC = st.columns([2, 1, 1])
with colA:
    out_name = st.text_input("Output file name", value="merged.pdf")
with colB:
    add_bookmarks = st.checkbox("Add bookmarks per source file", value=True)
with colC:
    keep_file_order = st.selectbox("File order", ["Upload/scan order (default)", "Alphabetical by name"], index=0)

if file_entries and selections:
    if keep_file_order == "Alphabetical by name":
        file_entries = sorted(file_entries, key=lambda t: Path(t[0]).name.lower())
    if st.button("Merge PDFs"):
        writer = PdfWriter()
        for name, data in file_entries:
            reader = try_open_reader(data)
            if not reader:
                continue

            # Unlock with cached password if needed
            if reader.is_encrypted:
                pw = passwords.get(name, "")
                ok = False
                try:
                    ok = bool(reader.decrypt(pw))
                except DependencyError:
                    st.error(f"{CRYPTO_HINT}")
                    continue
                except Exception:
                    ok = False
                if not ok:
                    st.warning(f"Skipping encrypted file '{name}' (no/invalid password).")
                    continue

            maxp = get_num_pages_safe(reader)
            if maxp < 1:
                st.warning(f"Skipping '{name}' — no readable pages.")
                continue

            wanted = selections.get(name, list(range(1, maxp + 1)))
            if add_bookmarks and wanted:
                try:
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


# ----- Export selections -----
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
