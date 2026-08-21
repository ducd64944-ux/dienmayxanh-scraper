import io
import json
import re
import time
import zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DienmayXANH Product Scraper", page_icon="📦", layout="wide")

# ---------------------------------------------------------------------------
# Cookie mặc định (dienmayxanh.com) — nhúng sẵn để tool tự resolve link rút
# gọn /sp-{id} mà không cần tải cookies.json lên mỗi lần dùng (dễ quên / dễ
# nhầm file). Cookie có thể hết hạn theo thời gian — nếu resolve id bắt đầu
# lỗi 404 hàng loạt, hãy tải cookies.json mới (xuất từ trình duyệt đã đăng
# nhập dienmayxanh.com) lên ở mục "Cookie nâng cao" bên dưới để ghi đè.
# ---------------------------------------------------------------------------
DEFAULT_COOKIES = [
    {"name": "_ce.clock_data", "domain": ".dienmayxanh.com",
     "value": "-140%2C113.161.36.0%2C1%2C6967ec7261b3cbe6a91d798c6b951c60%2CChrome%2CVN%2CDirect%20Traffic%2C"},
    {"name": "_ce.s", "domain": ".dienmayxanh.com",
     "value": "v~fae3bb319ee26758400d38572a0b2badd124a3fd~vir~returning~lva~1787296884142~vpv~3~v11ls~ea633fc0-9d30-11f1-a170-29257f06f2e4~v11.cs~454123~v11.s~ea633fc0-9d30-11f1-a170-29257f06f2e4~v11.vs~fae3bb319ee26758400d38572a0b2badd124a3fd~v11.sla~1787296884159~v11.wss~1787296884160~v11.ss~1787296884163~lcw~1787296884165"},
    {"name": "_fbp", "domain": ".dienmayxanh.com", "value": "fb.1.1782793183066.926715296684095679"},
    {"name": "_ga", "domain": ".dienmayxanh.com", "value": "GA1.1.763391756.1782793167"},
    {"name": "_ga_Y7SWKJEHCE", "domain": ".dienmayxanh.com",
     "value": "GS2.1.s1787296883$o4$g0$t1787296884$j59$l0$h0"},
    {"name": "_gcl_au", "domain": ".dienmayxanh.com", "value": "1.1.1295786637.1782793167"},
    {"name": "cebs", "domain": ".dienmayxanh.com", "value": "1"},
    {"name": "cebsp_", "domain": ".dienmayxanh.com", "value": "1"},
    {"name": "DMX_Personal", "domain": ".dienmayxanh.com",
     "value": "%7B%22CustomerId%22%3A0%2C%22CustomerSex%22%3A-1%2C%22CustomerName%22%3Anull%2C%22CustomerPhone%22%3Anull%2C%22CustomerMail%22%3Anull%2C%22Lat%22%3A0.0%2C%22Lng%22%3A0.0%2C%22Address%22%3Anull%2C%22CurrentUrl%22%3Anull%2C%22ProvinceId%22%3A1027%2C%22ProvinceType%22%3Anull%2C%22ProvinceName%22%3A%22H%E1%BB%93%20Ch%C3%AD%20Minh%22%2C%22DistrictId%22%3A0%2C%22DistrictType%22%3Anull%2C%22DistrictName%22%3Anull%2C%22WardId%22%3A0%2C%22WardType%22%3Anull%2C%22WardName%22%3Anull%2C%22StoreId%22%3A0%2C%22CouponCode%22%3Anull%2C%22HasLocation%22%3Afalse%7D"},
    {"name": "mwgsp", "domain": ".dienmayxanh.com", "value": "1"},
    {"name": "__admUTMtime", "domain": ".www.dienmayxanh.com", "value": "1787293383"},
    {"name": "__iid", "domain": ".www.dienmayxanh.com", "value": ""},
    {"name": "__su", "domain": ".www.dienmayxanh.com", "value": "0"},
    {"name": "__uidac", "domain": ".www.dienmayxanh.com", "value": "016a4343cfc80706dd0a2cd05d400c31"},
    {"name": "__iid", "domain": "www.dienmayxanh.com", "value": ""},
    {"name": "__IP", "domain": "www.dienmayxanh.com", "value": "1906387520"},
    {"name": "__R", "domain": "www.dienmayxanh.com", "value": "3"},
    {"name": "__RC", "domain": "www.dienmayxanh.com", "value": "5"},
    {"name": "__su", "domain": "www.dienmayxanh.com", "value": "0"},
    {"name": "__tb", "domain": "www.dienmayxanh.com", "value": "0"},
    {"name": "__uif", "domain": "www.dienmayxanh.com",
     "value": "__uid%3A7105658824095331495%7C__ui%3A-1%7C__create%3A1780565884"},
    {"name": "_customerIdRecommend", "domain": "www.dienmayxanh.com", "value": "e30cda076bfe1ee5"},
    {"name": "chatmode", "domain": "www.dienmayxanh.com", "value": "1"},
    {"name": "popup_banner_home", "domain": "www.dienmayxanh.com", "value": "popup_banner_H_1days"},
    {"name": "SvID", "domain": "www.dienmayxanh.com", "value": "new26124|aof8d|aof8d"},
]

# ---------------------------------------------------------------------------
# UI styling
# ---------------------------------------------------------------------------
CUSTOM_CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root {
    --dmx-accent: #e30613;
    --dmx-accent-soft: #fff1f1;
    --dmx-accent-2: #ff7a45;
}
html, body, [class*="css"] { font-family: 'Inter', -apple-system, 'Segoe UI', sans-serif; }
.dmx-hero {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, #ff7a45 0%, #e30613 55%, #8f0210 100%);
    padding: 30px 34px;
    border-radius: 20px;
    color: #fff;
    margin-bottom: 16px;
    box-shadow: 0 14px 34px rgba(227,6,19,0.30);
}
.dmx-hero::after {
    content: "";
    position: absolute; top: -60px; right: -60px;
    width: 220px; height: 220px; border-radius: 50%;
    background: radial-gradient(circle, rgba(255,255,255,0.16) 0%, rgba(255,255,255,0) 70%);
}
.dmx-hero h1 { margin: 0 0 6px 0; font-size: 26px; font-weight: 800; letter-spacing: -0.02em; }
.dmx-hero p { margin: 0; opacity: 0.95; font-size: 14px; line-height: 1.55; max-width: 720px; }
.dmx-chips { display: flex; flex-wrap: wrap; gap: 8px; margin: 14px 0 20px 0; }
.dmx-chip {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 6px 14px; border-radius: 999px; font-size: 12.5px; font-weight: 600;
    background: var(--dmx-accent-soft); color: var(--dmx-accent);
    border: 1px solid rgba(227,6,19,0.12);
}
.dmx-badge {
    display: inline-block; padding: 2px 10px; border-radius: 999px;
    font-size: 11px; font-weight: 700; background: var(--dmx-accent-soft);
    color: var(--dmx-accent); margin-left: 8px; vertical-align: middle;
}
.dmx-cat-header {
    display: flex; align-items: center; gap: 8px;
    padding: 12px 18px; border-radius: 14px;
    background: linear-gradient(90deg, var(--dmx-accent-soft) 0%, rgba(255,255,255,0) 100%);
    border-left: 4px solid var(--dmx-accent);
    margin: 24px 0 14px 0;
}
.dmx-cat-header h3 { margin: 0; font-size: 16.5px; font-weight: 700; color: #b5030f; }
.dmx-cat-icon { font-size: 20px; }
.dmx-card {
    border: 1px solid rgba(120,120,120,0.16);
    border-radius: 16px;
    padding: 14px;
    margin-bottom: 18px;
    background: rgba(127,127,127,0.02);
    transition: box-shadow .18s ease, transform .18s ease, border-color .18s ease;
}
.dmx-card:hover {
    box-shadow: 0 10px 26px rgba(0,0,0,0.12);
    transform: translateY(-2px);
    border-color: rgba(227,6,19,0.35);
}
.dmx-card-name { font-weight: 700; font-size: 14.5px; margin: 8px 0 3px 0; line-height: 1.35; }
.dmx-card-meta { font-size: 12px; opacity: 0.62; margin-bottom: 8px; }
.dmx-card-meta code {
    background: var(--dmx-accent-soft); color: var(--dmx-accent);
    padding: 1px 6px; border-radius: 6px; font-weight: 600;
}
[data-testid="stImage"] img { border-radius: 12px; }
[data-testid="stMetric"] {
    background: rgba(227,6,19,0.045);
    border: 1px solid rgba(227,6,19,0.10);
    border-radius: 14px;
    padding: 10px 6px 6px 6px;
}
button[kind="primary"] { border-radius: 10px !important; font-weight: 700 !important; }
</style>"""

CATEGORY_ICON_RULES = [
    (("máy lạnh", "điều hòa", "điều hoà"), "❄️"),
    (("tủ lạnh", "tủ đông", "tủ mát"), "🧊"),
    (("máy giặt", "sấy"), "🌀"),
    (("tivi", "tv", "smart tivi"), "📺"),
    (("quạt",), "🌬️"),
    (("bếp", "lò vi sóng", "lò nướng"), "🍳"),
    (("máy lọc nước", "lọc nước"), "💧"),
    (("máy lọc không khí", "lọc không khí"), "🌫️"),
    (("bình nóng lạnh", "nước nóng"), "🚿"),
    (("loa", "âm thanh"), "🔊"),
    (("điện thoại", "laptop", "máy tính"), "💻"),
]


def category_icon(cate_name: str) -> str:
    low = (cate_name or "").lower()
    for keywords, icon in CATEGORY_ICON_RULES:
        if any(k in low for k in keywords):
            return icon
    return "📂"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
}


def build_session(cookie_file) -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    cookies_to_set = DEFAULT_COOKIES
    if cookie_file is not None:
        try:
            cookie_file.seek(0)
            uploaded = json.load(cookie_file)
            if isinstance(uploaded, list) and uploaded:
                cookies_to_set = uploaded
        except Exception as e:  # noqa: BLE001
            st.warning(f"Không đọc được file cookie vừa tải lên, dùng cookie mặc định thay thế. Lỗi: {e}")
    for c in cookies_to_set:
        name = c.get("name")
        value = c.get("value")
        domain = (c.get("domain") or "").lstrip(".") or "www.dienmayxanh.com"
        if name and value is not None:
            s.cookies.set(name, value, domain=domain)
    return s


def strip_size_suffix(url: str) -> str:
    if not url:
        return url
    return re.sub(r"-\d+x\d+(?=\.\w+(\?.*)?$)", "", url)


def safe_folder_name(name: str, fallback: str) -> str:
    name = (name or fallback or "san-pham").strip()
    name = re.sub(r'[\\/:*?"<>|]+', "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:100] if name else fallback


ACCENT_FILL = PatternFill(start_color="FFE30613", end_color="FFE30613", fill_type="solid")
SOFT_FILL = PatternFill(start_color="FFFFF1F1", end_color="FFFFF1F1", fill_type="solid")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF", size=13)
LABEL_BOLD = Font(bold=True)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")


def build_specs_workbook(items) -> bytes:
    """1 sản phẩm = 1 khối: tiêu đề + Mã SP/URL + bảng Thông số/Giá trị — trình bày dọc,
    giống hệt cách hiển thị 'Thông số | Giá trị' trên web, dễ đọc hơn nhiều so với việc
    nhồi hàng chục cột thông số ngang trên 1 dòng."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Thong so ky thuat"
    ws.column_dimensions[get_column_letter(1)].width = 34
    ws.column_dimensions[get_column_letter(2)].width = 85

    row = 1
    for r in items:
        ws.cell(row=row, column=1, value=r["name"] or r["input"]).font = WHITE_BOLD
        ws.cell(row=row, column=1).fill = ACCENT_FILL
        ws.cell(row=row, column=2).fill = ACCENT_FILL
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.cell(row=row, column=1).alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 24
        row += 1

        for label, value in (
            ("Mã sản phẩm", r.get("product_id") or r["input"]),
            ("URL sản phẩm", r["url"]),
            ("Category", r.get("cate_name") or ""),
            ("Số ảnh gallery", len(r["gallery"])),
        ):
            ws.cell(row=row, column=1, value=label).font = LABEL_BOLD
            ws.cell(row=row, column=2, value=value)
            row += 1

        row += 1
        ws.cell(row=row, column=1, value="Thông số").font = LABEL_BOLD
        ws.cell(row=row, column=2, value="Giá trị").font = LABEL_BOLD
        ws.cell(row=row, column=1).fill = SOFT_FILL
        ws.cell(row=row, column=2).fill = SOFT_FILL
        row += 1

        if r["specs"]:
            for k, v in r["specs"].items():
                ws.cell(row=row, column=1, value=k).alignment = WRAP_TOP
                ws.cell(row=row, column=2, value=v).alignment = WRAP_TOP
                row += 1
        else:
            ws.cell(row=row, column=1, value="(không lấy được thông số)")
            row += 1

        row += 2  # dòng trống ngăn cách giữa các sản phẩm

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_images_workbook(items) -> bytes:
    """Danh sách link ảnh riêng biệt — mỗi dòng 1 ảnh, tách hẳn khỏi bảng thông số để
    không bị dồn chung thành 1 bảng ngang rối mắt."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Link hinh anh"
    headers = ["STT", "Mã SP", "Tên sản phẩm", "STT ảnh", "Link ảnh"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = ACCENT_FILL
    stt = 1
    for r in items:
        for gi, url in enumerate(r["gallery"], 1):
            ws.append([stt, r.get("product_id") or r["input"], r["name"], gi, url])
            stt += 1
    for col, w in zip(range(1, 6), [6, 14, 42, 8, 95]):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def scrape_one(session: requests.Session, entry: str, retries: int = 2):
    entry = entry.strip()
    if not entry:
        return None

    if entry.isdigit():
        url = f"https://www.dienmayxanh.com/sp-{entry}"
    elif entry.startswith("http"):
        url = entry
    else:
        return {"input": entry, "error": "Không nhận diện được định dạng (không phải id số hoặc link http)."}

    last_err = None
    for attempt in range(retries + 1):
        try:
            resp = session.get(url, timeout=15, allow_redirects=True)
            break
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(0.5)
    else:
        return {"input": entry, "error": f"Lỗi kết nối sau {retries + 1} lần thử: {last_err}"}

    if resp.status_code == 404 or "ERROR 404" in resp.text[:2000]:
        return {
            "input": entry,
            "error": (
                "404 - không resolve được (thường do link rút gọn /sp-{id} cần cookie/trình "
                "duyệt thật mới redirect đúng, hoặc cookie mặc định đã hết hạn). Hãy mở link "
                "này 1 lần trên trình duyệt và dán URL đầy đủ thay cho id, hoặc tải cookies.json "
                "mới ở mục Cookie nâng cao."
            ),
        }
    if resp.status_code != 200:
        return {"input": entry, "error": f"HTTP {resp.status_code}"}

    soup = BeautifulSoup(resp.text, "html.parser")

    h1 = soup.find("h1")
    name = h1.get_text(strip=True) if h1 else None

    # --- Product id (internal) + category id, from <section data-id=".." data-cate-id="..">
    product_id = None
    cate_id = None
    section = soup.select_one("section.detail[data-id]") or soup.select_one("section[data-cate-id]")
    if section:
        product_id = section.get("data-id")
        cate_id = section.get("data-cate-id")

    # --- Category name + slug, from breadcrumb
    cate_name = None
    cate_slug = None
    breadcrumb = soup.select_one("ul.breadcrumb") or soup.select_one(".breadcrumb")
    if breadcrumb:
        links = breadcrumb.select("a[href]")
        candidates = [a for a in links if a.get("href") and a.get("href") != "/"]
        if candidates:
            last = candidates[-1]
            cate_name = last.get_text(strip=True)
            cate_slug = last.get("href")

    # Gallery images: main slider container
    gallery = []
    container = soup.find(id="slider-default")
    if container:
        seen = set()
        for img in container.find_all("img"):
            for attr in ("src", "data-src", "data-thumb"):
                v = img.get(attr)
                if v and "/Products/Images/" in v:
                    base = strip_size_suffix(v)
                    if base not in seen:
                        seen.add(base)
                        gallery.append(base)
    og_img = soup.find("meta", property="og:image")
    if og_img and og_img.get("content"):
        og_url = strip_size_suffix(og_img["content"])
        seen_urls = {strip_size_suffix(u) for u in gallery}
        if og_url not in seen_urls:
            gallery.insert(0, og_url)

    # Specs
    # Each spec row is normally <li><aside><strong>Label:</strong></aside><aside>...value(s)...</aside></li>.
    # Some rows (e.g. "Tiện ích") pack several values as separate <span class="circle"> tags with no
    # punctuation between them — those are joined explicitly with ", ".
    specs = {}
    spec_root = soup.select_one("#tab-2 .specification-item") or soup.select_one(".specification-item")
    if spec_root:
        for box in spec_root.select(".box-specifi"):
            for li in box.select("ul.text-specifi > li"):
                asides = li.find_all("aside", recursive=False)
                if len(asides) >= 2:
                    label = asides[0].get_text(" ", strip=True).rstrip(":").strip()
                    value_container = asides[1]
                    parts = [
                        c.get_text(" ", strip=True)
                        for c in value_container.find_all(["span", "a"])
                        if c.get_text(strip=True)
                    ]
                    value = ", ".join(parts) if len(parts) > 1 else value_container.get_text(" ", strip=True)
                    if label:
                        specs[label] = re.sub(r"\s+", " ", value.strip())
                else:
                    text = li.get_text(" ", strip=True)
                    if ":" in text:
                        label, value = text.split(":", 1)
                        specs[label.strip()] = re.sub(r"\s+", " ", value.strip())

    return {
        "input": entry,
        "url": resp.url,
        "name": name,
        "product_id": product_id,
        "cate_id": cate_id,
        "cate_name": cate_name,
        "cate_slug": cate_slug,
        "gallery": gallery,
        "specs": specs,
        "error": None,
    }


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(
    """
<div class="dmx-hero">
  <h1>📦 DienmayXANH Product Scraper</h1>
  <p>Lấy tên, mã sản phẩm, category, toàn bộ ảnh gallery và thông số kỹ thuật từ dienmayxanh.com —
  tự nhóm theo category, tải song song nhiều luồng, xuất dữ liệu tách riêng theo category chỉ trong vài cú click.</p>
</div>
<div class="dmx-chips">
  <span class="dmx-chip">🍪 Cookie có sẵn</span>
  <span class="dmx-chip">⚡ Tải đa luồng</span>
  <span class="dmx-chip">🗂️ Tự nhóm Category</span>
  <span class="dmx-chip">📁 Xuất riêng từng Category</span>
  <span class="dmx-chip">📦 ZIP ảnh gốc</span>
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### ⚙️ Cấu hình")
    ids_text = st.text_area(
        "Danh sách ID hoặc link (mỗi dòng 1 mục)",
        height=180,
        placeholder="370579\n370580\nhttps://www.dienmayxanh.com/may-lanh/...",
    )
    max_workers = st.slider("⚡ Số luồng tải song song", min_value=1, max_value=10, value=5)

    with st.expander("🍪 Cookie nâng cao", expanded=False):
        st.caption(
            "Tool đã có sẵn cookie mặc định để resolve link rút gọn `/sp-{id}` — không cần tải "
            "cookies.json mỗi lần dùng nữa. Nếu cookie mặc định hết hạn (id bắt đầu lỗi 404 hàng "
            "loạt), tải cookies.json mới lên đây để ghi đè."
        )
        cookies_file = st.file_uploader(
            "Tuỳ chọn: tải cookies.json mới (ghi đè cookie mặc định)", type=["json"]
        )

    run = st.button("🚀 Lấy dữ liệu", type="primary", use_container_width=True)

    with st.expander("ℹ️ Hướng dẫn / lưu ý", expanded=False):
        st.markdown(
            """
- Mỗi dòng nhập **1 ID** (vd: `370579`) hoặc **1 link đầy đủ**.
- Với **ID**, tool resolve qua `https://www.dienmayxanh.com/sp-{id}` bằng cookie mặc định đã
  nhúng sẵn. Nếu id nào không resolve được, mở link đó 1 lần trên trình duyệt rồi dán URL đầy
  đủ thay cho id.
- Kết quả được **tự động nhóm theo Category** (từ `data-cate-id` + breadcrumb).
- Mỗi sản phẩm hiển thị **Mã sản phẩm** đi kèm **URL**.
- Có nút **tải ZIP ảnh** — mỗi sản phẩm 1 thư mục con, nhóm theo category.
            """
        )

if run:
    entries = [line for line in ids_text.splitlines() if line.strip()]
    if not entries:
        st.warning("Hãy nhập ít nhất 1 id hoặc link.")
        st.stop()

    session = build_session(cookies_file)
    results = [None] * len(entries)
    progress = st.progress(0.0, text="Đang lấy dữ liệu...")
    done_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(scrape_one, session, entry): i for i, entry in enumerate(entries)}
        for future in as_completed(future_map):
            idx = future_map[future]
            try:
                results[idx] = future.result()
            except Exception as exc:  # noqa: BLE001
                results[idx] = {"input": entries[idx], "error": f"Lỗi không xác định: {exc}"}
            done_count += 1
            progress.progress(done_count / len(entries), text=f"Đã xử lý {done_count}/{len(entries)}")

    progress.empty()
    st.session_state["results"] = results
    st.session_state.pop("image_zip", None)
    st.session_state.pop("export_zip", None)

if "results" in st.session_state:
    results = st.session_state["results"]
    ok_results = [r for r in results if r and not r.get("error")]
    err_results = [r for r in results if r and r.get("error")]

    groups = defaultdict(list)
    for r in ok_results:
        key = (r.get("cate_id") or "?", r.get("cate_name") or "Chưa xác định danh mục")
        groups[key].append(r)

    total_imgs = sum(len(r["gallery"]) for r in ok_results)

    with st.container(border=True):
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("📄 Tổng mục nhập", len(results))
        m2.metric("✅ Thành công", len(ok_results))
        m3.metric("❌ Lỗi", len(err_results))
        m4.metric("🗂️ Category", len(groups))
        m5.metric("🖼️ Tổng ảnh", total_imgs)

    if err_results:
        with st.expander(f"⚠️ {len(err_results)} mục lỗi", expanded=False):
            for r in err_results:
                st.write(f"- **{r['input']}**: {r['error']}")

    if ok_results:
        spec_keys = []
        for r in ok_results:
            for k in r["specs"].keys():
                if k not in spec_keys:
                    spec_keys.append(k)

        tab_cards, tab_table, tab_export = st.tabs(
            ["🗂️ Theo Category", "📋 Bảng dữ liệu", "⬇️ Xuất dữ liệu"]
        )

        # ---- Tab 1: card grid grouped by category, with live search ----
        with tab_cards:
            query = st.text_input("🔍 Tìm theo tên sản phẩm / category / mã SP", "")
            q = query.strip().lower()
            shown_any = False
            for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
                filtered = [
                    r for r in items
                    if not q
                    or q in (r["name"] or "").lower()
                    or q in cate_name.lower()
                    or q in str(r.get("product_id") or "").lower()
                ]
                if not filtered:
                    continue
                shown_any = True
                st.markdown(
                    f'<div class="dmx-cat-header"><span class="dmx-cat-icon">{category_icon(cate_name)}</span>'
                    f'<h3>{cate_name}'
                    f'<span class="dmx-badge">cate_id {cate_id}</span>'
                    f'<span class="dmx-badge">{len(filtered)} sản phẩm</span></h3></div>',
                    unsafe_allow_html=True,
                )
                cols = st.columns(3)
                for i, r in enumerate(filtered):
                    with cols[i % 3]:
                        st.markdown('<div class="dmx-card">', unsafe_allow_html=True)
                        if r["gallery"]:
                            st.image(r["gallery"][0], use_container_width=True)
                        st.markdown(f'<div class="dmx-card-name">{r["name"]}</div>', unsafe_allow_html=True)
                        st.markdown(
                            f'<div class="dmx-card-meta">Mã SP: <code>{r.get("product_id") or r["input"]}</code>'
                            f' · {len(r["gallery"])} ảnh · {len(r["specs"])} thông số</div>',
                            unsafe_allow_html=True,
                        )
                        st.markdown(f"[🔗 Xem trên dienmayxanh.com]({r['url']})")
                        with st.expander("Xem chi tiết & thông số"):
                            if len(r["gallery"]) > 1:
                                gcols = st.columns(min(len(r["gallery"]), 4))
                                for gi, gurl in enumerate(r["gallery"]):
                                    gcols[gi % len(gcols)].image(gurl, use_container_width=True)
                            if r["specs"]:
                                st.table(pd.DataFrame(r["specs"].items(), columns=["Thông số", "Giá trị"]))
                            else:
                                st.caption("Không lấy được thông số kỹ thuật.")
                        st.markdown("</div>", unsafe_allow_html=True)
            if not shown_any:
                st.info("Không có sản phẩm nào khớp với từ khoá tìm kiếm.")

        # ---- Tab 2: flat data table ----
        with tab_table:
            rows = []
            for r in ok_results:
                row = {
                    "Mã SP": r.get("product_id") or r["input"],
                    "Tên sản phẩm": r["name"],
                    "URL sản phẩm": r["url"],
                    "Category ID": r.get("cate_id") or "",
                    "Category": r.get("cate_name") or "",
                    "Số ảnh": len(r["gallery"]),
                    **{k: r["specs"].get(k, "") for k in spec_keys},
                }
                rows.append(row)
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.markdown("**📎 Copy nhanh toàn bộ URL sản phẩm**")
            st.code("\n".join(r["url"] for r in ok_results), language=None)

        # ---- Tab 3: exports ----
        with tab_export:
            st.markdown("**📁 Xuất dữ liệu — tách riêng theo từng Category**")
            st.caption(
                "Mỗi category ra 1 thư mục riêng trong file ZIP, gồm 2 file tách biệt: "
                "**thong_so_ky_thuat.xlsx** (trình bày dọc từng sản phẩm, giống hệt bảng "
                "Thông số / Giá trị trên web — dễ đọc, không dồn hàng chục cột ngang) và "
                "**link_hinh_anh.xlsx** (chỉ danh sách link ảnh, tách riêng khỏi thông số "
                "cho khỏi rối). Ngoài ra có thêm 1 file tổng quan nhanh ở ngoài cùng."
            )
            if st.button("📁 Chuẩn bị file xuất theo category", use_container_width=True):
                export_buf = io.BytesIO()
                used_folders = set()
                with zipfile.ZipFile(export_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    overview_buf = io.BytesIO()
                    with pd.ExcelWriter(overview_buf, engine="openpyxl") as writer:
                        for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
                            sheet_name = safe_folder_name(cate_name, f"cate-{cate_id}")[:31] or f"cate-{cate_id}"
                            overview_rows = [
                                {
                                    "Mã SP": r.get("product_id") or r["input"],
                                    "Tên sản phẩm": r["name"],
                                    "URL": r["url"],
                                    "Số ảnh": len(r["gallery"]),
                                    "Số thông số": len(r["specs"]),
                                }
                                for r in items
                            ]
                            pd.DataFrame(overview_rows).to_excel(writer, index=False, sheet_name=sheet_name)
                    zf.writestr("00_tong_quan_theo_category.xlsx", overview_buf.getvalue())

                    for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
                        folder = safe_folder_name(cate_name, f"category-{cate_id}")
                        candidate = folder
                        n = 1
                        while candidate in used_folders:
                            n += 1
                            candidate = f"{folder} ({n})"
                        used_folders.add(candidate)
                        zf.writestr(f"{candidate}/thong_so_ky_thuat.xlsx", build_specs_workbook(items))
                        zf.writestr(f"{candidate}/link_hinh_anh.xlsx", build_images_workbook(items))

                export_buf.seek(0)
                st.session_state["export_zip"] = export_buf.getvalue()
                st.success("Đã chuẩn bị xong, bấm nút bên dưới để tải về.")

            if "export_zip" in st.session_state:
                st.download_button(
                    "⬇️ Tải file xuất (ZIP, tách theo category)",
                    data=st.session_state["export_zip"],
                    file_name="dienmayxanh_export_theo_category.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

            st.download_button(
                "⬇️ Tải JSON (toàn bộ dữ liệu thô, cho dev)",
                data=json.dumps(ok_results, ensure_ascii=False, indent=2),
                file_name="dienmayxanh_products.json",
                mime="application/json",
            )

            st.divider()
            st.markdown("**📦 Tải ảnh gốc về (ZIP file ảnh thật)**")
            st.caption(
                "Tải toàn bộ ảnh gallery của các sản phẩm đã lấy thành 1 file ZIP duy nhất — mỗi "
                "sản phẩm 1 thư mục con đặt theo tên sản phẩm, được nhóm theo category. Việc tải "
                "ảnh chạy ngay trong trình duyệt của bạn (không qua server) vì CDN ảnh của "
                "dienmayxanh.com chặn tải tự động từ server ngoài."
            )

            zip_items = []
            used_names_js = set()
            for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
                cate_folder = safe_folder_name(cate_name, f"category-{cate_id}")
                for r in items:
                    if not r["gallery"]:
                        continue
                    base_name = safe_folder_name(r["name"], r.get("product_id") or r["input"])
                    folder_name = f"{base_name} (ID {r.get('product_id') or r['input']})"
                    n = 1
                    candidate = folder_name
                    while candidate in used_names_js:
                        n += 1
                        candidate = f"{folder_name} ({n})"
                    used_names_js.add(candidate)
                    zip_items.append(
                        {"cate": cate_folder, "product": candidate, "images": r["gallery"]}
                    )

            zip_items_json = json.dumps(zip_items, ensure_ascii=False)
            components.html(
                f"""
<div style="font-family:Inter,-apple-system,'Segoe UI',sans-serif;">
  <button id="dmxZipBtn" style="background:#e30613;color:#fff;border:none;padding:11px 20px;
    border-radius:10px;font-weight:700;font-size:14px;cursor:pointer;width:100%;">
    🖼️ Tải ZIP ảnh (chạy trong trình duyệt của bạn)
  </button>
  <div id="dmxZipStatus" style="margin-top:8px;font-size:13px;color:#666;"></div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script>
const DMX_ITEMS = {zip_items_json};
document.getElementById('dmxZipBtn').addEventListener('click', async () => {{
  const btn = document.getElementById('dmxZipBtn');
  const status = document.getElementById('dmxZipStatus');
  btn.disabled = true;
  btn.style.opacity = '0.6';
  const zip = new JSZip();
  let total = 0, done = 0, failed = 0;
  DMX_ITEMS.forEach(it => total += it.images.length);
  if (total === 0) {{
    status.textContent = 'Không có ảnh nào để tải.';
    btn.disabled = false;
    btn.style.opacity = '1';
    return;
  }}
  for (const it of DMX_ITEMS) {{
    for (let i = 0; i < it.images.length; i++) {{
      const url = it.images[i];
      try {{
        const resp = await fetch(url);
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        const blob = await resp.blob();
        const extMatch = url.match(/\\.(\\w+)(?:\\?.*)?$/);
        const ext = extMatch ? extMatch[1] : 'jpg';
        const num = String(i + 1).padStart(2, '0');
        zip.file(it.cate + '/' + it.product + '/' + num + '.' + ext, blob);
      }} catch (e) {{
        failed++;
      }}
      done++;
      status.textContent = 'Đang tải ảnh... ' + done + '/' + total + (failed ? ' (' + failed + ' lỗi)' : '');
    }}
  }}
  status.textContent = 'Đang nén file ZIP...';
  const content = await zip.generateAsync({{type: 'blob'}});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(content);
  a.download = 'dienmayxanh_images.zip';
  document.body.appendChild(a);
  a.click();
  a.remove();
  status.textContent = 'Xong! Đã tải ' + (total - failed) + '/' + total + ' ảnh vào file ZIP.'
    + (failed ? ' (' + failed + ' ảnh lỗi.)' : '');
  btn.disabled = false;
  btn.style.opacity = '1';
}});
</script>
""",
                height=90,
            )
