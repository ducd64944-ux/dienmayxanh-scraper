import io
import json
import re
import time

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="DienmayXANH Product Scraper", page_icon="📦", layout="wide")

st.title("📦 DienmayXANH Product Scraper")
st.caption(
    "Nhập danh sách ID sản phẩm (hoặc dán thẳng link sản phẩm đầy đủ), tool sẽ lấy tên, "
    "toàn bộ ảnh gallery, và bảng thông số kỹ thuật."
)

with st.expander("ℹ️ Hướng dẫn / lưu ý", expanded=False):
    st.markdown(
        """
- Mỗi dòng nhập **1 ID** (vd: `370579`) hoặc **1 link đầy đủ**
  (vd: `https://www.dienmayxanh.com/may-lanh/...`).
- Với **ID**, tool thử resolve qua `https://www.dienmayxanh.com/sp-{id}`. Link rút gọn này
  đôi khi cần cookie hợp lệ / trình duyệt thật mới redirect đúng — nếu id nào không resolve
  được, tool sẽ báo lỗi, bạn hãy tự mở link đó **1 lần trên trình duyệt** rồi copy URL đầy đủ
  sau khi trang chuyển hướng, dán thay cho id đó.
- Có thể tải lên file cookie xuất từ trình duyệt (định dạng JSON kiểu Chrome/EditThisCookie)
  để tăng khả năng resolve id thành công.
        """
    )

cookies_file = st.file_uploader(
    "Tuỳ chọn: tải lên file cookies.json (xuất từ trình duyệt, domain dienmayxanh.com)",
    type=["json"],
)

ids_text = st.text_area(
    "Danh sách ID hoặc link (mỗi dòng 1 mục)",
    height=180,
    placeholder="370579\n370580\nhttps://www.dienmayxanh.com/may-lanh/...",
)

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
    if cookie_file is not None:
        try:
            cookie_file.seek(0)
            data = json.load(cookie_file)
            for c in data:
                name = c.get("name")
                value = c.get("value")
                domain = c.get("domain", "").lstrip(".") or "www.dienmayxanh.com"
                if name and value is not None:
                    s.cookies.set(name, value, domain=domain)
        except Exception as e:  # noqa: BLE001
            st.warning(f"Không đọc được file cookie: {e}")
    return s


def strip_size_suffix(url: str) -> str:
    if not url:
        return url
    return re.sub(r"-\d+x\d+(?=\.\w+(\?.*)?$)", "", url)


def scrape_one(session: requests.Session, entry: str):
    entry = entry.strip()
    if not entry:
        return None

    if entry.isdigit():
        url = f"https://www.dienmayxanh.com/sp-{entry}"
    elif entry.startswith("http"):
        url = entry
    else:
        return {"input": entry, "error": "Không nhận diện được định dạng (không phải id số hoặc link http)."}

    try:
        resp = session.get(url, timeout=15, allow_redirects=True)
    except Exception as e:  # noqa: BLE001
        return {"input": entry, "error": f"Lỗi kết nối: {e}"}

    if resp.status_code == 404 or "ERROR 404" in resp.text[:2000]:
        return {
            "input": entry,
            "error": (
                "404 - không resolve được (thường do link rút gọn /sp-{id} cần trình duyệt "
                "thật). Hãy mở link này 1 lần trên trình duyệt và dán URL đầy đủ thay cho id."
            ),
        }
    if resp.status_code != 200:
        return {"input": entry, "error": f"HTTP {resp.status_code}"}

    soup = BeautifulSoup(resp.text, "html.parser")

    h1 = soup.find("h1")
    name = h1.get_text(strip=True) if h1 else None

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
    if not gallery:
        og_img = soup.find("meta", property="og:image")
        if og_img and og_img.get("content"):
            gallery = [strip_size_suffix(og_img["content"])]

    # Specs
    specs = {}
    spec_root = soup.select_one("#tab-2 .specification-item") or soup.select_one(".specification-item")
    if spec_root:
        for box in spec_root.select(".box-specifi"):
            for li in box.select("ul.text-specifi > li"):
                text = li.get_text(" ", strip=True)
                if ":" in text:
                    label, value = text.split(":", 1)
                    specs[label.strip()] = re.sub(r"\s+", " ", value.strip())

    return {
        "input": entry,
        "url": resp.url,
        "name": name,
        "gallery": gallery,
        "specs": specs,
        "error": None,
    }


if st.button("🚀 Lấy dữ liệu", type="primary"):
    entries = [line for line in ids_text.splitlines() if line.strip()]
    if not entries:
        st.warning("Hãy nhập ít nhất 1 id hoặc link.")
        st.stop()

    session = build_session(cookies_file)
    results = []
    progress = st.progress(0.0, text="Đang lấy dữ liệu...")

    for i, entry in enumerate(entries, 1):
        results.append(scrape_one(session, entry))
        progress.progress(i / len(entries), text=f"Đã xử lý {i}/{len(entries)}: {entry}")
        time.sleep(0.3)

    progress.empty()
    st.session_state["results"] = results

if "results" in st.session_state:
    results = st.session_state["results"]
    ok_results = [r for r in results if r and not r.get("error")]
    err_results = [r for r in results if r and r.get("error")]

    st.success(f"Lấy thành công {len(ok_results)}/{len(results)} sản phẩm.")

    if err_results:
        with st.expander(f"⚠️ {len(err_results)} mục lỗi", expanded=True):
            for r in err_results:
                st.write(f"- **{r['input']}**: {r['error']}")

    # Build spec union for table
    spec_keys = []
    for r in ok_results:
        for k in r["specs"].keys():
            if k not in spec_keys:
                spec_keys.append(k)

    if ok_results:
        rows = []
        for r in ok_results:
            row = {
                "Tên sản phẩm": r["name"],
                "URL": r["url"],
                "Số ảnh": len(r["gallery"]),
                **{k: r["specs"].get(k, "") for k in spec_keys},
            }
            rows.append(row)
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True)

        # Excel export
        max_gallery = max((len(r["gallery"]) for r in ok_results), default=0)
        excel_rows = []
        for r in ok_results:
            row = {"Tên sản phẩm": r["name"], "URL": r["url"]}
            for i in range(max_gallery):
                row[f"Ảnh {i+1}"] = r["gallery"][i] if i < len(r["gallery"]) else ""
            for k in spec_keys:
                row[k] = r["specs"].get(k, "")
            excel_rows.append(row)
        excel_df = pd.DataFrame(excel_rows)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            excel_df.to_excel(writer, index=False, sheet_name="San pham")
        st.download_button(
            "⬇️ Tải Excel",
            data=buf.getvalue(),
            file_name="dienmayxanh_products.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "⬇️ Tải JSON",
            data=json.dumps(ok_results, ensure_ascii=False, indent=2),
            file_name="dienmayxanh_products.json",
            mime="application/json",
        )

        st.subheader("Xem trước từng sản phẩm")
        for r in ok_results:
            with st.expander(r["name"] or r["input"]):
                cols = st.columns(min(len(r["gallery"]), 4) or 1)
                for i, img_url in enumerate(r["gallery"]):
                    cols[i % len(cols)].image(img_url, use_container_width=True)
                st.table(pd.DataFrame(r["specs"].items(), columns=["Thông số", "Giá trị"]))
