import io
import json
import re
import time
import zipfile
from collections import defaultdict

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="DienmayXANH Product Scraper", page_icon="📦", layout="wide")

st.title("📦 DienmayXANH Product Scraper")
st.caption(
    "Nhập danh sách ID sản phẩm (hoặc dán thẳng link sản phẩm đầy đủ), tool sẽ lấy tên, "
    "mã sản phẩm, danh mục (category), toàn bộ ảnh gallery, và bảng thông số kỹ thuật. "
    "Kết quả được nhóm theo category và có thể tải ảnh về dạng file ZIP."
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
- Kết quả được **tự động nhóm theo Category** (lấy từ `data-cate-id` + tên category trên
  breadcrumb của trang sản phẩm).
- Mỗi sản phẩm hiển thị **Mã sản phẩm (product id nội bộ)** đi kèm **URL sản phẩm**.
- Có nút **tải ZIP** để tải toàn bộ ảnh gallery về máy, ảnh được lưu vào thư mục con đặt
  theo tên sản phẩm.
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


def safe_folder_name(name: str, fallback: str) -> str:
    name = (name or fallback or "san-pham").strip()
    name = re.sub(r'[\\/:*?"<>|]+', "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:100] if name else fallback


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
        # Skip the first "Trang chủ" link if present; take the last category-like link
        # (the one right before the (optional) product name / current page).
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
    if not gallery:
        og_img = soup.find("meta", property="og:image")
        if og_img and og_img.get("content"):
            gallery = [strip_size_suffix(og_img["content"])]

    # Specs
    # Each spec row is normally <li><aside><strong>Label:</strong></aside><aside>...value(s)...</aside></li>.
    # Some rows (e.g. "Tiện ích") pack several values as separate <span class="circle"> tags with no
    # punctuation between them (e.g. "Khoá cửa tủ" "Giỏ đựng đồ" "Bánh xe") — get_text() on the whole
    # <li> would glue them together with no separator, so those are joined explicitly with ", ".
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
                    if len(parts) > 1:
                        value = ", ".join(parts)
                    else:
                        value = value_container.get_text(" ", strip=True)
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
    # Clear any previously built image zip since the data set changed
    st.session_state.pop("image_zip", None)

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
        # ---- Group by category ----
        groups = defaultdict(list)
        for r in ok_results:
            key = (r.get("cate_id") or "?", r.get("cate_name") or "Chưa xác định danh mục")
            groups[key].append(r)

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
        st.dataframe(df, use_container_width=True)

        # ---- Excel export (grouped by category, with product id + url together) ----
        max_gallery = max((len(r["gallery"]) for r in ok_results), default=0)
        excel_rows = []
        for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
            for r in items:
                row = {
                    "Category ID": cate_id,
                    "Category": cate_name,
                    "Mã SP": r.get("product_id") or r["input"],
                    "Tên sản phẩm": r["name"],
                    "URL sản phẩm": r["url"],
                }
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
            "⬇️ Tải Excel (theo category)",
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

        # ---- Download all gallery images as ZIP, folders named per product ----
        st.subheader("📦 Tải ảnh về (ZIP)")
        st.caption(
            "Tải toàn bộ ảnh gallery của các sản phẩm đã lấy thành 1 file ZIP duy nhất — mỗi "
            "sản phẩm 1 thư mục con đặt theo tên sản phẩm, được nhóm theo category."
        )
        if st.button("🖼️ Chuẩn bị file ZIP ảnh"):
            zip_buf = io.BytesIO()
            used_names = set()
            total_imgs = sum(len(r["gallery"]) for r in ok_results)
            dl_progress = st.progress(0.0, text="Đang tải ảnh...")
            done = 0
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
                    cate_folder = safe_folder_name(cate_name, f"category-{cate_id}")
                    for r in items:
                        base_name = safe_folder_name(
                            r["name"], r.get("product_id") or r["input"]
                        )
                        folder_name = f"{base_name} (ID {r.get('product_id') or r['input']})"
                        n = 1
                        candidate = folder_name
                        while candidate in used_names:
                            n += 1
                            candidate = f"{folder_name} ({n})"
                        used_names.add(candidate)
                        product_folder = candidate

                        for idx, img_url in enumerate(r["gallery"], 1):
                            try:
                                img_resp = session.get(img_url, timeout=20)
                                if img_resp.status_code == 200:
                                    ext = re.search(r"\.(\w+)(?:\?.*)?$", img_url)
                                    ext = ext.group(1) if ext else "jpg"
                                    arcname = f"{cate_folder}/{product_folder}/{idx:02d}.{ext}"
                                    zf.writestr(arcname, img_resp.content)
                            except Exception:  # noqa: BLE001
                                pass
                            done += 1
                            if total_imgs:
                                dl_progress.progress(
                                    done / total_imgs, text=f"Đang tải ảnh... {done}/{total_imgs}"
                                )
            dl_progress.empty()
            zip_buf.seek(0)
            st.session_state["image_zip"] = zip_buf.getvalue()
            st.success("Đã chuẩn bị xong file ZIP ảnh, bấm nút bên dưới để tải về.")

        if "image_zip" in st.session_state:
            st.download_button(
                "⬇️ Tải ZIP ảnh (theo category / tên sản phẩm)",
                data=st.session_state["image_zip"],
                file_name="dienmayxanh_images.zip",
                mime="application/zip",
            )

        # ---- Preview grouped by category ----
        st.subheader("Xem trước từng sản phẩm (nhóm theo Category)")
        for (cate_id, cate_name), items in sorted(groups.items(), key=lambda kv: kv[0][1]):
            st.markdown(f"### 🗂️ {cate_name}  `(cate_id: {cate_id})` — {len(items)} sản phẩm")
            for r in items:
                label = f"{r['name']}  —  Mã SP: {r.get('product_id') or r['input']}"
                with st.expander(label):
                    st.markdown(f"**Mã sản phẩm:** `{r.get('product_id') or r['input']}`")
                    st.markdown(f"**URL:** {r['url']}")
                    cols = st.columns(min(len(r["gallery"]), 4) or 1)
                    for i, img_url in enumerate(r["gallery"]):
                        cols[i % len(cols)].image(img_url, use_container_width=True)
                    st.table(pd.DataFrame(r["specs"].items(), columns=["Thông số", "Giá trị"]))
