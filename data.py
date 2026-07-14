import asyncio
import os
import random
import re

import pandas as pd
from bs4 import BeautifulSoup
from tkinter import Tk, filedialog
from tqdm import tqdm

from playwright.async_api import async_playwright


# ==============================
# تنظیمات
# ==============================

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

SESSION_PATH = "session"

OUTPUT_SUFFIX = "_rasmio_full.xlsx"

# ستون‌های دستی که در سایت رسمیو وجود ندارند و باید از اکسل ورودی حفظ شوند
MANUAL_COLUMNS = [
    "نوع عرضه",
    "عملکرد کاری",
    "نوع تسویه",
    "تعداد خرید",
    "ارتباطات",
]


# ==============================
# انتخاب اکسل
# ==============================

def select_excel():

    root = Tk()
    root.withdraw()

    file = filedialog.askopenfilename(
        title="فایل اکسل شرکت ها",
        filetypes=[("Excel Files", "*.xlsx")],
    )

    return file


# ==============================
# استخراج مقدار یک فیلد ساده بر اساس برچسب
# (پیدا کردن <p>برچسب</p> و برگرداندن <p> مقدارِ هم‌سطحِ بعدی)
# ==============================

def get_field(soup, label):

    try:

        label_tags = soup.find_all(
            lambda tag: tag.name == "p"
            and tag.get_text(strip=True) == label
        )

        for label_tag in label_tags:

            label_row = label_tag.parent

            if label_row is None:
                continue

            outer = label_row.parent

            if outer is None:
                continue

            direct_ps = outer.find_all("p", recursive=False)

            for p in direct_ps:

                text = p.get_text(strip=True)

                if text and text != label:
                    return text

    except Exception:
        pass

    return ""


# ==============================
# استخراج بخش‌های متنی طولانی‌تر (مثل دارندگان حق امضا)
# که ساختارشان "برچسب / منابع: / متن / تاریخ آگهی:" است
# ==============================

def get_section_value(soup, label):

    try:

        tags = soup.find_all(
            lambda t: t.name == "p" and t.get_text(strip=True) == label
        )

        for tag in tags:

            node = tag

            for _ in range(6):

                if node.parent is None:
                    break

                node = node.parent

                text = node.get_text(" ", strip=True)

                if (
                    text.startswith(label)
                    and "منابع" in text
                    and len(text) > len(label) + 15
                ):

                    rest = text[len(label):]
                    rest = re.sub(r"^\s*منابع\s*:\s*", "", rest)

                    m = re.search(r"تاریخ آگهی:\s*(\S+)", rest)
                    citation_date = m.group(1) if m else ""

                    value = re.sub(
                        r"تاریخ آگهی:\s*\S+", "", rest
                    ).strip()

                    return value, citation_date

    except Exception:
        pass

    return "", ""


# ==============================
# استخراج نام شرکت
# ==============================

def get_company_name(soup):

    try:

        if soup.title:

            title = soup.title.get_text(strip=True)
            title = re.sub(r"^[^\w\u0600-\u06FF]+", "", title).strip()
            title = title.split(" - ")[0].strip()

            if title:
                return title

    except Exception:
        pass

    return ""


# ==============================
# استخراج آدرس، کدپستی، استان و شهرستان
# ==============================

def get_address_info(soup):

    info = {"استان": "", "شهرستان": "", "آدرس": "", "کدپستی": ""}

    try:

        addr_el = soup.find(
            attrs={
                "aria-label": lambda v: v and v.startswith("کپی آدرس")
            }
        )

        if not addr_el:
            return info

        full = addr_el.get("aria-label")
        full = re.sub(r"^کپی آدرس:\s*", "", full).strip()

        parts = full.rsplit(" ", 1)

        address = full
        postal = ""

        if len(parts) == 2 and re.fullmatch(r"\d{10}", parts[1]):
            address = parts[0].strip()
            postal = parts[1]

        info["آدرس"] = address
        info["کدپستی"] = postal

        m = re.search(
            r"استان\s+([^\u060C,]+)[،,]\s*شهرستان\s+([^\u060C,]+)",
            address,
        )

        if m:
            info["استان"] = m.group(1).strip()
            info["شهرستان"] = m.group(2).strip()

    except Exception:
        pass

    return info


# ==============================
# استخراج عمومی جداول بر اساس عنوان ستون‌ها
# ==============================

def extract_table(soup, required_headers):

    try:

        for table in soup.find_all("table"):

            header_row = table.find("tr")

            if not header_row:
                continue

            headers = [
                th.get_text(strip=True)
                for th in header_row.find_all(["td", "th"])
            ]

            headers_set = {h for h in headers if h}

            if required_headers.issubset(headers_set):

                rows = []

                for tr in table.find_all("tr")[1:]:

                    cells = [
                        td.get_text(strip=True)
                        for td in tr.find_all("td")
                    ]

                    if not any(cells):
                        continue

                    row = dict(zip(headers, cells))
                    row.pop("", None)  # حذف ستون خالی (لینک آگهی مرتبط)

                    rows.append(row)

                return rows

    except Exception:
        pass

    return []


# ==============================
# استخراج آگهی‌های روزنامه رسمی که در صفحه بارگذاری شده‌اند
# ==============================

def extract_announcements(soup):

    announcements = []

    try:

        els = soup.find_all(
            string=lambda s: s and s.strip() == "روزنامه رسمی کشور"
        )

        seen_ids = set()

        for e in els:

            node = e.parent

            for _ in range(3):

                if node.parent is None:
                    break

                node = node.parent

            if id(node) in seen_ids:
                continue

            seen_ids.add(id(node))

            ps = node.find_all("p")
            texts = [p.get_text(strip=True) for p in ps]

            d = {
                "منبع": "",
                "تاریخ چاپ": "",
                "تاریخ نامه ثبت": "",
                "عنوان آگهی": "",
                "متن آگهی": "",
            }

            if len(texts) >= 1:
                d["منبع"] = texts[0]

            if len(texts) >= 3:
                d["تاریخ چاپ"] = texts[2]

            if len(texts) >= 5:
                d["تاریخ نامه ثبت"] = texts[4]

            if len(texts) >= 6:
                d["عنوان آگهی"] = texts[5]

            if len(texts) >= 7:
                d["متن آگهی"] = texts[6]

            announcements.append(d)

    except Exception:
        pass

    return announcements


# ==============================
# تعداد کل آگهی‌ها (شامل مواردی که بدون کلیک روی صفحه نیستند)
# ==============================

def get_total_announcements_hint(soup):

    try:

        matches = soup.find_all(
            attrs={
                "aria-label": lambda v: v
                and re.fullmatch(r"مشاهده \d+ مورد بیشتر", v or "")
            }
        )

        if matches:

            last = matches[-1].get("aria-label")
            m = re.search(r"\d+", last)

            if m:
                return m.group(0)

    except Exception:
        pass

    return ""


# ==============================
# استخراج کامل اطلاعات یک شرکت از HTML صفحه
# ==============================

def parse_company_html(html, national_id):

    data = {
        "basic": {},
        "board": [],
        "risks": [],
        "subsidiaries": [],
        "inspectors": [],
        "licenses": [],
        "capital_changes": [],
        "announcements": [],
        "error": "",
    }

    try:

        soup = BeautifulSoup(html, "lxml")

        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()

        company_name = get_company_name(soup)

        site_national_id = get_field(soup, "شناسه ملی") or national_id

        address_info = get_address_info(soup)

        sig_text, sig_date = get_section_value(soup, "دارندگان حق امضا")

        board_rows = extract_table(
            soup, {"نام", "سمت", "تاریخ شروع", "تاریخ پایان"}
        )

        risk_rows = extract_table(
            soup, {"عنوان", "نوع", "شرکت/شخص مرتبط", "شرح"}
        )

        subsidiary_rows = extract_table(
            soup, {"نام", "نماینده", "نوع ارتباط", "تاریخ شروع", "تاریخ پایان"}
        )

        inspector_rows = extract_table(
            soup, {"نام", "سمت", "درجه", "سال مالی"}
        )

        license_rows = extract_table(
            soup, {"مجوز", "مرجع صادر کننده", "تاریخ شروع مجوز"}
        )

        capital_rows = extract_table(
            soup,
            {"مبلغ سرمایه (ریال)", "نوع", "محل تغییر سرمایه", "تاریخ ثبت"},
        )

        announcements = extract_announcements(soup)

        # استخراج مدیرعامل / رئیس / نایب رئیس از جدول هیئت مدیره
        ceo, chairman, vice_chairman, all_members = [], [], [], []

        for row in board_rows:

            name = row.get("نام", "")
            role = row.get("سمت", "")

            if not name:
                continue

            all_members.append(f"{name} ({role})")

            if "نائب" in role or "نایب" in role:
                vice_chairman.append(name)
            elif "رئیس" in role or "رییس" in role:
                chairman.append(name)

            if "مدیرعامل" in role:
                ceo.append(name)

        basic = {
            "اسم شرکت": company_name,
            "شناسه ملی": site_national_id,
            "شماره ثبت": get_field(soup, "شماره ثبت"),
            "کد اقتصادی": get_field(soup, "کد اقتصادی"),
            "وضعیت شرکت": get_field(soup, "وضعیت شرکت"),
            "نوع شرکت": get_field(soup, "نوع شرکت"),
            "تاریخ تاسیس": get_field(soup, "تاریخ تاسیس"),
            "آخرین آگهی روزنامه رسمی": get_field(
                soup, "آخرین آگهی روزنامه رسمی"
            ),
            "سرمایه ثبتی": get_field(soup, "سرمایه ثبتی"),
            "استان": address_info["استان"],
            "شهرستان": address_info["شهرستان"],
            "آدرس کامل": address_info["آدرس"],
            "کدپستی": address_info["کدپستی"],
            "مدیرعامل": " | ".join(dict.fromkeys(ceo)),
            "رئیس هیئت مدیره": " | ".join(dict.fromkeys(chairman)),
            "نایب رئیس هیئت مدیره": " | ".join(dict.fromkeys(vice_chairman)),
            "تعداد اعضای هیئت مدیره": len(board_rows),
            "اعضای هیئت مدیره (خلاصه)": " | ".join(all_members),
            "دارندگان حق امضا": sig_text,
            "تاریخ آگهی حق امضا": sig_date,
            "تعداد ریسک‌های ثبت‌شده": len(risk_rows),
            "تعداد شرکت‌های مادر و زیرمجموعه": len(subsidiary_rows),
            "تعداد بازرسین": len(inspector_rows),
            "تعداد مجوزها": len(license_rows),
            "تعداد تغییرات سرمایه": len(capital_rows),
            "تعداد آگهی نمایش‌داده‌شده": len(announcements),
            "تعداد کل آگهی‌ها (تخمینی از صفحه)": get_total_announcements_hint(
                soup
            ),
        }

        data["basic"] = basic
        data["board"] = board_rows
        data["risks"] = risk_rows
        data["subsidiaries"] = subsidiary_rows
        data["inspectors"] = inspector_rows
        data["licenses"] = license_rows
        data["capital_changes"] = capital_rows
        data["announcements"] = announcements

        if not company_name and not basic["تاریخ تاسیس"]:
            data["error"] = "اطلاعاتی از صفحه استخراج نشد"

    except Exception as e:
        data["error"] = str(e)

    return data


# ==============================
# جستجو و رفتن به صفحه شرکت
# ==============================

# نتیجه جستجو یک <a href> معمولی نیست؛ یک المان React با کلیک جاوااسکریپتی
# است. برای اینکه اشتباهی روی آیکون‌های فوتر/شبکه‌های اجتماعی (لینکدین و..)
# کلیک نکنیم، فقط روی کارتی که متن «سرمایه ثبتی» را دارد هدف می‌گیریم؛ این
# عبارت فقط داخل کارت نتیجه جستجوی شرکت‌ها ظاهر می‌شود.

SOCIAL_KEYWORDS = [
    "لینکدین", "توییتر", "تلگرام", "اینستاگرام", "بله",
    "linkedin", "twitter", "telegram", "instagram", "bale",
]


async def is_safe_click_target(locator):

    try:
        aria = await locator.get_attribute("aria-label")
    except Exception:
        aria = None

    if aria:
        for kw in SOCIAL_KEYWORDS:
            if kw in aria:
                return False

    try:
        href = await locator.get_attribute("href")
    except Exception:
        href = None

    if href:
        for kw in SOCIAL_KEYWORDS:
            if kw in href.lower():
                return False

    return True


async def click_first_result(page):

    # اولویت اول: لینک واقعی اگر وجود داشت
    real_link = page.locator('a[href*="/company/"]').first

    try:
        real_link_count = await page.locator('a[href*="/company/"]').count()
    except Exception:
        real_link_count = 0

    candidates = []

    if real_link_count > 0:
        candidates.append(real_link)

    # اولویت دوم: کارتی که متن «سرمایه ثبتی» دارد (مشخصه‌ی کارت نتیجه جستجو)
    card = page.locator('div:has(p:has-text("سرمایه ثبتی"))').first

    try:
        card_count = await page.locator(
            'div:has(p:has-text("سرمایه ثبتی"))'
        ).count()
    except Exception:
        card_count = 0

    if card_count > 0:
        candidates.append(card)

    for start_node in candidates:

        node = start_node

        for _level in range(4):

            if not await is_safe_click_target(node):
                break

            try:
                await node.scroll_into_view_if_needed(timeout=5000)
            except Exception:
                pass

            try:
                await node.click(timeout=4000)
            except Exception:
                pass

            try:
                await page.wait_for_url("**/company/**", timeout=3500)
                return True
            except Exception:
                pass

            try:
                node = node.locator("xpath=..")
            except Exception:
                break

    return False


async def goto_company_page(page, national_id):

    search_url = "https://rasmio.com/search/companies/?q=" + national_id

    await page.goto(search_url, wait_until="networkidle", timeout=60000)

    await page.wait_for_timeout(random.randint(2000, 3500))

    try:
        await page.wait_for_selector(
            'p:has-text("سرمایه ثبتی"), a[href*="/company/"]',
            timeout=15000,
        )
    except Exception:
        return False, ""

    clicked = await click_first_result(page)

    if not clicked:
        return False, ""

    try:
        await page.wait_for_load_state("networkidle", timeout=60000)
    except Exception:
        pass

    await page.wait_for_timeout(random.randint(2500, 4000))

    current_url = page.url

    mismatch_warning = ""

    if national_id not in current_url:
        mismatch_warning = (
            "هشدار: شناسه ملی در URL صفحه یافت نشد، احتمال ورود به "
            "شرکت اشتباه - " + current_url
        )

    return True, mismatch_warning


# ==============================
# نوشتن همه‌ی شیت‌ها در فایل اکسل
# ==============================

def write_workbook(output_path, all_data):

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        pd.DataFrame(all_data["basic"]).to_excel(
            writer, sheet_name="اطلاعات پایه", index=False
        )

        sheet_map = {
            "board": "اعضای هیئت مدیره",
            "risks": "ریسک ها",
            "subsidiaries": "شرکت های مادر و زیرمجموعه",
            "inspectors": "بازرسین",
            "licenses": "مجوزها",
            "capital_changes": "تغییرات سرمایه",
            "announcements": "آگهی ها",
        }

        for key, sheet_name in sheet_map.items():

            rows = all_data[key]

            if rows:
                pd.DataFrame(rows).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
            else:
                pd.DataFrame(
                    columns=["شناسه ملی", "اسم شرکت"]
                ).to_excel(writer, sheet_name=sheet_name, index=False)


# ==============================
# برنامه اصلی
# ==============================

async def main():

    excel = select_excel()

    if not excel:
        print("فایلی انتخاب نشد")
        return

    df = pd.read_excel(excel)

    if "شناسه ملی" not in df.columns:
        print("ستون شناسه ملی وجود ندارد")
        return

    output = os.path.splitext(excel)[0] + OUTPUT_SUFFIX

    # داده‌های همه‌ی شیت‌ها را در حافظه نگه می‌داریم
    all_data = {
        "basic": [],
        "board": [],
        "risks": [],
        "subsidiaries": [],
        "inspectors": [],
        "licenses": [],
        "capital_changes": [],
        "announcements": [],
    }

    done_ids = set()

    # اگر فایل خروجی قبلی وجود دارد، برای ازسرگیری بارگذاری کن
    if os.path.exists(output):

        try:

            existing_basic = pd.read_excel(
                output, sheet_name="اطلاعات پایه"
            )

            all_data["basic"] = existing_basic.to_dict("records")

            done_ids = {
                str(x["شناسه ملی"])
                for x in all_data["basic"]
                if not x.get("خطا")
            }

            for key, sheet_name in {
                "board": "اعضای هیئت مدیره",
                "risks": "ریسک ها",
                "subsidiaries": "شرکت های مادر و زیرمجموعه",
                "inspectors": "بازرسین",
                "licenses": "مجوزها",
                "capital_changes": "تغییرات سرمایه",
                "announcements": "آگهی ها",
            }.items():

                try:
                    existing = pd.read_excel(output, sheet_name=sheet_name)
                    all_data[key] = existing.to_dict("records")
                except Exception:
                    pass

        except Exception:
            pass

    async with async_playwright() as p:

        browser = await p.chromium.launch_persistent_context(
            user_data_dir=SESSION_PATH,
            headless=False,
            executable_path=CHROME_PATH,
            viewport={"width": 1400, "height": 900},
        )

        page = await browser.new_page()

        await page.goto("https://rasmio.com")

        print()
        print("=================================")
        print("وارد حساب رسمیو شوید")
        print("بعد از ورود Enter بزنید")
        print("=================================")

        input()

        for _, row in tqdm(df.iterrows(), total=len(df)):

            nid = str(row["شناسه ملی"]).strip()

            if not nid or nid == "nan":
                continue

            if nid in done_ids:
                continue

            manual_values = {}

            for col in MANUAL_COLUMNS:

                if col in df.columns:
                    val = row.get(col, "")
                    manual_values[col] = "" if pd.isna(val) else val

            basic_row = {"شناسه ملی": nid, "خطا": ""}
            basic_row.update(manual_values)

            try:

                found, warning = await goto_company_page(page, nid)

                if not found:

                    basic_row["خطا"] = "شرکت پیدا نشد یا کلیک روی نتیجه ناموفق بود"

                else:

                    html = await page.content()

                    parsed = parse_company_html(html, nid)

                    basic_row.update(parsed["basic"])
                    basic_row["خطا"] = parsed["error"] or warning
                    basic_row.update(manual_values)  # اولویت با مقادیر دستی

                    company_name = parsed["basic"].get("اسم شرکت", "")

                    for key in [
                        "board",
                        "risks",
                        "subsidiaries",
                        "inspectors",
                        "licenses",
                        "capital_changes",
                        "announcements",
                    ]:

                        for item in parsed[key]:

                            item_row = {
                                "شناسه ملی": nid,
                                "اسم شرکت": company_name,
                            }

                            item_row.update(item)

                            all_data[key].append(item_row)

            except Exception as e:

                basic_row["خطا"] = str(e)

            all_data["basic"].append(basic_row)

            write_workbook(output, all_data)

            await page.wait_for_timeout(random.randint(1500, 3000))

        await browser.close()

    print()
    print("تمام شد")
    print(output)


if __name__ == "__main__":
    asyncio.run(main())