from flask import Flask, request, render_template_string, send_file
import pandas as pd
from io import BytesIO
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

app = Flask(__name__)
scrape_log = []
excel_data = None

HTML_TEMPLATE = """
<!doctype html>
<title>ICAEW Firm Scraper</title>
<h1>🔍 ICAEW Firm Scraper</h1>
<p>This tool scrapes firm data from ICAEW and exports to Excel.</p>
<form method="post" action="/scrape">
  <button type="submit">Start Scraping</button>
</form>
{% if log %}
  <h2>Scraping Log:</h2>
  <pre>{{ log }}</pre>
{% endif %}
{% if download %}
  <a href="/download">📥 Download Excel</a>
{% endif %}
"""

def is_last_page(soup):
    nav = soup.select_one('ul.pagination')
    if nav:
        li_tags = nav.find_all('li')
        if li_tags:
            last_li = li_tags[-1]
            return 'current' in last_li.get('class', [])
    return True

def scrape_page_items(soup):
    ul = soup.select_one('.search-results')
    return ul.find_all('li') if ul else []

def get_dd_by_dt_text(soup, dt_text):
    dl = soup.find('dl', class_='title-list')
    if not dl:
        return None
    for dt in dl.find_all('dt'):
        if dt.get_text(strip=True) == dt_text:
            dd = dt.find_next_sibling('dd')
            return dd.get_text(strip=True) if dd else None
    return None

def scrape_all_pages():
    global excel_data
    scrape_log.clear()
    base_url = "https://find.icaew.com/search?searchType=firm&term=&location_freetext=e11+1jz&page={}"
    page = 1
    data = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1280, "height": 720},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
        )
        page_obj = context.new_page()
        detail_page = context.new_page()

        while True:
            scrape_log.append(f"Scraping page {page}...")
            url = base_url.format(page)
            page_obj.goto(url)
            page_obj.wait_for_selector('#results')
            soup = BeautifulSoup(page_obj.content(), 'html.parser')
            items = scrape_page_items(soup)

            for item in items:
                link = 'https://find.icaew.com' + item.find('a').get('href')
                try:
                    detail_page.goto(link)
                    detail_page.wait_for_selector('.title-list')
                    soup1 = BeautifulSoup(detail_page.content(), 'html.parser')
                    name = soup1.find('h1').get_text(strip=True)
                    address = get_dd_by_dt_text(soup1, "Address")
                    website = get_dd_by_dt_text(soup1, "Website")
                    email = get_dd_by_dt_text(soup1, "Email address")
                    data.append({"Name": name, "Address": address, "Website": website, "Email": email})
                    scrape_log.append(f"✔ Scraped firm: {name}")
                except Exception:
                    scrape_log.append(f"⚠ Error scraping firm at {link}, skipping...")
                    continue

            if is_last_page(soup):
                break
            page += 1

        browser.close()

    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Firms', index=False)
        worksheet = writer.sheets['Firms']
        for i, column in enumerate(df.columns, 1):
            max_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
            col_letter = get_column_letter(i)
            worksheet.column_dimensions[col_letter].width = max_length
        header_font = Font(bold=True)
        for cell in worksheet[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

    output.seek(0)
    excel_data = output

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_TEMPLATE, log=None, download=False)

@app.route("/scrape", methods=["POST"])
def start_scrape():
    scrape_all_pages()
    return render_template_string(HTML_TEMPLATE, log="\n".join(scrape_log), download=True)

@app.route("/download")
def download():
    return send_file(excel_data, download_name="firms.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8080)
