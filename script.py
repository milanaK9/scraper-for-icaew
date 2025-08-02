from flask import Flask, send_file, render_template_string
import io
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

app = Flask(__name__)

# Your scraper code wrapped as a function returning bytes
def scrape_all_pages():
    # ... your existing scraping code here but changed to return bytes instead of saving file
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
        page1 = context.new_page()

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

        while True:
            url = base_url.format(page)
            page_obj.goto(url)
            page_obj.wait_for_selector('#results')
            html = page_obj.content()
            soup = BeautifulSoup(html, 'html.parser')

            items = scrape_page_items(soup)

            for item in items:
                link = 'https://find.icaew.com' + item.find('a').get('href')
                try:
                    page1.goto(link)
                    page1.wait_for_selector('.title-list')
                    html = page1.content()
                    soup1 = BeautifulSoup(html, 'html.parser')
                    name = soup1.find('h1').get_text(strip=True)
                    address = get_dd_by_dt_text(soup1, "Address")
                    website = get_dd_by_dt_text(soup1, "Website")
                    email = get_dd_by_dt_text(soup1, "Email address")

                    data.append({
                        "Name": name,
                        "Address": address,
                        "Website": website,
                        "Email": email
                    })
                except Exception:
                    pass

            if is_last_page(soup):
                break

            page += 1

        browser.close()

    df = pd.DataFrame(data)

    # Save Excel to bytes buffer with formatting
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Firms', index=False)
        worksheet = writer.sheets['Firms']

        for i, column in enumerate(df.columns, start=1):
            max_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
            col_letter = get_column_letter(i)
            worksheet.column_dimensions[col_letter].width = max_length

        header_font = Font(bold=True)
        for cell in worksheet[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

    output.seek(0)
    return output

# Flask routes
@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>ICAEW Scraper</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body class="bg-light d-flex flex-column justify-content-center align-items-center" style="height:100vh;">
        <div class="card p-5 shadow-sm" style="min-width: 300px;">
            <h2 class="mb-4 text-center">ICAEW Scraper</h2>
            <form action="/download" method="post">
                <button type="submit" class="btn btn-primary btn-lg w-100">Run Scraper & Download Excel</button>
            </form>
        </div>
    </body>
    </html>
    ''')

@app.route('/download', methods=['POST'])
def download():
    excel_data = scrape_all_pages()
    return send_file(
        excel_data,
        download_name="icaew_firms.xlsx",
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
