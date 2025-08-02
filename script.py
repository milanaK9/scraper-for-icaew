from flask import Flask, send_file, render_template_string, request, jsonify
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from io import BytesIO
import threading

app = Flask(__name__)

excel_data = None
scraping_in_progress = False
current_page = 0
scrape_log = []

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
    global excel_data, scraping_in_progress, current_page, scrape_log
    scrape_log = []  # clear log on start
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

        while True:
            current_page = page
            scrape_log.append(f"Scraping page {page}...")
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
                    scrape_log.append(f"Scraped firm: {name}")
                except Exception:
                    scrape_log.append("Error scraping a firm, skipping...")
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

        for i, column in enumerate(df.columns, start=1):
            max_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
            col_letter = get_column_letter(i)
            worksheet.column_dimensions[col_letter].width = max_length

        header_font = Font(bold=True)
        for cell in worksheet[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

    output.seek(0)
    excel_data = output.read()
    scraping_in_progress = False
    current_page = 0
    scrape_log.append("Scraping complete.")

@app.route('/')
def index():
    return render_template_string('''
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>ICAEW Scraper</title>
<style>
  body {
    font-family: Arial, sans-serif;
    background: #121212;
    color: #eee;
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    height: 100vh;
    margin: 0;
  }
  h1 { color: #4caf50; }
  button {
    background-color: #4caf50;
    color: #121212;
    border: none;
    padding: 15px 30px;
    font-size: 18px;
    border-radius: 8px;
    cursor: pointer;
    box-shadow: 0 4px 8px rgba(76, 175, 80, 0.4);
    transition: background-color 0.3s ease;
  }
  button:hover {
    background-color: #388e3c;
  }
  #status {
    margin-top: 20px;
    color: #aaa;
    font-size: 16px;
  }
  #log {
    max-height: 200px;
    overflow-y: auto;
    background: #222;
    padding: 10px;
    border-radius: 5px;
    margin-top: 10px;
    color: #4caf50;
    white-space: pre-wrap;
    font-family: monospace;
    width: 400px;
  }
</style>
</head>
<body>
<h1>ICAEW Scraper</h1>
<button id="startBtn">Start Scraping & Download Excel</button>
<div id="status"></div>
<pre id="log"></pre>

<script>
  const btn = document.getElementById('startBtn');
  const status = document.getElementById('status');
  const log = document.getElementById('log');

  btn.onclick = () => {
    btn.disabled = true;
    status.textContent = 'Starting scraping...';
    log.textContent = '';

    fetch('/start_scraping', { method: 'POST' })
      .then(response => response.json())
      .then(data => {
        if (data.status === 'started') {
          checkProgress();
        } else if (data.status === 'already_running') {
          status.textContent = 'Scraping already in progress. Please wait...';
          btn.disabled = true;
        } else {
          status.textContent = 'Error starting scraping.';
          btn.disabled = false;
        }
      })
      .catch(() => {
        status.textContent = 'Network error.';
        btn.disabled = false;
      });
  };

  function checkProgress() {
    const poll = setInterval(() => {
      fetch('/scraping_status')
        .then(res => res.json())
        .then(data => {
          if (!data.in_progress) {
            clearInterval(poll);
            status.textContent = 'Scraping complete! Download will start shortly.';
            log.textContent = data.log.join('\\n');
            downloadFile();
            btn.disabled = false;
          } else {
            status.textContent = 'Scraping page ' + data.current_page + '...';
            log.textContent = data.log.join('\\n');
            log.scrollTop = log.scrollHeight; // auto scroll to bottom
          }
        })
        .catch(() => {
          clearInterval(poll);
          status.textContent = 'Error checking status.';
          btn.disabled = false;
        });
    }, 2000);
  }

  function downloadFile() {
    const link = document.createElement('a');
    link.href = '/download_excel';
    link.download = 'icaew_firms.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    status.textContent = 'Download started!';
  }
</script>
</body>
</html>
''')

@app.route('/start_scraping', methods=['POST'])
def start_scraping():
    global scraping_in_progress
    if scraping_in_progress:
        return jsonify({"status": "already_running"})

    scraping_in_progress = True
    threading.Thread(target=scrape_all_pages).start()

    return jsonify({"status": "started"})

@app.route('/scraping_status')
def scraping_status():
    global scraping_in_progress, current_page, scrape_log
    last_logs = scrape_log[-20:]  # send last 20 lines
    return jsonify({
        "in_progress": scraping_in_progress,
        "current_page": current_page,
        "log": last_logs
    })

@app.route('/download_excel')
def download_excel():
    global excel_data
    if excel_data:
        return send_file(
            BytesIO(excel_data),
            download_name="icaew_firms.xlsx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        return "File not ready", 404

if __name__ == '__main__':
    app.run(debug=True)
