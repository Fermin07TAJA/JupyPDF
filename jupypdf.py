import os
import sys
import time
import subprocess
from base64 import b64decode
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def convert_notebook_to_pdf(notebook_path, pdf_name=None, write_log=True):
    notebook_path = os.path.abspath(notebook_path)
    notebook_dir = os.path.dirname(notebook_path)
    notebook_name = os.path.splitext(os.path.basename(notebook_path))[0]

    # PDF name fallback
    if not pdf_name or not str(pdf_name).strip():
        pdf_name = notebook_name

    html_path = os.path.join(notebook_dir, notebook_name + ".html")
    pdf_path = os.path.join(notebook_dir, pdf_name + ".pdf")

    print("1/4: Exporting to HTML...")
    subprocess.run(
        [
            sys.executable,
            "-m",
            "jupyter",
            "nbconvert",
            "--execute",
            "--to",
            "html",
            "--HTMLExporter.exclude_input_prompt=True",
            notebook_path,
        ],
        check=True,
    )

    print("2/4: Chrome Conversion")
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-print-preview")
    options.add_argument("--disable-software-rasterizer")

    driver = webdriver.Chrome(options=options)

    try:
        print("3/4: Printing to PDF...")
        driver.get(f"file:///{html_path}")

        driver.execute_script(
            """
            const style = document.createElement('style');
            style.innerHTML = `
                body {
                    transform: scale(0.7647);
                    transform-origin: top left;
                    width: 100%;
                    margin-left: 11.76%;
                }
                header, footer, .header, .footer {
                    display: none !important;
                    visibility: hidden !important;
                    height: 0 !important;
                }
            `;
            document.head.appendChild(style);
            """
        )

        time.sleep(5)

        pdf_data = driver.execute_cdp_cmd(
            "Page.printToPDF",
            {
                "displayHeaderFooter": False,
                "printBackground": True,
                "preferCSSPageSize": True,
            },
        )

        with open(pdf_path, "wb") as f:
            f.write(b64decode(pdf_data["data"]))

    finally:
        driver.quit()

    if os.path.exists(html_path):
        os.remove(html_path)

    # LOG Log LOOOGGGG
    if write_log:
        tool_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(tool_dir, "conversion_log.txt")

        line = (
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
            f"Converted: {notebook_path} -> {pdf_path}"
        )
        with open(log_path, "w", encoding="utf-8", errors="ignore") as log:
            log.write(line + "\n")

    print("@Complete")
    return pdf_path


notebook_path = sys.argv[1]
pdf_name_arg = sys.argv[2] if len(sys.argv) > 2 else None
convert_notebook_to_pdf(notebook_path, pdf_name_arg)
