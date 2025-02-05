import os
import time
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, StringVar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from base64 import b64decode
from datetime import datetime
import sys

# Hardcoded Chrome Driver Path
CHROME_DRIVER_PATH = "chromedriver.exe"

# Global variables for UI elements
selected_file = None
log_enabled = False

def convert_notebook_to_pdf(notebook_path, progress_bar, status_label, pdf_name_input, log_enabled):
    """
    Converts a Jupyter Notebook (.ipynb) to a PDF using Chrome headless printing.

    :param notebook_path: Full path to the .ipynb file
    :param progress_bar: Tkinter progress bar widget
    :param status_label: Label to display current processing status
    :param pdf_name_input: User input for the desired PDF name
    :param log_enabled: Boolean indicating if logging is enabled
    """
    status_label.configure(text="Processing: Extracting file details...")
    progress_bar.set(0.1)

    # Get directory and filename
    notebook_dir = os.path.dirname(notebook_path)
    notebook_name = os.path.splitext(os.path.basename(notebook_path))[0]

    # Get custom filename or use default
    pdf_name = pdf_name_input.get().strip()
    if not pdf_name:
        pdf_name = notebook_name  # Keep original name if no custom name provided

    html_path = os.path.join(notebook_dir, f"{notebook_name}.html")
    pdf_path = os.path.join(notebook_dir, f"{pdf_name}.pdf")

    # Step 1: Convert Jupyter Notebook to HTML (execute it first)
    status_label.configure(text="Executing notebook and converting to HTML...")
    progress_bar.set(0.2)
    subprocess.run(["jupyter", "nbconvert", "--execute", "--to", "html", "--HTMLExporter.exclude_input_prompt=True", notebook_path], check=True)

    # Step 2: Configure Chrome options
    status_label.configure(text="Configuring browser settings...")
    progress_bar.set(0.3)
    options = Options()
    options.add_argument("--headless=new")  # Headless mode ensures no UI interaction
    options.add_argument("--disable-gpu")  # Prevent hardware acceleration issues
    options.add_argument("--start-maximized")  # Ensure full rendering
    options.add_argument("--disable-print-preview")  # Skip print preview
    options.add_argument("--disable-software-rasterizer")  # Fix rendering issues

    # Step 3: Initialize WebDriver
    status_label.configure(text="Starting browser...")
    progress_bar.set(0.4)
    service = Service(CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    # Step 4: Open the generated HTML file
    status_label.configure(text="Opening generated HTML file...")
    progress_bar.set(0.5)
    driver.get(f"file:///{html_path}")

    # Step 5: Inject CSS for scaling and margin correction
    status_label.configure(text="Applying document formatting...")
    driver.execute_script("""
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
    """)

    # Step 6: Wait for MathJax rendering
    status_label.configure(text="Waiting for LaTeX rendering...")
    time.sleep(5)
    progress_bar.set(0.7)

    # Step 7: Use Chrome DevTools Protocol to print PDF
    status_label.configure(text="Generating PDF...")
    pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {
        "displayHeaderFooter": False,
        "printBackground": True,
        "preferCSSPageSize": True
    })

    # Step 8: Save the PDF
    with open(pdf_path, "wb") as f:
        f.write(b64decode(pdf_data["data"]))

    # Step 9: Close browser
    driver.quit()
    status_label.configure(text="Finalizing...")
    progress_bar.set(0.9)

    # Step 10: Delete the original HTML file
    if os.path.exists(html_path):
        os.remove(html_path)

    progress_bar.set(1.0)
    status_label.configure(text=f"PDF saved as: {pdf_path}")

    # Step 11: Write to log file if enabled
    if log_enabled:
        log_file = os.path.join(notebook_dir, "conversion_log.txt")
        with open(log_file, "a") as log:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Converted: {notebook_path} → {pdf_path}\n")
        status_label.configure(text=f"PDF saved as: {pdf_path} (Logged)")

def open_file_dialog():
    """
    Opens a file dialog to select a Jupyter Notebook.
    """
    global selected_file
    file_path = filedialog.askopenfilename(title="Select Jupyter Notebook",
                                           filetypes=[("Jupyter Notebooks", "*.ipynb")],
                                           initialdir=os.path.expanduser("~/Downloads"))
    if file_path:
        selected_file = file_path
        file_label.configure(text=f"Selected: {os.path.basename(file_path)}")

def toggle_logging():
    """
    Toggles the logging feature on and off.
    """
    global log_enabled
    log_enabled = not log_enabled
    log_toggle_button.configure(text=f"Logging: {'ON' if log_enabled else 'OFF'}")

def start_conversion():
    """
    Starts the conversion process in a separate thread to prevent UI freezing.
    """
    global selected_file
    if selected_file:
        progress_bar.set(0.0)
        status_label.configure(text="Starting conversion...")
        threading.Thread(target=convert_notebook_to_pdf, args=(selected_file, progress_bar, status_label, pdf_name_input, log_enabled), daemon=True).start()
    else:
        file_label.configure(text="No file selected!")

# GUI Setup
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Jupyter to PDF Converter")

script_dir = os.path.dirname(sys.argv[0])  # Get the script's directory
icon_path = os.path.join(script_dir, "fysh.ico")
root.iconbitmap(icon_path)  # Set the icon

root.geometry("500x300")

label = ctk.CTkLabel(root, text="Select a Jupyter Notebook (.ipynb) to convert to PDF")
label.pack(pady=10)

file_button = ctk.CTkButton(root, text="Choose File", command=open_file_dialog, fg_color="darkgreen")
file_button.pack(pady=5)

file_label = ctk.CTkLabel(root, text="No file selected")
file_label.pack(pady=5)

pdf_name_input = ctk.CTkEntry(root, placeholder_text="Enter PDF name (optional)")
pdf_name_input.pack(pady=5, padx=20, fill="x")

log_toggle_button = ctk.CTkButton(root, text="Logging: OFF", command=toggle_logging, fg_color="darkgreen")
log_toggle_button.pack(pady=5)

start_button = ctk.CTkButton(root, text="Start Conversion", command=start_conversion, fg_color="darkgreen")
start_button.pack(pady=10)

progress_bar = ctk.CTkProgressBar(root)
progress_bar.pack(pady=5, fill="x", padx=20)
progress_bar.set(0.0)

status_label = ctk.CTkLabel(root, text="Awaiting user action...")
status_label.pack(pady=5)

root.mainloop()
