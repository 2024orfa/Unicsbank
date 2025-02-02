import os
import asyncio
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright
import markdownify
import requests
import fitz  # PyMuPDF
import docx
import openpyxl
from pptx import Presentation
import markdown

# Global list to store all extracted Markdown content
all_markdown_content = []

async def fetch_page_content(url):
    """Fetches the HTML content of a page using Playwright's async API."""
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(url)
        await asyncio.sleep(2)  # Wait for JavaScript to load content
        content = await page.content()
        await browser.close()
    return content

def extract_links(base_url, html_content):
    """Extracts and returns a set of internal links and file links from the HTML content."""
    soup = BeautifulSoup(html_content, 'html.parser')
    links = set()
    file_links = set()
    for a_tag in soup.find_all('a', href=True):
        href = a_tag['href']
        full_url = urljoin(base_url, href)
        if urlparse(full_url).netloc == urlparse(base_url).netloc:
            if full_url.lower().endswith(('.pdf', '.docx', '.xlsx', '.pptx')):
                file_links.add(full_url)
            else:
                links.add(full_url)
    return links, file_links

def download_file(url, output_dir):
    """Downloads a file from a URL to the specified output directory."""
    response = requests.get(url)
    if response.status_code == 200:
        os.makedirs(output_dir, exist_ok=True)
        filename = os.path.join(output_dir, os.path.basename(urlparse(url).path))
        with open(filename, 'wb') as file:
            file.write(response.content)
        return filename
    return None

def extract_text_from_pdf(file_path):
    """Extracts text from a PDF file."""
    text = ""
    with fitz.open(file_path) as doc:
        for page in doc:
            text += page.get_text()
    return text

def extract_text_from_docx(file_path):
    """Extracts text from a Word (.docx) file."""
    doc = docx.Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_xlsx(file_path):
    """Extracts text from an Excel (.xlsx) file."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    text = ""
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            text += "\t".join([str(cell) if cell is not None else "" for cell in row]) + "\n"
    return text

def extract_text_from_pptx(file_path):
    """Extracts text from a PowerPoint (.pptx) file."""
    prs = Presentation(file_path)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

def save_text_as_markdown(text, filename, output_dir):
    """Saves extracted text as a Markdown file and appends it to all_markdown_content."""
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(filename))[0]
    md_filename = f"{base_name}.md"
    filepath = os.path.join(output_dir, md_filename)
    with open(filepath, 'w', encoding='utf-8') as file:
        file.write(text)
    # Append the content to all_markdown_content
    all_markdown_content.append(text)

def convert_to_markdown(html_content):
    """Converts HTML content to Markdown format."""
    return markdownify.markdownify(html_content, heading_style="ATX")

def save_markdown(content, url, output_dir):
    """Saves the Markdown content to a file named after the URL path and appends it to all_markdown_content."""
    parsed_url = urlparse(url)
    path = parsed_url.path.strip('/')
    if not path:
        path = 'index'
    filename = f"{path.replace('/', '_')}.md"
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, filename)
    with open(filepath, 'w', encoding='utf-8') as file:
        file.write(content)
    # Append the content to all_markdown_content
    all_markdown_content.append(content)

def markdown_to_word(markdown_content, output_filename):
    """Converts combined Markdown content to Word (DOCX)."""
    # Convert markdown to HTML
    html_content = markdown.markdown(markdown_content)
    
    # Create a new Word document
    doc = docx.Document()
    
    # Add HTML content as text to the Word document (simple approach)
    doc.add_paragraph(html_content)
    
    # Save the document as a .docx file
    doc.save(output_filename)

async def scrape_site(url, depth, visited=None, output_dir='scraped_content'):
    """Recursively scrapes a website up to a specified depth."""
    if visited is None:
        visited = set()
    if depth == 0 or url in visited:
        return
    visited.add(url)
    print(f"Scraping: {url}")
    html_content = await fetch_page_content(url)
    markdown_content = convert_to_markdown(html_content)
    save_markdown(markdown_content, url, output_dir)
    links, file_links = extract_links(url, html_content)
    
    # Process file links
    for file_url in file_links:
        print(f"Processing file: {file_url}")
        file_path = download_file(file_url, output_dir)
        if file_path:
            if file_path.lower().endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
            elif file_path.lower().endswith('.docx'):
                text = extract_text_from_docx(file_path)
            elif file_path.lower().endswith('.xlsx'):
                text = extract_text_from_xlsx(file_path)
            elif file_path.lower().endswith('.pptx'):
                text = extract_text_from_pptx(file_path)
            else:
                continue
            save_text_as_markdown(text, file_path, output_dir)
    
    for link in links:
        await scrape_site(link, depth - 1, visited, output_dir)

if __name__ == "__main__":
    start_url = "https://unicsgroup.com"  # Replace with your target URL
    max_depth = 2  # Set the desired depth
    output_dir = "scraped_content"

    # Run the scraper
    asyncio.run(scrape_site(start_url, max_depth, output_dir=output_dir))

    # Save combined Markdown content after scraping
    combined_markdown_content = "\n\n".join(all_markdown_content)
    markdown_filename = os.path.join(output_dir, "combined_output.md")
    with open(markdown_filename, 'w', encoding='utf-8') as file:
        file.write(combined_markdown_content)
    print(f"All scraped content saved in {markdown_filename}")

    # Convert the combined Markdown content to Word
    word_filename = os.path.join(output_dir, "combined_output.docx")
    markdown_to_word(combined_markdown_content, word_filename)
    print(f"All scraped content saved in {word_filename}")
# import os
# import asyncio
# from urllib.parse import urljoin, urlparse
# from bs4 import BeautifulSoup
# from playwright.async_api import async_playwright
# import markdownify
# import requests
# import fitz  # PyMuPDF
# import docx
# import openpyxl
# from pptx import Presentation
# import markdown

# # Global variable to store combined Markdown content
# combined_markdown_content = ""

# async def fetch_page_content(url):
#     """Fetches the HTML content of a page using Playwright's async API."""
#     async with async_playwright() as p:
#         browser = await p.chromium.launch()
#         page = await browser.new_page()
#         await page.goto(url)
#         await asyncio.sleep(2)  # Wait for JavaScript to load content
#         content = await page.content()
#         await browser.close()
#     return content

# def extract_links(base_url, html_content):
#     """Extracts and returns a set of internal links and file links from the HTML content."""
#     soup = BeautifulSoup(html_content, 'html.parser')
#     links = set()
#     file_links = set()
#     for a_tag in soup.find_all('a', href=True):
#         href = a_tag['href']
#         full_url = urljoin(base_url, href)
#         if urlparse(full_url).netloc == urlparse(base_url).netloc:
#             if full_url.lower().endswith(('.pdf', '.docx', '.xlsx', '.pptx')):
#                 file_links.add(full_url)
#             else:
#                 links.add(full_url)
#     return links, file_links

# def download_file(url, output_dir):
#     """Downloads a file from a URL to the specified output directory."""
#     response = requests.get(url)
#     if response.status_code == 200:
#         os.makedirs(output_dir, exist_ok=True)
#         filename = os.path.join(output_dir, os.path.basename(urlparse(url).path))
#         with open(filename, 'wb') as file:
#             file.write(response.content)
#         return filename
#     return None

# def extract_text_from_pdf(file_path):
#     """Extracts text from a PDF file."""
#     text = ""
#     with fitz.open(file_path) as doc:
#         for page in doc:
#             text += page.get_text()
#     return text

# def extract_text_from_docx(file_path):
#     """Extracts text from a Word (.docx) file."""
#     doc = docx.Document(file_path)
#     return "\n".join([para.text for para in doc.paragraphs])

# def extract_text_from_xlsx(file_path):
#     """Extracts text from an Excel (.xlsx) file."""
#     wb = openpyxl.load_workbook(file_path, data_only=True)
#     text = ""
#     for sheet in wb.worksheets:
#         for row in sheet.iter_rows(values_only=True):
#             text += "\t".join([str(cell) if cell is not None else "" for cell in row]) + "\n"
#     return text

# def extract_text_from_pptx(file_path):
#     """Extracts text from a PowerPoint (.pptx) file."""
#     prs = Presentation(file_path)
#     text = ""
#     for slide in prs.slides:
#         for shape in slide.shapes:
#             if hasattr(shape, "text"):
#                 text += shape.text + "\n"
#     return text

# def save_text_as_markdown(text, filename, output_dir):
#     """Saves extracted text as a Markdown file."""
#     os.makedirs(output_dir, exist_ok=True)
#     base_name = os.path.splitext(os.path.basename(filename))[0]
#     md_filename = f"{base_name}.md"
#     filepath = os.path.join(output_dir, md_filename)
#     with open(filepath, 'w', encoding='utf-8') as file:
#         file.write(text)

# def convert_to_markdown(html_content):
#     """Converts HTML content to Markdown format."""
#     return markdownify.markdownify(html_content, heading_style="ATX")

# def save_markdown(content, url, output_dir):
#     """Saves the Markdown content to a file named after the URL path."""
#     parsed_url = urlparse(url)
#     path = parsed_url.path.strip('/')
#     if not path:
#         path = 'index'
#     filename = f"{path.replace('/', '_')}.md"
#     os.makedirs(output_dir, exist_ok=True)
#     filepath = os.path.join(output_dir, filename)
#     with open(filepath, 'w', encoding='utf-8') as file:
#         file.write(content)
    
#     # Append the content to the combined Markdown content
#     global combined_markdown_content
#     combined_markdown_content += content + "\n\n"  # Add extra newline for separation

# def markdown_to_word(markdown_content, output_filename):
#     """Converts combined Markdown content to Word (DOCX)."""
#     # Convert markdown to HTML
#     html_content = markdown.markdown(markdown_content)
    
#     # Create a new Word document
#     doc = docx.Document()
    
#     # Add HTML content as text to the Word document (simple approach)
#     doc.add_paragraph(html_content)
    
#     # Save the document as a .docx file
#     doc.save(output_filename)

# async def scrape_site(url, depth, visited=None, output_dir='scraped_content'):
#     """Recursively scrapes a website up to a specified depth."""
#     if visited is None:
#         visited = set()
#     if depth == 0 or url in visited:
#         return
#     visited.add(url)
#     print(f"Scraping: {url}")
#     html_content = await fetch_page_content(url)
#     markdown_content = convert_to_markdown(html_content)
#     save_markdown(markdown_content, url, output_dir)
#     links, file_links = extract_links(url, html_content)
    
#     # Process file links
#     for file_url in file_links:
#         print(f"Processing file: {file_url}")
#         file_path = download_file(file_url, output_dir)
#         if file_path:
#             if file_path.lower().endswith('.pdf'):
#                 text = extract_text_from_pdf(file_path)
#             elif file_path.lower().endswith('.docx'):
#                 text = extract_text_from_docx(file_path)
#             elif file_path.lower().endswith('.xlsx'):
#                 text = extract_text_from_xlsx(file_path)
#             elif file_path.lower().endswith('.pptx'):
#                 text = extract_text_from_pptx(file_path)
#             else:
#                 continue
#             save_text_as_markdown(text, file_path, output_dir)
    
#     for link in links:
#         await scrape_site(link, depth - 1, visited, output_dir)

# if __name__ == "__main__":
#     start_url = "https://unicsgroup.com"  # Replace with your target URL
#     max_depth = 2  # Set the desired depth
#     output_dir = "scraped_content"

#     # Run the scraper
#     asyncio.run(scrape_site(start_url, max_depth, output_dir=output_dir))

#     # Save combined Markdown content after scraping
#     markdown_filename = os.path.join(output_dir, "combined_output.md")
#     with open(markdown_filename, 'w', encoding='utf-8') as file:
#         file.write(combined_markdown_content)
#     print(f"All scraped content saved in {markdown_filename}")

#     # Convert the combined Markdown content to Word
#     word_filename = os.path.join(output_dir, "combined_output.docx")
#     markdown_to_word(combined_markdown_content, word_filename)
#     print(f"All scraped content saved in {word_filename}")