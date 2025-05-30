import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import time 

def scrape_scholar_articles(query, year_start, year_end, max_pages=5):
    articles = []
    for page in range(max_pages):
        url = f"https://scholar.google.com/scholar?start={page*10}&q={query}&hl=en&as_sdt=0,5&as_ylo={year_start}&as_yhi={year_end}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
        }

        response = requests.get(url, headers=headers)
        if "Our systems have detected unusual traffic" in response.text:
            print("🚫 Google Scholar mendeteksi scraping. Diblokir sementara.")
            break

        soup = BeautifulSoup(response.text, "html.parser")
        results = soup.find_all("div", class_="gs_ri")

        if not results:
            break

        for result in results:
            title_tag = result.find("h3", class_="gs_rt")
            if not title_tag:
                continue
            title = title_tag.text.strip()
            link = title_tag.find("a")
            link = link["href"] if link else "No link available"

            authors_info = result.find("div", class_="gs_a").text.strip()
            source_info = authors_info.split(" - ")[-1] if " - " in authors_info else "Unknown"

            year = "Unknown"
            for part in authors_info.split(" "):
                if part.isdigit() and len(part) == 4:
                    year = part
                    break

            citations = "0"
            cite_section = result.find("div", class_="gs_fl")
            if cite_section:
                for cite in cite_section.find_all("a"):
                    if "Cited by" in cite.text:
                        citations = cite.text.split(" ")[-1]
                        break

            pdf_link = "No PDF available"
            pdf_result = result.find_previous_sibling("div", class_="gs_or_ggsm")
            if pdf_result and pdf_result.find("a"):
                pdf_link = pdf_result.find("a")["href"]

            articles.append({
                "Title": title,
                "Authors & Source": authors_info,
                "Year": year,
                "Journal/Source": source_info,
                "Citations": citations,
                "Link": link,
                "PDF Link": pdf_link
            })

        time.sleep(3)  # Delay antar request untuk menghindari blokir

    return articles


def save_to_excel(articles, filename):
    df = pd.DataFrame(articles)
    df.to_excel(filename, index=False, engine="openpyxl")

def browse_folder():
    folder_path = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(tk.END, folder_path)

def scrape_articles():
    query_input = entry_query.get()
    queries = [q.strip() for q in query_input.split(",")]
    year_start = entry_year_start.get()
    year_end = entry_year_end.get()

    if not year_start.isdigit() or not year_end.isdigit():
        label_status.config(text="Please enter a valid year range.")
        return

    label_status.config(text="Scraping in progress... Please wait.")
    window.update()

    all_articles = []
    for query in queries:
        label_status.config(text=f"Scraping for keyword: {query}")
        window.update()
        articles = scrape_scholar_articles(query, int(year_start), int(year_end))
        for article in articles:
            article["Keyword"] = query
        all_articles.extend(articles)

    folder_path = entry_folder.get()
    filename = f"{folder_path}/scholar_articles_{year_start}_{year_end}.xlsx" if folder_path else f"scholar_articles_{year_start}_{year_end}.xlsx"

    save_to_excel(all_articles, filename)
    label_status.config(text=f"Extraction complete. Saved to {filename}")

# GUI Setup
window = tk.Tk()
window.title("Google Scholar Scraper")
window.geometry("450x370")

label_query = tk.Label(window, text="🔎 Article Titles or Keywords (comma separated):")
label_query.pack()
entry_query = tk.Entry(window, width=50)
entry_query.pack()

label_year_range = tk.Label(window, text="📅 Publication Year Range (e.g., 2018 - 2023):")
label_year_range.pack()

frame_years = tk.Frame(window)
frame_years.pack()

entry_year_start = tk.Entry(frame_years, width=10)
entry_year_start.pack(side=tk.LEFT)
tk.Label(frame_years, text="to").pack(side=tk.LEFT)
entry_year_end = tk.Entry(frame_years, width=10)
entry_year_end.pack(side=tk.LEFT)

label_folder = tk.Label(window, text="💾 Output Folder (optional):")
label_folder.pack()
entry_folder = tk.Entry(window, width=50)
entry_folder.pack()
button_browse = tk.Button(window, text="📁 Browse", command=browse_folder)
button_browse.pack()

button_extract = tk.Button(window, text="🚀 Extract Data", command=scrape_articles)
button_extract.pack(pady=10)

label_status = tk.Label(window, text="")
label_status.pack()

window.mainloop()
