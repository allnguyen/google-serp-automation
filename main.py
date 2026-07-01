# Google SERP Automation 
import random
import time
import pandas as pd
import urllib.parse
import asyncio
from crawl4ai import AsyncWebCrawler

# 1. Paths
input_path = r"C:\Users\allan\Projects\automation\google_search_counter\Results\Queries Test.xlsx"
output_path = r"C:\Users\allan\Projects\automation\google_search_counter\Queries Test RESULTS.xlsx"

# Async function to check for sponsored content
async def check_sponsored(query):
    async with AsyncWebCrawler() as crawler:
        url = f"https://www.google.com/search?q={urllib.parse.quote(str(query))}"
        result = await crawler.arun(url=url)
        if hasattr(result, 'markdown') and result.markdown:
            return result.markdown.lower().count('sponsored')
    return 0

# Async function to process queries in a DataFrame
async def process_sheet(df, target_col):
    sponsored_counts = []
    sponsored_presence = []
    for query in df[target_col]:
        sponsored_count = await check_sponsored(query)
        sponsored_counts.append(sponsored_count)
        sponsored_presence.append("TRUE" if sponsored_count > 0 else "FALSE")
        await asyncio.sleep(random.uniform(1, 3)) # Random human-like delay
    return sponsored_counts, sponsored_presence

async def main():
    print("Loading Excel sheets...")
    all_sheets = pd.read_excel(input_path, sheet_name=None)
    output_sheets = {}

    for sheet_name, df in all_sheets.items():
        print(f"Checking sheet: {sheet_name}")
        
        # Case-insensitive header match
        query_col = next((col for col in df.columns if str(col).strip().lower() == "queries"), None)
        
        if query_col is None:
            print(f"⚠️ Skipping '{sheet_name}' (No 'Queries' column found.)")
            continue

        print(f"🚀 Processing '{sheet_name}'...")
        
        # FIXED: Create standard clean URL strings
        df["Search URL String"] = df[query_col].apply(lambda q: f"https://google.com/search?q={urllib.parse.quote(str(q))}")
        
        # Run scraper loop safely
        sponsored_counts, sponsored_presence = await process_sheet(df, query_col)
        df["Sponsored Count"] = sponsored_counts
        df["Sponsored Presence"] = sponsored_presence
        
        # FIXED: Use Excel's HYPERLINK formula instead of forcing openpyxl modifications
        # This yields a safe file, using "GG" as display text, native blue link formatting included.
        df["Search URL"] = df["Search URL String"].apply(lambda url: f'=HYPERLINK("{url}", "GG")')
        
        # Drop temporary unformatted string column
        df = df.drop(columns=["Search URL String"])
        
        output_sheets[sheet_name] = df

    if not output_sheets:
        print("❌ Error: No sheets were processed.")
        return

    # FIXED: Single-pass file save. No re-opening files.
    print("Saving clean, uncorrupted Excel file...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in output_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✨ Success! Safe, uncorrupted Excel saved to: {output_path}")

if __name__ == "__main__":
    asyncio.run(main())
