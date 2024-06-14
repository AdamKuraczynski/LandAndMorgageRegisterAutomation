import pandas as pd
from playwright.sync_api import sync_playwright

def run(playwright, register_numbers):
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    results = []

    for register_number in register_numbers:
        parts = register_number.split('/')
        part1 = parts[0]
        part2 = parts[1]
        part3 = parts[2]

        page.goto("https://przegladarka-ekw.ms.gov.pl/eukw_prz/KsiegiWieczyste/wyszukiwanieKW?komunikaty=true&kontakt=true&okienkoSerwisowe=false")
        page.get_by_role("textbox", name="Lista rozwijana wybierz kod").click()
        page.get_by_role("textbox", name="Lista rozwijana wybierz kod").fill(part1)
        page.get_by_label("Wpisz numer księgi wieczystej").click()
        page.get_by_label("Wpisz numer księgi wieczystej").fill(part2)
        page.get_by_label("Wpisz cyfrę kontrolną").click()
        page.get_by_label("Wpisz cyfrę kontrolną").fill(part3)
        page.get_by_role("button", name="Wyszukaj Księgę").click()
        
        try:
            page.get_by_role("button", name="Przeglądanie zupełnej treści").click()
            page.get_by_role("button", name="Dział IV").click()

            page_content = page.content()
            
            if "Wpisujący" in page_content and "Chwila wpisu" in page_content:
                results.append((register_number, "YES"))
            else:
                results.append((register_number, "NO"))
                
        except Exception as e:
            results.append((register_number, "ERROR"))

    context.close()
    browser.close()

    return results

input_file_path = 'input_file.xlsx'
df = pd.read_excel(input_file_path)

column_name = 'RegisterNumber'
register_numbers = df[column_name].tolist()

with sync_playwright() as playwright:
    results = run(playwright, register_numbers)

results_df = pd.DataFrame(results, columns=["RegisterNumber", "Status"])
output_file_path = 'output_file.xlsx'
results_df.to_excel(output_file_path, index=False)

print(f"Results saved to {output_file_path}")
