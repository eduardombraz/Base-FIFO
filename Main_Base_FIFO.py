import asyncio
from playwright.async_api import async_playwright
import time
import datetime
import os
import shutil
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import zipfile
from gspread_dataframe import set_with_dataframe  # Import para upload de DataFrame

DOWNLOAD_DIR = "/tmp/shopee_automation"

def rename_downloaded_file(download_dir, download_path):
    """Renames the downloaded file to include the current hour."""
    try:
        current_hour = datetime.datetime.now().strftime("%H")
        new_file_name = f"TO-Packed{current_hour}.zip"
        new_file_path = os.path.join(download_dir, new_file_name)
        if os.path.exists(new_file_path):
            os.remove(new_file_path)
        shutil.move(download_path, new_file_path)
        print(f"Arquivo salvo como: {new_file_path}")
        return new_file_path
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return None

def unzip_and_process_data(zip_path, extract_to_dir):
    """
    Unzips a file, merges all CSVs, and processes the data according to the specified logic.
    """
    try:
        unzip_folder = os.path.join(extract_to_dir, "extracted_files")
        os.makedirs(unzip_folder, exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(unzip_folder)
        print(f"Arquivo '{os.path.basename(zip_path)}' descompactado.")

        csv_files = [os.path.join(unzip_folder, f) for f in os.listdir(unzip_folder) if f.lower().endswith('.csv')]
        
        if not csv_files:
            print("Nenhum arquivo CSV encontrado no ZIP.")
            shutil.rmtree(unzip_folder)
            return None

        print(f"Lendo e unificando {len(csv_files)} arquivos CSV...")
        all_dfs = [pd.read_csv(file, encoding='utf-8') for file in csv_files]
        df_final = pd.concat(all_dfs, ignore_index=True)

        print("Iniciando processamento dos dados...")

        # Filtrar colunas desejadas pela posição
        indices_para_manter = [0, 14, 39, 40, 48]
        df_final = df_final.iloc[:, indices_para_manter]

        print("Processamento de dados concluído com sucesso.")
        
        shutil.rmtree(unzip_folder)  # Limpa pasta temporária

        return df_final
    except Exception as e:
        print(f"Erro ao descompactar ou processar os dados: {e}")
        return None

def update_google_sheet_with_dataframe(df_to_upload):
    """Updates a Google Sheet with the content of a pandas DataFrame."""
    if df_to_upload is None or df_to_upload.empty:
        print("Nenhum dado para enviar ao Google Sheets.")
        return
        
    try:
        print("Enviando dados processados para o Google Sheets...")

        # Limpar valores incompatíveis
        df_to_upload = df_to_upload.fillna("").astype(str)

        scope = [
            "https://spreadsheets.google.com/feeds",
            'https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
        client = gspread.authorize(creds)
        
        planilha = client.open("FIFO INBOUND SP5")

        # Verifica se a aba existe, senão cria
        try:
            aba = planilha.worksheet("Base")
        except gspread.exceptions.WorksheetNotFound:
            aba = planilha.add_worksheet(title="Base", rows="1000", cols="20")
        
        aba.clear()
        set_with_dataframe(aba, df_to_upload)
        
        print("✅ Dados enviados para o Google Sheets com sucesso!")
        time.sleep(5)

    except Exception as e:
        import traceback
        print(f"❌ Erro ao enviar para o Google Sheets:\n{traceback.format_exc()}")

async def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage", "--window-size=1920,1080"])
        context = await browser.new_context(accept_downloads=True, viewport={"width": 1920, "height": 1080})
        page = await context.new_page()
        try:
            # LOGIN
            await page.goto("https://spx.shopee.com.br/")
            await page.wait_for_selector('xpath=//*[@placeholder="Ops ID"]', timeout=15000)
            await page.locator('xpath=//*[@placeholder="Ops ID"]').fill('Ops35673')
            await page.locator('xpath=//*[@placeholder="Senha"]').fill('@Porpeta2025')
            await page.locator('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button').click()
            await page.wait_for_timeout(15000)
            try:
                await page.locator('.ssc-dialog-close').click(timeout=5000)
            except:
                print("Nenhum pop-up de diálogo foi encontrado.")
                await page.keyboard.press("Escape")
            
            # NAVEGAÇÃO E DOWNLOAD
            await page.goto("https://spx.shopee.com.br/#/orderTracking")
            await page.wait_for_timeout(8000)
            await page.get_by_role('button', name='Exportar').click()
            await page.wait_for_timeout(8000)
            await page.locator('xpath=/html[1]/body[1]/span[6]/div[1]/div[1]/div[1]').click()
            await page.wait_for_timeout(8000)
            await page.get_by_role("treeitem", name="SOC_Received", exact=True).click()
            await page.wait_for_timeout(8000)
            await page.locator(".ssc-dialog-body > .ssc-form > div:nth-child(9) > .ssc-form-item-content").click()
            await page.wait_for_timeout(8000)
            await page.get_by_role('textbox', name='procurar por').fill('SoC_SP_Cravinhos')
            await page.wait_for_timeout(8000)
            await page.get_by_role('listitem', name='SoC_SP_Cravinhos').click()
            await page.wait_for_timeout(8000)
            await page.get_by_role("button", name="Confirmar").click()
            await page.wait_for_timeout(480000)
            
            # DOWNLOAD
            async with page.expect_download() as download_info:
            await page.goto("https://spx.shopee.com.br/#/orderTracking")
            await page.wait_for_timeout(8000)
            await page.locator('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[8]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/table[1]/tbody[2]/tr[1]/td[7]/div[1]/div[1]/button[1]/span[1]/span[1]').first.click()
            #await page.get_by_role("button", name="Baixar").first.click()
            
            download = await download_info.value
            download_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
            await download.save_as(download_path)
            print(f"Download concluído: {download_path}")


            # --- PROCESSA E ENVIA PARA GOOGLE SHEETS ---
            renamed_zip_path = rename_downloaded_file(DOWNLOAD_DIR, download_path)
            
            if renamed_zip_path:
                final_dataframe = unzip_and_process_data(renamed_zip_path, DOWNLOAD_DIR)
                update_google_sheet_with_dataframe(final_dataframe)

        except Exception as e:
            print(f"Erro durante o processo principal: {e}")
        finally:
            await browser.close()
            if os.path.exists(DOWNLOAD_DIR):
                shutil.rmtree(DOWNLOAD_DIR)
                print(f"Diretório de trabalho '{DOWNLOAD_DIR}' limpo.")

if __name__ == "__main__":
    asyncio.run(main())
