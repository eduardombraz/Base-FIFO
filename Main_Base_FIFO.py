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
from gspread_dataframe import set_with_dataframe  
import re  
  
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
        # Configuração otimizada para GitHub Actions  
        browser = await p.chromium.launch(  
            headless=True,  
            args=[  
                "--no-sandbox",  
                "--disable-dev-shm-usage",  
                "--disable-gpu",  
                "--single-process",  
                "--disable-setuid-sandbox"  
            ]  
        )  
        context = await browser.new_context(  
            accept_downloads=True,  
            viewport={"width": 1920, "height": 1080},  
            user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36"  
        )  
        page = await context.new_page()  
        try:  
            # ETAPA DE LOGIN - ABORDAGEM ROBUSTA  
            print("Iniciando processo de login...")  
            await page.goto("https://spx.shopee.com.br/", wait_until="domcontentloaded")  
              
            # Preenche credenciais com verificações  
            await page.wait_for_selector('input[placeholder="Ops ID"]', timeout=60000)  
            await page.fill('input[placeholder="Ops ID"]', 'Ops35673')  
            await page.fill('input[placeholder="Senha"]', '@Porpeta2025')  
              
            # Clique no botão de login com tratamento especial  
            login_button = page.locator('button:has-text("Entrar")')  
            await login_button.wait_for(state="visible", timeout=30000)  
            await login_button.click()  
              
            # Estratégia de espera aprimorada pós-login  
            print("Aguardando redirecionamento pós-login...")  
            try:  
                # Esperar por qualquer elemento que indique login bem-sucedido  
                await page.wait_for_selector('div.ssc-layout-header, text=Sair', timeout=90000)  
                print("✅ Elemento pós-login encontrado")  
            except Exception as e:  
                print(f"❌ Falha na verificação pós-login: {str(e)}")  
                # Tentar verificar pela URL  
                if "spx.shopee.com.br/#/" in page.url:  
                    print("⚠️ URL de dashboard detectada, continuando...")  
                else:  
                    raise  
              
            # Fecha pop-up se existir  
            try:  
                await page.locator('.ssc-dialog-close').click(timeout=10000)  
                print("Pop-up fechado")  
            except:  
                print("Nenhum pop-up encontrado")  
              
            # NAVEGAÇÃO PARA A PÁGINA DE EXPORTAÇÃO  
            print("Navegando para página de rastreamento...")  
            await page.goto("https://spx.shopee.com.br/#/orderTracking", wait_until="networkidle", timeout=120000)  
              
            # VERIFICAÇÃO FLEXÍVEL DA PÁGINA  
            print("Verificando carregamento da página...")  
            try:  
                # Tentar vários elementos possíveis  
                await page.wait_for_selector(  
                    'h1:has-text("Rastreamento de Pedidos"), '  
                    'button:has-text("Exportar"), '  
                    'div.ssc-layout-content',  
                    timeout=60000  
                )  
                print("✅ Elemento de confirmação encontrado")  
            except:  
                print("⚠️ Elementos primários não encontrados, tentando verificar por qualquer conteúdo...")  
                # Verificação de fallback  
                content_selector = "body"  
                await page.wait_for_selector(content_selector, timeout=30000)  
                if await page.inner_html(content_selector):  
                    print("✅ Conteúdo HTML detectado, continuando...")  
                else:  
                    raise Exception("Falha crítica: página não carregou conteúdo")  
              
            # Botão Exportar  
            export_button = page.locator('button:has-text("Exportar")')  
            await export_button.wait_for(state="visible", timeout=60000)  
            await export_button.click()  
            print("Clicou em Exportar")  
              
            # Seleção de filtro  
            await page.wait_for_selector('.ssc-dropdown', timeout=30000)  
            await page.click('.ssc-dropdown >> nth=0')  
            print("Abriu dropdown de filtros")  
              
            # Seleção de SOC_Received  
            await page.wait_for_selector('text=SOC_Received', timeout=30000)  
            await page.click('text=SOC_Received')  
            print("Selecionou SOC_Received")  
              
            # Seleção de warehouse  
            warehouse_input = page.locator('input[placeholder="procurar por"]')  
            await warehouse_input.wait_for(state="visible", timeout=30000)  
            await warehouse_input.fill('SoC_SP_Cravinhos')  
            print("Digitou nome do warehouse")  
              
            # Aguarda e clica na opção  
            warehouse_option = page.locator('text=SoC_SP_Cravinhos')  
            await warehouse_option.wait_for(state="visible", timeout=30000)  
            await warehouse_option.click()  
            print("Selecionou warehouse")  
              
            # Confirmação  
            confirm_button = page.locator('button:has-text("Confirmar")')  
            await confirm_button.wait_for(state="visible", timeout=30000)  
            await confirm_button.click()  
            print("Confirmou seleção")  
              
            # ESPERA DINÂMICA PARA PROCESSAMENTO  
            print("Aguardando processamento do relatório...")  
            start_time = time.time()  
            timeout = 600  # 10 minutos  
              
            while time.time() - start_time < timeout:  
                download_button = page.locator('button:has-text("Baixar")').first  
                if await download_button.is_visible() and await download_button.is_enabled():  
                    print("Botão Baixar habilitado!")  
                    break  
                await asyncio.sleep(10)  
                print("Aguardando...")  
            else:  
                raise TimeoutError("Tempo excedido aguardando botão Baixar")  
              
            # DOWNLOAD  
            async with page.expect_download() as download_info:  
                await download_button.click()  
              
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
            print(f"❌ Erro durante o processo principal: {str(e)}")  
            import traceback  
            print(traceback.format_exc())  
        finally:  
            await browser.close()  
            if os.path.exists(DOWNLOAD_DIR):  
                shutil.rmtree(DOWNLOAD_DIR)  
                print(f"Diretório de trabalho '{DOWNLOAD_DIR}' limpo.")  
  
if __name__ == "__main__":  
    asyncio.run(main()) 
