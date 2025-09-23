async def main():
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
async with async_playwright() as p:
browser = await p.chromium.launch(
headless=True,
args=["--no-sandbox", "--disable-dev-shm-usage", "--window-size=1920,1080"]
)
context = await browser.new_context(accept_downloads=True, viewport={"width": 1920, "height": 1080})
page = await context.new_page()
try:
# LOGIN
await page.goto("https://spx.shopee.com.br/", wait_until="domcontentloaded")
await page.wait_for_selector('xpath=//[@placeholder="Ops ID"]', timeout=60000)
await page.locator('xpath=//[@placeholder="Ops ID"]').fill('Ops35673')
await page.locator('xpath=//*[@placeholder="Senha"]').fill('@Porpeta2025')
await page.locator('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button').click()

        # fecha pop-up se houver  
        try:  
            await page.locator('.ssc-dialog-close').click(timeout=5000)  
        except:  
            try:  
                await page.keyboard.press("Escape")  
            except:  
                pass  

        # NAVEGAÇÃO E EXPORT  
        await page.goto("https://spx.shopee.com.br/#/orderTracking", wait_until="domcontentloaded")  
        # aguarda a página ficar ociosa (muitas vezes há várias requisições)  
        try:  
            await page.wait_for_load_state("networkidle", timeout=120000)  
        except:  
            pass  

        # abrir exportação  
        await page.get_by_role('button', name=re.compile('Exportar', re.I)).click()  

        # aguarda modal de export  
        export_modal = page.locator(".ssc-dialog").filter(has_text=re.compile("Exportar|Export", re.I)).last  
        await export_modal.wait_for(state="visible", timeout=120000)  

        # abrir seletor de tipo de relatório (caso necessário)  
        # clique no dropdown (ajuste se o seletor mudar)  
        await page.locator('xpath=/html[1]/body[1]/span[contains(@class,"ssc")]/div[1]/div[1]/div[1]').click()  

        # selecionar tipo SOC_Received  
        await page.get_by_role("treeitem", name="SOC_Received", exact=True).click()  

        # abrir seletor de SOC  
        await page.locator(".ssc-dialog-body > .ssc-form > div:nth-child(9) > .ssc-form-item-content").click()  

        # pesquisar e selecionar SoC_SP_Cravinhos  
        procurar_box = page.get_by_role('textbox', name=re.compile('procurar por', re.I))  
        await procurar_box.fill('SoC_SP_Cravinhos')  
        await page.get_by_role('listitem', name='SoC_SP_Cravinhos').click()  

        # confirmar  
        await page.get_by_role("button", name=re.compile("Confirmar", re.I)).click()  

        # esperar processamento do job/preview: spinners/redenetwork  
        # tente aguardar spinner sumir se existir  
        try:  
            await page.locator(".ssc-loading, .spinner").wait_for(state="hidden", timeout=180000)  
        except:  
            pass  
        try:  
            await page.wait_for_load_state("networkidle", timeout=180000)  
        except:  
            pass  

        # localizar botão Baixar no contexto do modal  
        baixar_btn = export_modal.get_by_role("button", name=re.compile("Baixar|Download|Exportar arquivo", re.I))  

        # esperar botão visível  
        await baixar_btn.wait_for(state="visible", timeout=180000)  

        # garantir que está habilitado  
        btn_handle = await baixar_btn.element_handle()  
        await page.wait_for_function(  
            "(btn) => !!btn && !btn.hasAttribute('disabled') && !btn.classList.contains('is-disabled')",  
            arg=btn_handle,  
            timeout=180000  
        )  

        # DEBUG: log de atributos  
        disabled_attr = await baixar_btn.get_attribute("disabled")  
        class_attr = await baixar_btn.get_attribute("class")  
        print(f"[DEBUG] Baixar attrs => disabled={disabled_attr}, class={class_attr}")  

        # DOWNLOAD  
        async with page.expect_download(timeout=180000) as download_info:  
            await baixar_btn.click()  

        download = await download_info.value  
        download_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)  
        await download.save_as(download_path)  
        print(f"Download concluído: {download_path}")  

        # PROCESSA E ENVIA PARA GOOGLE SHEETS  
        renamed_zip_path = rename_downloaded_file(DOWNLOAD_DIR, download_path)  

        if renamed_zip_path:  
            final_dataframe = unzip_and_process_data(renamed_zip_path, DOWNLOAD_DIR)  
            update_google_sheet_with_dataframe(final_dataframe)  

    except Exception as e:  
        # dump da página para debug se falhar no clique/fluxo  
        try:  
            html = await page.content()  
            with open("/tmp/page_dump.html", "w", encoding="utf-8") as f:  
                f.write(html)  
            print("Dump de HTML salvo em /tmp/page_dump.html")  
        except:  
            pass  
        print(f"Erro durante o processo principal: {e}")  
    finally:  
        await browser.close()  
        if os.path.exists(DOWNLOAD_DIR):  
            shutil.rmtree(DOWNLOAD_DIR)  
            print(f"Diretório de trabalho '{DOWNLOAD_DIR}' limpo.")  
