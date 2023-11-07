

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # Altere para headless=True para execução sem interface gráfica
    context = browser.new_context()
    page = context.new_page()

    # Navegar para uma URL
    page.goto(os.getenv("URL"))

    # Definir o tamanho da viewport
    page.set_viewport_size({'width': 1920, 'height': 1080})
    #page.wait_for_selector('text="Que bom te ver por aqui!"', timeout=60000)