from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unicodedata
import pandas as pd  # Importar pandas para manipulação dos dados

# Função para remover acentos e normalizar o texto
def remove_accent(text):
    return ''.join(
        c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn'
    ).lower()

# Configurar o ChromeDriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

# Inicializar o driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Acessar o site
driver.get("https://cepea.esalq.usp.br/br/")

# Lista de culturas a serem procuradas
culturas = [
    "açúcar", "algodão", "arroz", "café", 
    "milho", 
    "soja", "trigo"
]

# Lista de frutas/vegetais que mapeiam para "hortifruti"
frutas_e_vegetais = [
    "banana", "cenoura", "mamao", "melancia", "uva", "batata", 
    "citros", "manga", "melao", "cebola", "folhosas", "maca", "tomate"
]

# Criar uma lista para armazenar todos os dados extraídos
dados_tabela = []

def processar_cultura(cultura):
    try:
        # Aguardar a div com id 'imagenet-wrap-categoria' estar visível
        imagenet_wrap_categoria = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "imagenet-wrap-categoria"))
        )
        print(f"Div com id 'imagenet-wrap-categoria' encontrada para a cultura {cultura}!")

        # Encontrar a div com id 'imagenet-categoria' dentro da div 'imagenet-wrap-categoria'
        imagenet_categoria = imagenet_wrap_categoria.find_element(By.ID, "imagenet-categoria")
        
        # Encontrar a div com class 'imagenet-col-max imagenet-ma' dentro da div 'imagenet-categoria'
        imagenet_col_max = imagenet_categoria.find_element(By.CLASS_NAME, "imagenet-col-max.imagenet-ma")
        
        # Encontrar a ul com class 'imagenet-seg-menu-indicador' dentro da div 'imagenet-col-max imagenet-ma'
        imagenet_seg_menu_indicador = imagenet_col_max.find_element(By.CLASS_NAME, "imagenet-seg-menu-indicador")

        # Normalizar a cultura para comparação
        cultura_normalizada = remove_accent(cultura)

        # Definir qual cultura procurar
        if cultura_normalizada in frutas_e_vegetais:
            # Caso a cultura seja da lista de frutas/vegetais, procuramos "hortifruti"
            list_items = imagenet_seg_menu_indicador.find_elements(By.TAG_NAME, "li")
            for item in list_items:
                item_text = remove_accent(item.text.strip())
                if "hortifruti" in item_text:  # Verifica se contém a palavra "hortifruti"
                    link = item.find_element(By.TAG_NAME, "a")
                    link.click()
                    print(f"Clicado no item 'hortifruti', pois a cultura '{cultura}' está na lista de frutas/vegetais.")
                    break
        else:
            # Caso a cultura não seja da lista de frutas/vegetais, procurar o item correspondente à cultura
            list_items = imagenet_seg_menu_indicador.find_elements(By.TAG_NAME, "li")
            for item in list_items:
                item_text = remove_accent(item.text.strip())
                if item_text == cultura_normalizada:
                    link = item.find_element(By.TAG_NAME, "a")
                    link.click()
                    print(f"Clicado no item: '{cultura}'")
                    break

        # Aguardar a div 'imagenet-links-after-table' e clicar em 'Mais valores'
        imagenet_links_after_table = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "imagenet-links-after-table.imagenet-col-2.imagenet-pa-l.imagenet-bb.imagenet-fl"))
        )
        print(f"Div com class 'imagenet-links-after-table' encontrada para a cultura {cultura}!")

        # Procurar e clicar no link "Mais valores"
        mais_valores = imagenet_links_after_table.find_element(By.XPATH, ".//a[contains(text(), 'Mais valores')]")
        mais_valores.click()
        print(f"Clicado em 'Mais valores' para a cultura {cultura}.")

        # Agora, procurar pela tabela com id 'imagenet-indicador1'
        table = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.ID, "imagenet-indicador1"))
        )

        # Procurar todos os 'tr' dentro do 'tbody' da tabela
        tbody = table.find_element(By.TAG_NAME, "tbody")
        trs = tbody.find_elements(By.TAG_NAME, "tr")
        print(f"Encontrados {len(trs)} tr(s) para a cultura {cultura}.")

        for i, tr in enumerate(trs, start=1):  # enumerate fornece o índice começando em 1
            # Para cada tr, pegar os td e imprimir as informações
            tds = tr.find_elements(By.TAG_NAME, "td")
            if len(tds) > 4:  # Garantir que há ao menos 5 td
                data_atualizacao = tds[0].text.strip()
                valor_rs = tds[1].text.strip()
                variacao_dia = tds[2].text.strip()
                variacao_mes = tds[3].text.strip()
                valor_dolar = tds[4].text.strip()

                # Armazenar os dados na lista 'dados_tabela'
                dados_tabela.append({
                    "Cultura": cultura,
                    "Data de Atualização": data_atualizacao,
                    "Valor em R$": valor_rs,
                    "Variação por dia": variacao_dia,
                    "Variação por mês": variacao_mes,
                    "Valor em Dólar": valor_dolar
                })

                print(f"ID {i}")
                print(f"Cultura: {cultura}")
                print(f"Data de Atualização: {data_atualizacao}")
                print(f"Valor em R$: {valor_rs}")
                print(f"Variação por dia: {variacao_dia}")
                print(f"Variação por mês: {variacao_mes}")
                print(f"Valor em Dólar: {valor_dolar}")
                print("-" * 50)

    except Exception as e:
        print(f"Erro ao localizar os elementos para a cultura '{cultura}': {e}")

# Rodar o processo para cada cultura
for cultura in culturas:
    processar_cultura(cultura)

# Fechar o driver
driver.quit()

# Criar um DataFrame com os dados coletados
df = pd.DataFrame(dados_tabela)

# Salvar o DataFrame em um arquivo Excel
df.to_excel("dados_culturas.xlsx", index=False, engine='openpyxl')

print("Dados salvos em 'dados_culturas.xlsx'")
