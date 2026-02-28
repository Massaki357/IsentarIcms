import pyautogui
import pandas as pd
import pyperclip
import time

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# =========================
# FUNÇÃO PARA ESPERAR IMAGEM
# =========================
def esperar_e_clicar(imagem1, confianca=0.8, timeout=40):
    inicio = time.time()

    caminho1 = "imagens/" + imagem1
    caminho2 = "imagens/n.png"
    caminho3 = "imagens/6.png"

    while True:
        try:
            local1 = pyautogui.locateCenterOnScreen(caminho1, confidence=confianca)
            if local1:
                pyautogui.click(local1)
                return "clicou"


        except:
            pass

        if time.time() - inicio > timeout:
            print("Nenhuma das opções apareceu.")
            return "timeout"

        time.sleep(1)

# =========================
# LER EXCEL
# =========================
df = pd.read_excel("Pasta2.xlsx")

# =========================
# LOOP PRINCIPAL
# =========================
for index, row in df.iterrows():
    nome_produto = str(row["nome"])
    print(f"Alterando imposto: {nome_produto}")

    # 1️⃣ Clicar na barra de pesquisa
    esperar_e_clicar("1.png")

    # 2️⃣ Limpar campo
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("delete")

    # 3️⃣ Colar nome do produto
    pyperclip.copy(nome_produto)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.press("enter")

    # 4️⃣ Esperar botão editar aparecer
    esperar_e_clicar("2.png")

    # Clicar em tributações
    esperar_e_clicar("3.png")

    # 5️⃣ Clicar no campo imposto
    esperar_e_clicar("4.png")

    # 6️⃣ Alterar valor do imposto
    esperar_e_clicar("5.png")

    # 7️⃣ Salvar
    esperar_e_clicar("6.png")

    time.sleep(2)

print("Processo finalizado.")
