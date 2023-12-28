import openpyxl
import pyautogui
import time

# Carregar a planilha Excel
workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas = workbook['vendas']

# Iterar sobre as linhas da planilha
for linha in vendas.iter_rows(min_row=2):
    try:
        # CLIENTE
        print("Clicando em CLIENTE")
        pyautogui.click(944, 520, duration=0.5)
        print("Escrevendo nome do cliente:", linha[0].value)
        pyautogui.write(linha[0].value)

        # PRODUTO
        print("Clicando em PRODUTO")
        pyautogui.click(943, 541, duration=0.5)
        print("Escrevendo nome do produto:", linha[1].value)
        pyautogui.write(linha[1].value)

        # QUANTIDADECadeiraViolinoLmpadaLmpadaCaixa de SomHarpaGuarda-rCondicionadorMochilaoupa
        print("Clicando em QUANTIDADE")
        pyautogui.click(948, 571, duration=0.5)
        print("Escrevendo quantidade:", linha[2].value)
        pyautogui.write(str(linha[2].value))

        # CATEGORIA
        print("Clicando em CATEGORIA")
        pyautogui.click(1020, 599, duration=0.5)
        print("Escrevendo categoria:", linha[3].value)
        pyautogui.write(linha[3].value)

        # SALVAR
        print("Clicando em SALVAR")
        pyautogui.click(892, 632, duration=0.5)

        # OK
        print("Clicando em OK")
        pyautogui.click(949, 582, duration=0.5)

        # Adicione uma pausa para dar tempo ao sistema
        time.sleep(2)

    except Exception as e:
     print(f"Erro: {e}")

# Fechar a planilha Excel
workbook.close()

    
    
   

   

