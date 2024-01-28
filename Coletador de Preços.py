from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os


def coletarPreco(item: str, countLinks: int):
    if item == '' or pd.isna(item):
        return
    try:
        navegador = webdriver.Chrome()
        navegador.get(item)

        xpath = '/html/body/app-root/app-produto-detalhe/div/div/div[1]/div/div/app-tag-preco/div/div[2]'
        xpathPromo = '/html/body/app-root/app-produto-detalhe/div/div/div[1]/div/div/app-tag-preco/div[2]/div[2]'

        try:
            elemento = WebDriverWait(navegador, 30).until(
                    EC.visibility_of_element_located((By.XPATH, xpath))
                )
            preco = navegador.find_element(by='xpath', value=xpath).text.split(' ')[1].replace(',', '.')
        except IndexError:
            elemento = WebDriverWait(navegador, 30).until(
                    EC.visibility_of_element_located((By.XPATH, xpathPromo))
                )
            preco = navegador.find_element(by='xpath', value=xpathPromo).text.split(' ')[1].replace(',', '.')
        
        time.sleep(2)
        
        global contador
        print(f"Analisando {contador}/{countLinks}...")
        contador +=1

        navegador.quit()
        print(preco)
        return float(preco)
    
    except:
        print(f"Houve um erro ao executar esse o item:\n{item}")
        return

def atualiarPlanilha(semana: int, arquivo: str, modo: int):
    df = pd.read_excel(arquivo)
    df.columns = df.iloc[0]
    df = df.iloc[1:]

    df = df.drop(axis=1, columns=['Cálculo '])
    df = df.drop(axis=1, columns=[df.columns[-1]])

    identificaColuna = {
        1: df.columns[3],
        2: df.columns[4],
        3: df.columns[5],
        4: df.columns[6],
    }

    if modo:
        df = df[df[identificaColuna[semana]].isna()]

    countLinks = df[df.columns[-1]].count()
    df[identificaColuna[semana]] = df[df.columns[-1]].apply(lambda x: coletarPreco(x, countLinks))
        
    while True:
        try:
            with pd.ExcelWriter(arquivo, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
                df.to_excel(writer, index=False, sheet_name='Resultado')
            resposta = int(input('Deseja realizar outra ação?\n0 - Sim\n1 - Não\n'))
            return resposta
        
        except PermissionError:
            print("Houve um erro ao gravar os dados na planilha! Verifique se o aqruio está aberto em seu computador e tente novamente.")
            input("Tecle qualquer tecla para continuar...")
            continue

        except FileNotFoundError:
            print("Houve um erro ao gravar os dados na planilha! Verifique se o aqruio está junto com o executável do programa.")
            input("Tecle qualquer tecla para continuar...")
            continue
        except:
            return 1
        

def main():
    while True:
        arquivos = {}
        num = 1
        texto = ''
        for arquivo in os.listdir(os.getcwd()):
            arquivos[num] = arquivo
            texto += f"{num} - {arquivo}\n"
            num += 1
        
        try:
            numeroArquivo = int(input(f"Selecione o arquivo desejado:\n{texto}"))
            semana = int(input('Qual a semana que deseja preencher?\n1 - Primeira\n2 - Segunda\n3 - Terceira\n4 - Quarta\n'))
            modo = int(input('Qual o modo de execução?\n0 - Em todas as linhas\n1 - Somente nas linhas faltantes\n'))
            
            resposta = atualiarPlanilha(semana, arquivos[numeroArquivo], modo)
            if resposta:
                print('Finalizando...')
                time.sleep(1)
                break
            else:
                continue
        
        except ValueError:
            print("Insira valores corretamente.")
            continue
        
contador = 1
main()