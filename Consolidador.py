# ========================================
# CONSOLIDADOR DE BOLETAS - FUNDOS
# ========================================

import pandas as pd
import os
import requests
import win32com.client
from datetime import datetime, timedelta
import time
import tkinter as tk
from tkinter import messagebox
import unicodedata


# ========================================
# CONFIGURAÇÕES INICIAIS
# ========================================

assuntos = [
    "BOLETA DE MOVIMENTACAO FUNDOS", 
    "Aplic",
    "Aplicação Fundos", 
    "TEDs recebidas",
    "Resgate",  # Vai pegar "Resgate", "ENC: Resgate", "Resgate Fundos", etc.
    "LIQUIDAÇÃO"
]
Saida = r"C:\temp\boletas"
URL_API = [
    "http://nfappr003l:214/aplicacoes_pendentes",
    "http://nfappr003l:214/resgates_pendentes"
]

# Cria a pasta de saída se não existir
os.makedirs(Saida, exist_ok=True)


# ========================================
# FUNÇÕES AUXILIARES
# ========================================

def normalizar_texto(texto):
    """
    Remove acentos e caracteres especiais do texto.
    Exemplo: 'Aplicação' -> 'aplicacao'
    """
    if not texto:
        return ""
    
    # Remove acentos
    texto_normalizado = unicodedata.normalize('NFD', texto)
    texto_sem_acentos = ''.join(
        char for char in texto_normalizado 
        if unicodedata.category(char) != 'Mn'
    )
    
    return texto_sem_acentos.lower()


def conectar_outlook():
    """Conecta ao Outlook e retorna a caixa de entrada"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Caixa de entrada
        print("✓ Conectado ao Outlook com sucesso!")
        return inbox
    except Exception as e:
        raise Exception("Erro ao conectar ao Outlook. Verifique se o Outlook está aberto.") from e


def carregar_dados_api(url):
    """Carrega dados de uma API específica"""
    try:
        resposta = requests.get(url)
        resposta.raise_for_status()
        dados = resposta.json()

        if not dados:
            print(f"⚠️ API {url} retornou vazia.")
            return pd.DataFrame()

        df_api = pd.DataFrame(dados)
        df_api = df_api.rename(columns={
            "RECNUM": "ID",
            "NOME_FUNDO": "FUNDO",
            "SINACOR": "CODIGO",
            "USUARIONOME": "NOME",
            "NM_ASSESSOR": "NOME.1",
            "CNPJ_FUNDO": "CNPJ",
            "RESGATE_TOTAL": "RESGATE"
        })
        return df_api

    except Exception as e:
        print(f"❌ Erro ao consultar API {url}: {e}")
        return pd.DataFrame()

def verificar_destinatario(mail):
    #faz a verificação se o email foi enviado para distribuição Fundos
    try:
        destinatarios_to = mail.To.lower() if mail.To else ""
        destinatarios_CC = mail.CC.lower() if mail.CC else ""

        todos_dest =  destinatarios_to + " " + destinatarios_CC

        termos_busca = [
            "distribuição fundos",
            "distribuicao fundos",
            "dist fundos",
            "fundos@"
        ]

        termos_normalizados = [normalizar_texto(b) for b in termos_busca]

        for termo in termos_normalizados:
            if normalizar_texto(termo) in normalizar_texto(todos_dest):
                return True

        return False
    
    except Exception as e:
        print(f"Erro ao verificar destinatários: {e}")
        return False
    

def buscar_emails_na_inbox(inbox, assuntos):
    """Busca emails na caixa de entrada do Outlook baseado nos assuntos"""
    emails = []
    try:
        itens = inbox.Items
        itens.Sort("[ReceivedTime]", True)

        # Filtro de data - desde ontem às 15h
        hoje = datetime.now().date()
        ontem = hoje - timedelta(days=1)
        data_inicio = ontem.strftime('%m/%d/%Y') + ' 15:00'

        # Aplica filtro de data
        restr = f"[ReceivedTime] >= '{data_inicio}'"
        itens_filtrados = itens.Restrict(restr)

        print(f"📧 Total de emails desde ontem 15h: {itens_filtrados.Count}")

        # Normaliza os assuntos de busca
        assuntos_normalizados = [normalizar_texto(a) for a in assuntos]

        # Filtra por assunto manualmente
        for item in itens_filtrados:
            try:
                if hasattr(item, "Subject") and item.Subject:
                    # Normaliza o assunto do email
                    subj_normalizado = normalizar_texto(item.Subject)

                    # Compara com assuntos normalizados
                    for assunto_norm in assuntos_normalizados:
                        if assunto_norm in subj_normalizado:
                            if verificar_destinatario(item):
                                emails.append(item)
                                print(f"  ✓ Encontrado: {item.Subject}")
                            else: 
                                print(f"Ignorado (não enviado para dist. fundos): {item.Subject}")
                            break
            except Exception as e:
                print(f"⚠️ Erro ao processar email: {e}")
                continue

    except Exception as e:
        print(f"❌ Erro na caixa de entrada: {e}")

    return emails


def salvar_anexos(emails, pasta_saida):
    """Salva os anexos dos emails na pasta especificada"""
    arquivos_salvos = []
    
    for mail in emails:
        if mail.Attachments.Count > 0:
            for anexo in mail.Attachments:
                try:
                    nome_arquivo = anexo.FileName
                    caminho_completo = os.path.join(pasta_saida, nome_arquivo)
                    anexo.SaveAsFile(caminho_completo)
                    arquivos_salvos.append(caminho_completo)
                    print(f"  ✓ Salvo: {nome_arquivo}")
                except Exception as e:
                    print(f"  ❌ Erro ao salvar {nome_arquivo}: {e}")

    return arquivos_salvos


def processar_dados_api(df_api, tipo_operacao):
    """Processa os dados da API conforme o tipo de operação"""
    if df_api.empty:
        return df_api

    # Definição de Status
    df_api["STATUS"] = df_api["STATUS"].apply(lambda x: "PENDENTE" if x == 1 else "ND")

    # Define operação
    if tipo_operacao == "aplicacao":
        df_api["OPERAÇÃO"] = "APLICA"
    elif tipo_operacao == "resgate":
        df_api["OPERAÇÃO"] = df_api["RESGATE"].apply(
            lambda x: "RESGATE TOTAL" if x is True else "RESGATE BRUTO"
        )
        df_api["VALOR"] = df_api.apply(
            lambda row: 0 if row["OPERAÇÃO"] == "RESGATE TOTAL" else row["VALOR"],
            axis=1
        )
        df_api = df_api.drop('RESGATE', axis=1)

    # Converte e filtra datas
    df_api['DATA'] = pd.to_datetime(df_api['DATA'], errors='coerce')
    
    hoje = pd.Timestamp.today().normalize()
    agora = pd.Timestamp.now()
    ontem = hoje - pd.Timedelta(days=1) + pd.Timedelta(hours=15)

    df_api = df_api[(df_api['DATA'] >= ontem) & (df_api["DATA"] < agora + pd.Timedelta(seconds=1))]
    df_api["DATA"] = df_api["DATA"].dt.normalize()

    # Ordena colunas
    ordem = ["ID", "FUNDO", "CNPJ", "CODIGO", "NOME", "NOME.1", "VALOR", "STATUS", "DATA", "OPERAÇÃO"]
    df_api = df_api[ordem]

    return df_api


def consolidar_dados_api(urls):
    """Consolida dados de todas as APIs"""
    df_total = pd.DataFrame()
    tipos = ["aplicacao", "resgate"]

    for idx, url in enumerate(urls):
        print(f"\n🔄 Carregando dados da API: {url}")
        df_api = carregar_dados_api(url)

        if not df_api.empty:
            df_api = processar_dados_api(df_api, tipos[idx])
            print(f"✓ API retornou {len(df_api)} registros após filtros")
            df_total = pd.concat([df_total, df_api], ignore_index=True)
        else:
            print("⚠️ Nenhum dado retornado pela API.")

    return df_total


def consolidar_anexos(arquivos_salvos, df_base):
    """Consolida dados dos anexos salvos"""
    df_total = df_base.copy()

    for arquivo in arquivos_salvos:
        try:
            if arquivo.endswith(".xlsx"):
                df = pd.read_excel(arquivo)
            elif arquivo.endswith(".csv"):
                df = pd.read_csv(arquivo, sep=";", encoding="utf-8")
            else:
                continue

            df_total = pd.concat([df_total, df], ignore_index=True)
            print(f"  ✓ Consolidado: {os.path.basename(arquivo)}")

        except Exception as e:
            print(f"  ❌ Erro ao ler {arquivo}: {e}")

    return df_total


def salvar_excel_com_popup(df, caminho):
    """Salva arquivo Excel com popup se estiver aberto"""
    while True:
        try:
            df.to_excel(caminho, index=False)
            print(f"✓ Arquivo salvo com sucesso em {caminho}")
            break
        except (PermissionError, OSError):
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning(
                "Aviso",
                "Feche a planilha de boletas e as planilhas extraídas antes de continuar."
            )
            root.destroy()
            time.sleep(3)


def adicionar_macro_vba(caminho_arquivo):
    """Adiciona botão e macro VBA ao arquivo Excel"""
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    wb = excel.Workbooks.Open(caminho_arquivo)
    ws = wb.Sheets(1)

    # Adiciona botão
    btn = ws.Buttons().Add(550, 10, 160, 30)
    btn.OnAction = "EnviarParaOrderManager"
    btn.Text = "Enviar para o Order Manager"

    # Adiciona módulo VBA
    vb_module = wb.VBProject.VBComponents.Add(1)
    vb_module.Name = "ModuloEnvio"

    codigo_macro = r'''
Sub EnviarParaOrderManager()
    Dim origem As Worksheet
    Dim destino As Workbook

    Set origem = ThisWorkbook.Sheets(1)
    Rows(1).Delete

    With ThisWorkbook.Sheets(1)
        ' Coluna G em formato financeiro
        .Columns("G").NumberFormat = "#,##0.00 [$R$-416]"
        ' Coluna I em formato de data abreviada (dd/mm/aa)
        .Columns("I").NumberFormat = "dd/mm/yy"
    End With

    Set destino = Workbooks.Open("C:\TEMP\boletas\Order Manager Fundos_V1.5 - Oficial (Britech).xlsm")

    origem.UsedRange.Copy Destination:=destino.Sheets(5).Range("A6")

    destino.Save
    destino.Close
    MsgBox "Dados enviados com sucesso!", vbInformation
End Sub
'''
    vb_module.CodeModule.AddFromString(codigo_macro)
    print("✓ Macro VBA adicionada com sucesso!")


def exibir_estatisticas(emails):
    """Exibe estatísticas detalhadas dos emails encontrados"""
    print("\n" + "="*60)
    print(f"TOTAL DE EMAILS ENCONTRADOS: {len(emails)}")
    print("="*60)

    if not emails:
        print("⚠️ Nenhum email encontrado com este filtro.")
        return

    # Contadores específicos
    assuntos_busca = {
        'BOLETA DE MOVIMENTACAO FUNDOS': 0,
        'Aplicação Fundos': 0,
        'Aplicação em Fundos': 0,
        'TEDs recebidas de Fundos': 0,
        'ENC: Resgate': 0,
        'Resgate': 0,
        'Resgate Fundos': 0,
        'ENC: Resgate Fundos': 0,
        'Aplic': 0,
        'TEDs recebidas': 0,
        'LIQUIDAÇÃO': 0
    }

    emails_com_anexo = 0
    emails_sem_anexo = 0

    print("\n📋 Detalhamento dos emails encontrados:")
    for idx, mail in enumerate(emails, 1):
        subj = mail.Subject
        tem_anexo = mail.Attachments.Count > 0
        
        if tem_anexo:
            emails_com_anexo += 1
            print(f"  {idx}. ✓ [{mail.ReceivedTime}] {subj} ({mail.Attachments.Count} anexos)")
        else:
            emails_sem_anexo += 1
            print(f"  {idx}. ✗ [{mail.ReceivedTime}] {subj} (SEM ANEXO)")
        
        # Conta por assunto
        subj_lower = subj.lower()
        for assunto in assuntos_busca.keys():
            if assunto.lower() in subj_lower:
                assuntos_busca[assunto] += 1
                break

    print("\n📊 Resumo por assunto:")
    for assunto, count in assuntos_busca.items():
        if count > 0:
            print(f"  • {assunto}: {count}")

    print(f"\n📎 Emails COM anexo: {emails_com_anexo}")
    print(f"📭 Emails SEM anexo: {emails_sem_anexo}")
    print("="*60 + "\n")


# ========================================
# EXECUÇÃO PRINCIPAL
# ========================================

def main():
    print("\n" + "="*60)
    print("CONSOLIDADOR DE BOLETAS - INICIANDO")
    print("="*60 + "\n")

    # 1. Conecta ao Outlook
    inbox = conectar_outlook()

    # 2. Busca emails
    print("\n🔍 Buscando emails...")
    emails = buscar_emails_na_inbox(inbox, assuntos)

    # 3. Exibe estatísticas
    exibir_estatisticas(emails)

    # 4. Salva anexos
    print("💾 Salvando anexos...")
    arquivos_salvos = salvar_anexos(emails, Saida)
    print(f"\n✓ Total de anexos salvos: {len(arquivos_salvos)}")

    if not arquivos_salvos:
        print("⚠️ Nenhum anexo foi salvo. Verifique os e-mails.")

    # 5. Consolida dados da API
    print("\n🔄 Consolidando dados das APIs...")
    df_total = consolidar_dados_api(URL_API)

    # 6. Consolida anexos
    print("\n📂 Consolidando anexos...")
    df_total = consolidar_anexos(arquivos_salvos, df_total)

    # 7. Salva arquivo consolidado
    caminho_final = os.path.join(Saida, "consolidado_boletas.xlsx")
    print(f"\n💾 Salvando arquivo consolidado...")
    salvar_excel_com_popup(df_total, caminho_final)

    # 8. Adiciona macro VBA
    print("\n🔧 Adicionando macro VBA...")
    adicionar_macro_vba(caminho_final)

    # 9. Mensagens finais
    print("\n" + "="*60)
    print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
    print("="*60)
    print("\n⚠️ Caso o consolidado venha vazio ou com pouca quantidade,")
    print("   verifique os arquivos da pasta boletas.")
    
    time.sleep(1000)


# ========================================
# EXECUTA O PROGRAMA
# ========================================

if __name__ == "__main__":
    main()