import pandas as pd
import os
import requests
import logging
from datetime import datetime
from math import ceil
import sys
import json

# --- Configuração de Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('abastecimento_automatico.log')
    ]
)
logger = logging.getLogger(__name__)

# --- Constantes do Script ---
CAMINHO_BASE_PLANILHAS = r"C:\Users\vinicius.magagnini\Desktop\psiAbastecSugest"
NOME_ARQUIVO_BACKLOG = 'backlog_cards_semana.xlsx'
CAMINHO_BACKLOG = os.path.join(CAMINHO_BASE_PLANILHAS, NOME_ARQUIVO_BACKLOG)
REQUESTER_ID = '7e48e47a-8c81-4777-a896-afb2d871ebc7'
URL_WMS = 'https://warehouse-inventory.mottu.cloud/Order/file'
MAX_ITENS_POR_CARD = 30
URL_MOTTU_MESSAGE_API = 'https://message-integration.mottu.cloud/api/v1/messages'

# --- Funções de API e Utilitários ---
def enviar_email_mottu_api(token, subject, body_html, recipient_list):
    """Envia um e-mail usando a API de integração de mensagens da Mottu."""
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {token}'
    }
    payload = {
        "recipients": recipient_list,
        "type": 2,
        "message": body_html,
        "title": subject
    }
    print("Enviando e-mail pela API da Mottu...")
    try:
        response = requests.post(URL_MOTTU_MESSAGE_API, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        print(f"E-mail enviado com sucesso para: {', '.join(recipient_list)}")
        logger.info("Relatório enviado por e-mail com sucesso via API Mottu.")
        return True
    except requests.exceptions.HTTPError as e:
        print(f"\nERRO DE E-MAIL (API Mottu): Falha ao enviar. Status: {e.response.status_code}")
        print(f"Detalhes: {e.response.text}")
        logger.error(f"Falha ao enviar e-mail via API Mottu: {e.response.status_code} - {e.response.text}")
        return False
    except Exception as e:
        print(f"\nERRO DE E-MAIL (API Mottu): Falha inesperada. Detalhe: {e}")
        logger.error(f"Falha inesperada ao enviar e-mail via API Mottu: {e}")
        return False

def get_token_mottu():
    """Obtém token de autenticação do sistema Mottu"""
    try:
        url = 'https://sso.mottu.cloud/realms/Internal/protocol/openid-connect/token'
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        data = {
            'username': 'victor.rodrigues@mottu.com.br',
            'password': '@@Welcome01',
            'client_id': 'mottu-admin',
            'grant_type': 'password'
        }
        logger.info("Obtendo token de autenticação Mottu...")
        response = requests.post(url, headers=headers, data=data, timeout=30)
        response.raise_for_status()
        token = response.json()['access_token']
        logger.info("Token Mottu obtido com sucesso.")
        return token
    except Exception as e:
        logger.error(f"Erro ao obter token Mottu: {e}")
        raise

def get_planilha_para_amanha():
    nome_planilha = 'separacaoAmanha.xlsx'
    caminho_completo = os.path.join(CAMINHO_BASE_PLANILHAS, nome_planilha)
    if not os.path.exists(caminho_completo):
        raise FileNotFoundError(f"Planilha não encontrada: {caminho_completo}")
    logger.info(f"Usando planilha: {nome_planilha}")
    return caminho_completo

def validar_dados_planilha(df):
    colunas_obrigatorias = ['cdAbastecimentoId', 'filialOperacaoId', 'originalCode', 'sugestaoAbastecimento', 'filial', 'abastecimento_cd', 'PesoTotal']
    colunas_faltantes = [col for col in colunas_obrigatorias if col not in df.columns]
    if colunas_faltantes:
        raise ValueError(f"Colunas obrigatórias não encontradas na planilha: {colunas_faltantes}")
    logger.info(f"Planilha válida com {len(df)} linhas")
    return True

def enviar_card_wms(df_card, token, warehouse_origin_id, warehouse_destiny_id):
    arquivo_temp = None
    file_handle = None
    try:
        arquivo_temp = f'pedido_card_temp_{datetime.now().strftime("%Y%m%d_%H%M%S%f")}.xlsx'
        df_card[['originalCode', 'sugestaoAbastecimento']].to_excel(arquivo_temp, index=False, header=['Code', 'Quantity'])
        file_handle = open(arquivo_temp, 'rb')
        files = {'file': file_handle}
        payload = {'warehouseOriginId': warehouse_origin_id, 'category': 3, 'SubCategory': 15, 'requesterId': REQUESTER_ID, 'warehouseDestinyId': warehouse_destiny_id}
        headers = {'Authorization': f'Bearer {token}'}
        logger.info(f"Enviando card para filial {warehouse_destiny_id} com {len(df_card)} itens")
        response = requests.post(URL_WMS, files=files, data=payload, headers=headers, timeout=60)
        response.raise_for_status()
        logger.info(f"[SUCESSO] Card enviado para filial {warehouse_destiny_id}")
        try:
            response_data = response.json()
            card_id = response_data.get('result', {}).get('code', 'ID_Nao_Encontrado')
            return {'success': True, 'card_id': card_id, 'error': None}
        except json.JSONDecodeError:
            logger.warning("Resposta da API não é um JSON válido, mas status foi 200.")
            return {'success': True, 'card_id': 'ID_Nao_Encontrado_JSON_Error', 'error': None}
    except Exception as e:
        error_message = f"Erro ao enviar card para filial {warehouse_destiny_id}: {e}"
        if hasattr(e, 'response') and e.response is not None:
             error_message += f" - Resposta: {e.response.text}"
        logger.error(f"[ERRO] {error_message}")
        return {'success': False, 'card_id': None, 'error': error_message}
    finally:
        if file_handle: file_handle.close()
        if arquivo_temp and os.path.exists(arquivo_temp):
            try: os.remove(arquivo_temp)
            except Exception as e: logger.warning(f"Não foi possível remover arquivo temporário {arquivo_temp}: {e}")

def gerar_relatorio_semanal(df_resultados):
    """Gera o resumo semanal, remove Sábado/Domingo e adiciona o Peso Total."""
    if df_resultados.empty:
        print("Nenhum card foi criado com sucesso para gerar o relatório semanal.")
        return None
    
    resumo_semanal = df_resultados.groupby(['abastecimento_cd', 'dia_separacao_nome']).agg(
        qtd_cards=('card_id', 'nunique'),
        qtd_filiais=('filial_nome', 'nunique'),
        qtd_sku_distintos=('qtd_skus', 'sum'),
        qtd_unit_total=('qtd_unidades', 'sum'),
        peso_total_kg=('peso_total', 'sum') 
    ).reset_index()

    resumo_semanal['resumo_formatado'] = resumo_semanal.apply(
        lambda row: (
            f"{row['qtd_cards']} Cards / {row['qtd_filiais']} Filiais / "
            f"{row['qtd_sku_distintos']} SKUs / {row['qtd_unit_total']} Unid. / "
            f"{row['peso_total_kg']:,.2f} Kg"
        ).replace(",", "X").replace(".", ",").replace("X", "."), 
        axis=1
    )
    
    dias_ordenados = ['SEGUNDA', 'TERCA', 'QUARTA', 'QUINTA', 'SEXTA']
    pivot = resumo_semanal.pivot_table(
        index='abastecimento_cd',
        columns='dia_separacao_nome',
        values='resumo_formatado',
        aggfunc='first'
    ).reindex(columns=dias_ordenados).fillna('-')

    peso_total_da_semana = df_resultados.groupby('abastecimento_cd')['peso_total'].sum()
    pivot['Peso Total da Semana'] = peso_total_da_semana.apply(
        lambda x: f"{x:,.2f} Kg".replace(",", "X").replace(".", ",").replace("X", ".")
    ).fillna('0,00 Kg')
    
    pivot.index.name = "CD de Abastecimento"
    print("\n" + "="*120 + "\nVISÃO GERENCIAL - RESUMO DA EXECUÇÃO POR CD\n" + "="*120)
    print(pivot.to_string())
    print("="*120 + "\n")
    return pivot

# --- Lógica Principal Refatorada ---

def criar_cards_e_salvar_backlog(token):
    """
    MODO 'criar': Lê a planilha, cria os cards, salva o relatório detalhado e o backlog.
    """
    print("MODO 'CRIAR': Iniciando criação de cards para a semana...")
    path_planilha = get_planilha_para_amanha()
    df = pd.read_excel(path_planilha, dtype={'originalCode': str})
    
    df['PesoTotal'] = pd.to_numeric(df['PesoTotal'], errors='coerce').fillna(0)
    
    validar_dados_planilha(df)
    df_filtrado = df[df['sugestaoAbastecimento'] > 0].copy()
    
    if df_filtrado.empty:
        print("Nenhum item com sugestão de abastecimento > 0 encontrado na planilha.")
        return None

    resultados_sucesso, erros_envio = [], []
    dados_para_excel_detalhado = []

    for (cd_id, filial_id), df_filial in df_filtrado.groupby(['cdAbastecimentoId', 'filialOperacaoId']):
        abastecimento_cd, filial_nome = df_filial['abastecimento_cd'].iloc[0], df_filial['filial'].iloc[0]
        
        for i in range(ceil(len(df_filial) / MAX_ITENS_POR_CARD)):
            inicio, fim = i * MAX_ITENS_POR_CARD, (i + 1) * MAX_ITENS_POR_CARD
            df_card = df_filial.iloc[inicio:fim]
            resultado_envio = enviar_card_wms(df_card, token, cd_id, filial_id)
            
            if resultado_envio['success']:
                card_id = resultado_envio['card_id']
                print(f"[SUCESSO] Card {card_id} criado para {filial_nome}")
                
                dia_separacao = df_card['DiasParaSeparacaoConvertido'].iloc[0]

                resultados_sucesso.append({
                    'abastecimento_cd': abastecimento_cd, 'filial_nome': filial_nome, 
                    'dia_separacao_nome': dia_separacao, 'card_id': card_id, 
                    'qtd_skus': df_card['originalCode'].nunique(), 
                    'qtd_unidades': df_card['sugestaoAbastecimento'].sum(),
                    'peso_total': df_card['PesoTotal'].sum()
                })

                df_card_com_id = df_card.copy()
                df_card_com_id['ID DO PEDIDO'] = card_id
                dados_para_excel_detalhado.append(df_card_com_id)

            else:
                erro_msg = resultado_envio.get('error', 'Erro desconhecido')
                print(f"[ERRO] Falha ao criar card para {filial_nome}: {erro_msg}")
                erros_envio.append({'filial': filial_nome, 'cd': abastecimento_cd, 'error': erro_msg})

    if dados_para_excel_detalhado:
        df_final_excel = pd.concat(dados_para_excel_detalhado, ignore_index=True)
        
        colunas_desejadas = [
            'abastecimento_cd', 'filial', 'originalCode', 'manutencaoInsumoNome', 
            'veiculoModelo', 'demandaMes', 'qtdEstoque', 'qtdTransito', 
            'estoqueCdabastecimento', 'estoqueCdSp', 'estoqueCdPe', 'estoqueCdSc', 
            'estoqueCDEx', 'sugestaoAbastecimento', 'PesoTotal', 
            'DiasParaSeparacaoConvertido', 'ID DO PEDIDO', 'Regiao'
        ]
        
        for col in colunas_desejadas:
            if col not in df_final_excel.columns:
                df_final_excel[col] = None
        
        df_final_excel = df_final_excel[colunas_desejadas]
        
        nome_arquivo_excel = f'Relatorio_Abastecimento_Detalhado_{datetime.now().strftime("%Y-%m-%d_%H%M%S")}.xlsx'
        caminho_arquivo_excel = os.path.join(CAMINHO_BASE_PLANILHAS, nome_arquivo_excel)
        
        try:
            df_excel_agg = df_final_excel.groupby(
                [col for col in colunas_desejadas if col != 'PesoTotal']
            )['PesoTotal'].sum().reset_index()

            df_excel_agg.to_excel(caminho_arquivo_excel, index=False)
            print(f"\n[SUCESSO] Relatório Excel detalhado salvo em: {caminho_arquivo_excel}")
            logger.info(f"Relatório Excel detalhado salvo com sucesso em {caminho_arquivo_excel}")
        except Exception as e:
            print(f"\n[ERRO] Não foi possível salvar o relatório Excel detalhado: {e}")
            logger.error(f"Falha ao salvar o relatório Excel detalhado: {e}")
    else:
        print("\nNenhum card criado com sucesso, o arquivo Excel detalhado não será gerado.")

    print("\n" + "="*30 + " CRIAÇÃO DE CARDS CONCLUÍDA " + "="*30)
    
    if erros_envio:
        print("\n" + "!"*80 + "\nCards com erro durante a criação:")
        for erro in erros_envio: print(f"  - CD: {erro['cd']}, Filial: {erro['filial']}, Motivo: {erro['error']}")
        print("!"*80)
    else:
        print("\nTodos os cards foram criados com sucesso!")

    if not resultados_sucesso:
        print("Nenhum card foi criado com sucesso. O arquivo de backlog não será gerado.")
        return None
    
    df_resultados = pd.DataFrame(resultados_sucesso)
    try:
        df_resultados.to_excel(CAMINHO_BACKLOG, index=False)
        logger.info(f"Arquivo de backlog '{NOME_ARQUIVO_BACKLOG}' salvo com sucesso com {len(df_resultados)} registros.")
        print(f"\n[SUCESSO] Arquivo de backlog salvo em: {CAMINHO_BACKLOG}")
    except Exception as e:
        logger.error(f"Falha ao salvar o arquivo de backlog: {e}")
        print(f"\n[ERRO] Não foi possível salvar o arquivo de backlog: {e}")
        return None
        
    return df_resultados

def enviar_relatorio_do_backlog(token):
    """
    MODO 'relatorio': Lê o arquivo de backlog, filtra para D+1 e envia o e-mail.
    """
    print("MODO 'RELATORIO': Lendo backlog para enviar relatório D+1...")
    
    if not os.path.exists(CAMINHO_BACKLOG):
        print(f"ERRO: Arquivo de backlog '{NOME_ARQUIVO_BACKLOG}' não encontrado.")
        print("Execute o script em modo 'criar' primeiro para gerar o backlog da semana.")
        return

    df_resultados = pd.read_excel(CAMINHO_BACKLOG)
    
    if df_resultados.empty:
        print("O arquivo de backlog está vazio. Nenhum relatório para enviar.")
        return

    hoje_str = datetime.now().strftime("%d/%m/%Y")
    
    dia_seguinte_index = (datetime.now().weekday() + 1) % 7
    mapa_dias = {0: 'SEGUNDA', 1: 'TERCA', 2: 'QUARTA', 3: 'QUINTA', 4: 'SEXTA', 5: 'SABADO', 6: 'DOMINGO'}
    nome_dia_d1 = mapa_dias[dia_seguinte_index]
    
    print(f"Hoje é {datetime.now().strftime('%A')}, o relatório D+1 será para: {nome_dia_d1}")
    
    df_resultados_d1 = df_resultados[df_resultados['dia_separacao_nome'].str.upper() == nome_dia_d1].copy()
    
    df_relatorio_semanal = gerar_relatorio_semanal(df_resultados)
    html_semanal = df_relatorio_semanal.to_html(classes='custom-table') if df_relatorio_semanal is not None else "<p>Nenhum dado para a visão gerencial.</p>"

    peso_total_do_dia = 0
    resumo_por_cd = []
    partes_html_diario = []
    
    if not df_resultados_d1.empty:
        # Calcula peso total do dia
        peso_total_do_dia = df_resultados_d1['peso_total'].sum()
        
        for cd_nome, df_cd in df_resultados_d1.groupby('abastecimento_cd'):
            # Calcula peso total do CD para o dia
            peso_total_cd = df_cd['peso_total'].sum()
            
            partes_html_diario.append(f"<h3>Cards de Separação para: {cd_nome}</h3>")
            
            resumo_cd = df_cd.groupby('filial_nome').agg(
                qtd_sku_distintos=('qtd_skus', 'sum'), 
                qtd_total_unidades=('qtd_unidades', 'sum'),
                peso_total_kg=('peso_total', 'sum'),
                qtd_cards=('card_id', 'nunique'), 
                card_ids=('card_id', lambda x: ', '.join(map(str, sorted(x.unique()))))
            ).reset_index()
            
            resumo_cd['peso_total_kg'] = resumo_cd['peso_total_kg'].apply(
                lambda x: f"{x:,.2f} Kg".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            
            resumo_cd.columns = ['Filial Separada', 'Qtd SKU Distintos', 'Qtd Total Unidades', 'Peso Total', 'Qtd Cards Gerados', 'IDs dos Cards']
            partes_html_diario.append(resumo_cd.to_html(index=False, classes='custom-table'))
            
            # Adiciona resumo do CD com peso total
            peso_cd_formatado = f"{peso_total_cd:,.2f} Kg".replace(",", "X").replace(".", ",").replace("X", ".")
            total_cards_cd = df_cd['card_id'].nunique()
            total_filiais_cd = df_cd['filial_nome'].nunique()
            
            partes_html_diario.append(f"""<div style="padding: 10px; margin: 10px 0; border: 1px solid #ccc;">
                <strong>RESUMO {cd_nome}:</strong> {total_cards_cd} Cards | {total_filiais_cd} Filiais | 
                <strong>PESO TOTAL: {peso_cd_formatado}</strong>
            </div>""")
            
            resumo_por_cd.append({
                'cd': cd_nome,
                'peso': peso_total_cd,
                'cards': total_cards_cd,
                'filiais': total_filiais_cd
            })
    else:
        partes_html_diario.append(f"<p>Nenhum card a ser separado para o dia seguinte ({nome_dia_d1}).</p>")

    # Adiciona resumo geral do dia no início
    if peso_total_do_dia > 0:
        peso_dia_formatado = f"{peso_total_do_dia:,.2f} Kg".replace(",", "X").replace(".", ",").replace("X", ".")
        total_cards_dia = df_resultados_d1['card_id'].nunique()
        total_cds_dia = len(resumo_por_cd)
        
        resumo_geral_dia = f"""<div style="padding: 15px; margin: 20px 0; border: 1px solid #ccc;">
            <h3 style="margin-top: 0;">RESUMO GERAL DO DIA ({nome_dia_d1})</h3>
            <p style="font-size: 16px; margin: 5px 0;"><strong>Total de CDs:</strong> {total_cds_dia}</p>
            <p style="font-size: 16px; margin: 5px 0;"><strong>Total de Cards:</strong> {total_cards_dia}</p>
            <p style="font-size: 18px; margin: 10px 0;"><strong>PESO TOTAL DO DIA: {peso_dia_formatado}</strong></p>
        </div>"""
        
        partes_html_diario.insert(0, resumo_geral_dia)
    
    html_diario_completo = "<br>".join(partes_html_diario)
    
    html_body = f"""
    <html><head><style>
        body {{ font-family: Arial, sans-serif; font-size: 14px; }}
        table.custom-table {{ border-collapse: collapse; width: auto; margin-bottom: 20px; border: 1px solid #cccccc; }}
        th, td {{ border: 1px solid #dddddd; text-align: left; padding: 8px; }}
        th {{ background-color: #f2f2f2; }}
        h2, h3 {{ color: #333333; }}
        hr {{ border: 0; border-top: 1px solid #cccccc; }}
    </style></head>
    <body>
        <h2>Relatório de Separação (D+1) - {hoje_str}</h2>
        {html_diario_completo}
        <br><hr><br>
        <h2>Visão Gerencial - Resumo da Execução por CD</h2>
        {html_semanal}
    </body></html>
    """

    destinatarios = list(set([
        # 'victor.rodrigues@mottu.com.br',
        'vinicius.magagnini@mottu.com.br',
        'fabio.canton@mottu.com.br', 
        # 'bruce.cardoso@mottu.com.br',
        # 'matheus.paula@mottu.com.br', 'joao.junqueira@mottu.com.br',
        # 'marcelo.paiva@mottu.com.br', 'rogerio.oliveira@mottu.com.br',
        # 'cdpe@mottu.com.br', 'martello@mottu.com.br', 'celina.melo@mottu.com.br'
    ]))
    
    enviar_email_mottu_api(token=token, subject=f"Automação Abastecimento: Relatório de {hoje_str}", body_html=html_body, recipient_list=destinatarios)

def main():
    """Função principal que direciona a execução com base na escolha do usuário."""
    
    modo = None
    while True:
        print("\n" + "="*50)
        print("Você deseja executar o script em qual modo?")
        print("  1 - MODO CRIAR (Execução de Domingo)")
        print("      - Gera todos os cards, o relatório detalhado e o backlog.")
        print("\n  2 - MODO RELATÓRIO (Execução de Segunda a Sexta)")
        print("      - Envia o e-mail D+1 com base no backlog existente.")
        print("="*50)
        
        escolha = input("Digite o número da opção desejada (1 ou 2): ")
        
        if escolha == '1':
            modo = 'criar'
            break
        elif escolha == '2':
            modo = 'relatorio'
            break
        else:
            print("\n!!! Opção inválida. Por favor, digite 1 ou 2. !!!")

    try:
        print(f"\nIniciando processamento em MODO '{modo.upper()}'...")
        token = get_token_mottu()
        
        if modo == 'criar':
            df_resultados = criar_cards_e_salvar_backlog(token)
            if df_resultados is not None:
                print("\nEnviando e-mail de confirmação da criação dos cards...")
                enviar_relatorio_do_backlog(token)
        
        elif modo == 'relatorio':
            enviar_relatorio_do_backlog(token)
            
    except Exception as e:
        logger.error(f"Erro fatal no script (Modo: {modo}): {e}", exc_info=True)
        print(f"\nERRO FATAL: {e}")
        sys.exit(1)

if __name__ == '__main__':
    try:
        main()
        input("\nProcesso concluído. Pressione ENTER para sair.")
    except KeyboardInterrupt:
        logger.info("Processo interrompido pelo usuário")
        print("\nProcesso interrompido.")