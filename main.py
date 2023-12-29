from imap_tools import MailBox, AND
from datetime import datetime
from elasticsearch import Elasticsearch, helpers

import pandas as pd
import csv
import numpy as np
import pytz
import os

################### INFORMAÇÕES IMPORTANTES ###################
# Emails de um destinatário
username = "destinatario@email.com"
password = "PASSWORD"

email_remetente = "noreply-gzc@info.bitdefender.com"

# Configurações do Elasticsearch (incluindo autenticação)
host = "172.16.1.87"
port = 9200
usuario = "elastic"
senha = 'PASSWORD ELASTIC'

# Nome do arquivo de anexo que deve ser verificado
nome_relatorio = "Relatório de Auditoria de Segurança"
###############################################################
def limpar_diretorio(caminho_diretorio):
    try:
        # Obtém a lista de arquivos no diretório
        arquivos = os.listdir(caminho_diretorio)

        # Itera sobre os arquivos e remove cada um
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(caminho_diretorio, arquivo)
            if os.path.isfile(caminho_arquivo):
                os.remove(caminho_arquivo)

        print(f'Limpeza do diretório {caminho_diretorio} concluída com sucesso.')

    except Exception as e:
        print(f"Ocorreu um erro durante a limpeza do diretório: {e}")


# Substitua 'caminho_do_seu_diretorio' pelo caminho do diretório que você deseja limpar
caminho_do_seu_diretorio = "./reportsBitdefender/"
limpar_diretorio(caminho_do_seu_diretorio)

# Obter a data atual
data_atual = datetime.now()

# Formatar a data no formato YYYY-MM-DD
data_formatada = data_atual.strftime('%Y-%m-%d')

# lista de imaps: https://www.systoolsgroup.com/imap/
meu_email = MailBox('172.16.1.9').login(username, password)

# criterios: https://github.com/ikvk/imap_tools#search-criteria
lista_emails = meu_email.fetch(AND(from_=email_remetente))
for email in lista_emails:
    print(email.subject)
    print(email.text)

# pegar emails com um anexo específico
anexo_relatorio = f"{nome_relatorio}-{data_formatada}"
lista_emails = meu_email.fetch(AND(from_=email_remetente))
for email in lista_emails:
    if len(email.attachments) > 0:
        for anexo in email.attachments:
            if anexo_relatorio in anexo.filename:
                print(anexo.content_type)
                print(anexo.payload)

                nome_do_arquivo = f"./reportsBitdefender/{anexo.filename}"
                with open(nome_do_arquivo, 'wb') as arquivo_excel:
                    arquivo_excel.write(anexo.payload)

                    # Configurações do Elasticsearch
                    nome_indice = "grupocaio_bitdefender"
                    data_index = datetime.now().strftime('%Y.%m')
                    nome_indice_com_data = f"{nome_indice}.{data_index}"

                    # Mapeamento de renomeação de colunas
                    mapeamento_renomeacao = {
                        "Nome do Endpoint": "endpointNome",
                        "FQDN do Endpoint": "endpointFQDN",
                        "Usuário": "endpointUsuario",
                        "Ocorrências": "endpointOcorrencia",
                        "Ultima ocorrência":"endpointUltimaOcorrencia",
                        "Módulo": "endpintModulo",
                        "Tipo de Evento": "endpointTipoDeEvento",
                        "Detalhes": "endpointDetalhes",
                        "SHA256 Hash": "endpointSHA256Hash",
                        "Ataque sem arquivo": "endpointAtaqueSemArquivo"
                    }

                    # Carregar dados do CSV usando pandas com tratamento de NaN
                    try:
                        dados_csv = pd.read_csv(nome_do_arquivo, quoting=csv.QUOTE_ALL)

                        # Renomear colunas, se necessário
                        dados_csv.rename(columns=mapeamento_renomeacao, inplace=True)

                        # Substituir NaN por None (ou numpy.nan, se preferir)
                        dados_csv.replace("NA", None, inplace=True)

                        # Converter dados para formato JSON
                        dados_csv.replace({np.nan: None}, inplace=True)
                        dados_json = dados_csv.to_dict(orient='records')

                        # Enviar dados para o Elasticsearch
                        with Elasticsearch([{'host': host, 'port': port, 'scheme': 'http'}], basic_auth=(usuario, senha)) as es:
                            for documento in dados_json:
                                # Adicione um campo personalizado
                                documento['nomeRelatorio'] = nome_relatorio

                                documento['@timestamp'] = datetime.now(pytz.utc).isoformat()
                                # Converta o timestamp para o fuso horário desejado (por exemplo, America/Sao_Paulo)
                                timestamp_str = documento['@timestamp'].replace('+00:00', 'Z')
                                timestamp_utc = datetime.strptime(timestamp_str, '%Y-%m-%dT%H:%M:%S.%fZ')
                                timestamp_utc = pytz.utc.localize(timestamp_utc)
                                documento['@timestamp'] = timestamp_utc.astimezone(pytz.timezone('America/Sao_Paulo')).isoformat()

                                es.index(index=nome_indice_com_data, body=documento)

                        print(f'Dados do CSV enviados para o Elasticsearch no índice "{nome_indice_com_data}"')

                    except Exception as e:
                        print(f"Ocorreu um erro: {e}")
