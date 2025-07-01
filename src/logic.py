import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
from tkinter import messagebox
import json
from requests.exceptions import RequestException
from typing import List, Dict, Any, Tuple
from queue import Queue
import pandas as pd


def validate_source_file_headers_thread(source_file: str, possible_name_cols: List[str], possible_email_cols: List[str],
                                        queue: Queue) -> None:
    """Lê apenas o cabeçalho do arquivo de origem e verifica as colunas."""
    try:
        if not source_file:
            return

        df_header = pd.read_excel(source_file, nrows=0) if source_file.endswith('.xlsx') else pd.read_csv(source_file,
                                                                                                          nrows=0)

        has_name = any(col in df_header.columns for col in possible_name_cols)
        has_email = any(col in df_header.columns for col in possible_email_cols)

        if has_name and has_email:
            queue.put(("log", ("Arquivo de contatos validado com sucesso!", "SUCCESS")))
        else:
            missing = []
            if not has_name: missing.append("NOME")
            if not has_email: missing.append("EMAIL")
            msg = f"O arquivo parece não conter a(s) coluna(s) obrigatória(s): {', '.join(missing)}."
            queue.put(("log", (msg, "WARNING")))
            messagebox.showwarning("Cabeçalho Inválido", msg)
    except Exception as e:
        msg = f"Não foi possível ler o arquivo de contatos: {e}"
        queue.put(("log", (msg, "ERROR")))
        messagebox.showerror("Erro de Arquivo", msg)


def check_and_clear_sheet_thread(json_path: str, mailmerge_url: str, queue: Queue) -> None:
    queue.put(("buttons_state", "disabled"))
    queue.put(("progress_start", "Verificando/Limpando planilha..."))
    queue.put(("log", ("Iniciando verificação/limpeza da planilha...", "INFO")))
    service_account_email = ''
    try:
        if not all([json_path, mailmerge_url]):
            raise ValueError("Os campos 'Arquivo de Chave JSON' e 'URL da Planilha' devem ser preenchidos.")
        with open(json_path, 'r') as f:
            sa_info = json.load(f)
        service_account_email = sa_info.get('client_email')
        if not service_account_email:
            raise ValueError("Arquivo JSON inválido ou não contém o campo 'client_email'.")
        queue.put(("log", ("Autenticando com Conta de Serviço...", "INFO")))
        scopes_svc = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = Credentials.from_service_account_file(json_path, scopes=scopes_svc)
        client = gspread.authorize(credentials)
        queue.put(("log", ("Acessando a planilha...", "INFO")))
        planilha_mailmerge = client.open_by_url(mailmerge_url)
        aba_mailmerge = planilha_mailmerge.get_worksheet(0)
        queue.put(("log", (f"Conexão bem-sucedida com a planilha: '{planilha_mailmerge.title}'", "SUCCESS")))
        num_registros = len(aba_mailmerge.get_all_records())
        if num_registros > 0:
            queue.put(("log", (f"A planilha contém {num_registros} registros.", "WARNING")))
            if messagebox.askyesno("Planilha Contém Dados",
                                   f"A planilha de destino contém {num_registros} registros.\n\nDeseja apagar TODOS os dados (mantendo o cabeçalho)?"):
                queue.put(("log", ("Usuário confirmou a limpeza. Apagando dados...", "INFO")))
                headers = aba_mailmerge.row_values(1) if aba_mailmerge.row_count > 0 else ['First name', 'Last name',
                                                                                           'Recipient', 'Description',
                                                                                           'Email Sent']
                aba_mailmerge.clear()
                time.sleep(1)
                aba_mailmerge.append_row(headers, value_input_option="USER_ENTERED")
                queue.put(("log", ("Planilha limpa com sucesso. Cabeçalhos mantidos.", "SUCCESS")))
                messagebox.showinfo("Sucesso", "A planilha foi limpa com sucesso!")
            else:
                queue.put(("log", ("Limpeza cancelada pelo usuário.", "WARNING")))
        else:
            queue.put(("log", ("A planilha de destino já está vazia.", "INFO")))
            messagebox.showinfo("Informação", "A planilha já está vazia. Nenhuma ação de limpeza foi necessária.")

    except RequestException:
        msg = "Falha de rede ao contatar a API do Google. Verifique sua conexão com a internet."
        queue.put(("log", (msg, "ERROR")))
        messagebox.showerror("Erro de Rede", msg)

    except Exception as e:
        level = "ERROR"
        msg = f"{type(e).__name__} - {e}"
        if isinstance(e, gspread.exceptions.APIError) and e.response.json().get('error', {}).get(
                'status') == 'PERMISSION_DENIED':
            queue.put(("log", ("ERRO: Permissão negada para a Conta de Serviço.", level)))
            queue.put(("permission_error", service_account_email))
        else:
            queue.put(("log", (msg, level)))
            messagebox.showerror("Erro na Operação", f"Não foi possível completar a operação.\n\nDetalhe: {msg}")
    finally:
        queue.put(("progress_stop", None))
        queue.put(("buttons_state", "normal"))


def analyze_data_thread(json_path: str, mailmerge_url: str, source_file: str, possible_name_cols: List[str],
                        possible_email_cols: List[str], queue: Queue) -> None:
    queue.put(("buttons_state", "disabled"))
    queue.put(("progress_start", "Analisando contatos..."))
    queue.put(("log", ("Iniciando processo de análise...", "INFO")))
    service_account_email = ''
    try:
        if not all([json_path, mailmerge_url, source_file]):
            raise ValueError("Todos os campos (JSON, URL e Arquivo de Origem) são obrigatórios.")
        with open(json_path, 'r') as f:
            sa_info = json.load(f)
        service_account_email = sa_info.get('client_email')

        queue.put(("log", ("Autenticando com Conta de Serviço...", "INFO")))
        scopes_svc = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = Credentials.from_service_account_file(json_path, scopes=scopes_svc)
        client = gspread.authorize(credentials)

        queue.put(("log", ("Acessando a planilha...", "INFO")))
        planilha_mailmerge = client.open_by_url(mailmerge_url)
        aba_mailmerge = planilha_mailmerge.get_worksheet(0)

        queue.put(("log", ("Otimização: Lendo apenas a coluna de e-mails existentes...", "INFO")))
        headers = aba_mailmerge.row_values(1)
        if 'Recipient' not in headers:
            raise ValueError("A planilha de destino deve ter uma coluna de cabeçalho chamada 'Recipient'.")
        recipient_col_index = headers.index('Recipient') + 1
        email_list = aba_mailmerge.col_values(recipient_col_index)[1:]
        emails_existentes = set(
            email.strip().lower() for email in email_list if isinstance(email, str) and email.strip())
        num_existentes = len(emails_existentes)
        queue.put(("log", (f"Encontrados {num_existentes} contatos únicos na planilha.", "INFO")))

        queue.put(("log", ("Lendo arquivo de origem...", "INFO")))
        agencias_df = pd.read_csv(source_file) if source_file.endswith('.csv') else pd.read_excel(source_file)
        num_origem = len(agencias_df)
        queue.put(("log", (f"Encontrados {num_origem} contatos no arquivo.", "INFO")))

        actual_name_col = next((c for c in possible_name_cols if c in agencias_df.columns), None)
        actual_email_col = next((c for c in possible_email_cols if c in agencias_df.columns), None)
        if not all([actual_name_col, actual_email_col]):
            raise ValueError(
                "Colunas de nome/e-mail não encontradas no arquivo de origem. Verifique o arquivo de contatos ou as configurações em config.json.")

        queue.put(("log", (f"Mapeando: '{actual_name_col}' -> First name, '{actual_email_col}' -> Recipient.", "INFO")))
        novos_dados = agencias_df[[actual_name_col, actual_email_col]].copy()
        novos_dados.columns = ['First name', 'Recipient']
        novos_dados.dropna(inplace=True)
        novos_filtrados = novos_dados[~novos_dados['Recipient'].str.lower().isin(emails_existentes)]
        num_novos = len(novos_filtrados)
        queue.put(("log", ("Análise concluída.", "SUCCESS")))

        analysis_result_text = f"Arquivo: {num_origem} contatos | Planilha: {num_existentes} contatos | NOVOS: {num_novos}"
        spinbox_config, sync_button_config = (
            {"state": "normal", "from_": 1, "to": num_novos, "start_value": 1, "end_value": num_novos},
            {"state": "normal"}) if num_novos > 0 else ({"state": "disabled"}, {"state": "disabled"})

        if num_novos == 0:
            queue.put(("log", ("Nenhum novo contato para adicionar.", "WARNING")))

        queue.put(("update_analysis", (novos_filtrados, analysis_result_text, spinbox_config, sync_button_config)))

    except RequestException:
        msg = "Falha de rede ao contatar a API do Google. Verifique sua conexão com a internet."
        queue.put(("log", (msg, "ERROR")))
        messagebox.showerror("Erro de Rede", msg)
        queue.put(("update_analysis",
                   (None, "Falha na análise. Verifique o log.", {"state": "disabled"}, {"state": "disabled"})))

    except Exception as e:
        level = "ERROR"
        msg = f"{type(e).__name__} - {e}"
        if isinstance(e, gspread.exceptions.APIError) and e.response.json().get('error', {}).get(
                'status') == 'PERMISSION_DENIED':
            queue.put(("log", ("ERRO: Permissão negada para a Conta de Serviço.", level)))
            queue.put(("permission_error", service_account_email))
        else:
            queue.put(("log", (msg, level)))
            messagebox.showerror("Erro na Análise", f"Não foi possível completar a análise.\n\nDetalhe: {msg}")
        queue.put(("update_analysis",
                   (None, "Falha na análise. Verifique o log.", {"state": "disabled"}, {"state": "disabled"})))
    finally:
        queue.put(("progress_stop", None))
        queue.put(("buttons_state", "normal"))


def sync_data_thread(json_path: str, mailmerge_url: str, is_dry_run: bool, contacts_df_to_sync: pd.DataFrame,
                     queue: Queue) -> None:
    queue.put(("buttons_state", "disabled"))
    queue.put(("progress_start", "Sincronizando contatos..."))
    log_prefix = "SIMULAÇÃO" if is_dry_run else "SINCRONIZAÇÃO"
    queue.put(("log", (f"Iniciando processo de {log_prefix.lower()}...", "INFO")))
    try:
        if contacts_df_to_sync is None or contacts_df_to_sync.empty:
            raise ValueError("Nenhum novo parceiro para sincronizar.")

        linhas_para_adicionar = [[row['First name'], '', row['Recipient']] for index, row in
                                 contacts_df_to_sync.iterrows()]
        num_linhas = len(linhas_para_adicionar)
        queue.put(("log", (f"Preparando para adicionar {num_linhas} contatos...", "INFO")))

        if is_dry_run:
            queue.put(("log", (f"MODO SIMULAÇÃO: {num_linhas} linhas seriam adicionadas.", "SUCCESS")))
            messagebox.showinfo("Simulação Concluída", f"{num_linhas} novos contatos seriam processados.")
        else:
            queue.put(("log", ("Conectando ao Google para escrever os dados...", "INFO")))
            scopes_svc = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            credentials = Credentials.from_service_account_file(json_path, scopes=scopes_svc)
            client = gspread.authorize(credentials)
            planilha_mailmerge = client.open_by_url(mailmerge_url)
            aba_mailmerge = planilha_mailmerge.get_worksheet(0)
            queue.put(("log", ("Adicionando novas linhas à planilha...", "INFO")))
            aba_mailmerge.append_rows(linhas_para_adicionar, value_input_option="USER_ENTERED")
            queue.put(("log", (f"SUCESSO! {num_linhas} novas linhas adicionadas.", "SUCCESS")))
            messagebox.showinfo("Sincronização Concluída", f"{num_linhas} novos contatos foram adicionados.")

    except RequestException:
        msg = f"Falha de rede durante a {log_prefix.lower()}. Verifique sua conexão com a internet."
        queue.put(("log", (msg, "ERROR")))
        messagebox.showerror("Erro de Rede", msg)

    except Exception as e:
        msg = f"ERRO NA {log_prefix}: {type(e).__name__} - {e}"
        queue.put(("log", (msg, "ERROR")))
        messagebox.showerror(f"Erro na {log_prefix}", f"Não foi possível sincronizar os dados.\n\nDetalhe: {e}")
    finally:
        queue.put(("progress_stop", None))
        queue.put(("buttons_state", "normal"))
        queue.put(("update_analysis",
                   ("Execute uma nova análise para continuar.", {"state": "disabled"}, {"state": "disabled"})))
        queue.put(("log", ("Processo finalizado.", "INFO")))