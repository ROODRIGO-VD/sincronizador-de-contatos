import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext, font
import threading
from src import logic
import webbrowser
import json
import os
from queue import Queue, Empty
from typing import Dict, Any, List, Optional, Callable
import pandas as pd


class HelpWindow(ttk.Toplevel):
    """Janela de ajuda que ensina a compartilhar a planilha."""

    def __init__(self, parent: tk.Widget, service_account_email: str) -> None:
        super().__init__(parent)
        self.title("Ajuda: Como Compartilhar a Planilha")
        self.geometry("800x600")
        self.transient(parent)
        self.grab_set()
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        title_label = ttk.Label(main_frame, text="Resolvendo Erros de Permissão", font=("-size 12 -weight bold"),
                                bootstyle="info")
        title_label.pack(pady=(0, 10))
        explanation_text = (
            "Se você encontrar um erro de 'PermissionError', significa que o aplicativo não tem acesso à sua planilha. Para resolver, compartilhe a planilha com o e-mail da Conta de Serviço abaixo:")
        explanation_label = ttk.Label(main_frame, text=explanation_text, wraplength=550, justify=LEFT)
        explanation_label.pack(pady=(0, 20))
        steps_frame = ttk.Labelframe(main_frame, text="Passo a Passo", padding=15)
        steps_frame.pack(fill=X, pady=5)
        steps = ["1. Clique no botão 'Copiar E-mail' abaixo.", "2. Abra sua Planilha Google no navegador.",
                 "3. Clique no botão azul 'Partilhar' (Share) no canto superior direito.",
                 "4. Cole o e-mail e defina a permissão como 'Editor'."]
        for step in steps:
            ttk.Label(steps_frame, text=step, wraplength=500, justify=LEFT).pack(anchor=W)
        email_frame = ttk.Frame(main_frame, padding=(0, 15, 0, 0))
        email_frame.pack(fill=X)
        ttk.Label(email_frame, text="E-mail da Conta de Serviço para Compartilhar:").pack(anchor=W)
        self.email_entry = ttk.Entry(email_frame, font=("Courier", 10))
        self.email_entry.insert(0, service_account_email)
        self.email_entry.config(state="readonly")
        self.email_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        self.copy_button = ttk.Button(email_frame, text="Copiar E-mail", bootstyle="outline-success",
                                      command=self.copy_to_clipboard)
        self.copy_button.pack(side=LEFT)
        ok_button = ttk.Button(main_frame, text="OK, Entendi", bootstyle="primary", command=self.destroy)
        ok_button.pack(pady=20)

    def copy_to_clipboard(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.email_entry.get())
        if hasattr(self.master, 'log'):
            self.master.log("E-mail da conta de serviço copiado!", "SUCCESS")


class EditContactWindow(ttk.Toplevel):
    """Uma janela pop-up para editar um contato selecionado."""

    def __init__(self, parent: tk.Widget, item_id: int, contact_data: pd.Series, save_callback: Callable):
        super().__init__(parent)
        self.title("Editar Contato")
        self.geometry("450x200")
        self.transient(parent)
        self.grab_set()

        self.item_id = item_id
        self.save_callback = save_callback

        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=True)

        ttk.Label(main_frame, text="Nome:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.name_var = tk.StringVar(value=contact_data['First name'])
        self.name_entry = ttk.Entry(main_frame, textvariable=self.name_var, width=50)
        self.name_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)
        self.name_entry.focus_set()

        ttk.Label(main_frame, text="Email:").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.email_var = tk.StringVar(value=contact_data['Recipient'])
        self.email_entry = ttk.Entry(main_frame, textvariable=self.email_var, width=50)
        self.email_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=1, pady=15, sticky=E)

        save_button = ttk.Button(button_frame, text="Salvar", command=self._on_save, bootstyle="success")
        save_button.pack(side=LEFT, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=self.destroy, bootstyle="secondary-outline")
        cancel_button.pack(side=LEFT, padx=5)

        main_frame.columnconfigure(1, weight=1)
        self.bind("<Return>", self._on_save)
        self.bind("<Escape>", lambda e: self.destroy())

    def _on_save(self, event: Any = None) -> None:
        """Coleta os dados e chama a função de callback para salvar."""
        new_data = {
            "name": self.name_var.get(),
            "email": self.email_var.get()
        }
        self.save_callback(self.item_id, new_data)
        self.destroy()


class AppGUI:
    def __init__(self, root: ttk.Window) -> None:
        self.root = root
        self.global_new_contacts_df: Optional[pd.DataFrame] = None

        self.queue: Queue = Queue()
        self.config_file: str = "config.json"
        self.config_data: Dict[str, Any] = {}

        self.style = ttk.Style()
        self.theme_var = tk.StringVar(value="superhero")

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.load_or_create_config()
        self._setup_menubar()
        self._setup_widgets()
        self.process_queue()
        self.apply_loaded_config()

    def load_or_create_config(self) -> None:
        default_config = {
            # Adiciona 'saved_mailmerge_urls' à configuração padrão
            "user_settings": {"json_path": "", "mailmerge_url": "", "source_file": "", "theme": "superhero",
                              "saved_mailmerge_urls": []},
            "app_settings": {
                "template_url": "https://docs.google.com/spreadsheets/d/1w8bnEEei0U5fYcOJXfA7ItdyXxnUGnQGJ4vFZrZE04Q/copy?hl=pt-br",
                "possible_name_cols": ["NOME", "First name", "Name", "Nome"],
                "possible_email_cols": ["EMAIL", "Last name", "Email", "E-mail", "E-MAIL", "EMAIL(MINUSCULOS)"]
            }
        }
        try:
            if not os.path.exists(self.config_file):
                self.config_data = default_config
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(self.config_data, f, indent=4, ensure_ascii=False)
            else:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config_data = json.load(f)
                # Garante que as novas chaves existam ao carregar configurações antigas
                if "user_settings" not in self.config_data:
                    self.config_data["user_settings"] = default_config["user_settings"]
                # Adiciona 'saved_mailmerge_urls' se não existir nas configurações carregadas
                if "saved_mailmerge_urls" not in self.config_data["user_settings"]:
                    self.config_data["user_settings"]["saved_mailmerge_urls"] = default_config["user_settings"][
                        "saved_mailmerge_urls"]
                if "app_settings" not in self.config_data:
                    self.config_data["app_settings"] = default_config["app_settings"]
        except Exception as e:
            self.config_data = default_config
            messagebox.showerror("Erro de Configuração", f"Não foi possível ler ou criar o config.json: {e}")

    def apply_loaded_config(self) -> None:
        user_cfg = self.config_data.get("user_settings", {})
        self.entry_json.insert(0, user_cfg.get("json_path", ""))

        # Configura o Combobox da URL do MailMerge
        saved_urls = user_cfg.get("saved_mailmerge_urls", [])
        current_url = user_cfg.get("mailmerge_url", "")  # Pega a última URL usada

        # Garante que a URL atual (se existir) esteja na lista e seja o valor inicial
        if current_url:
            if current_url not in saved_urls:
                saved_urls.insert(0, current_url)  # Adiciona ao início para que seja facilmente selecionável
            self.mailmerge_url_var.set(current_url)  # Define o valor do combobox para a última URL usada

        self.mailmerge_url_combobox['values'] = tuple(saved_urls)  # Define as opções do combobox como uma tupla

        self.entry_source_file.insert(0, user_cfg.get("source_file", ""))
        saved_theme = user_cfg.get("theme", "superhero")
        self.theme_var.set(saved_theme)
        self.change_theme(saved_theme)
        self.log("Configurações carregadas.", "SUCCESS")

    def save_config(self) -> None:
        user_cfg = {
            "json_path": self.entry_json.get(),
            "mailmerge_url": self.mailmerge_url_combobox.get(),  # Salva a URL atualmente selecionada/digitada
            "source_file": self.entry_source_file.get(),
            "theme": self.theme_var.get(),
            "saved_mailmerge_urls": list(self.mailmerge_url_combobox['values'])  # Salva a lista de URLs do combobox
        }
        self.config_data["user_settings"] = user_cfg
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log(f"Falha ao salvar o arquivo de configuração: {e}", "ERROR")

    def on_closing(self) -> None:
        self.log("Fechando e salvando configurações...", "INFO")
        self.save_config()
        self.root.destroy()

    def _setup_menubar(self) -> None:
        menubar = ttk.Menu(self.root)
        options_menu = ttk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="Opções", menu=options_menu)
        theme_menu = ttk.Menu(options_menu, tearoff=False)
        options_menu.add_cascade(label="Temas", menu=theme_menu)
        for theme in sorted(self.style.theme_names()):
            theme_menu.add_radiobutton(label=theme, variable=self.theme_var,
                                       command=lambda t=theme: self.change_theme(t))
        # Adiciona a nova opção de limpar informações
        options_menu.add_separator()
        options_menu.add_command(label="Limpar Todas as Informações", command=self.clear_all_information)
        self.root.config(menu=menubar)

    def change_theme(self, theme_name: str) -> None:
        self.style.theme_use(theme_name)
        if hasattr(self, 'log_text'): self.log(f"Tema alterado para '{theme_name}'", "INFO")

    def _setup_widgets(self) -> None:
        template_url = self.config_data.get("app_settings", {}).get("template_url", "")

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Seção 1
        frame1 = ttk.Labelframe(main_frame, text="1. Autenticação e Planilha de Destino", padding="10")
        frame1.pack(fill=X, pady=5)
        ttk.Label(frame1, text="Arquivo de Chave JSON:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.entry_json = ttk.Entry(frame1)
        self.entry_json.grid(row=0, column=1, sticky=EW, padx=5)
        self.browse_json_button = ttk.Button(frame1, text="Procurar...", bootstyle="info-outline",
                                             command=self._on_browse_json_click)
        self.browse_json_button.grid(row=0, column=2, padx=5)

        # Frame para o Combobox da URL do MailMerge e o botão Salvar
        mailmerge_url_frame = ttk.Frame(frame1)
        mailmerge_url_frame.grid(row=1, column=0, columnspan=3, sticky=EW, padx=5, pady=9)
        ttk.Label(mailmerge_url_frame, text="URL da Planilha MailMerge:").pack(side=LEFT, padx=(0, 5))

        self.mailmerge_url_var = tk.StringVar()
        self.mailmerge_url_combobox = ttk.Combobox(mailmerge_url_frame, textvariable=self.mailmerge_url_var)
        self.mailmerge_url_combobox.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        #botão para salvar o link da planilha
        self.save_mailmerge_link_button = ttk.Button(mailmerge_url_frame, text="Salvar Link",
                                                     bootstyle="outline-primary",
                                                     command=self._on_save_mailmerge_link)
        self.save_mailmerge_link_button.pack(side=LEFT)

        mailmerge_url_frame.columnconfigure(1, weight=1)  # Faz o combobox expandir

        action_frame1 = ttk.Frame(frame1)
        action_frame1.grid(row=2, column=1, columnspan=3, pady=10, sticky=EW)
        link_font = font.nametofont("TkDefaultFont").copy()
        link_font.configure(underline=True)
        link_label = ttk.Label(action_frame1, text="Não tem uma planilha? Clique aqui para gerar o modelo.",
                               bootstyle="info", cursor="hand2", font=link_font)
        link_label.pack(side=LEFT)
        link_label.bind("<Button-1>", lambda e, url=template_url: webbrowser.open_new_tab(url))
        self.check_clear_button = ttk.Button(action_frame1, text="Verificar / Limpar Planilha",
                                             bootstyle="outline-danger", command=self.start_check_and_clear_thread)
        self.check_clear_button.pack(side=LEFT, padx=(10, 5))
        self.help_button = ttk.Button(action_frame1, text="Ajuda com Permissões", bootstyle="outline-info",
                                      command=self._on_help_button_click)
        self.help_button.pack(side=LEFT, padx=(10, 10))
        frame1.columnconfigure(1, weight=1)

        # Seção 2
        frame2 = ttk.Labelframe(main_frame, text="2. Fonte dos Novos Contatos (para Sincronização)", padding="10")
        frame2.pack(fill=X, pady=5)
        ttk.Label(frame2, text="Arquivo (.xlsx, .csv):").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.entry_source_file = ttk.Entry(frame2)
        self.entry_source_file.grid(row=0, column=1, sticky=EW, padx=5)
        self.browse_source_button = ttk.Button(frame2, text="Procurar...", bootstyle="info-outline",
                                               command=self._on_browse_source_click)
        self.browse_source_button.grid(row=0, column=2, padx=5)
        info_label = ttk.Label(frame2,
                               text='Atenção: O arquivo de contatos DEVE conter as colunas "NOME" e "EMAIL".',
                               bootstyle="warning", font=("Segoe UI", 8, "italic"), wraplength=500, justify=LEFT)
        info_label.grid(row=1, column=1, columnspan=2, sticky=W, padx=5, pady=(5, 0))
        frame2.columnconfigure(1, weight=1)

        # Seção 3
        frame3 = ttk.Labelframe(main_frame, text="3. Sincronizar Novos Contatos", padding="10")
        frame3.pack(fill=X, pady=5)
        self.analyze_button = ttk.Button(frame3, text="1. Analisar Novos Contatos", bootstyle="primary",
                                         command=self.start_analysis_thread)
        self.analyze_button.pack(pady=10, ipady=5, fill=X)
        results_frame = ttk.Frame(frame3, padding="5")
        results_frame.pack(fill=X, expand=True)
        self.analysis_result_var = tk.StringVar(value="Aguardando análise...")
        analysis_result_label = ttk.Label(results_frame, textvariable=self.analysis_result_var,
                                          font=("Segoe UI", 9, "italic"), wraplength=700)
        analysis_result_label.pack(pady=5)
        preview_frame = ttk.Frame(results_frame)
        preview_frame.pack(fill=BOTH, expand=True, pady=(10, 0))
        columns = ('first_name', 'recipient')
        self.preview_table = ttk.Treeview(preview_frame, columns=columns, show='headings', height=5)
        self.preview_table.heading('first_name', text='Nº. Nome')
        self.preview_table.heading('recipient', text='Recipient')
        self.preview_table.column('first_name', width=250)
        self.preview_table.column('recipient', width=400)
        self.preview_table.bind("<Double-1>", self._open_edit_dialog)
        scrollbar = ttk.Scrollbar(preview_frame, orient=VERTICAL, command=self.preview_table.yview)
        self.preview_table.configure(yscroll=scrollbar.set)
        self.preview_table.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        sync_control_frame = ttk.Frame(results_frame)
        sync_control_frame.pack(pady=10)
        ttk.Label(sync_control_frame, text="Adicionar contatos do:").pack(side=LEFT, padx=(0, 5))
        self.spinbox_start_var = tk.StringVar(value="0")
        self.spinbox_start = ttk.Spinbox(sync_control_frame, from_=0, to=0, width=8,
                                         textvariable=self.spinbox_start_var, state=DISABLED)
        self.spinbox_start.pack(side=LEFT, padx=(0, 5))
        ttk.Label(sync_control_frame, text="até o:").pack(side=LEFT, padx=(0, 5))
        self.spinbox_end_var = tk.StringVar(value="0")
        self.spinbox_end = ttk.Spinbox(sync_control_frame, from_=0, to=0, width=8, textvariable=self.spinbox_end_var,
                                       state=DISABLED)
        self.spinbox_end.pack(side=LEFT, padx=(0, 10))
        self.update_button = ttk.Button(sync_control_frame, text="Atualizar Pré-visualização",
                                        bootstyle="primary-outline", state=DISABLED, command=self.update_preview_table)
        self.update_button.pack(side=LEFT, padx=(0, 20))
        self.sync_button = ttk.Button(sync_control_frame, text="2. SINCRONIZAR AGORA", bootstyle="success",
                                      state=DISABLED, command=self.start_sync_thread)
        self.sync_button.pack(side=LEFT)
        self.dry_run_var = tk.BooleanVar(value=False)
        dry_run_check = ttk.Checkbutton(frame3, text="Modo Simulação", variable=self.dry_run_var,
                                        bootstyle="round-toggle")
        dry_run_check.pack(pady=5)

        # Seção 4
        frame4 = ttk.Labelframe(main_frame, text="4. Status e Log", padding="10")
        frame4.pack(fill=BOTH, expand=True, pady=5)
        self.status_frame = ttk.Frame(frame4)
        self.status_label_var = tk.StringVar(value="Aguardando...")
        status_label = ttk.Label(self.status_frame, textvariable=self.status_label_var)
        status_label.pack(side=LEFT, padx=(0, 10))
        self.progress_bar = ttk.Progressbar(self.status_frame, mode='indeterminate')
        self.progress_bar.pack(side=LEFT, fill=X, expand=True)
        self.log_text = scrolledtext.ScrolledText(frame4, height=5, wrap=WORD, font=("Consolas", 10))
        self.log_text.pack(fill=BOTH, expand=True)
        self.log_text.config(state=DISABLED)
        self.log_text.tag_config("SUCCESS", foreground="#4CAF50")
        self.log_text.tag_config("ERROR", foreground="#F44336")
        self.log_text.tag_config("WARNING", foreground="#FFC107")
        self.log_text.tag_config("INFO", foreground="#2196F3")
        self.log_text.tag_config("DEFAULT", foreground="#d8d8d8")

    def log(self, message: str, level: str = "DEFAULT") -> None:
        try:
            self.log_text.config(state=NORMAL)
            self.log_text.insert(tk.END, f"> {message}\n", (level.upper(),))
            self.log_text.config(state=DISABLED)
            self.log_text.see(tk.END)
        except tk.TclError:
            pass

    def process_queue(self) -> None:
        try:
            while True:
                msg_type, data = self.queue.get_nowait()
                if msg_type == "log":
                    self.log(data[0], data[1])
                elif msg_type == "progress_start":
                    self.status_label_var.set(data)
                    self.status_frame.pack(fill=X, pady=(0, 5), before=self.log_text)
                    self.progress_bar.start(10)
                elif msg_type == "progress_stop":
                    self.progress_bar.stop()
                    self.status_frame.pack_forget()
                elif msg_type == "buttons_state":
                    self.set_buttons_state(data)
                elif msg_type == "update_analysis":
                    df, result_text, spin_config, sync_config = data
                    if df is not None: df.reset_index(drop=True, inplace=True)
                    self.global_new_contacts_df = df
                    self.update_analysis_results(result_text, spin_config, sync_config)
                    if df is not None: self.populate_preview_table(df)
                elif msg_type == "permission_error":
                    self.show_permission_error_dialog(data)
        except Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

    def _on_browse_json_click(self) -> None:
        filename = filedialog.askopenfilename(title="Selecione o arquivo JSON da Conta de Serviço",
                                              filetypes=[("JSON files", "*.json")])
        if filename:
            self.entry_json.delete(0, tk.END)
            self.entry_json.insert(0, filename)
            self.log(f"Arquivo de chave selecionado: {os.path.basename(filename)}", "INFO")

    def _on_browse_source_click(self) -> None:
        filename = filedialog.askopenfilename(title="Selecione o arquivo de contatos",
                                              filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if filename:
            self.entry_source_file.delete(0, tk.END)
            self.entry_source_file.insert(0, filename)
            self.log(f"Arquivo de contatos selecionado: {os.path.basename(filename)}", "INFO")
            self.log("Validando cabeçalhos do arquivo...", "INFO")
            app_cfg = self.config_data.get("app_settings", {})
            threading.Thread(target=logic.validate_source_file_headers_thread, args=(
                filename, app_cfg.get("possible_name_cols", []), app_cfg.get("possible_email_cols", []), self.queue
            ), daemon=True).start()

    def start_check_and_clear_thread(self) -> None:
        # Pega a URL do combobox agora
        mailmerge_url = self.mailmerge_url_combobox.get()
        threading.Thread(target=logic.check_and_clear_sheet_thread,
                         args=(self.entry_json.get(), mailmerge_url, self.queue), daemon=True).start()

    def start_analysis_thread(self) -> None:
        self.populate_preview_table(None)
        app_cfg = self.config_data.get("app_settings", {})
        # Pega a URL do combobox agora
        mailmerge_url = self.mailmerge_url_combobox.get()
        threading.Thread(target=logic.analyze_data_thread, args=(
            self.entry_json.get(), mailmerge_url, self.entry_source_file.get(),
            app_cfg.get("possible_name_cols", []), app_cfg.get("possible_email_cols", []), self.queue
        ), daemon=True).start()

    def start_sync_thread(self) -> None:
        try:
            start_val, end_val = int(self.spinbox_start_var.get()), int(self.spinbox_end_var.get())
            if start_val > end_val:
                self.log("O número inicial não pode ser maior que o final.", "ERROR");
                return
            contacts_to_sync = self.global_new_contacts_df.iloc[start_val - 1:end_val]
            if not self.dry_run_var.get():
                if not messagebox.askyesno("Confirmar Sincronização",
                                           f"Você tem certeza que deseja adicionar {len(contacts_to_sync)} contatos à planilha?"):
                    self.log("Sincronização cancelada pelo usuário.", "WARNING");
                    return
            # Pega a URL do combobox agora
            mailmerge_url = self.mailmerge_url_combobox.get()
            threading.Thread(target=logic.sync_data_thread, args=(
                self.entry_json.get(), mailmerge_url, self.dry_run_var.get(), contacts_to_sync,
                self.queue
            ), daemon=True).start()
        except (ValueError, TypeError):
            self.log("Valores de intervalo inválidos para sincronização.", "ERROR")
        except Exception as e:
            self.log(f"Erro ao preparar sincronização: {e}", "ERROR")

    def _on_help_button_click(self) -> None:
        json_path = self.entry_json.get()
        service_account_email = "[Selecione um arquivo JSON para ver o e-mail]"
        if json_path and os.path.exists(json_path):
            try:
                with open(json_path, 'r') as f:
                    sa_info = json.load(f)
                service_account_email = sa_info.get('client_email', service_account_email)
            except Exception as e:
                self.log(f"AVISO: Erro ao ler o arquivo JSON: {e}", "WARNING")
        HelpWindow(self.root, service_account_email)

    def show_permission_error_dialog(self, service_account_email: str) -> None:
        if service_account_email:
            HelpWindow(self.root, service_account_email)
        else:
            messagebox.showerror("Erro de Permissão",
                                 "Permissão negada. Verifique se o arquivo de chave JSON é válido.")

    def _open_edit_dialog(self, event: Any) -> None:
        selected_item_id_str = self.preview_table.focus()
        if not selected_item_id_str: return
        try:
            item_id = int(selected_item_id_str)
            contact_data = self.global_new_contacts_df.loc[item_id]
            EditContactWindow(self.root, item_id, contact_data, self._save_edited_contact)
        except (ValueError, KeyError):
            self.log(f"Não foi possível encontrar os dados para o item selecionado.", "ERROR")

    def _save_edited_contact(self, item_id: int, new_data: Dict[str, str]) -> None:
        try:
            self.global_new_contacts_df.loc[item_id, 'First name'] = new_data['name']
            self.global_new_contacts_df.loc[item_id, 'Recipient'] = new_data['email']
            self.populate_preview_table(self.global_new_contacts_df)
            self.log(f"Contato Nº {item_id + 1} atualizado.", "SUCCESS")
        except Exception as e:
            self.log(f"Erro ao salvar a edição do contato: {e}", "ERROR")

    def populate_preview_table(self, dataframe: Optional[pd.DataFrame]) -> None:
        for item in self.preview_table.get_children(): self.preview_table.delete(item)
        if dataframe is not None:
            for index, row in dataframe.iterrows():
                display_name = f"{index + 1}. {row['First name']}"
                self.preview_table.insert('', END, iid=str(index), values=(display_name, row['Recipient']))

    def update_preview_table(self) -> None:
        if self.global_new_contacts_df is None or self.global_new_contacts_df.empty: return
        try:
            start_val, end_val = int(self.spinbox_start_var.get()), int(self.spinbox_end_var.get())
            if start_val > end_val:
                self.log("ERRO: O número inicial não pode ser maior que o final.", "ERROR");
                return
            df_slice = self.global_new_contacts_df.iloc[start_val - 1:end_val]
            self.populate_preview_table(df_slice)
            self.log(f"Pré-visualização atualizada para mostrar contatos de {start_val} a {end_val}.", "INFO")
        except (ValueError, TypeError):
            self.log("ERRO: Valores inválidos para o intervalo de pré-visualização.", "ERROR")

    def set_buttons_state(self, state: str) -> None:
        for btn in [self.check_clear_button, self.analyze_button, self.sync_button, self.browse_json_button,
                    self.browse_source_button, self.update_button, self.help_button, self.save_mailmerge_link_button]:
            if btn.winfo_exists(): btn.config(state=state)

    def update_analysis_results(self, result_text: str, spinbox_config: Dict[str, Any],
                                sync_button_config: Dict[str, Any]) -> None:
        self.analysis_result_var.set(result_text)
        self.spinbox_start.config(state=spinbox_config.get("state", "disabled"), from_=spinbox_config.get("from_", 0),
                                  to=spinbox_config.get("to", 0))
        self.spinbox_end.config(state=spinbox_config.get("state", "disabled"), from_=spinbox_config.get("from_", 0),
                                to=spinbox_config.get("to", 0))
        if "start_value" in spinbox_config: self.spinbox_start_var.set(str(spinbox_config.get("start_value")))
        if "end_value" in spinbox_config: self.spinbox_end_var.set(str(spinbox_config.get("end_value")))
        sync_state = sync_button_config.get("state", "disabled")
        self.sync_button.config(state=sync_state)
        self.update_button.config(state=sync_state)

    def clear_all_information(self) -> None:
        """Limpa todas as informações de configuração e pré-visualização do app, exceto os links salvos."""
        if messagebox.askyesno("Confirmar Limpeza",
                               "Você tem certeza que deseja limpar TODAS as informações (arquivos, URLs atualmente selecionada e pré-visualização)? Os links de planilha SALVOS não serão removidos."):
            # Limpa os campos de entrada e a URL atual no combobox
            self.entry_json.delete(0, tk.END)
            self.mailmerge_url_var.set("")  # Limpa o texto atual no Combobox
            # self.mailmerge_url_combobox['values'] = () # <-- REMOVIDO: Não zera os valores salvos
            self.entry_source_file.delete(0, tk.END)

            # Limpa as variáveis de configuração em memória, exceto a lista de URLs salvas
            self.config_data["user_settings"]["json_path"] = ""
            self.config_data["user_settings"]["mailmerge_url"] = ""  # Limpa a última URL usada
            self.config_data["user_settings"]["source_file"] = ""
            # self.config_data["user_settings"]["saved_mailmerge_urls"]
            self.save_config()  # Salva as configurações vazias, mas mantém saved_mailmerge_urls

            # Limpa os resultados da análise e pré-visualização
            self.global_new_contacts_df = None
            self.analysis_result_var.set("Aguardando análise...")
            self.populate_preview_table(None)

            # Desativa os controles de sincronização/pré-visualização
            self.spinbox_start.config(state=DISABLED)
            self.spinbox_end.config(state=DISABLED)
            self.spinbox_start_var.set("0")
            self.spinbox_end_var.set("0")
            self.sync_button.config(state=DISABLED)
            self.update_button.config(state=DISABLED)

            self.log("Todas as informações do aplicativo (exceto links salvos) foram limpas.", "INFO")
        else:
            self.log("Limpeza de informações cancelada.", "WARNING")

    def _on_save_mailmerge_link(self) -> None:
        """Salva a URL atual do MailMerge para uso futuro."""
        current_url = self.mailmerge_url_combobox.get().strip()
        if not current_url:
            self.log("A URL da planilha está vazia. Não é possível salvar.", "WARNING")
            messagebox.showwarning("URL Vazia", "Por favor, insira uma URL antes de tentar salvar.")
            return

        saved_urls = list(self.mailmerge_url_combobox['values'])  # Pega os valores atuais do combobox
        if current_url not in saved_urls:
            saved_urls.insert(0, current_url)  # Adiciona ao início da lista
            self.mailmerge_url_combobox['values'] = tuple(saved_urls)  # Atualiza as opções do combobox
            self.save_config()  # Salva as configurações com a nova URL
            self.log(f"Link da planilha '{current_url}' salvo com sucesso!", "SUCCESS")
            messagebox.showinfo("Link Salvo", "O link da planilha foi salvo para acesso rápido futuramente.")
        else:
            self.log(f"Link da planilha '{current_url}' já está salvo.", "INFO")
            messagebox.showinfo("Link Já Existe", "Este link já está salvo na sua lista.")