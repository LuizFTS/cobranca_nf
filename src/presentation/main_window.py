import tkinter as tk
import os
import shutil
import tempfile
from tkinter import ttk, messagebox, filedialog

from src.presentation.store_dialog import StoreFormDialog
from src.application.spreadsheet_service import SpreadsheetService
from src.infrastructure.config_manager import EmailConfigManager
from src.infrastructure.email_sender import SendEmail

class StoreEmailConfigUI:
    """Interface para gerenciar e-mails por filial"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Parâmetros - Filiais cadastradas")
        self.root.geometry("900x600")
        
        self.config_manager = EmailConfigManager()
        self.create_spreadsheet = SpreadsheetService()
        self.email_sender = SendEmail()
        
        self._create_widgets()
        self._load_stores()
    
    def _create_widgets(self):
        """Cria os widgets da interface"""
        
        # Frame principal
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Treeview para listar filiais
        columns = ("Número", "Email Adm", "Email Coord")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=20)
        
        # Configurar colunas
        self.tree.heading("Número", text="Número")
        self.tree.heading("Email Adm", text="Email Adm")
        self.tree.heading("Email Coord", text="Email Coord")
        
        self.tree.column("Número", width=100, anchor=tk.CENTER)
        self.tree.column("Email Adm", width=350)
        self.tree.column("Email Coord", width=350)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame de botões
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        btn_add = tk.Button(button_frame, text="Adicionar Filial", command=self._add_store, width=20)
        btn_add.pack(side=tk.LEFT, padx=5)
        
        btn_edit = tk.Button(button_frame, text="Editar Filial", command=self._edit_store, width=20)
        btn_edit.pack(side=tk.LEFT, padx=5)
        
        btn_delete = tk.Button(button_frame, text="Excluir Filial", command=self._delete_store, width=20)
        btn_delete.pack(side=tk.LEFT, padx=5)
        
        # Separador
        tk.Frame(button_frame, width=20).pack(side=tk.LEFT)
        
        btn_send = tk.Button(button_frame, text="Enviar E-mails", command=self._send_emails, width=20, bg="#4CAF50", fg="white")
        btn_send.pack(side=tk.LEFT, padx=5)
    
    def _load_stores(self):
        """Carrega as filiais na tabela"""
        # Limpa a tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Carrega filiais do config
        stores = self.config_manager.get_all_stores()
        
        # Ordena por número da filial
        sorted_stores = sorted(stores.items(), key=lambda x: x[0])
        
        for store_code, store_data in sorted_stores:
            admins = ", ".join(store_data.get("admins", []))
            coords = ", ".join(store_data.get("coordinators", []))
            
            self.tree.insert("", tk.END, values=(store_code, admins, coords))
    
    def _add_store(self):
        """Adiciona uma nova filial"""
        dialog = StoreFormDialog(self.root, "Adicionar Filial")
        
        if dialog.result:
            store_code = dialog.result["store_code"]
            admins = dialog.result["admins"]
            coords = dialog.result["coordinators"]
            
            # Verifica se já existe
            if store_code in self.config_manager.get_all_stores():
                messagebox.showerror("Erro", f"Filial {store_code} já cadastrada!")
                return
            
            self.config_manager.add_store(store_code, admins, coords)
            self._load_stores()
            messagebox.showinfo("Sucesso", "Filial adicionada com sucesso!")
    
    def _edit_store(self):
        """Edita uma filial existente"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma filial para editar!")
            return
        
        item = self.tree.item(selection[0])
        store_code = str(item["values"][0])
        # Garante que números de 1 dígito tenham o zero à esquerda para bater com o JSON
        if store_code.isdigit() and len(store_code) < 2:
            store_code = store_code.zfill(2)
        
        # Carrega dados atuais
        store_data = self.config_manager.get_store(store_code)
        
        dialog = StoreFormDialog(
            self.root, 
            "Editar Filial",
            store_code=store_code,
            admins=store_data.get("admins", []),
            coordinators=store_data.get("coordinators", [])
        )
        
        if dialog.result:
            admins = dialog.result["admins"]
            coords = dialog.result["coordinators"]
            
            self.config_manager.update_store(store_code, admins, coords)
            self._load_stores()
            messagebox.showinfo("Sucesso", "Filial atualizada com sucesso!")
    
    def _delete_store(self):
        """Exclui uma filial"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma filial para excluir!")
            return
        
        item = self.tree.item(selection[0])
        store_code = str(item["values"][0])
        if store_code.isdigit() and len(store_code) < 2:
            store_code = store_code.zfill(2)
        
        confirm = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Deseja realmente excluir a filial {store_code}?"
        )
        
        if confirm:
            self.config_manager.delete_store(store_code)
            self._load_stores()
            messagebox.showinfo("Sucesso", "Filial excluída com sucesso!")
    
    def _send_emails(self):
        """Executa o script main.py ou o executável CobrancaNF.exe para enviar e-mails"""
        
        file = filedialog.askopenfilename(
            title="Selecione o arquivo com as notas fiscais e fretes",
            filetypes=[("Arquivos CSV", "*.csv")]
        )

        if file:
            try:
                # cria diretório temporário
                temp_dir = tempfile.gettempdir()
                
                # mantém o mesmo nome do arquivo
                file_name = os.path.basename(file)
                temp_file_path = os.path.join(temp_dir, file_name)

                # copia o arquivo selecionado para o temp
                shutil.copy2(file, temp_file_path)

            except Exception as e:
                print("Erro ao processar o arquivo:", e)
                return

            messagebox.showinfo("Atenção!", "Selecione a pasta em que será gerado as planilhas individuais de cada loja.")

            output_folder = filedialog.askdirectory(
                title="Selecione a pasta"
            )

            try:
                reports = self.create_spreadsheet.execute(temp_file_path, output_folder)

                for report in reports:
                    self.email_sender.execute(
                        report.branch,
                        report.period_initial,
                        report.period_final,
                        self.email_sender.get_admins(report.branch),
                        self.email_sender.get_coordinators(report.branch),
                        report.excel_path,
                        report.quantity,
                        report.total_value,
                        report.table
                    )
                messagebox.showinfo("Sucesso!", "E-mails enviados com sucesso!")
            except Exception as e:
                print("Erro ao processar o arquivo:", e)
                messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")
