import tkinter as tk
from tkinter import simpledialog, messagebox

class StoreFormDialog:
    """Diálogo para adicionar/editar filial"""
    
    def __init__(self, parent, title, store_code="", admins=None, coordinators=None):
        self.result = None
        self.admins_list = admins if admins is not None else []
        self.coords_list = coordinators if coordinators is not None else []
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("600x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Código da filial
        tk.Label(self.dialog, text="Número da Filial:").pack(pady=(10, 0))
        self.store_entry = tk.Entry(self.dialog, width=20)
        self.store_entry.pack(pady=5)
        
        if store_code:
            self.store_entry.insert(0, store_code)
            self.store_entry.config(state="readonly")
        
        # E-mails de Administradores
        tk.Label(self.dialog, text="E-mails dos Administradores:").pack(pady=(20, 0))
        
        admin_frame = tk.Frame(self.dialog)
        admin_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.admin_listbox = tk.Listbox(admin_frame, height=6)
        self.admin_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        admin_scroll = tk.Scrollbar(admin_frame, command=self.admin_listbox.yview)
        admin_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.admin_listbox.config(yscrollcommand=admin_scroll.set)
        
        for email in self.admins_list:
            self.admin_listbox.insert(tk.END, email)
        
        admin_btn_frame = tk.Frame(self.dialog)
        admin_btn_frame.pack()
        
        tk.Button(admin_btn_frame, text="Adicionar", command=self._add_admin, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(admin_btn_frame, text="Remover", command=self._remove_admin, width=15).pack(side=tk.LEFT, padx=5)
        
        # E-mails de Coordenadores
        tk.Label(self.dialog, text="E-mails dos Coordenadores:").pack(pady=(20, 0))
        
        coord_frame = tk.Frame(self.dialog)
        coord_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.coord_listbox = tk.Listbox(coord_frame, height=6)
        self.coord_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        coord_scroll = tk.Scrollbar(coord_frame, command=self.coord_listbox.yview)
        coord_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.coord_listbox.config(yscrollcommand=coord_scroll.set)
        
        for email in self.coords_list:
            self.coord_listbox.insert(tk.END, email)
        
        coord_btn_frame = tk.Frame(self.dialog)
        coord_btn_frame.pack()
        
        tk.Button(coord_btn_frame, text="Adicionar", command=self._add_coord, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(coord_btn_frame, text="Remover", command=self._remove_coord, width=15).pack(side=tk.LEFT, padx=5)
        
        # Botões de ação
        action_frame = tk.Frame(self.dialog)
        action_frame.pack(pady=20)
        
        tk.Button(action_frame, text="Salvar", command=self._save, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(action_frame, text="Cancelar", command=self.dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)
        
        self.dialog.wait_window()
    
    def _add_admin(self):
        """Adiciona e-mail de administrador"""
        email = simpledialog.askstring("Adicionar E-mail", "Digite o e-mail do administrador:")
        if email:
            self.admin_listbox.insert(tk.END, email)
    
    def _remove_admin(self):
        """Remove e-mail de administrador"""
        selection = self.admin_listbox.curselection()
        if selection:
            self.admin_listbox.delete(selection[0])
    
    def _add_coord(self):
        """Adiciona e-mail de coordenador"""
        email = simpledialog.askstring("Adicionar E-mail", "Digite o e-mail do coordenador:")
        if email:
            self.coord_listbox.insert(tk.END, email)
    
    def _remove_coord(self):
        """Remove e-mail de coordenador"""
        selection = self.coord_listbox.curselection()
        if selection:
            self.coord_listbox.delete(selection[0])
    
    def _save(self):
        """Salva os dados"""
        store_code = self.store_entry.get().strip()
        
        if not store_code:
            messagebox.showerror("Erro", "Número da filial é obrigatório!")
            return
        
        admins = list(self.admin_listbox.get(0, tk.END))
        coords = list(self.coord_listbox.get(0, tk.END))
        
        self.result = {
            "store_code": store_code,
            "admins": admins,
            "coordinators": coords
        }
        
        self.dialog.destroy()
