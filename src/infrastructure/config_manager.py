import json
from pathlib import Path
from typing import Dict, List

class EmailConfigManager:
    """Gerenciador de configurações de e-mail por filial armazenadas em JSON"""
    
    def __init__(self):
        """Inicializa o gerenciador e garante que o arquivo de configuração existe"""
        self.config_dir = Path.home() / "Documents" / "CobrancaNF"
        self.config_file = self.config_dir / "email_config.json"
        self._ensure_config_exists()
    
    def _ensure_config_exists(self):
        """Cria o diretório e arquivo de configuração se não existirem"""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        
        if not self.config_file.exists():
            default_config = {
                "stores": {}
            }
            self._save_config(default_config)
    
    def _load_config(self) -> Dict:
        """Carrega a configuração do arquivo JSON"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Erro ao carregar configuração: {e}")
            return {"stores": {}}
    
    def _save_config(self, config: Dict):
        """Salva a configuração no arquivo JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar configuração: {e}")
            raise
    
    def get_all_stores(self) -> Dict[str, Dict]:
        """Retorna todas as filiais configuradas"""
        config = self._load_config()
        return config.get("stores", {})
    
    def get_store(self, store_code: str) -> Dict:
        """Retorna configuração de uma filial específica"""
        config = self._load_config()
        return config.get("stores", {}).get(store_code, {"admins": [], "coordinators": []})
    
    def add_store(self, store_code: str, admins: List[str] = None, coordinators: List[str] = None):
        """Adiciona uma nova filial"""
        config = self._load_config()
        if "stores" not in config:
            config["stores"] = {}
        
        config["stores"][store_code] = {
            "admins": admins or [],
            "coordinators": coordinators or []
        }
        self._save_config(config)
    
    def update_store(self, store_code: str, admins: List[str], coordinators: List[str]):
        """Atualiza os e-mails de uma filial"""
        config = self._load_config()
        if "stores" not in config:
            config["stores"] = {}
        
        config["stores"][store_code] = {
            "admins": admins,
            "coordinators": coordinators
        }
        self._save_config(config)
    
    def delete_store(self, store_code: str):
        """Remove uma filial"""
        config = self._load_config()
        if store_code in config.get("stores", {}):
            del config["stores"][store_code]
            self._save_config(config)
    
    def get_store_admins(self, store_code: str) -> List[str]:
        """Retorna lista de e-mails dos administradores de uma filial"""
        store = self.get_store(store_code)
        return store.get("admins", [])
    
    def get_store_coordinators(self, store_code: str) -> List[str]:
        """Retorna lista de e-mails dos coordenadores de uma filial"""
        store = self.get_store(store_code)
        return store.get("coordinators", [])
    
    def get_config_path(self) -> str:
        """Retorna o caminho do arquivo de configuração"""
        return str(self.config_file)
