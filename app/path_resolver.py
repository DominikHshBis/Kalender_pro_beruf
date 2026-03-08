from pathlib import Path
import os

class PathResolver:
    """Zentralisiert Pfad-Verwaltung für das Projekt."""
    
    def __init__(self, project_root: Path):
        self.project_root = project_root
    
    def find_file(self, filename: str) -> Path | None:
        """Sucht Datei rekursiv im Projekt."""
        for path in self.project_root.rglob(filename):
            return path
        return None
    
    def env_path(self, var_name: str, default: str) -> Path:
        """Nimmt ENV-Wert oder Default-Pfad."""
        value = os.getenv(var_name)
        if value:
            return self.project_root / value
        found = self.find_file(default)
        return found if found else self.project_root / default  # Fallback
    
    @property
    def config_path(self) -> Path:
        return self.env_path("CONFIG_PATH", "config.json")
    
    @property
    def service_account_file(self) -> Path:
        return self.env_path("SERVICE_ACCOUNT_FILE", "credentials.json")
    
    @property
    def excel_load_path(self) -> Path:
        return self.env_path("EXCEL_LOAD_PATH", "Muster_Honorarrechnung-Lehrkräfte_pytest.xlsx")
    
    @property
    def excel_load_path_ho(self) -> Path:
        return self.env_path("EXCEL_LOAD_PATH_HO", "Homeoffice_zaehler.xlsx")
    
    @property
    def output_dir(self) -> Path:
        output_dir = Path(os.getenv("OUTPUT_DIR", self.project_root / "output"))
        output_dir.mkdir(exist_ok=True)
        return output_dir