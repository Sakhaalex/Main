import os
from datetime import datetime

class CommCentre:
    """Singleton class for managing System State and Paths."""
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(CommCentre, cls).__new__(cls)
            cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        self.PCB_ID = ""
        self.tester_name = ""
        self.reviewed_by = ""
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.date = datetime.now().strftime("%Y-%m-%d")
        
        # P0 through P9 path variables
        self.paths = {f"P{i}": "" for i in range(10)}

    def update_pcb_id(self, pcb_id: str):
        self.PCB_ID = pcb_id
        
    def update_credentials(self, tester_name: str, reviewed_by: str):
        self.tester_name = tester_name
        self.reviewed_by = reviewed_by
        
    def set_path(self, index: int, path: str):
        if 0 <= index <= 9:
            self.paths[f"P{index}"] = path
            
    def resolve_path(self, path_str: str) -> str:
        """Resolves placeholders like P0/MSO or PCB_ID/Results.csv."""
        if not path_str:
            return ""
        
        resolved = path_str
        
        if "PCB_ID" in resolved and self.PCB_ID:
            resolved = resolved.replace("PCB_ID", self.PCB_ID)
            
        for p_key, p_val in self.paths.items():
            if p_val and p_key in resolved:
                resolved = resolved.replace(f"{p_key}/", f"{p_val}/")
                resolved = resolved.replace(p_key, p_val)
                
        return os.path.normpath(resolved)

# Global instance
LX_CommCentre = CommCentre()
