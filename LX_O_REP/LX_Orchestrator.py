import sys
import os
# Ensure the script's own directory is on sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import csv
import time
from collections import defaultdict
import xlsxwriter

from LX_CommCentre import LX_CommCentre
from LX_BaseFrame import LXLockState
from frames.F_IMGR import F_IMGR
from frames.F_MIR import F_MIR
from frames.F_HDCREP import F_HDCREP
from frames.F_REPRF import F_REPRF
from frames.F_REPT import F_REPT

class Orchestrator:
    def __init__(self, log_callback=None):
        self.registry = {
            "IMG": F_IMGR(),
            "MIR": F_MIR(),
            "HDC": F_HDCREP(),
            "VAL": F_REPRF(),
            "TXT": F_REPRF(), 
            "STR": F_REPT()
        }
        self.log_callback = log_callback

    def log(self, msg):
        print(msg)
        if self.log_callback:
            self.log_callback(msg)

    def run_sequence(self, source_of_truth_file: str):
        if not os.path.exists(source_of_truth_file):
            self.log(f"[O_Main] Source of Truth file not found: {source_of_truth_file}")
            return

        rows = []
        with open(source_of_truth_file, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(row)

        file_groups = defaultdict(list)
        for row in rows:
            if str(row.get("LX", "")).upper() == "X":
                continue
                
            dest_path = LX_CommCentre.resolve_path(row.get("D_Path", ""))
            dest_file = row.get("D_File_Name", "")
            if not dest_file:
                continue
                
            full_dest = os.path.join(dest_path, dest_file)
            file_groups[full_dest].append(row)

        for dest_file, file_rows in file_groups.items():
            self.log(f"\n[O_Main] Processing Report: {dest_file}")
            os.makedirs(os.path.dirname(dest_file), exist_ok=True)
            
            # Using xlsxwriter exclusively for all output generation
            wb_xlsx = xlsxwriter.Workbook(dest_file)
            ws_dict = {}

            # Execution order: IMG -> MIR -> HDC -> VAL/TXT/STR
            frame_order = ["IMG", "MIR", "HDC", "VAL", "TXT", "STR"]
            
            for frame_type in frame_order:
                type_rows = [r for r in file_rows if r.get("Type", "").upper() == frame_type]
                for row in type_rows:
                    dest_sheet = row.get("D_Sheet", "Sheet1")
                    if dest_sheet not in ws_dict:
                        ws_dict[dest_sheet] = wb_xlsx.add_worksheet(dest_sheet)
                    
                    row["_wb"] = wb_xlsx
                    row["_ws_dict"] = ws_dict
                    row["_log_callback"] = self.log
                    self._execute_row(row)

            wb_xlsx.close()
            self.log(f"[O_Main] Finished Report: {dest_file}")

    def _execute_row(self, row: dict):
        frame_type = row.get("Type", "").upper()
        if frame_type not in self.registry:
            self.log(f"[O_Main] Unknown Frame Type '{frame_type}'. Skipping row.")
            return

        frame = self.registry[frame_type]
        success = frame.execute(row)

        while True:
            status = frame.get_lock_status()
            if status == LXLockState.FREE or status == LXLockState.SKIP:
                break
            elif status == LXLockState.BUSY:
                time.sleep(0.05)
            elif status == LXLockState.ERROR:
                self.log(f"[O_Main] CRITICAL ERROR on Row with Type {frame_type}. Halting!")
                return
