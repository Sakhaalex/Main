import os
from LX_BaseFrame import LX_BaseFrame, LXLockState
from LX_CommCentre import LX_CommCentre

class F_IMGR(LX_BaseFrame):
    def __init__(self):
        super().__init__(alias="IMG")

    def execute(self, row_data: dict) -> bool:
        if str(row_data.get("LX", "")).upper() == "X":
            self.set_lock_status(LXLockState.SKIP)
            return True

        self.set_lock_status(LXLockState.BUSY)
        log = row_data.get("_log_callback", print)
        log(f"[{self.alias}] External Lock ENGAGED. Populating Images...")

        try:
            source_path = LX_CommCentre.resolve_path(row_data.get("S_Path", ""))
            source_file = row_data.get("S_File_Name", "")
            image_full_path = os.path.join(source_path, source_file)

            dest_sheet = row_data.get("D_Sheet", "Sheet1")
            dest_start_cell = row_data.get("D_Start_Cell_Range", "")

            ws = row_data.get("_ws_dict").get(dest_sheet)

            if not os.path.exists(image_full_path):
                log(f"[{self.alias}] WARNING: Image not found at {image_full_path}")
            else:
                ws.insert_image(dest_start_cell, image_full_path, 
                                {'x_scale': 0.0938, 'y_scale': 0.0926})
                log(f"[{self.alias}] Inserted {source_file} at {dest_start_cell}")

            self.set_lock_status(LXLockState.FREE)
            return True

        except Exception as err:
            log(f"[{self.alias}] CRITICAL ERROR: {err}")
            self.set_lock_status(LXLockState.ERROR)
            return False
