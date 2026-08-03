from LX_BaseFrame import LX_BaseFrame, LXLockState

class F_REPT(LX_BaseFrame):
    def __init__(self):
        super().__init__(alias="STR")

    def execute(self, row_data: dict) -> bool:
        if str(row_data.get("LX", "")).upper() == "X":
            self.set_lock_status(LXLockState.SKIP)
            return True

        self.set_lock_status(LXLockState.BUSY)
        log = row_data.get("_log_callback", print)
        log(f"[{self.alias}] External Lock ENGAGED. Writing String Data...")

        try:
            dest_sheet_name = row_data.get("D_Sheet", "Sheet1")
            ws = row_data.get("_ws_dict").get(dest_sheet_name)

            raw_value = row_data.get("M", "")
            dest_start_cell = row_data.get("D_Start_Cell_Range", "")
            
            if dest_start_cell:
                ws.write_string(dest_start_cell, str(raw_value))
                log(f"[{self.alias}] Wrote string '{raw_value}' at {dest_start_cell}")

            self.set_lock_status(LXLockState.FREE)
            return True

        except Exception as err:
            log(f"[{self.alias}] CRITICAL ERROR: {err}")
            self.set_lock_status(LXLockState.ERROR)
            return False
