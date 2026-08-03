from LX_BaseFrame import LX_BaseFrame, LXLockState
from LX_CommCentre import LX_CommCentre

class F_HDCREP(LX_BaseFrame):
    def __init__(self):
        super().__init__(alias="HDC")

    def execute(self, row_data: dict) -> bool:
        if str(row_data.get("LX", "")).upper() == "X":
            self.set_lock_status(LXLockState.SKIP)
            return True

        self.set_lock_status(LXLockState.BUSY)
        log = row_data.get("_log_callback", print)
        log(f"[{self.alias}] External Lock ENGAGED. Injecting Header Data...")

        try:
            dest_sheet_name = row_data.get("D_Sheet", "Sheet1")
            ws = row_data.get("_ws_dict").get(dest_sheet_name)

            variable_name = row_data.get("M", "")
            dest_start_cell = row_data.get("D_Start_Cell_Range", "")

            value_to_inject = ""
            if variable_name.lower() == "tester_name":
                value_to_inject = LX_CommCentre.tester_name
            elif variable_name.lower() == "reviewed_by":
                value_to_inject = LX_CommCentre.reviewed_by
            elif variable_name.lower() == "timestamp":
                value_to_inject = LX_CommCentre.timestamp
            elif variable_name.lower() == "date":
                value_to_inject = LX_CommCentre.date
            elif variable_name.lower() == "pcb_id":
                value_to_inject = LX_CommCentre.PCB_ID
            else:
                value_to_inject = variable_name

            if dest_start_cell:
                ws.write(dest_start_cell, value_to_inject)
                log(f"[{self.alias}] Injected {variable_name} ('{value_to_inject}') at {dest_start_cell}")

            self.set_lock_status(LXLockState.FREE)
            return True

        except Exception as err:
            log(f"[{self.alias}] CRITICAL ERROR: {err}")
            self.set_lock_status(LXLockState.ERROR)
            return False
