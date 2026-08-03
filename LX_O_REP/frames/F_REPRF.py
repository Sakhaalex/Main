from LX_BaseFrame import LX_BaseFrame, LXLockState

class F_REPRF(LX_BaseFrame):
    def __init__(self):
        super().__init__(alias="VAL")

    def execute(self, row_data: dict) -> bool:
        if str(row_data.get("LX", "")).upper() == "X":
            self.set_lock_status(LXLockState.SKIP)
            return True

        self.set_lock_status(LXLockState.BUSY)
        log = row_data.get("_log_callback", print)
        log(f"[{self.alias}] External Lock ENGAGED. Writing Values/Text...")

        try:
            dest_sheet_name = row_data.get("D_Sheet", "Sheet1")
            ws = row_data.get("_ws_dict").get(dest_sheet_name)

            raw_value = row_data.get("M", "")
            dest_start_cell = row_data.get("D_Start_Cell_Range", "")
            row_type = row_data.get("Type", "").upper()

            if dest_start_cell:
                if row_type == "VAL":
                    parsed_val = raw_value
                    try:
                        if "." in str(raw_value):
                            parsed_val = float(raw_value)
                        else:
                            parsed_val = int(raw_value)
                    except ValueError:
                        pass
                    ws.write(dest_start_cell, parsed_val)
                elif row_type == "TXT":
                    ws.write(dest_start_cell, str(raw_value))
                    
                log(f"[{self.alias}] Wrote {row_type} '{raw_value}' at {dest_start_cell}")

            self.set_lock_status(LXLockState.FREE)
            return True

        except Exception as err:
            log(f"[{self.alias}] CRITICAL ERROR: {err}")
            self.set_lock_status(LXLockState.ERROR)
            return False
