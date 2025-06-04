import pandas as pd
from tank.platform.qt import QtGui
def on_checkbox_clicked(iomanager_instance, row, xlsx_path):
    df = pd.read_excel(xlsx_path)

    seq = df.at[row, "seq_name"]
    shot = df.at[row, "shot_name"]

    if pd.isna(seq) or pd.isna(shot) or str(seq).strip() == "" or str(shot).strip() == "":
        QtGui.QMessageBox.warning(
        iomanager_instance,
        "Missing Data",
        f"Row[{row + 1}] missing 'seq_name' or 'shot_name'.\nThis row cannot be selected.",
        QtGui.QMessageBox.Ok
        )

        checkbox = iomanager_instance.table.cellWidget(row, 0)
        if isinstance(checkbox, QtGui.QCheckBox):
            checkbox.setChecked(False)
