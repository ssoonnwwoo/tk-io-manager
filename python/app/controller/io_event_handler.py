from PySide6.QtWidgets import QFileDialog, QTableWidget, QMessageBox
import os

def select_xlsx_file(line_edit_widget):
    # default_dir = os.path.join(os.path.expanduser("~"), "show")
    current_path = line_edit_widget.text()
    if not os.path.isdir(current_path):
        return None

    file_path, _ = QFileDialog.getOpenFileName(
        None,  # 부모 위젯 없음
        "Select XLSX file",
        current_path,
        "XLSX files (*.xlsx)"
    )
    if file_path:
        if not os.path.abspath(file_path).startswith(os.path.abspath(current_path)):
            QMessageBox.warning(None, "Warning", f"'{date_path}' Out of boundary")
            return None
    # line_edit_widget.setText(os.path.dirname(file_path))
    return file_path

def toggle_edit_mode(table, edit_mode):
    if edit_mode:
        table.setEditTriggers(QTableWidget.NoEditTriggers)
    else:
        table.setEditTriggers(QTableWidget.AllEditTriggers)
    return not edit_mode