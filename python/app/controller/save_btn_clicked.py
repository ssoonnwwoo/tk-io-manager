from tank.platform.qt import QtGui
import os

def on_save_clicked(iomanager_instance):
    new_xl = iomanager_instance.excel_manager.get_new_xl()
    if not new_xl:
        return

    reply = QtGui.QMessageBox.question(
        iomanager_instance,
        "Confirm Save",
        f"The file will be saved as\n{os.path.basename(new_xl)}\nDo you want to continue?",
        QtGui.QMessageBox.StandardButton.Yes | QtGui.QMessageBox.StandardButton.No
    )

    if reply == QtGui.QMessageBox.StandardButton.Yes:
        iomanager_instance.excel_manager.save_table_to_xlsx(new_xl)
        iomanager_instance.excel_label.setText(new_xl)
        iomanager_instance.update_table(new_xl)
        print(f"[COMPLETE] Version up file saved : {new_xl}")
    else:
        pass