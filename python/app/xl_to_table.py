from openpyxl import load_workbook
from sgtk.platform.qt import QtGui

def update_table(app_instance):
    """
    Update table widget based on excel file

    Args:
        str: Path of reference excel file
    """
    xlsx_path = app_instance.xl_manager.get_xl_path()
    wb = load_workbook(xlsx_path)
    ws = wb.active

    # Making headers : First row of xlsx file -> Header label of table widget 
    headers = []
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        headers.append(cell.value)
    header_list = ["check"] + headers

    app_instance.iomanager_ui.table.setColumnCount(len(header_list))
    data_rows = list(ws.iter_rows(min_row=2, values_only=False))
    app_instance.iomanager_ui.table.setRowCount(len(data_rows))
    app_instance.iomanager_ui.table.setHorizontalHeaderLabels(header_list)

    for row_idx, row_data in enumerate(data_rows):
        checkbox = QtGui.QCheckBox()
        checkbox.clicked.connect(lambda _, row=row_idx: app_instance.on_checkbox_clicked(row, xlsx_path))
        app_instance.iomanager_ui.table.setCellWidget(row_idx, 0, checkbox)

        for col_idx, cell in enumerate(row_data):
            header = headers[col_idx]
            value = cell.value
            if value == None:
                value = ""

            if header == "thumbnail":
                set_thumbnail_cell(app_instance, row_idx, col_idx + 1, value)
            else:
                item = QtGui.QTableWidgetItem(str(value))
                item.setFlags(item.flags() | QtGui.Qt.ItemIsEditable)
                app_instance.iomanager_ui.table.setItem(row_idx, col_idx + 1, item)

def set_thumbnail_cell(app_instance, row, col, img_path):
    """
    Set thumbnail cell of table widget using QPixmap

    Args:
        row(int)
        col(int)
        img_path(str)
    """
    label = QtGui.QLabel()
    pixmap = QtGui.QPixmap(img_path)
    if not pixmap.isNull():
        label.setPixmap(pixmap.scaledToWidth(400, QtGui.Qt.SmoothTransformation))
        label.setAlignment(QtGui.Qt.AlignLeft)
        app_instance.iomanager_ui.table.setCellWidget(row, col, label)
        app_instance.iomanager_ui.table.setRowHeight(row, 230)
        app_instance.iomanager_ui.table.setColumnWidth(col, 400)