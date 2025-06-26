# Copyright (c) 2013 Shotgun Software Inc.
#
# CONFIDENTIAL AND PROPRIETARY
#
# This work is provided "AS IS" and subject to the Shotgun Pipeline Toolkit
# Source Code License included in this distribution package. See LICENSE.
# By accessing, using, copying or modifying this work you indicate your
# agreement to the Shotgun Pipeline Toolkit Source Code License. All rights
# not expressly granted therein are reserved by Shotgun Software Inc.

import sgtk
import os
import re
from .get_xl import get_latest_xlsx, get_new_xlsx, save_table_to_xlsx
from .xl_to_table import update_table
from .sg_upload import get_publish_list, upload_shotgrid, table_to_publish_list
from sgtk.platform.qt import QtGui
from .excel_manager import ExcelManager
from .ui.iomanager_ui import IOManagerWidget
from .shot_converter import ShotConverter


# standard toolkit logger
logger = sgtk.platform.get_logger(__name__)


def show_dialog(app_instance):
    """
    Shows the main dialog window.
    """
    app_instance.engine.show_dialog("IO Manager", app_instance, AppDialog)

class AppDialog(QtGui.QWidget):
    """
    Main application dialog window
    """

    def __init__(self):
        """
        Constructor
        """
        QtGui.QWidget.__init__(self)
        self.current_engine = sgtk.platform.current_engine()
        self.sg = self.get_shotgun_api()
        self.project_name = self.get_project_name()
        self.home_path = os.path.expanduser("~")
        self.project_path = os.path.join(self.home_path, "show", self.project_name)
        self.default_path = os.path.join(self.project_path, "product", "scan")
        # ex : default_path = {self.home_path}/show/{project_name}/product/scan
        
        # View
        self.iomanager_ui = IOManagerWidget()
        self.connect_event_handlers()
        self.iomanager_ui.project_label.setText(self.project_name)
        
        # Model
        self.xl_manager = ExcelManager()

        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.iomanager_ui)
        self.setLayout(layout)
        self.resize(1400, 800)

        self.shot_converter = ShotConverter()

        logger.info("Launching IO Manager Dev...")


    def connect_event_handlers(self):
        self.iomanager_ui.shot_select_btn.clicked.connect(self.on_select_btn_clicked)
        self.iomanager_ui.excel_save_btn.clicked.connect(self.on_save_btn_click)
        self.iomanager_ui.select_excel_btn.clicked.connect(self.on_select_xl_btn_clicked)
        self.iomanager_ui.publish_btn.clicked.connect(self.on_publish_btn_clicked)
        self.iomanager_ui.validate_version_btn.clicked.connect(self.on_version_btn_clicked)

    def on_select_btn_clicked(self):
        selected_date_path = self.select_date_directory()
        if not selected_date_path:
            return
        self.xl_manager.set_date_path(selected_date_path)
        self.iomanager_ui.file_path_le.setText(selected_date_path)
        latest_xlsx_path = get_latest_xlsx(self)
        self.xl_manager.set_xl_path(latest_xlsx_path)
        update_table(self)
        self.iomanager_ui.excel_label.setText(latest_xlsx_path)

    def on_save_btn_click(self):
        new_xl_path = get_new_xlsx(self)
        if not new_xl_path:
            return
        
        reply = QtGui.QMessageBox.question(
            self.iomanager_ui,
            "Confirm save",
            f"The file will be saved as\n{os.path.basename(new_xl_path)}\nDo you wand to continue?",
            QtGui.QMessageBox.StandardButton.Yes | QtGui.QMessageBox.StandardButton.No
        )

        if reply == QtGui.QMessageBox.StandardButton.Yes:
            save_table_to_xlsx(self, new_xl_path)
            self.xl_manager.set_xl_path(new_xl_path)
            update_table(self)
            self.iomanager_ui.excel_label.setText(new_xl_path)
            logger.info(f"Excel file version up and saved {new_xl_path}")
        else:
            return

    def on_select_xl_btn_clicked(self):
        date_directory_path = self.xl_manager.get_date_path()
        if not os.path.isdir(date_directory_path):
            return
        
        xl_file_path, _ = QtGui.QFileDialog.getOpenFileName(
            self.iomanager_ui,
            "Select XLSX file",
            date_directory_path,
            "XLSX files (*.xlsx)"
        )

        if not xl_file_path:
            return
        
        if not xl_file_path.startswith(self.project_path):
            QtGui.QMessageBox.warning(
                    self.iomanager_ui, 
                    "Warning", 
                    f"'{xl_file_path}' Out of boundary : Not Selected Project"
                )
            return
        
        self.xl_manager.set_xl_path(xl_file_path)
        update_table(self)
        self.iomanager_ui.excel_label.setText(xl_file_path)

    def on_publish_btn_clicked(self):
        xl_path = self.xl_manager.get_xl_path()
        if xl_path == "":
            logger.error("No selected excel file for publish")
            self.show_error_dialog("Error: No selected excel file for publish")
            return
        
        #publish_list = get_publish_list(self, xl_path)
        publish_list = table_to_publish_list(self)
        shot_success = 0
        for shot_info in publish_list:
            seq = shot_info["seq name"]
            shot = shot_info["shot name"]
            scan_path = shot_info["scan path"]
            fps = shot_info["fps"]
            ver = shot_info["version"]
            plate_dir_path = os.path.join(self.project_path, "seq", seq, shot, "plate")
            org_path = os.path.join(plate_dir_path, "org", ver)
            sg_path = os.path.join(plate_dir_path, "shotgrid_upload_datas", ver)
            shot_info["org path"] = org_path

            self.shot_converter.set_paths(scan_path, org_path, sg_path, fps)
            ext = self.shot_converter.get_ext()

            if ext == ".exr":
                paths = self.shot_converter.make_sg_datas()
                # Upload
                upload_result = upload_shotgrid(self, paths, shot_info)
                if upload_result: 
                    shot_success += 1
                else: return

            elif ext == ".mov":
                pass

        self.show_info_dialog(f"Publish completed {shot_success}/{len(publish_list)}")

    def on_version_btn_clicked(self):
        table = self.iomanager_ui.table
        row_count = table.rowCount()
        check_col = 0  
        headers     = [table.horizontalHeaderItem(i).text() for i in range(table.columnCount())] # Get headers from table
        seq_col     = headers.index("seq name")
        shot_col    = headers.index("shot name")
        version_col = headers.index("version")

        checked_rows = []
        for row in range(row_count):
            widget = table.cellWidget(row, check_col)
            if isinstance(widget, QtGui.QCheckBox) and widget.isChecked():
                checked_rows.append(row)

        if not checked_rows:
            self.show_info_dialog("No checked row")
            return
        
        for row in checked_rows:
            seq  = table.item(row, seq_col).text()
            shot = table.item(row, shot_col).text()
            shot_dir = os.path.join(self.project_path, "seq", seq, shot, "plate", "org")
            if os.path.exists(shot_dir):
                versions = os.listdir(shot_dir)
            else:
                self.show_error_dialog("No exist path")
                return

            latest_version = self.get_latest_plate_version(versions)
            item = QtGui.QTableWidgetItem(latest_version)
            table.setItem(row, version_col, item)

        
    def get_latest_plate_version(self, dir_list):
        version_nums = []
        for name in dir_list:
            m = re.match(r"v(\d{3})", name)
            if m:
                version_nums.append(int(m.group(1)))
        if version_nums:
            latest = max(version_nums) + 1
            return f"v{latest:03d}"
        else:
            return "v001"

    def select_date_directory(self):
        default_path = self.default_path
        if not os.path.isdir(default_path):
            default_path = self.project_path
        
        selected_date_path = QtGui.QFileDialog.getExistingDirectory(
            self.iomanager_ui, # Parent Widget
            "Select scan data folder",
            default_path, 
            QtGui.QFileDialog.ShowDirsOnly
        )

        # return date directory path
        # ex) date directory path : /home/rapa/show/my_project/product/scan/250529
        if selected_date_path:
            if not selected_date_path.startswith(self.project_path):
                self.show_error_dialog(f"'{selected_date_path}' Out of boundary : Not Selected Project")
                return None
            return selected_date_path
        return None

    def get_project_name(self):
        """
        Get the project name from engine & context that is currently running

        Returns: 
            str: project name
        """
        context = self.current_engine.context
        project_name = context.project.get("name")
        return project_name

    def get_shotgun_api(self):
        """
        Get Shotgun API from engine & sgtk

        Returns: 
            object: shotgun api
        """
        # Grab the already created Sgtk instance from the current engine.
        tk = self.current_engine.sgtk
        sg = tk.shotgun
        return sg
    
    def show_error_dialog(self, message):
        QtGui.QMessageBox.warning(
            self.iomanager_ui,
            "Error",
            message,
            QtGui.QMessageBox.Ok
        )

    def show_info_dialog(self, message):
        QtGui.QMessageBox.information(
            self.iomanager_ui,
            "Info",
            message,
            QtGui.QMessageBox.Ok
        )