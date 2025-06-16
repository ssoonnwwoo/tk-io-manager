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
from .get_xl import get_latest_xlsx, get_new_xlsx, save_table_to_xlsx
from .xl_to_table import update_table
from .sg_upload import get_publish_list
from sgtk.platform.qt import QtGui
from .model.excel_manager import ExcelManager
from .view.iomanager_ui import IOManagerWidget

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


        logger.info("Launching IO Manager Dev...")


    def connect_event_handlers(self):
        self.iomanager_ui.shot_select_btn.clicked.connect(self.on_select_btn_clicked)
        self.iomanager_ui.excel_save_btn.clicked.connect(self.on_save_btn_click)
        self.iomanager_ui.select_excel_btn.clicked.connect(self.on_select_xl_btn_clicked)
        self.iomanager_ui.publish_btn.clicked.connect(self.on_publish_btn_clicked)

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
            self.show_error_dialog("Error: No selected excel file for publish")
            return
        
        publish_list = get_publish_list(self, xl_path)
        #[{'seq': 'S001', 'shot': 'S001_0010', 'version': 'v001', 'directory': '/home/rapa/show/Vamos/product/scan/20250530_test/002_A206C024_240315_R29Q'}]

        for shot_info in publish_list:
            seq = shot_info["seq"]
            shot = shot_info["shot"]
            ver = shot_info["version"]
            scan_dir_path = shot_info["directory"]

            plate_dir_path = os.path.join(self.project_path, "seq", seq, shot, "plate")
            org_path = os.path.join(plate_dir_path, "org", ver)
            sg_path = os.path.join(plate_dir_path, "shotgrid_upload_datas", ver)
            jpg_path = org_path+"_jpg"

            os.makedirs(org_path, exist_ok=True)
            os.makedirs(jpg_path, exist_ok=True)
            os.makedirs(sg_path, exist_ok=True)

            # seq, shot, ver, scandir, org path, sg path, jpg path


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
                QtGui.QMessageBox.warning(
                    self.iomanager_ui, 
                    "Warning", 
                    f"'{selected_date_path}' Out of boundary : Not Selected Project"
                )
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