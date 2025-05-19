from tank.platform.qt import QtCore
for name, cls in QtCore.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls

from tank.platform.qt import QtGui
for name, cls in QtGui.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls


from ..event.io_event_handler import select_directory, toggle_edit_mode, select_xlsx_file
from ..tools.export_metadata import export_metadata
from ..tools.save_as_xlsx import save_as_xlsx
from ..tools.get_latest_xlsx_file import get_latest_version_file
from ..tools.get_new_version_file import get_new_version_name
from ..tools.table_to_metalist import save_table_to_xlsx
from ..tools.extract_directory_column import extract_directory_column
from ..tools.generate_directory_list import generate_directory_list
from ..tools.get_publish_info import get_publish_info
from ..tools.rename import rename_sequence
from ..tools.convert import exrs_to_jpgs, mov_to_exrs, exrs_to_video, exrs_to_montage, exrs_to_thumbnail
import os
import pandas as pd
import shotgun_api3

sg = shotgun_api3.Shotgun("https://westworld5th.shotgrid.autodesk.com",
                            script_name="ssoonnwwoo",
                            api_key="hueyvBxznqw&op8wvuzljtljv")

class IOManagerWidget(QWidget):
    def __init__(self):
        super().__init__()
        show_dir = os.path.join(os.path.expanduser("~"), "show")
        project_list = os.listdir(show_dir)
        project_list.insert(0, "Select your project first") 
        # Widgets
        project_label = QLabel("Select project: ")
        self.project_cb = QComboBox()
        self.project_cb.setCurrentIndex(0)
        self.project_cb.addItems(project_list)
        file_path_label = QLabel("File path:")
        self.file_path_le = QLineEdit()
        self.file_path_le.setPlaceholderText("Input your shot path")
        shot_select_btn = QPushButton("Select")
        # shot_load_btn = QPushButton("Load")

        self.table = QTableWidget()

        current_excel_label = QLabel("Currently displayed Excel file:")
        self.excel_label = QLabel("Ready to load")
        excel_save_btn = QPushButton("Version and Save")
        self.excel_edit_btn = QPushButton("Enable Edit")
        select_excel_btn = QPushButton("Select Excel")

        publish_btn = QPushButton("Publish")

        # Event Handle
        # shot_select_btn.clicked.connect(
        #     lambda: select_directory(self.file_path_le)
        # )
        # shot_load_btn.clicked.connect(self.on_load_clicked)
        shot_select_btn.clicked.connect(self.on_select_clicked)
        self.excel_edit_btn.clicked.connect(self.on_edit_clicked)
        excel_save_btn.clicked.connect(self.on_save_clicked)
        self.project_cb.currentTextChanged.connect(self.on_project_selected)
        select_excel_btn.clicked.connect(self.on_select_excel_clicked)
        publish_btn.clicked.connect(self.on_publish_clicked)

        # Layout
        main_layout = QVBoxLayout()
        shot_select_container = QHBoxLayout()
        bottom_layout = QHBoxLayout()
        excel_group = QGroupBox("Excel")
        excel_btn_layout = QHBoxLayout()
        excel_btn_layout2 = QHBoxLayout()
        excel_container = QVBoxLayout()
        excel_label_conatainer = QVBoxLayout()

        shot_select_container.addWidget(project_label)
        shot_select_container.addWidget(self.project_cb)
        shot_select_container.addWidget(file_path_label)
        shot_select_container.addWidget(self.file_path_le)
        shot_select_container.addWidget(shot_select_btn)
        #shot_select_container.addWidget(shot_load_btn)
        
        excel_label_conatainer.addWidget(current_excel_label, alignment=Qt.AlignTop)
        excel_label_conatainer.addWidget(self.excel_label, alignment=Qt.AlignTop)
        excel_btn_layout.addWidget(excel_save_btn)
        excel_btn_layout.addWidget(self.excel_edit_btn)
        excel_btn_layout2.addWidget(select_excel_btn)
        excel_container.addLayout(excel_btn_layout)
        excel_container.addLayout(excel_btn_layout2)
        excel_group.setLayout(excel_container)
        bottom_layout.addLayout(excel_label_conatainer)
        bottom_layout.addWidget(excel_group)
        bottom_layout.addWidget(publish_btn)

        main_layout.addLayout(shot_select_container)
        main_layout.addWidget(self.table)
        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)
        self.edit_mode = False

    def on_project_selected(self, project_name):
        base_path = os.path.expanduser("~")
        scan_path = os.path.join(base_path, "show", project_name, "product", "scan")
        self.file_path_le.setText(scan_path)

    def on_select_clicked(self):
        selected_path = select_directory(self.file_path_le)
        if not selected_path:
            return
        
        date_path = self.file_path_le.text()
        # First get latest version of xlsx file
        latest_xlsx_path = get_latest_version_file(date_path)
        # If not xlsx file, export metadata and save as {prefix_date}_list_v001.xlsx
        if not latest_xlsx_path:
            meta_data = export_metadata(date_path)
            latest_xlsx_path = save_as_xlsx(date_path, latest_xlsx_path, meta_data)
            self.update_table(latest_xlsx_path)
            self.excel_label.setText(latest_xlsx_path)
            return
        
        # If xlsx file exists
        meta_data_list = export_metadata(date_path)

        # Compare "Directory" column
        dirs_from_xlsx = sorted(extract_directory_column(latest_xlsx_path))
        current_dirs = sorted(generate_directory_list(meta_data_list))

        if dirs_from_xlsx == current_dirs:
            self.update_table(latest_xlsx_path)
            self.excel_label.setText(latest_xlsx_path)
        else:
            self.show_update_dialog(meta_data_list, date_path)

    def show_update_dialog(self, meta_data_list, date_path):
        reply = QMessageBox.question(
            self,
            f"Update Detected in {os.path.basename(date_path)}",
            "Something changed!\nDo you want to update & open the xlsx file",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            new_name = get_new_version_name(date_path)
            new_path = os.path.join(date_path, new_name)
            save_as_xlsx(date_path, new_name, meta_data_list)
            self.update_table(new_path)
            self.excel_label.setText(new_path)
        else:
            print("[CANCEL] Update canceled")

            latest_xlsx_path = get_latest_version_file(date_path)
            if latest_xlsx_path:
                self.update_table(latest_xlsx_path)
                self.excel_label.setText(latest_xlsx_path)

    def on_edit_clicked(self):
        self.edit_mode = toggle_edit_mode(self.table, self.edit_mode)
        if self.edit_mode:
            self.excel_edit_btn.setText("Disable Edit")
        else:
            self.excel_edit_btn.setText("Enable Edit")

    def on_save_clicked(self):
        # csv_path = self.excel_label.text()
        # save_table_to_csv(self.table, csv_path, parent=self)
        if not os.path.exists(self.file_path_le.text()):
            return
        xlsx_path = get_new_version_name(self.file_path_le.text())

        reply = QMessageBox.question(
            self,
            "Confirm Save",
            f"The file will be saved as\n{os.path.basename(xlsx_path)}\nDo you want to continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            save_table_to_xlsx(self.table, xlsx_path)
            self.excel_label.setText(xlsx_path)
            self.update_table(xlsx_path)
            print(f"[COMPLETE] Version up file saved : {xlsx_path}")
        else:
            pass

    def update_table(self, xlsx_path):
        self.edit_mode = False
        df = pd.read_excel(xlsx_path)
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns) + 1)
        header_list = ["check"] + df.columns.tolist()
        self.table.setHorizontalHeaderLabels(header_list)

        for row_idx, row_data in df.iterrows():
            checkbox = QCheckBox()
            checkbox.clicked.connect(lambda _, row = row_idx: self.on_checkbox_clicked(row, xlsx_path))
            self.table.setCellWidget(row_idx, 0, checkbox)

            for col_idx, header in enumerate(df.columns):
                value = row_data[header]

                # NaN 처리
                if pd.isna(value):
                    value = ""

                # 썸네일 처리
                if header == "thumbnail":
                    value = row_data["thumbnail_path"]
                    self.set_thumbnail_cell(row_idx, col_idx + 1, value)
                    #item = QTableWidgetItem(value)
                    #self.table.setItem(row_idx, col_idx + 1, item)
                else:
                    item = QTableWidgetItem(str(value))
                    self.table.setItem(row_idx, col_idx + 1, item)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)

    def on_checkbox_clicked(self, row, xlsx_path):
        # print(f"[TEST] Checkbox clicked on row {row}")
        df = pd.read_excel(xlsx_path)

        seq = df.at[row, "seq_name"]
        shot = df.at[row, "shot_name"]

        if pd.isna(seq) or pd.isna(shot) or str(seq).strip() == "" or str(shot).strip() == "":
            QMessageBox.warning(
            self,
            "Missing Data",
            f"Row[{row + 1}] missing 'seq_name' or 'shot_name'.\nThis row cannot be selected.",
            QMessageBox.Ok
            )

            checkbox = self.table.cellWidget(row, 0)
            if isinstance(checkbox, QCheckBox):
                checkbox.setChecked(False)
            return

        else:
            print(f"[OK] seq: {seq}, shot: {shot} at row {row + 1}")

    def set_thumbnail_cell(self, row, col, img_path):
        label = QLabel()
        pixmap = QPixmap(img_path)
        if not pixmap.isNull():
            label.setPixmap(pixmap.scaledToWidth(400, Qt.SmoothTransformation))
            label.setAlignment(Qt.AlignLeft)
            self.table.setCellWidget(row, col, label)
            self.table.setRowHeight(row, 230)
            self.table.setColumnWidth(col, 400)

    def on_select_excel_clicked(self):
        xlsx_file_path = select_xlsx_file(self.file_path_le)
        if xlsx_file_path:
            self.update_table(xlsx_file_path)
            self.excel_label.setText(xlsx_file_path)

    def get_checked_rows(self):
        checked_rows = []
        for row in range(self.table.rowCount()):
            checkbox = self.table.cellWidget(row, 0)
            if isinstance(checkbox, QCheckBox) and checkbox.isChecked():
                checked_rows.append(row + 2)
        return checked_rows

    def on_publish_clicked(self):
        xlsx_file_path = self.excel_label.text()
        checked_rows = self.get_checked_rows()
        if not checked_rows:
            QMessageBox.information(self, "No Selection", "Please check at least one item to publish.", QMessageBox.Ok)
            return
        shot_info_list = get_publish_info(xlsx_file_path, checked_rows)
        home_dir = os.path.expanduser("~")
        base_path = os.path.join(home_dir, "show")
        project_name = self.project_cb.currentText()
        for shot_info in shot_info_list:
            seq = shot_info["sequence"]
            shot = shot_info["shot"]
            typ = shot_info["type"]
            ver = shot_info["version"]
            ver_jpg = f"{ver}_jpg"
            directory = shot_info["directory"]

            print(f"[INFO] Preparing to publish shot : {shot}")
            org_path = os.path.join(
            base_path, project_name,"seq",
            seq,shot, "plate", typ, ver
            )

            jpg_path = os.path.join(
            base_path, project_name,"seq",
            seq, shot, "plate", typ, ver_jpg
            )

            sg_path = os.path.join(
                base_path, project_name, "seq",
                seq, shot, "plate", "shotgrid_upload_datas", ver
            )

            os.makedirs(org_path, exist_ok=True)
            os.makedirs(jpg_path, exist_ok=True)
            os.makedirs(sg_path, exist_ok=True)

            files = []
            for file in os.listdir(directory):
                # full_path = os.path.join(directory, file)
                files.append(file)

            # 확장자 조사
            exts = set()
            for file in files:
                _, ext = os.path.splitext(file)
                exts.add(ext.lower())
            
            # exr -> rename + exr sequence to jpg sequence
            if ".exr" in exts:
                rename_sequence(directory, org_path)
                exrs_to_jpgs(org_path, jpg_path)
                mp4_path = exrs_to_video(org_path, sg_path, vformat='mp4')
                webm_path = exrs_to_video(org_path, sg_path, vformat='webm')
                filmstrip_path = exrs_to_montage(org_path, sg_path)
                thumbnail_path = exrs_to_thumbnail(org_path, sg_path)
            
            # mov -> mov to exr sequence + exr sequence to jpg sequence
            elif ".mov" in exts:
                mov_path = os.path.join(directory,files[0])
                success = mov_to_exrs(mov_path, org_path)
                if success:
                    exrs_to_jpgs(org_path, jpg_path)
                    mp4_path = exrs_to_video(org_path, sg_path, vformat='mp4')
                    webm_path = exrs_to_video(org_path, sg_path, vformat='webm')
                    filmstrip_path = exrs_to_montage(org_path, sg_path)
                    thumbnail_path = exrs_to_thumbnail(org_path, sg_path)

            if all(os.path.isfile(path) for path in [mp4_path, webm_path, filmstrip_path, thumbnail_path]):
                project_id = sg.find_one("Project", [["name", "is", project_name]], ["id"])
                # Create sequence
                # 시퀀스 존재 여부 확인
                sequence_name = seq
                sequence = sg.find_one("Sequence", [
                    ["project", "is", project_id],
                    ["code", "is", sequence_name]
                ], ['code'])

                # 없으면 생성
                if not sequence:
                    sequence_data = {
                        "project": project_id,
                        "code": sequence_name
                    }
                    sequence = sg.create("Sequence", sequence_data)
                    print(f"[COMPLETE] Sequence created: {sequence['code']}")
                else:
                    print(f"[INFO] Sequence already exists: {sequence['code']}")

                # Create shot
                shot_name = f"{shot}_{ver}"
                shot_data = {
                    "project": project_id,
                    "code": shot_name,          
                    "sg_sequence": sequence
                }
                created_shot = sg.create("Shot", shot_data)
                print(f"[COMPLETE] Created Shot: {created_shot['code']}")

                # thumbnail upload to shot
                sg.upload_thumbnail("Shot", created_shot["id"], thumbnail_path)
                print(f"[COMPLETE] Uploaded thumbnail: {thumbnail_path}")

                # filmstrip upload to shot
                sg.upload_filmstrip_thumbnail("Shot", created_shot["id"], filmstrip_path)
                print(f"[COMPLETE] Uploaded filmstrip thumbnail: {filmstrip_path}")

                # create version and link to shot
                version_name = f"{shot}_plate_{ver}"
                version_data = {
                    "project": project_id,
                    "code": version_name,
                    "entity": created_shot,
                    "description": f"Description for {shot}"
                }
                version = sg.create("Version", version_data)
                print(f"[COMPLETE] Created Version: {version['code']}")
                
                # mp4 upload to version
                sg.upload("Version", version["id"], mp4_path, field_name="sg_uploaded_movie")
                print(f"[COMPLETE] Uploaded mp4 movie: {mp4_path}")

                # jpg, filmstrip upload to Version
                sg.upload_thumbnail("Version", version["id"], thumbnail_path)
                sg.upload_filmstrip_thumbnail("Version", version["id"], filmstrip_path)

                print(f"[COMPLETE]: {shot} publish completed\n\n")

            else:
                print("[ERROR] One or more expected output files are missing:")
                for path in [mp4_path, webm_path, filmstrip_path, thumbnail_path]:
                    if not os.path.isfile(path):
                        print(f" - Missing: {path}")