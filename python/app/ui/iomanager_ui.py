from tank.platform.qt import QtCore
for name, cls in QtCore.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls

from tank.platform.qt import QtGui
for name, cls in QtGui.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls

import os
import sys

# Import resource binding python code
resource_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "resources"))
if resource_path not in sys.path:
    sys.path.append(resource_path)
import resources_rc

class IOManagerWidget(QWidget): 
    def __init__(self):
        super().__init__()
        # Widgets Top
        p_label = QLabel("Current project: ")
        self.project_label = QLabel()
        file_path_label = QLabel("File path:")
        self.file_path_le = QLineEdit()
        self.file_path_le.setPlaceholderText("Input your shot path")
        self.shot_select_btn = QPushButton("Select")
        
        # Widgets Mid(Table)
        self.table = QTableWidget()
        
        # Widgets Bottom
        current_excel_label = QLabel("Currently displayed Excel file:")
        self.excel_label = QLabel("Ready to load")
        self.excel_save_btn = QPushButton("Save Excel")
        self.select_excel_btn = QPushButton("Select Excel")
        self.publish_btn = QPushButton("Publish")

        # Layout
        main_layout = QVBoxLayout()
        shot_select_container = QHBoxLayout()
        bottom_layout = QHBoxLayout()
        excel_group = QGroupBox("Excel")
        excel_btn_layout = QHBoxLayout()
        excel_btn_layout2 = QHBoxLayout()
        excel_container = QVBoxLayout()
        excel_label_conatainer = QVBoxLayout()
        
        shot_select_container.addWidget(p_label)
        shot_select_container.addWidget(self.project_label)
        shot_select_container.addStretch()
        shot_select_container.addWidget(file_path_label)
        shot_select_container.addWidget(self.file_path_le)
        shot_select_container.addWidget(self.shot_select_btn)
        
        excel_label_conatainer.addWidget(current_excel_label, alignment=Qt.AlignTop)
        excel_label_conatainer.addWidget(self.excel_label, alignment=Qt.AlignTop)
        excel_btn_layout.addWidget(self.excel_save_btn)
        excel_btn_layout2.addWidget(self.select_excel_btn)
        excel_container.addLayout(excel_btn_layout)
        excel_container.addLayout(excel_btn_layout2)
        excel_group.setLayout(excel_container)
        bottom_layout.addLayout(excel_label_conatainer)
        bottom_layout.addWidget(excel_group)
        bottom_layout.addWidget(self.publish_btn)

        main_layout.addLayout(shot_select_container)
        main_layout.addWidget(self.table)
        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)