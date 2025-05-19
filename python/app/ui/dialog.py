# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'dialog.ui'
##
## Created by: Qt User Interface Compiler version 5.15.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from tank.platform.qt import QtCore
for name, cls in QtCore.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls

from tank.platform.qt import QtGui
for name, cls in QtGui.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls


from  . import resources_rc

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        if not Dialog.objectName():
            Dialog.setObjectName(u"Dialog")
        Dialog.resize(431, 392)
        self.horizontalLayout = QHBoxLayout(Dialog)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.contextContainer = QVBoxLayout(Dialog)
        self.contextContainer.setObjectName(u"contextContainer")
        self.logo_example = QLabel(Dialog)
        self.logo_example.setObjectName(u"logo_example")
        self.logo_example.setPixmap(QPixmap(u":/res/sg_logo.png"))

        self.horizontalLayout.addWidget(self.logo_example)

        self.context = QLabel(Dialog)
        self.context.setObjectName(u"context")
        self.context_user = QLabel(Dialog)
        self.context_user.setObjectName(u"context_user")
        self.project_dir = QLabel(Dialog)
        self.project_dir.setObjectName(u"project_dir")

        #sizePolicy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        #sizePolicy.setHorizontalStretch(0)
        #sizePolicy.setVerticalStretch(0)
        #sizePolicy.setHeightForWidth(self.context.sizePolicy().hasHeightForWidth())
        #sizePolicy.setHeightForWidth(self.context_user.sizePolicy().hasHeightForWidth())
        #self.context.setSizePolicy(sizePolicy)
        #self.context.setAlignment(Qt.AlignLeading|Qt.AlignLeft|Qt.AlignVCenter)
        #self.context_user.setSizePolicy(sizePolicy)
        #self.context_user.setAlignment(Qt.AlignLeading|Qt.AlignLeft|Qt.AlignVCenter)

        self.contextContainer.addWidget(self.context)
        self.contextContainer.addWidget(self.context_user)
        self.contextContainer.addWidget(self.project_dir)

        self.horizontalLayout.addLayout(self.contextContainer)

        self.retranslateUi(Dialog)

        QMetaObject.connectSlotsByName(Dialog)
    # setupUi

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"The Current Sgtk Environment", None))
        self.logo_example.setText("")
        self.context.setText(QCoreApplication.translate("Dialog", u"Your Current Context: ", None))
    # retranslateUi
