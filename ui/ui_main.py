from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QMessageBox


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("동물매크로")
        Dialog.resize(300, 400)
        Dialog.setMinimumSize(QtCore.QSize(300, 430))
        Dialog.setMaximumSize(QtCore.QSize(300, 430))
        Dialog.setAutoFillBackground(False)
        Dialog.setSizeGripEnabled(False)

        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(50, 10, 200, 30))
        font = QtGui.QFont()
        font.setPointSize(20)
        font.setBold(True)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")

        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(20, 50, 50, 20))
        self.fontComboBox = QtWidgets.QFontComboBox(Dialog)
        self.fontComboBox.setGeometry(QtCore.QRect(20, 70, 120, 25))

        self.font_size = QtWidgets.QSpinBox(Dialog)
        self.font_size.setGeometry(QtCore.QRect(150, 70, 50, 25))
        self.font_size.setProperty("value", 9)

        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(210, 50, 50, 20))
        self.line = QtWidgets.QSpinBox(Dialog)
        self.line.setGeometry(QtCore.QRect(210, 70, 70, 25))
        self.line.setMinimum(100)
        self.line.setMaximum(300)
        self.line.setProperty("value", 130)

        self.groupBox = QtWidgets.QGroupBox(Dialog)
        self.groupBox.setGeometry(QtCore.QRect(10, 110, 280, 110))
        self.groupBox.setTitle("전체 작업")

        self.excel_to = QtWidgets.QPushButton(self.groupBox)
        self.excel_to.setGeometry(QtCore.QRect(10, 20, 125, 60))
        self.style_to = QtWidgets.QPushButton(self.groupBox)
        self.style_to.setGeometry(QtCore.QRect(145, 20, 125, 60))
        self.Cvar1 = QtWidgets.QCheckBox(self.groupBox)
        self.Cvar1.setGeometry(QtCore.QRect(145, 85, 125, 20))
        self.Cvar1.setChecked(True)

        self.groupBox_2 = QtWidgets.QGroupBox(Dialog)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 230, 280, 110))
        self.groupBox_2.setTitle("개별 작업")

        self.one_fam = QtWidgets.QPushButton(self.groupBox_2)
        self.one_fam.setGeometry(QtCore.QRect(10, 20, 125, 35))
        self.one_bird = QtWidgets.QPushButton(self.groupBox_2)
        self.one_bird.setGeometry(QtCore.QRect(145, 20, 125, 35))
        self.remove_superscript_btn = QtWidgets.QPushButton(self.groupBox_2)
        self.remove_superscript_btn.setGeometry(QtCore.QRect(10, 65, 260, 35))

        self.groupBox_3 = QtWidgets.QGroupBox(Dialog)
        self.groupBox_3.setGeometry(QtCore.QRect(10, 350, 280, 65))
        self.groupBox_3.setTitle("곤충만 작업(엑셀에서 자르기 사용했을때)")

        self.insect_to = QtWidgets.QPushButton(self.groupBox_3)
        self.insect_to.setGeometry(QtCore.QRect(10, 20, 125, 35))
        self.only_insect = QtWidgets.QPushButton(self.groupBox_3)
        self.only_insect.setGeometry(QtCore.QRect(145, 20, 125, 35))

        button_style = """
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #c0c0c0;
                border-radius: 5px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """
        self.excel_to.setStyleSheet(button_style)
        self.style_to.setStyleSheet(button_style)
        self.one_fam.setStyleSheet(button_style)
        self.one_bird.setStyleSheet(button_style)
        self.remove_superscript_btn.setStyleSheet(button_style)
        self.insect_to.setStyleSheet(button_style)
        self.only_insect.setStyleSheet(button_style)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "동물 매크로"))
        self.label_3.setText(_translate("Dialog", "동물 매크로"))
        self.label.setText(_translate("Dialog", "폰트"))
        self.label_2.setText(_translate("Dialog", "줄 간격"))
        self.excel_to.setText(_translate("Dialog", "엑셀에서\n한글로 변환"))
        self.style_to.setText(_translate("Dialog", "한글 종목록\n스타일 적용"))
        self.Cvar1.setText(_translate("Dialog", "조류군집표 포함"))
        self.one_fam.setText(_translate("Dialog", "종"))
        self.one_bird.setText(_translate("Dialog", "조류군집표"))
        self.remove_superscript_btn.setText(_translate("Dialog", "윗첨자 제거"))
        self.insect_to.setText(_translate("Dialog", "곤충 목록"))
        self.only_insect.setText(_translate("Dialog", "곤충 스타일"))


class MainWindow(QtWidgets.QDialog, Ui_Dialog):
    show_message_signal = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show_message_signal.connect(self.show_message_box)
        
        from core.thread_runner import run_in_thread
        from logic.job_trigger import excel_to_hwp, insect_to_hwp, 표, 곤충
        from logic.hwp_formatting import 표정리, 조류군집표
        from logic.styles import remove_superscript

        self.excel_to.clicked.connect(lambda: run_in_thread(excel_to_hwp))
        self.style_to.clicked.connect(lambda: run_in_thread(표))
        self.one_fam.clicked.connect(lambda: run_in_thread(표정리))
        self.one_bird.clicked.connect(lambda: run_in_thread(조류군집표))
        self.remove_superscript_btn.clicked.connect(lambda: run_in_thread(remove_superscript))
        self.insect_to.clicked.connect(lambda: run_in_thread(insect_to_hwp))
        self.only_insect.clicked.connect(lambda: run_in_thread(곤충))
        
        self.Cvar1.stateChanged.connect(self.update_checkbox)

        self.font_size.setValue(9)
        self.line.setValue(130)

    def update_checkbox(self, state):
        from core.global_state import set_Cv
        var = state
        set_Cv(var)

    def get_selected_font(self):
        return self.fontComboBox.currentText()

    def get_font_size(self):
        return self.font_size.value()

    def show_message_box(self, msg):
        box = QMessageBox(QMessageBox.Information, "작업 완료", msg, parent=self)
        box.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint, True)  # ✅ 항상 맨 위에
        box.exec()