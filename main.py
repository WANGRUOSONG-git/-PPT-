import sys
import os
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QVBoxLayout, QWidget, QProgressBar, \
    QLabel, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal
from translator import PPTTranslator

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, file_path, save_path):
        super().__init__()
        self.file_path = file_path
        self.save_path = save_path

    def run(self):
        try:
            translator = PPTTranslator(self.file_path, self.save_path)
            for progress in translator.translate():
                self.progress.emit(progress)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class PPTTranslatorApp(QMainWindow):
    RESOURCE_DIR = os.path.join(os.path.dirname(__file__), 'resources')

    def __init__(self):
        super().__init__()
        self.setWindowTitle("麦子PPT翻译")
        self.setGeometry(100, 100, 600, 400)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.importButton = QPushButton("导入PPT")
        self.importButton.clicked.connect(self.importPPT)
        layout.addWidget(self.importButton)

        self.saveButton = QPushButton("选择保存位置")
        self.saveButton.clicked.connect(self.saveFile)
        layout.addWidget(self.saveButton)

        self.translateButton = QPushButton("开始翻译")
        self.translateButton.clicked.connect(self.startTranslation)
        layout.addWidget(self.translateButton)

        self.progressBar = QProgressBar()
        layout.addWidget(self.progressBar)

        self.statusLabel = QLabel()
        layout.addWidget(self.statusLabel)

        self.reminderLabel1 = QLabel("1、原中文PPT要给翻译后的韩语留足位置，使用单倍行距。")
        self.reminderLabel2 = QLabel("2、翻译完成后检查生成文件，对阿拉伯数字不会进行翻译。")
        self.reminderLabel3 = QLabel("3、请勿使用复杂的PPT模板，回车检查可能出现的文字重叠。")
        layout.addWidget(self.reminderLabel1)
        layout.addWidget(self.reminderLabel2)
        layout.addWidget(self.reminderLabel3)

        self.designerLabel = QLabel("Designed by Wang Ruosong")
        self.feedbackLabel = QLabel("Problem feedback: 652046917@qq.com")
        layout.addWidget(self.designerLabel)
        layout.addWidget(self.feedbackLabel)

        self.versionLabel = QLabel("Version：0.2.1")
        layout.addWidget(self.versionLabel)

        self.manualButton = QPushButton("用户手册")
        self.manualButton.clicked.connect(self.downloadUserManual)
        layout.addWidget(self.manualButton)

        self.exampleButton = QPushButton("示范PPT")
        self.exampleButton.clicked.connect(self.downloadExampleFile)
        layout.addWidget(self.exampleButton)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.filePath = ""
        self.savePath = ""
        self.thread = None

    def importPPT(self):
        options = QFileDialog.Options()
        self.filePath, _ = QFileDialog.getOpenFileName(self, "导入PPT文件", "", "PowerPoint Files (*.pptx);;All Files (*)", options=options)
        if self.filePath:
            self.statusLabel.setText(f"导入成功: {self.filePath}")

    def saveFile(self):
        options = QFileDialog.Options()
        self.savePath, _ = QFileDialog.getSaveFileName(self, "选择保存位置", "", "PowerPoint Files (*.pptx);;All Files (*)", options=options)

    def startTranslation(self):
        if not self.filePath or not self.savePath:
            QMessageBox.warning(self, "警告", "请确保已导入PPT文件并选择保存位置。")
            return

        if self.thread and self.thread.isRunning():
            QMessageBox.warning(self, "警告", "翻译任务正在进行中，请稍后再试。")
            return

        self.thread = Worker(self.filePath, self.savePath)
        self.thread.progress.connect(self.updateProgress)
        self.thread.finished.connect(self.translationFinished)
        self.thread.error.connect(self.handleError)
        self.thread.start()

    def updateProgress(self, value):
        self.progressBar.setValue(value)

    def translationFinished(self):
        if not self.thread:
            return
        QMessageBox.information(self, "完成", "翻译完成！")
        self.thread = None

    def handleError(self, message):
        QMessageBox.critical(self, "错误", f"出现错误: {message}")

    def downloadUserManual(self):
        self._downloadFile('user_manual.pdf', "PDF Files (*.pdf)", "用户手册下载成功！")

    def downloadExampleFile(self):
        self._downloadFile('example_PPT.zip', "ZIP Files (*.zip)", "示范文件下载成功！")

    def _downloadFile(self, filename, file_filter, success_message):
        try:
            file_path = os.path.join(self.RESOURCE_DIR, filename)
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"{filename} 文件不存在。")

            save_path, _ = QFileDialog.getSaveFileName(self, "选择保存位置", "", file_filter)
            if save_path:
                shutil.copy(file_path, save_path)
                QMessageBox.information(self, "成功", success_message)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"下载文件时出错: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PPTTranslatorApp()
    window.show()
    sys.exit(app.exec_())
