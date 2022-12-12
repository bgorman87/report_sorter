import os
import shutil
import ctypes

import regex as re
import win32com.client as win32
from PIL import Image
from PyQt5 import QtCore, QtGui, QtWidgets
from pdf2image import convert_from_path

from functions.analysis import WorkerAnalyzeThread, detect_package_number
from functions.project_info import project_info, json_setup

debug = False

# set current working directory to variable to save files to
home_dir = os.getcwd()

# hard coded poppler path from current working directory
poppler_path = str(os.path.abspath(os.path.join(os.getcwd(), r"poppler\bin")))


def output(self):
    self.output_box.appendPlainText("Analyzing...\n")


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.threadpool = QtCore.QThreadPool()
        print("Multithreading with maximum %d threads" %
              self.threadpool.maxThreadCount())
        self.analyzeWorker = None  # Worker
        self.thread = None  # Thread
        self.fileNames = None
        self.analyzed = False
        self.progress = 0
        QtCore.QMetaObject.connectSlotsByName(self)

        self.central_widget = QtWidgets.QWidget()
        self.central_widget = QtWidgets.QWidget()
        self.central_widget.setObjectName("centralwidget")
        self.layout_grid = QtWidgets.QGridLayout(self.central_widget)

        self.tab_widget = QtWidgets.QTabWidget()
        self.label = QtWidgets.QLabel()
        self.tab = QtWidgets.QWidget()
        self.progress_bar = QtWidgets.QProgressBar()
        self.grid_layout = QtWidgets.QGridLayout(self.tab)
        self.select_files = QtWidgets.QPushButton(self.tab)
        self.line = QtWidgets.QFrame(self.tab)
        self.line_2 = QtWidgets.QFrame(self.tab)
        self.analyze_button = QtWidgets.QPushButton(self.tab)
        self.email_button = QtWidgets.QPushButton(self.tab)
        self.test_box = QtWidgets.QComboBox(self.tab)
        self.debug_box = QtWidgets.QCheckBox(self.tab)
        self.tab_2 = QtWidgets.QWidget()
        self.grid_layout_2 = QtWidgets.QGridLayout(self.tab_2)
        self.output_box = QtWidgets.QPlainTextEdit(self.tab)
        self.label_4 = QtWidgets.QLabel(self.tab_2)
        self.list_widget = QtWidgets.QListWidget(self.tab_2)
        self.file_rename = QtWidgets.QLineEdit(self.tab_2)
        self.file_rename_button = QtWidgets.QPushButton(self.tab_2)
        self.label_3 = QtWidgets.QLabel(self.tab_2)
        self.graphics_view = QtWidgets.QGraphicsView(self.tab_2)
        self.status_bar = QtWidgets.QStatusBar()
        self.dialog = QtWidgets.QFileDialog()
        self.list_widget_item = QtWidgets.QListWidgetItem()
        self.project_numbers = []
        self.project_numbers_short = []

    # def setup_ui(self, main_window):
        self.setObjectName("MainWindow")
        self.resize(850, 850)
        self.statusBar().setSizeGripEnabled(False)

        self.tab_widget = QtWidgets.QTabWidget(self.central_widget)
        self.tab_widget.setObjectName("tabWidget")
        self.tab_widget.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.layout_grid.addWidget(self.tab_widget, 0, 0, 1, 1)

        self.label = QtWidgets.QLabel(self.central_widget)
        self.label.setObjectName("creatorLabel")
        self.layout_grid.addWidget(self.label, 1, 0, 1, 3)

        self.tab.setObjectName("tab")
        self.grid_layout.setObjectName("gridLayout")

        self.select_files.setObjectName("SelectFiles")
        self.grid_layout.addWidget(self.select_files, 3, 0, 1, 2)

        self.output_box.setObjectName("outputBox")
        self.output_box.setReadOnly(True)
        self.grid_layout.addWidget(self.output_box, 5, 0, 1, 8)

        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.grid_layout.addWidget(self.line, 2, 0, 1, 8)

        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.grid_layout.addWidget(self.line_2, 4, 0, 1, 8)

        self.analyze_button.setObjectName("analyzeButton")
        self.grid_layout.addWidget(self.analyze_button, 3, 2, 1, 2)

        self.email_button.setObjectName("emailButton")
        self.grid_layout.addWidget(self.email_button, 3, 4, 1, 2)

        self.test_box.setObjectName("testBox")
        self.grid_layout.addWidget(self.test_box, 3, 6, 1, 1)
        self.test_box.setEditable(True)
        self.test_box.lineEdit().setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.test_box.addItems(["Live", "Test"])
        self.test_box.setEditable(False)

        self.debug_box.setObjectName("debugBox")
        self.grid_layout.addWidget(self.debug_box, 3, 7, 1, 1)

        self.progress_bar.setObjectName("progressBar")
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setStyle(QtWidgets.QStyleFactory.create("GTK"))
        self.progress_bar.setTextVisible(False)
        self.grid_layout.addWidget(self.progress_bar, 6, 0, 1, 8)

        self.tab_widget.addTab(self.tab, "")
        self.tab_2.setObjectName("tab_2")
        self.grid_layout_2.setObjectName("gridLayout_2")

        self.label_4.setGeometry(QtCore.QRect(10, 10, 81, 16))
        self.label_4.setObjectName("combinedFilesLabel")
        self.grid_layout_2.addWidget(self.label_4, 0, 0, 1, 2)

        self.list_widget.setGeometry(QtCore.QRect(10, 30, 320, 100))
        self.list_widget.setObjectName("listWidget")
        self.grid_layout_2.addWidget(self.list_widget, 1, 0, 5, 5)

        self.file_rename.setObjectName("file rename")
        self.grid_layout_2.addWidget(self.file_rename, 6, 0, 1, 4)

        self.file_rename_button.setObjectName("fileRenameButton")
        self.grid_layout_2.addWidget(self.file_rename_button, 6, 4, 1, 1)

        self.label_3.setGeometry(QtCore.QRect(10, 140, 100, 16))
        self.label_3.setObjectName("pdfOutputLabel")
        self.grid_layout_2.addWidget(self.label_3, 7, 0, 1, 2)

        self.graphics_view.setGeometry(QtCore.QRect(10, 160, 320, 400))
        self.graphics_view.setObjectName("graphicsView")
        self.grid_layout_2.addWidget(self.graphics_view, 8, 0, 20, 5)
        self.graphics_view.setViewportUpdateMode(
            QtWidgets.QGraphicsView.FullViewportUpdate)

        self.tab_widget.addTab(self.tab_2, "")
        self.tab_widget.raise_()
        self.label.raise_()

        self.setCentralWidget(self.central_widget)
        self.status_bar = QtWidgets.QStatusBar()
        self.status_bar.setObjectName("status bar")
        self.setStatusBar(self.status_bar)

        self.translate_ui()
        self.tab_widget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(self)

        self.setTabOrder(self.select_files, self.analyze_button)
        self.setTabOrder(self.analyze_button, self.output_box)
        self.setTabOrder(self.output_box, self.tab)
        self.setTabOrder(self.tab, self.tab_2)

        self.show()

    progress_update = QtCore.pyqtSignal(int)

    def translate_ui(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "Englobe Sorter"))
        self.label.setText(_translate(
            "MainWindow", "Created By Brandon Gorman"))
        self.select_files.setText(_translate("MainWindow", "Select Files"))
        self.select_files.clicked.connect(self.select_files_handler)
        self.file_rename_button.setWhatsThis(_translate(
            "MainWindow", "Rename the currently selected file"))
        self.file_rename_button.setText(_translate("MainWindow", "Rename"))
        self.file_rename_button.clicked.connect(
            self.file_rename_button_handler)
        self.analyze_button.setText(_translate("MainWindow", "Analyze"))
        self.analyze_button.clicked.connect(self.analyze_button_handler)
        self.email_button.setText(_translate("MainWindow", "E-Mail"))
        self.debug_box.setText(_translate("MainWindow", "Debug"))
        self.email_button.clicked.connect(self.email_button_handler)
        self.tab_widget.setTabText(self.tab_widget.indexOf(
            self.tab), _translate("MainWindow", "Input"))
        self.label_3.setText(_translate("MainWindow", "File Output Viewer:"))
        self.label_4.setText(_translate("MainWindow", "Combined Files:"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(
            self.tab_2), _translate("MainWindow", "Output"))
        self.list_widget.itemClicked.connect(self.list_widget_handler)
        self.list_widget.itemDoubleClicked.connect(self.rename_file_handler)

    def email_button_handler(self):
        signature_path = os.path.abspath(os.path.join(
            home_dir + r"\\Signature\\concrete.htm"))
        signature_path_28 = os.path.abspath(os.path.join(
            home_dir + r"\\Signature\\concrete28.htm"))
        if not self.analyzed and self.fileNames:
            if os.path.isfile(signature_path):
                with open(signature_path, "r") as file:
                    body_text = file.read()
                with open(signature_path_28, "r") as file:
                    body_text_28 = file.read()

            else:
                print("Signature File Not Found")
                body_text = ""
                pass
            msg = QtWidgets.QMessageBox()
            button_reply = msg.question(msg, "", "Do you want to create e-mails for non-analyzed files?",
                                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if button_reply == QtWidgets.QMessageBox.No:
                self.output_box.appendPlainText("E-Mails not generated\n")
            elif button_reply == QtWidgets.QMessageBox.Yes:
                json_setup(self.test_box.currentText())
                for file in self.fileNames:
                    description = file.split("_")
                    description = description[1]
                    project_number, project_number_short, email_recipient_to, \
                        email_recipient_cc, email_recipient_subject = project_info(f=file,
                                                                                   analyzed=self.analyzed,
                                                                                   description=description)
                    title = file.split("/").pop()
                    attachment = file
                    if "Dexter" in email_recipient_subject:
                        dexter_number = "NA"
                        if re.search(r"Dexter_([\d-]*)[_\dA-z]", email_recipient_subject, re.I) is not None:
                            dexter_number = re.search(r"Dexter_([\d-]*)[_\dA-z]", email_recipient_subject,
                                                      re.I).groups()
                            dexter_number = dexter_number[-1]
                        email_recipient_subject = email_recipient_subject.replace(
                            "%%", dexter_number)
                    try:
                        outlook = win32.Dispatch('outlook.application')
                        mail = outlook.CreateItem(0)
                        mail.To = email_recipient_to
                        mail.CC = email_recipient_cc
                        mail.Subject = email_recipient_subject
                        if "28d" in title:
                            mail.HtmlBody = body_text_28
                        else:
                            mail.HtmlBody = body_text
                        mail.Attachments.Add(attachment)
                        mail.Save()
                        e = "Drafted email for: {0}".format(title)
                        self.output_box.appendPlainText(e)
                    except Exception as e:
                        print(e)
                        self.output_box.appendPlainText(e)
                        pass
        if self.analyzed:
            if os.path.isfile(signature_path):
                with open(signature_path, "r") as file:
                    body_text = file.read()
                with open(signature_path_28, "r") as file:
                    body_text_28 = file.read()
            else:
                print("Signature File Not Found")
                body_text = ""
                pass
            all_list_titles = []
            all_list_data = []
            for i in range(self.list_widget.count()):
                all_list_data.append(self.list_widget.item(
                    i).data(QtCore.Qt.UserRole).split("%%")[0])
                all_list_titles.append(self.list_widget.item(i).text())
            for i, project_number in enumerate(self.project_numbers):
                if self.project_numbers_short[i][0] == "P":
                    self.project_numbers_short[i] = self.project_numbers_short[i].replace(
                        "P-", "P-00")
                project_number, project_number_short, \
                    recipients, recipients_cc, subject = project_info(project_number, self.project_numbers_short[i],
                                                                      all_list_data[i], None, self.analyzed)
                attachment = all_list_data[i]
                if "Dexter" in subject:
                    dexter_number = "NA"
                    if re.search(r"Dexter_([\d-]*)[_\dA-z]", all_list_titles[i], re.I) is not None:
                        dexter_number = re.search(
                            r"Dexter_([\d-]*)[_\dA-z]", all_list_titles[i], re.I).groups()
                        dexter_number = dexter_number[-1]
                    subject = subject.replace("%%", dexter_number)
                try:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = recipients
                    mail.CC = recipients_cc
                    mail.Subject = subject
                    if "28d" in attachment:
                        mail.HtmlBody = body_text_28
                    else:
                        mail.HtmlBody = body_text
                    mail.Attachments.Add(attachment)
                    mail.Save()
                    e = "Drafted email for: {0}".format(all_list_titles[i])
                    self.output_box.appendPlainText(e)
                except Exception as e:
                    print(e)
                    self.output_box.appendPlainText(str(e))
                    pass

    def debug_check(self):
        global debug
        if self.debug_box.isChecked():
            debug = True
        else:
            debug = False

    def select_files_handler(self):
        self.open_file_dialog()

    def open_file_dialog(self):
        self.dialog = QtWidgets.QFileDialog(directory=str(
            os.path.abspath(os.path.join(os.getcwd(), r"..\.."))))
        self.fileNames, filters = QtWidgets.QFileDialog.getOpenFileNames()

        tuple(self.fileNames)
        if len(self.fileNames) == 1:
            file_names_string = "(" + str(len(self.fileNames)) + \
                ")" + " file has been selected: \n"
        else:
            file_names_string = "(" + str(len(self.fileNames)) + \
                ")" + " files have been selected: \n"
        for item in self.fileNames:
            file_names_string = file_names_string + item + "\n"
        self.output_box.appendPlainText(file_names_string)

    def rename_file_handler(self):
        if self.list_widget.isPersistentEditorOpen(self.list_widget.currentItem()):
            self.list_widget.closePersistentEditor(
                self.list_widget.currentItem())
            self.list_widget.editItem(self.list_widget.currentItem())
        else:
            self.list_widget.editItem(self.list_widget.currentItem())

    def file_rename_button_handler(self):
        file_path = self.list_widget.currentItem().data(QtCore.Qt.UserRole).split("%%")
        file_path_transit_src = file_path[0]
        # Project path may be changed if project number updated so declare up here
        file_path_project_src = file_path[1]

        # See if project number is the edited string. If it is and description is == "SomeProjectDescription"
        # Then the project was previously not detected properly so assume the project edit is correct and find
        # details in the JSON file.
        # Before renaming occurs Old data = entry in the listWidget
        #                        New data = entry in the text edit box
        description = "SomeProjectDescription"
        old_project_number = ""
        new_project_number = ""
        project_number = ""
        project_number_short = ""
        old_package = ""
        old_title = self.list_widget.currentItem().text()
        project_details_changed = False
        new_title = self.file_rename.text()
        if re.search(r"-([\dPBpb\.-]+)_", old_title, re.I) is not None:
            old_project_number = re.search(
                r"-([\dPBpb\.-]+)_", old_title, re.I).groups()
            old_project_number = old_project_number[-1]
        elif re.search(r"-(NA)_", old_title, re.I) is not None:
            old_project_number = "NA"
        if re.search(r"-([\dPBpb\.-]+)_", new_title, re.I) is not None:
            new_project_number = re.search(
                r"-([\dPBpb\.-]+)_", new_title, re.I).groups()
            new_project_number = new_project_number[-1]
        if old_project_number != new_project_number:
            project_number, project_number_short, \
                description, file_path_project_src = project_info(new_project_number, new_project_number,
                                                                  file_path_transit_src, None, False)
            project_details_changed = True
        if re.search(r"(\d+)-[\dA-z]", old_title, re.I) is not None:
            old_package = re.search(r"(\d+)-[\dA-z]", old_title, re.I).groups()
            old_package = old_package[-1]

        if project_details_changed:
            updated_file_details = old_title.replace(
                "SomeProjectDescription", description)
            updated_package = detect_package_number(
                file_path_project_src, debug)[0]
            updated_file_details = updated_file_details.replace(
                old_package, updated_package)
            updated_file_details = updated_file_details.replace(
                old_project_number, project_number_short)
            rename_transit_len = 260 - len(
                str(file_path_transit_src.replace(file_path_transit_src.split("\\").pop(), "")))
            rename_project_length = 260 - len(str(file_path_project_src))
            if len(updated_file_details) > rename_transit_len or len(updated_file_details) > rename_project_length:
                updated_file_details = updated_file_details.replace(
                    "Concrete", "Conc")
            if len(updated_file_details) > rename_transit_len or len(updated_file_details) > rename_project_length:
                updated_file_details = updated_file_details.replace(
                    "-2022", "")
                updated_file_details = updated_file_details.replace(
                    "-2021", "")
            if len(updated_file_details) > rename_transit_len or len(updated_file_details) > rename_project_length:
                if rename_project_length > rename_transit_len:
                    cut = rename_project_length + 4
                else:
                    cut = rename_transit_len + 4
                updated_file_details = updated_file_details.replace(".pdf", "")
                updated_file_details = updated_file_details[:-cut] + "LONG.pdf"
            rename_path_transit = os.path.abspath(os.path.join(
                file_path_transit_src.replace(
                    file_path_transit_src.split("\\").pop(), ""),
                updated_file_details + ".pdf"))
            rename_path_project = os.path.abspath(os.path.join(
                file_path_project_src, updated_file_details + ".pdf"))
            os.rename(file_path_transit_src, rename_path_transit)
            if not os.path.isfile(file_path_project_src):
                file_path_project_src = rename_path_project
            if os.path.isfile(file_path_project_src):
                if file_path_project_src != file_path_transit_src:
                    os.rename(file_path_project_src, rename_path_project)
            else:
                shutil.copy(rename_path_transit, rename_path_project)
            self.list_widget.currentItem().setText(updated_file_details)
            data = rename_path_transit + "%%" + rename_path_project
            self.list_widget.currentItem().setData(QtCore.Qt.UserRole, data)
            self.file_rename.setText(updated_file_details)
            self.project_numbers_short[self.list_widget.currentRow(
            )] = project_number_short
            self.project_numbers[self.list_widget.currentRow()
                                 ] = project_number
        else:
            # 254 to accommodate the .pdf
            rename_transit_len = 254 - \
                len(file_path_transit_src.replace(
                    file_path_transit_src.split("\\").pop(), ""))
            rename_project_len = 254 - \
                len(file_path_project_src.replace(
                    file_path_project_src.split("\\").pop(), ""))
            if len(self.file_rename.text()) > rename_transit_len or len(self.file_rename.text()) > rename_project_len:
                print("Filename too long")
                if rename_transit_len > rename_project_len:
                    msg_string = f"Filename too long. Reduce by {len(self.file_rename.text()) - rename_transit_len}"
                else:
                    msg_string = f"Filename too long. Reduce by {len(self.file_rename.text()) - rename_project_len}"
                ctypes.windll.user32.MessageBoxW(
                    0, msg_string, "Filename Too Long", 1)
            else:
                rename_path_transit = os.path.abspath(os.path.join(
                    file_path_transit_src.replace(
                        file_path_transit_src.split("\\").pop(), ""),
                    str(self.file_rename.text()) + ".pdf"))
                rename_path_project = os.path.abspath(os.path.join(
                    file_path_project_src.replace(
                        file_path_project_src.split("\\").pop(), ""),
                    str(self.file_rename.text()) + ".pdf"))
                try:
                    os.rename(file_path_transit_src, rename_path_transit)
                except Exception:
                    pass
                if file_path_project_src != file_path_transit_src:  # If project and transit aren't the same, rename
                    try:
                        os.rename(file_path_project_src, rename_path_project)
                    except Exception:
                        pass
                self.list_widget.currentItem().setText(self.file_rename.text())
                if debug:
                    print('Renamed File Path: \n{0}\n{1}'.format(
                        file_path_transit_src, file_path_project_src))
                data = rename_path_transit + "%%" + rename_path_project
                self.list_widget.currentItem().setData(QtCore.Qt.UserRole, data)

    def evt_analyze_complete(self, results):
        print_string = results[0]
        file_title = results[1]
        data = results[2]
        project_number = results[3]
        project_number_short = results[4]
        self.output_box.appendPlainText(print_string)
        self.list_widget_item = QtWidgets.QListWidgetItem(file_title)
        self.list_widget_item.setData(QtCore.Qt.UserRole, data)
        self.list_widget.addItem(self.list_widget_item)
        self.project_numbers.append(project_number)
        self.project_numbers_short.append(project_number_short)

    def evt_analyze_progress(self, val):
        self.progress += val
        self.progress_bar.setValue(int(self.progress / len(self.fileNames)))

    def analyze_button_handler(self):
        if self.fileNames is not None:
            json_setup(self.test_box.currentText())
            self.debug_check()
            self.analyzed = False
            self.progress = 0
            self.progress_bar.setValue(0)
            self.output_box.appendPlainText("Analysis Started...\n")
            self.data_processing()
            self.analyzed = True
        else:
            self.output_box.appendPlainText(
                "Please select at least 1 file to analyze...\n")
        self.analyze_button.setEnabled(True)

    def analyze_queue_button_handler(self):
        self.output_box.appendPlainText("Analyzing Queue Folder...\n")

    def list_widget_handler(self):
        file_path = str(self.list_widget.currentItem().data(
            QtCore.Qt.UserRole)).split("%%")
        file_path_transit_src = file_path[0]
        image_jpeg = []
        try:
            image_jpeg = convert_from_path(
                file_path_transit_src, fmt="jpeg", poppler_path=poppler_path)
        except Exception as e:
            print(e)
        if image_jpeg:
            result = Image.new("RGB", (1700, len(image_jpeg) * 2200))
            scene = QtWidgets.QGraphicsScene()
            for count, temp in enumerate(image_jpeg, 1):
                x = 0
                y = (count - 1) * 2200
                result.paste(temp, (x, y))
            name_jpeg = file_path_transit_src.replace(".pdf", ".jpg")
            result.save(name_jpeg, 'JPEG')
            pix = QtGui.QPixmap(name_jpeg)
            pix = pix.scaledToWidth(self.graphics_view.width())
            item = QtWidgets.QGraphicsPixmapItem(pix)
            scene.addItem(item)
            self.graphics_view.setScene(scene)
            os.remove(name_jpeg)
            set_text = file_path_transit_src.split(
                "\\").pop().replace(".pdf", "")
            self.file_rename.setText(set_text)
        else:
            print("image_jpeg list is empty")

    def data_processing(self):
        # iterate through all input files
        # for each file scan top right of sheet ((w/2, 0, w, h/8))
        # if top right of sheet contains "test" it's a concrete break sheet
        # pre-process entire image
        # from preprocessed image - crop to (1100,320, 1550, 360)
        # search resultant tesseract data for project_number
        # from preprocessed image - crop to (100, 675, 300, 750)
        # search resultant tesseract data for set_no
        # from preprocessed image - crop to (1260, 710, 1475, 750)
        # search resultant tesseract data for date_cast
        # from preprocessed image - crop to (1150, 830, 1350, 1100)
        # search resultant tesseract data for compressive strengths
        # split results by \n, last value in stored split should be the most recent break result
        # from preprocessed image - crop to (450, 830, 620, 1100)
        # search resultant tesseract data for age of cylinders when broken
        # split results by \n, find result in split equal to len(compressive_strength[splitdata])
        # this should return age of most recent broken cylinder

        # Import images from file path "f" using pdf to image to open
        for f in self.fileNames:
            self.analyzeWorker = WorkerAnalyzeThread(
                fileName=f, debug=debug, analyzed=self.analyzed)
            self.analyzeWorker.signals.progress.connect(
                self.evt_analyze_progress)
            self.analyzeWorker.signals.result.connect(
                self.evt_analyze_complete)
            # thread_pool.start(WorkerAnalyzeThread)
            self.threadpool.start(self.analyzeWorker)


if __name__ == "__main__":

    app = QtWidgets.QApplication([])
    window = MainWindow()
    app.exec_()
