# pyinstaller.exe --onedir --windowed --name "Grand Champion Calculator" .\src\main.py

import sys
import os
import json
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

from calculator import calculate


class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.output_overridden = False

        # import settings
        if not os.path.exists("settings.json"):
            with open("settings.json", 'w') as f:
                json.dump({"class hierarchy": ["utility", "open", "novice"],
                           "defaults": {"break ties by class": False, "write to new file": False}}, f)

        with open("settings.json", 'r') as f:
            settings = json.load(f)

        # input box setup
        input_box = QGroupBox("Input File")
        input_layout = QHBoxLayout()
        self.input_file_field = QLineEdit()
        self.input_file_field.editingFinished.connect(self.update_output_field)
        select_file_button = QPushButton(QIcon("icons8-xlsx-160.png"), "Open")
        select_file_button.setToolTip("XLSX icon by Icons8")
        select_file_button.clicked.connect(self.select_file)
        input_layout.addWidget(self.input_file_field)
        input_layout.addWidget(select_file_button)
        input_box.setLayout(input_layout)

        # output box setup
        output_box = QGroupBox("Output File")
        output_layout = QVBoxLayout()
        self.output_file_field = QLineEdit()
        self.output_file_field.setDisabled(True)
        self.output_file_field.editingFinished.connect(self.output_manually_changed)
        output_layout.addWidget(self.output_file_field)
        output_box.setLayout(output_layout)

        # competition type box setup
        competition_box = QGroupBox("Competition")
        competition_layout = QHBoxLayout()
        self.obedience_button = QRadioButton("Obedience")
        competition_layout.addWidget(self.obedience_button)
        self.rally_button = QRadioButton("Rally")
        competition_layout.addWidget(self.rally_button)
        competition_box.setLayout(competition_layout)

        # options box setup
        options_box = QGroupBox("Options")
        options_layout = QVBoxLayout()
        self.return_multiple_away_winners = QCheckBox("Break ties by class")
        if "defaults" in settings and "break ties by class" in settings["defaults"]:
            self.return_multiple_away_winners.setChecked(settings["defaults"]["break ties by class"])
        options_layout.addWidget(self.return_multiple_away_winners)
        self.new_file = QCheckBox("Write output to new file")
        if "defaults" in settings and "write to new file" in settings["defaults"]:
            self.new_file.setChecked(settings["defaults"]["write to new file"])
        self.new_file.clicked.connect(self.new_file_changed)
        options_layout.addWidget(self.new_file)
        options_box.setLayout(options_layout)

        # remaining widgets
        calculate_button = QPushButton('Calculate')
        calculate_button.clicked.connect(self.calculate_function)
        self.status_field = QLabel("Status: Idle", alignment=Qt.AlignCenter)

        # add widgets/layouts to the main layout
        layout = QVBoxLayout()
        layout.addWidget(input_box)
        layout.addWidget(output_box)
        layout.addWidget(competition_box)
        layout.addWidget(options_box)
        status_layout = QHBoxLayout()
        status_layout.addWidget(calculate_button)
        status_layout.addWidget(self.status_field)
        layout.addLayout(status_layout)

        # set layout
        self.setLayout(layout)

        # finish setting up some other settings
        self.setWindowTitle("Grand Champion Calculator v1.0")
        self.setWindowIcon(QIcon("trip.ico"))
        self.setGeometry(100, 100, 128 * 3, 0)

    def select_file(self) -> None:
        """
        Opens a window to allow the user to select a file, then updates the output field.
        """
        self.status_field.setText("Status: Waiting for input file")
        self.input_file_field.setText(QFileDialog.getOpenFileName(self, 'Open File', os.getenv('HOME'))[0])
        self.update_output_field()
        self.status_field.setText("Status: Idle")

    def new_file_changed(self) -> None:
        if self.new_file.isChecked():
            self.output_file_field.setDisabled(False)
        else:
            self.output_file_field.setDisabled(True)

        self.update_output_field()

    def update_output_field(self) -> None:
        """
        Automatically updates the output field to correspond to the input file if the output field has not been
        manually overridden.
        """
        if self.new_file.isChecked():
            if not self.output_overridden:
                input_file = self.input_file_field.text().rsplit('.', 1)
                if len(input_file) > 1:
                    self.output_file_field.setText(f"{input_file[0]}_scores.{input_file[1]}")
        else:
            self.output_file_field.setText(self.input_file_field.text())

    def output_manually_changed(self) -> None:
        """
        Indicates that the output field has been manually changed, preventing it from automatically being changed again.
        """
        self.output_overridden = True

    def calculate_function(self) -> None:
        """
        Calculates the scores of the input file and updates the status in the window.
        """
        if not self.obedience_button.isChecked() and not self.rally_button.isChecked():
            self.status_field.setText(f"Error: Competition type required")
            return
        elif self.obedience_button.isChecked():
            type = "obedience"
        else:
            type = "rally"

        self.status_field.setText("Status: Calculating")
        try:
            status = calculate(self.input_file_field.text(), self.output_file_field.text(), type)

            if status:
                self.status_field.setText("Status: Complete")
            else:
                self.status_field.setText(f"Error: Unknown")

        except Exception as e:
            self.status_field.setText(f"Error: {e.args[-1]}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    app.exec_()
