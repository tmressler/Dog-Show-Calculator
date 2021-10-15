import sys
import os
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PyQt5.QtGui import QIcon

from calculator import calculate


class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.output_overridden = False

        # set up widgets
        self.input_file_field = QLineEdit()
        self.input_file_field.editingFinished.connect(self.update_output_field)
        select_file_button = QPushButton()
        select_file_button.setIcon(QIcon("document-open.svg"))
        select_file_button.clicked.connect(self.select_file)
        self.output_file_field = QLineEdit()
        self.output_file_field.editingFinished.connect(self.output_manually_changed)
        self.return_multiple_away_winners = QCheckBox()
        self.return_multiple_away_winners.setText('Return multiple award winners (when tied)')
        self.return_multiple_away_winners.setChecked(True)
        calculate_button = QPushButton('Calculate')
        calculate_button.clicked.connect(self.calculate_function)
        self.status_field = QLabel()
        self.status_field.setAlignment(QtCore.Qt.AlignCenter)
        self.status_field.setText("Status: Idle")

        # add widgets to the layout
        layout = QVBoxLayout()
        layout.addWidget(QLabel('Input File'))
        input_field_layout = QHBoxLayout()
        input_field_layout.addWidget(self.input_file_field)
        input_field_layout.addWidget(select_file_button)
        layout.addLayout(input_field_layout)
        layout.addWidget(QLabel('Output File'))
        layout.addWidget(self.output_file_field)
        layout.addWidget(QLabel('Options'))
        layout.addWidget(self.return_multiple_away_winners)
        status_layout = QHBoxLayout()
        status_layout.addWidget(calculate_button)
        status_layout.addWidget(self.status_field)
        layout.addLayout(status_layout)

        # set layout
        self.setLayout(layout)

        # finish setting up some other settings
        self.setWindowTitle("Grand Champion Calculator")
        self.setWindowIcon(QIcon("trip.ico"))
        self.setGeometry(100, 100, 128*3, 0)

    def select_file(self) -> None:
        """
        Opens a window to allow the user to select a file, then updates the output field.
        """
        self.status_field.setText("Status: Waiting for input file")
        self.input_file_field.setText(QFileDialog.getOpenFileName(self, 'Open File', os.getenv('HOME'))[0])
        self.update_output_field()
        self.status_field.setText("Status: Idle")

    def update_output_field(self) -> None:
        """
        Automatically updates the output field to correspond to the input file if the output field has not been
        manually overriden.
        """
        if not self.output_overridden:
            input_file = self.input_file_field.text().rsplit('.', 1)
            if len(input_file) > 1:
                self.output_file_field.setText(f"{input_file[0]}_scores.{input_file[1]}")

    def output_manually_changed(self) -> None:
        """
        Indicates that the output field has been manually changed, preventing it from automatically being changed again.
        """
        self.output_overridden = True

    def calculate_function(self) -> None:
        """
        Calculates the scores of the input file and updates the status in the window.
        """
        self.status_field.setText("Status: Calculating")
        status = calculate(self.input_file_field.text(), self.output_file_field.text())
        if status:
            self.status_field.setText("Status: Complete")
        else:
            self.status_field.setText("Error")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()
