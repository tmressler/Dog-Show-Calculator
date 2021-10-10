import sys
from os.path import expanduser
from PyQt5.QtWidgets import *
from PyQt5.Qt import QRect

from calculator import calculate


class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.output_overridden = False

        # set up widgets
        self.input_file_field = QLineEdit()
        self.input_file_field.editingFinished.connect(self.update_output_field)
        select_file_button = QPushButton('open')
        select_file_button.clicked.connect(self.select_file)
        self.output_file_field = QLineEdit()
        self.output_file_field.editingFinished.connect(self.output_manually_changed)
        calculate_button = QPushButton('Calculate')
        calculate_button.clicked.connect(self.calculate_function)

        # add widgets to the layout
        layout = QVBoxLayout()
        layout.addWidget(QLabel('Input File'))
        layout.addWidget(self.input_file_field)
        layout.addWidget(select_file_button)
        layout.addWidget(QLabel('Output File'))
        layout.addWidget(self.output_file_field)
        layout.addWidget(calculate_button)

        # set layout
        self.setLayout(layout)

    def select_file(self) -> None:
        window = FileSelectionWindow()
        window.setGeometry(QRect(100, 100, 100, 100))
        window.show()
        print("select complete")

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
        if calculate(self.input_file_field.text(), self.output_file_field.text()):
            print("Success")
        else:
            print("Failed")


class FileSelectionWindow(QWidget):
    def __init__(self):
        QWidget.__init__(self)

        home_directory = expanduser('~')

        layout = QVBoxLayout()
        model = QDirModel()
        view = QTreeView()
        view.setModel(model)
        view.setRootIndex(model.index(home_directory))

        layout.addWidget(view)

        self.setLayout(layout)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()
