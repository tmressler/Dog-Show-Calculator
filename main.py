from os.path import expanduser
from PyQt5.QtWidgets import *

from calculator import calculate

if __name__ == '__main__':
    home_directory = expanduser('~')

    def selectFile():
        file_window = QWidget()
        file_layout = QVBoxLayout()
        model = QDirModel()
        view = QTreeView()
        view.setModel(model)
        view.setRootIndex(model.index(home_directory))

        file_layout.addWidget(view)

        file_window.setLayout(file_layout)
        file_window.show()

    app = QApplication([])
    window = QWidget()
    layout = QVBoxLayout()

    select_file_button = QPushButton('open')
    select_file_button.clicked.connect(selectFile)
    calculate_button = QPushButton('Calculate')

    layout.addWidget(QLabel('Input File'))
    layout.addWidget(QLineEdit())
    layout.addWidget(select_file_button)
    layout.addWidget(QLabel('Output File'))
    layout.addWidget(QLineEdit())
    layout.addWidget(calculate_button)

    window.setLayout(layout)
    window.show()
    app.exec()