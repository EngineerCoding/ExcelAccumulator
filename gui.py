import os
import sys
import traceback

from PyQt5.QtCore import pyqtSlot, QByteArray, Qt
from PyQt5.QtGui import QMovie, QIcon
from PyQt5.QtWidgets import (
    QWidget, QApplication, QDesktopWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QSizePolicy, QDialog, QPlainTextEdit, QHBoxLayout)

from accumulator import accumulate_sheets


def get_path(path: str) -> str:
    base_path = getattr(sys, '_MEIPASS', os.getcwd())
    return os.path.join(base_path, path)


class ImagePlayer(QWidget):
    """ src: https://gist.github.com/Svenito/4000025 """

    def __init__(self, filename, parent=None):
        super().__init__(parent)

        # Load the file into a QMovie
        self.movie = QMovie(filename, QByteArray(), self)

        size = self.movie.scaledSize()
        self.setGeometry(200, 200, size.width(), size.height())

        self.movie_screen = QLabel()
        # Make label fit the gif
        self.movie_screen.setSizePolicy(
            QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.movie_screen.setAlignment(Qt.AlignCenter)

        # Create the layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.movie_screen)

        self.setLayout(main_layout)

        # Add the QMovie object to the label
        self.movie.setCacheMode(QMovie.CacheAll)
        self.movie.setSpeed(100)
        self.movie_screen.setMovie(self.movie)
        self.movie.loopCount = -1
        self.movie.start()


class BaseDialog(QDialog):

    def __init__(self, *args, close_text='Ok', **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowIcon(QIcon(get_path('icon.ico')))
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.close_button = QPushButton(close_text)
        self.close_button.clicked.connect(self.do_close)

    @pyqtSlot()
    def do_close(self):
        self.close()


class DoneDialog(BaseDialog):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle('Klaar!')
        self.setMinimumWidth(250)
        self.setMaximumWidth(500)
        self.setMaximumHeight(500)

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Uitvoeren is afgerond'))
        layout.addWidget(self.close_button)
        self.setLayout(layout)


class ErrorDialog(BaseDialog):

    def __init__(self, *args, exception_text=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle('Error!')

        main_layout = QVBoxLayout()
        main_layout.addWidget(
            QLabel('Er heeft zich een error voorgedaan. Email de volgende '
                   'tekst naar de ontwikkelaar:'))
        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlainText(exception_text)
        self.text_edit.setReadOnly(True)
        main_layout.addWidget(self.text_edit)

        copy_button = QPushButton('Kopiëren')
        copy_button.clicked.connect(self.copy_text)

        button_layout = QHBoxLayout()
        button_layout.addWidget(copy_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    @pyqtSlot()
    def copy_text(self):
        QApplication.clipboard().setText(self.text_edit.toPlainText())


class ExcelAccumulator(QWidget):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle('Excel accumulator')
        self.setWindowIcon(QIcon(get_path('icon.ico')))
        self.setMinimumWidth(640)

        # Actual data used
        self._input_file_path = None
        self._output_file_path = None

        layout = QVBoxLayout()
        # Input
        self._selected_file_header = QLabel('Geselecteerd bestand:')
        self._selected_file_label = QLabel('')
        self._input_button = QPushButton('Selecteer XLS(X) bestand')
        self._input_button.clicked.connect(self.select_input_excel)
        layout.addWidget(self._selected_file_header)
        layout.addWidget(self._selected_file_label)
        layout.addWidget(self._input_button)
        # Output
        self._output_file_header = QLabel('Opslaan als:')
        self._output_file_label = QLabel()
        self._output_button = QPushButton('Opslaan als..')
        self._output_button.clicked.connect(self.select_output)
        layout.addWidget(self._output_file_header)
        layout.addWidget(self._output_file_label)
        layout.addWidget(self._output_button)
        # Run
        self._run_button = QPushButton('Uitvoeren')
        self._run_button.clicked.connect(self.run)
        self._movie = ImagePlayer(get_path('loading.gif'))
        layout.addWidget(self._run_button)
        layout.addWidget(self._movie)

        self.setLayout(layout)
        self.reset()
        self.show()

    def reset(self):
        self._selected_file_header.hide()
        self._selected_file_label.hide()
        self._output_file_header.hide()
        self._output_file_label.hide()
        self._output_button.setEnabled(False)
        self._run_button.setEnabled(False)
        self._movie.hide()

        self._input_button.show()
        self._output_button.show()
        self._run_button.show()

        self.adjustSize()
        self.center()

    @staticmethod
    def _update_label(header: QLabel, value_label: QLabel, value: str):
        header.show()
        value_label.setText(value)
        value_label.show()

    @pyqtSlot()
    def select_input_excel(self):
        file_name = QFileDialog.getOpenFileName(
            self, 'Selecteer XLS(X) bestand', '', '*.xls *.xlsx')[0]
        if file_name:
            ExcelAccumulator._update_label(
                self._selected_file_header, self._selected_file_label,
                file_name)
            self._input_file_path = file_name
            self._output_button.setEnabled(True)

    @pyqtSlot()
    def select_output(self):
        file_name = QFileDialog.getSaveFileName(
            self, 'Opslaan als..', '', '*.xlsx')[0]
        if file_name:
            if not file_name.endswith('.xlsx'):
                if file_name.endswith('.xls'):
                    file_name += 'x'
                else:
                    file_name += '.xlsx'
            ExcelAccumulator._update_label(
                self._output_file_header, self._output_file_label,
                file_name)
            self._output_file_path = file_name
            self._run_button.setEnabled(True)

    @pyqtSlot()
    def run(self):
        # Show the loading gif
        self.reset()
        self._input_button.hide()
        self._output_button.hide()
        self._run_button.hide()
        self._movie.show()
        self.center()
        # execute
        try:
            accumulate_sheets(self._input_file_path, self._output_file_path)
        except Exception:
            formatted_exception = traceback.format_exc()
            ErrorDialog(exception_text=formatted_exception).exec_()
        else:
            DoneDialog().exec_()
        self.reset()

    def center(self):
        resolution = QDesktopWidget().screenGeometry()
        window_resolution = self.frameSize()
        self.move(
            resolution.width() / 2 - window_resolution.width() / 2,
            resolution.height() / 2 - window_resolution.height() / 2)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelAccumulator()
    sys.exit(app.exec_())
