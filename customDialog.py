from PyQt6.QtWidgets import QDialog, QLabel, QVBoxLayout, QPushButton


class CustomDialog(QDialog):
    def __init__(self, parent, title="Диалог", message="", show_yes_no=False):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.label = QLabel(message, self)
        layout = QVBoxLayout(self)
        layout.addWidget(self.label)
        if not show_yes_no:
            self.button_ok = QPushButton("Ok", self)
            self.button_ok.clicked.connect(self.accept)
            layout.addWidget(self.button_ok)

        if show_yes_no:
            self.button_yes = QPushButton("Да", self)
            self.button_no = QPushButton("Нет", self)
            self.button_yes.clicked.connect(self.yes_clicked)
            self.button_no.clicked.connect(self.no_clicked)

            layout.addWidget(self.button_yes)
            layout.addWidget(self.button_no)

        self.setMinimumWidth(300)
        self.adjustSize()

    def yes_clicked(self):
        print("Пользователь выбрал 'Да'")
        self.accept()

    def no_clicked(self):
        print("Пользователь выбрал 'Нет'")
        self.reject()

    def close(self):
        self.reject()