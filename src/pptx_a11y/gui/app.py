"""Application entry point for pptx-a11y GUI."""
import sys
from PySide6.QtWidgets import QApplication
from .main_window import MainWindow


def create_app(argv=None) -> QApplication:
    app = QApplication.instance()
    if not app or not isinstance(app, QApplication):
        app = QApplication(argv or sys.argv)
    app.setApplicationName("pptx-a11y")
    app.setOrganizationName("ThirstyHead")
    return app


def main():
    app = create_app()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
