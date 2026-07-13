import sys
import time
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QRect
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QVBoxLayout, QWidget, QFileDialog, QStatusBar, QToolBar, QFontDialog, QAction, QComboBox, QSpinBox, QPushButton, QDialog, QLabel, QHBoxLayout, QScrollBar, QMenu, QInputDialog, QTabWidget, QTextBrowser, QLineEdit
from PyQt5.QtGui import QFont, QColor, QTextCharFormat, QIcon, QTextCursor, QFontDatabase
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QIcon
from docx import Document  # for exporting to Word


# Copyright 2026 Jasper K 
# Any distribution of this software is prohibited.


class CustomFontDialog(QDialog):
    def __init__(self, current_font, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Font")
        self.setStyleSheet("""
            QDialog {
                background-color: #1e1e1e;
                color: #d4d4d4;
                border-radius: 10px;
            }
            QLabel {
                color: #d4d4d4;
                font-size: 14px;
            }
            QComboBox {
                background-color: #2e2e2e;
                color: #d4d4d4;
                border: 1px solid #444;
                border-radius: 5px;
                padding: 5px;
            }
            QComboBox QAbstractItemView {
                background-color: #2e2e2e;
                color: #d4d4d4;
                selection-background-color: #444;
                selection-color: #fff;
            }
            QSpinBox {
                background-color: #2e2e2e;
                color: #d4d4d4;
                border: 1px solid #444;
                border-radius: 5px;
                padding: 5px;
            }
            QPushButton {
                background-color: #3e3e3e;
                color: #d4d4d4;
                border: 1px solid #555;
                border-radius: 5px;
                padding: 5px 10px;
            }
            QPushButton:hover {
                background-color: #555;
            }
        """)

        layout = QVBoxLayout(self)

        # Font selection
        self.font_label = QLabel("Choose Font:", self)
        layout.addWidget(self.font_label)

        self.font_combobox = QComboBox(self)
        self.font_combobox.addItems(QFontDatabase().families())  # List all system fonts
        self.font_combobox.setCurrentText(current_font.family())
        layout.addWidget(self.font_combobox)

        # Font size selection
        self.font_size_label = QLabel("Font Size:", self)
        layout.addWidget(self.font_size_label)

        self.font_size_spinbox = QSpinBox(self)
        self.font_size_spinbox.setRange(5, 72)
        self.font_size_spinbox.setValue(current_font.pointSize())
        layout.addWidget(self.font_size_spinbox)

        # Apply button
        self.apply_button = QPushButton("Apply", self)
        self.apply_button.clicked.connect(self.accept)
        layout.addWidget(self.apply_button)

    def get_selected_font(self):
        font_name = self.font_combobox.currentText()
        font_size = self.font_size_spinbox.value()
        return QFont(font_name, font_size)


class Crystal(QMainWindow):
    def __init__(self):
        super().__init__()

        self.dark_mode = True  # track current theme
        self.fullscreen = False

        self.setMinimumWidth(1800)
        self.setMinimumHeight(1000)

        self.setWindowTitle('crystal')
        self.setWindowIcon(QIcon('crystal.ico'))

        # Initial dark theme for main window
        self.setStyleSheet("background-color: #111;")
        self.init_text_edit()
        self.init_central_widget()
        self.init_status_bar()
        self.create_toolbar()
        self.toolbar.hide()  # Initially hide the toolbar

        # Timer to auto-hide the toolbar after inactivity
        self.toolbar_hide_timer = QTimer()
        self.toolbar_hide_timer.setSingleShot(True)
        self.toolbar_hide_timer.timeout.connect(self.hide_toolbar)

        # Set up a timer for showing the toolbar when mouse is near the top
        self.toolbar_timer = QTimer(self)
        self.toolbar_timer.setInterval(50)
        self.toolbar_timer.timeout.connect(self.check_mouse_position)
        self.toolbar_timer.start()

        # Variable to track whether the toolbar is currently shown
        self.toolbar_shown = False

        self.init_shortcuts()  # Initialize keyboard shortcuts

    def init_text_edit(self):
        self.text_edit = QTextEdit(self)
        self.text_edit.setFont(QFont("Arial", 18))  # Set default font to Arial, size 18
        # Dark theme style for text edit.
        self.text_edit.setStyleSheet("background-color: #111; color: #eee; border: none;")
        self.text_edit.setAcceptRichText(True)  # Enable rich text for links
        self.text_edit.cursorPositionChanged.connect(self.highlight_current_line)
        self.text_edit.textChanged.connect(self.update_status_bar)
        self.text_edit.setContextMenuPolicy(Qt.CustomContextMenu)
        self.text_edit.customContextMenuRequested.connect(self.show_floating_menu)

    def init_central_widget(self):
        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(50, 50, 50, 50)  # Add ample margins for focus
        layout.addWidget(self.text_edit)
        self.setCentralWidget(central_widget)

    def init_status_bar(self):
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("color: #ccc; background-color: #222;")
        self.setStatusBar(self.status_bar)

    def create_toolbar(self):
        self.toolbar = QToolBar()
        self.toolbar.setMovable(False)
        # Modern VS Code-like styling for the toolbar.
        self.toolbar.setStyleSheet("""
            QToolBar {
                background: #1e1e1e;
                spacing: 10px;
                border: none;
                padding: 5px;
            }
            QToolButton {
                color: #d4d4d4;
                background: transparent;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-size: 14px;
            }
            QToolButton:hover {
                background: #333333;
            }
            QToolButton:pressed {
                background: #444444;
            }
        """)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        # File actions
        open_action = QAction("Open", self)
        open_action.triggered.connect(self.open_file)
        self.toolbar.addAction(open_action)

        save_action = QAction("Save", self)
        save_action.triggered.connect(self.save_file)
        self.toolbar.addAction(save_action)

        export_txt_action = QAction("Export as TXT", self)
        export_txt_action.triggered.connect(self.export_txt)
        self.toolbar.addAction(export_txt_action)

        export_pdf_action = QAction("Export as PDF", self)
        export_pdf_action.triggered.connect(self.export_pdf)
        self.toolbar.addAction(export_pdf_action)

        export_word_action = QAction("Export as Word", self)
        export_word_action.triggered.connect(self.export_word)
        self.toolbar.addAction(export_word_action)

        self.toolbar.addSeparator()

        # Settings actions
        font_action = QAction("Change Font", self)
        font_action.triggered.connect(self.change_font)
        self.toolbar.addAction(font_action)

        theme_action = QAction("Toggle Theme", self)
        theme_action.triggered.connect(self.toggle_theme)
        self.toolbar.addAction(theme_action)

        self.toolbar.addSeparator()

        fullscreen_action = QAction("Toggle Fullscreen", self)
        fullscreen_action.triggered.connect(self.toggle_fullscreen)
        self.toolbar.addAction(fullscreen_action)

        self.toolbar.addSeparator()

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.close)
        self.toolbar.addAction(quit_action)

    def init_shortcuts(self):
        # Keyboard shortcuts
        save_shortcut = QAction(self)
        save_shortcut.setShortcut("Ctrl+S")
        save_shortcut.triggered.connect(self.save_file)
        self.addAction(save_shortcut)

        open_shortcut = QAction(self)
        open_shortcut.setShortcut("Ctrl+O")
        open_shortcut.triggered.connect(self.open_file)
        self.addAction(open_shortcut)

        bold_shortcut = QAction(self)
        bold_shortcut.setShortcut("Ctrl+B")
        bold_shortcut.triggered.connect(self.toggle_bold)
        self.addAction(bold_shortcut)

        italic_shortcut = QAction(self)
        italic_shortcut.setShortcut("Ctrl+I")
        italic_shortcut.triggered.connect(self.toggle_italic)
        self.addAction(italic_shortcut)

        underline_shortcut = QAction(self)
        underline_shortcut.setShortcut("Ctrl+U")
        underline_shortcut.triggered.connect(self.toggle_underline)
        self.addAction(underline_shortcut)

    def toggle_bold(self):
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            fmt = cursor.charFormat()
            fmt.setFontWeight(QFont.Bold if fmt.fontWeight() != QFont.Bold else QFont.Normal)
            cursor.mergeCharFormat(fmt)

    def toggle_italic(self):
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            fmt = cursor.charFormat()
            fmt.setFontItalic(not fmt.fontItalic())
            cursor.mergeCharFormat(fmt)

    def toggle_underline(self):
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            fmt = cursor.charFormat()
            fmt.setFontUnderline(not fmt.fontUnderline())
            cursor.mergeCharFormat(fmt)

    def toggle_strikethrough(self):
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            fmt = cursor.charFormat()
            fmt.setFontStrikeOut(not fmt.fontStrikeOut())
            cursor.mergeCharFormat(fmt)

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Text Files (*.txt)")
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    self.text_edit.setPlainText(file.read())
            except Exception as e:
                print("Error opening file:", e)

    def save_file(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Text Files (*.txt)")
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(self.text_edit.toPlainText())
            except Exception as e:
                print("Error saving file:", e)

    def toggle_fullscreen(self):
        if self.fullscreen:
            self.showNormal()
        else:
            self.showFullScreen()
        self.fullscreen = not self.fullscreen

    def export_txt(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Export File as TXT", "", "Text Files (*.txt)")
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(self.text_edit.toPlainText())
            except Exception as e:
                print("Error exporting file:", e)

    def export_pdf(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Export File as PDF", "", "PDF Files (*.pdf)")
        if file_path:
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_path)
            self.text_edit.document().print_(printer)

    def export_word(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Export File as Word", "", "Word Documents (*.docx)")
        if file_path:
            try:
                doc = Document()
                # Add plain text into the Word document.
                doc.add_paragraph(self.text_edit.toPlainText())
                doc.save(file_path)
            except Exception as e:
                print("Error exporting Word file:", e)

    def change_font(self):
        current_font = self.text_edit.font()
        font_dialog = CustomFontDialog(current_font, self)
        if font_dialog.exec_():
            selected_font = font_dialog.get_selected_font()
            self.text_edit.setFont(selected_font)

    def toggle_theme(self):
        if self.dark_mode:
            # Switch to light theme.
            self.dark_mode = False
            self.setStyleSheet("background-color: #fff;")
            self.text_edit.setStyleSheet("background-color: #fff; color: #111; border: none;")
            self.status_bar.setStyleSheet("color: #111; background-color: #ddd;")
            self.toolbar.setStyleSheet("""
                QToolBar {
                    background: #f3f3f3;
                    spacing: 10px;
                    border: none;
                    padding: 5px;
                }
                QToolButton {
                    color: #333;
                    background: transparent;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 10px;
                    font-size: 14px;
                }
                QToolButton:hover {
                    background: #e0e0e0;
                }
                QToolButton:pressed {
                    background: #d0d0d0;
                }
            """)
        else:
            # Switch back to dark theme.
            self.dark_mode = True
            self.setStyleSheet("background-color: #111;")
            self.text_edit.setStyleSheet("background-color: #111; color: #eee; border: none;")
            self.status_bar.setStyleSheet("color: #ccc; background-color: #222;")
            self.toolbar.setStyleSheet("""
                QToolBar {
                    background: #1e1e1e;
                    spacing: 10px;
                    border: none;
                    padding: 5px;
                }
                QToolButton {
                    color: #d4d4d4;
                    background: transparent;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 10px;
                    font-size: 14px;
                }
                QToolButton:hover {
                    background: #333333;
                }
                QToolButton:pressed {
                    background: #444444;
                }
            """)

    def highlight_current_line(self):
        extra_selections = []
        selection = QTextEdit.ExtraSelection()
        selection.format.setBackground(QColor(80, 80, 120, 100))  # semi-transparent highlight
        selection.format.setProperty(QTextCharFormat.FullWidthSelection, True)
        selection.cursor = self.text_edit.textCursor()
        extra_selections.append(selection)
        self.text_edit.setExtraSelections(extra_selections)

    def update_status_bar(self):
        text = self.text_edit.toPlainText()
        char_count = len(text)
        word_count = len(text.split())
        reading_time = self.calculate_reading_time(char_count)
        self.status_bar.showMessage(
            f"Words: {word_count} | Characters: {char_count} | Reading Time: {reading_time}"
        )

    def calculate_reading_time(self, char_count):
        # Assuming average reading speed is 300 characters per minute
        total_seconds = (char_count / 300) * 60
        return time.strftime("%H:%M:%S", time.gmtime(total_seconds))

    def check_mouse_position(self):
        # Track mouse position and show toolbar if near top
        cursor_pos = QApplication.desktop().cursor().pos()
        if (cursor_pos.y() < 50):  # Show toolbar when mouse is near the top
            if not self.toolbar_shown:
                self.show_toolbar()
                self.toolbar_shown = True
            if self.toolbar_hide_timer.isActive():
                self.toolbar_hide_timer.stop()  # Stop hide timer if mouse is near top
        else:
            if self.toolbar_shown:
                if not self.toolbar_hide_timer.isActive():
                    self.toolbar_hide_timer.start(6000)  # Start the timer to hide the toolbar after 6 seconds

    def show_toolbar(self):
        self.toolbar.show()
        # Slide down animation for toolbar
        self.toolbar_animation = QPropertyAnimation(self.toolbar, b"geometry")
        self.toolbar_animation.setDuration(300)
        self.toolbar_animation.setStartValue(QRect(0, -self.toolbar.height(), self.width(), self.toolbar.height()))
        self.toolbar_animation.setEndValue(QRect(0, 0, self.width(), self.toolbar.height()))
        self.toolbar_animation.start()

    def hide_toolbar(self):
        # Slide up animation for toolbar
        self.toolbar_animation = QPropertyAnimation(self.toolbar, b"geometry")
        self.toolbar_animation.setDuration(300)
        self.toolbar_animation.setStartValue(self.toolbar.geometry())
        self.toolbar_animation.setEndValue(QRect(0, -self.toolbar.height(), self.width(), self.toolbar.height()))
        self.toolbar_animation.start()
        self.toolbar_shown = False

    def show_floating_menu(self, position):
        cursor = self.text_edit.textCursor()
        if not cursor.hasSelection():
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #333;
                color: #eee;
                border: 1px solid #444;
                border-radius: 8px;
            }
            QMenu::item {
                padding: 8px 16px;
            }
            QMenu::item:selected {
                background-color: #555;
            }
        """)

        # Add formatting options
        h1_action = QAction("H1", self)
        h1_action.triggered.connect(lambda: self.apply_heading(24))
        menu.addAction(h1_action)

        h2_action = QAction("H2", self)
        h2_action.triggered.connect(lambda: self.apply_heading(20))
        menu.addAction(h2_action)

        h3_action = QAction("H3", self)
        h3_action.triggered.connect(lambda: self.apply_heading(16))
        menu.addAction(h3_action)

        menu.addSeparator()

        bold_action = QAction("Bold", self)
        bold_action.triggered.connect(self.toggle_bold)
        menu.addAction(bold_action)

        italic_action = QAction("Italic", self)
        italic_action.triggered.connect(self.toggle_italic)
        menu.addAction(italic_action)

        underline_action = QAction("Underline", self)
        underline_action.triggered.connect(self.toggle_underline)
        menu.addAction(underline_action)

        strikethrough_action = QAction("Strikethrough", self)
        strikethrough_action.triggered.connect(self.toggle_strikethrough)
        menu.addAction(strikethrough_action)

        menu.addSeparator()

        link_action = QAction("Insert Link", self)
        link_action.triggered.connect(self.insert_link)
        menu.addAction(link_action)

        # Adjust menu position to appear below the selected text
        cursor_rect = self.text_edit.cursorRect(cursor)
        menu_position = self.text_edit.viewport().mapToGlobal(cursor_rect.bottomLeft())
        menu.exec_(menu_position)

    def apply_heading(self, font_size):
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            fmt = cursor.charFormat()
            fmt.setFontPointSize(font_size)
            cursor.mergeCharFormat(fmt)

    def insert_link(self):
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            dialog = QDialog(self)
            dialog.setWindowTitle("Insert Link")
            dialog.setStyleSheet("""
                QDialog {
                    background-color: #222;
                    color: #eee;
                    border-radius: 10px;
                }
                QLabel {
                    color: #eee;
                }
                QLineEdit {
                    background-color: #333;
                    color: #eee;
                    border: 1px solid #444;
                    border-radius: 5px;
                    padding: 5px;
                }
                QPushButton {
                    background-color: #444;
                    color: #eee;
                    border: none;
                    border-radius: 5px;
                    padding: 5px 10px;
                }
                QPushButton:hover {
                    background-color: #555;
                }
            """)

            layout = QVBoxLayout(dialog)

            label = QLabel("Enter URL:", dialog)
            layout.addWidget(label)

            url_input = QLineEdit(dialog)
            layout.addWidget(url_input)

            button_layout = QHBoxLayout()
            apply_button = QPushButton("Apply", dialog)
            apply_button.clicked.connect(lambda: self.apply_link(cursor, url_input.text(), dialog))
            button_layout.addWidget(apply_button)

            cancel_button = QPushButton("Cancel", dialog)
            cancel_button.clicked.connect(dialog.reject)
            button_layout.addWidget(cancel_button)

            layout.addLayout(button_layout)
            dialog.exec_()

    def apply_link(self, cursor, url, dialog):
        if url:
            fmt = cursor.charFormat()
            fmt.setAnchor(True)
            fmt.setAnchorHref(url)
            fmt.setForeground(QColor("blue"))
            fmt.setFontUnderline(True)
            cursor.mergeCharFormat(fmt)
        dialog.accept()

    def process_markup(self):
        text = self.text_edit.toPlainText()
        # Replace markup patterns with corresponding rich text
        text = text.replace("~", "<s>") if "~" in text else text.replace("</s>", "~")
        text = text.replace("*", "<b>") if "*" in text else text.replace("</b>", "*")
        text = text.replace("_", "<i>") if "_" in text else text.replace("</i>", "_")
        self.text_edit.setHtml(text)

    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            self.show_floating_menu(event.globalPos())

    def mousePressEvent(self, event):
        if event.modifiers() == Qt.ControlModifier:  # Check if Ctrl is held
            cursor = self.text_edit.cursorForPosition(self.text_edit.mapFromGlobal(event.globalPos()))
            if cursor.charFormat().isAnchor():  # Check if the clicked text is a link
                url = cursor.charFormat().anchorHref()
                self.open_link(url)
                return  # Prevent further processing of the event
        super().mousePressEvent(event)

    def open_link(self, url):
        # Open the link in the default web browser
        import webbrowser
        webbrowser.open(url)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Crystal()
    window.show()
    sys.exit(app.exec_())
