import sys
import os
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit, 
                             QFileDialog, QWidget, QGridLayout, QLineEdit, QScrollArea, QStatusBar, QSizePolicy, QComboBox, QCompleter,QSplitter, QMenu, QAction
)
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon
from PyQt5.QtCore import Qt, QTimer, QSize, QEvent  # Import QEvent
from functools import partial

class MessageTool(QMainWindow):
    def __init__(self):
        super().__init__()

        self.current_index = -1  # Özniteliği burada tanımlayın

        self.initUI()
        
        self.input_folder = ""
        self.output_folder = ""
        self.files = []
        
        self.initial_load_done = False  # Flag to track if initial load is done

        self.folder_indices = {}  # Dictionary to store the current index for each folder
       
        self.read_files = []  # List to store read files

        self.product_name_list = []  # List to store unique product names

        self.brand_list = []  # List to store brand names
        self.category_list = []  # List to store category names
        self.color_list = []  # List to store color names

        self.load_table_headers()

        # Timer for auto save (initialize but do not start)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.auto_save)

        self.load_state()  # Load state after initializing timer
        self.update_status()
        self.start_auto_save()
        self.initial_load_done = True  # Set flag to True after initial load


    def initUI(self):
        self.setWindowTitle('Message Tool')
        self.setGeometry(100, 100, 1200, 600)
        
        # Set palette for the main window
        self.setAutoFillBackground(True)
        p = self.palette()
        p.setColor(QPalette.Window, QColor(220,220,220))
        self.setPalette(p)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # Main layout
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        self.setCentralWidget(main_widget)

        # Splitter for main layout
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # Left container with layout
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(10)
        left_layout.setContentsMargins(10, 10, 10, 10)
        

        self.message_label = QLabel('Message Content')
        self.message_label.setFont(QFont('Arial', 15, QFont.Bold))
        left_layout.addWidget(self.message_label)

        # View mode combo box
        self.view_mode = QComboBox(self)
        self.view_mode.addItems(["Tablo", "Processed Text"])
        self.view_mode.currentTextChanged.connect(self.switch_view_mode)
        left_layout.addWidget(self.view_mode)

        # Using QScrollArea for the message_text
        self.message_text = QTextEdit(self)
        self.message_text.setFont(QFont('Seoul', 14))
        self.message_text.setStyleSheet("""
            padding: 15px; 
            background-color: #FFFFFF; 
            border: 1px solid #DDDDDD; 
            border-radius: 8px;
            margin-bottom: 10px;
            font-family: 'Roboto', sans-serif;
            font-size: 18px;
        """)

        self.message_text.setLineWrapMode(QTextEdit.NoWrap)
        self.message_text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.message_text.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        message_scroll = QScrollArea()
        message_scroll.setWidget(self.message_text)
        message_scroll.setWidgetResizable(True)
        left_layout.addWidget(message_scroll)

        # Processed Text area (initially hidden)
        self.processed_text = QTextEdit(self)
        self.processed_text.setFont(QFont('Seoul', 14))
        self.processed_text.setStyleSheet("""
            padding: 15px; 
            background-color: #FFFFFF; 
            border: 1px solid #DDDDDD; 
            border-radius: 8px;
            margin-bottom: 10px;
            font-family: 'Roboto', sans-serif;
            font-size: 18px;
        """)
        self.processed_text.setLineWrapMode(QTextEdit.NoWrap)
        self.processed_text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.processed_text.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_layout.addWidget(self.processed_text)
        self.processed_text.hide()  # Initially hide the processed text area
        self.processed_text.textChanged.connect(self.save_state)  # Processed Text değişikliklerini kaydetmek için


        self.file_path_label = QLabel('')
        self.file_path_label.setFont(QFont('Seoul', 12))
        left_layout.addWidget(self.file_path_label)
        self.file_path_label.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_path_label.customContextMenuRequested.connect(self.show_context_menu)

        # Read button and label
        self.read_button = self.create_read_button()
        left_layout.addWidget(self.read_button)

        self.read_label = QLabel('')
        self.read_label.setFont(QFont('Arial', 10))
        left_layout.addWidget(self.read_label)

        self.load_button = self.create_icon_button('icons/load_icon.png', self.select_input_folder)
        left_layout.addWidget(self.load_button)

        self.prev_button = self.create_icon_button('icons/prev_icon.png', self.show_prev_message)
        left_layout.addWidget(self.prev_button)

        self.next_button = self.create_icon_button('icons/next_icon.png', self.show_next_message)
        left_layout.addWidget(self.next_button)

        # QComboBox for file selection
        self.file_selector = QComboBox(self)
        self.file_selector.currentIndexChanged.connect(self.go_to_selected_file)
        left_layout.addWidget(self.file_selector)

        splitter.addWidget(left_container)

        # Right container with layout
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(10)
        right_layout.setContentsMargins(10, 10, 10, 10)

        self.scroll_area = QScrollArea()
        self.scroll_widget = QWidget()
        self.table_layout = QGridLayout(self.scroll_widget)
        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setWidgetResizable(True)
        right_layout.addWidget(self.scroll_area)

        button_layout = QHBoxLayout()
        self.add_row_button = self.create_icon_button('icons/add_icon.png', self.add_table_row, 40)
        button_layout.addWidget(self.add_row_button)

        self.save_button = self.create_icon_button('icons/save_icon.png', self.save_state)
        button_layout.addWidget(self.save_button)

        self.export_button = self.create_icon_button('icons/export_icon.png', self.export_table)
        button_layout.addWidget(self.export_button)

        right_layout.addLayout(button_layout)

        splitter.addWidget(right_container)
        
        # Set initial sizes of the splitter
        splitter.setSizes([300, 900])  # Adjust these values to set the initial width of the left and right panels

        self.update_read_button()  # Initialize button and label text


    def switch_view_mode(self):
        selected_mode = self.view_mode.currentText()
        if selected_mode == "Tablo":
            self.scroll_area.show()
            self.processed_text.hide()
        elif selected_mode == "Processed Text":
            self.scroll_area.hide()
            self.processed_text.show()
            # Processed text kutusunu temizle ve yükle
            if 0 <= self.current_index < len(self.files):
                current_file = self.files[self.current_index]
                file_path = os.path.join(self.input_folder, current_file + '.json')
                if os.path.exists(file_path):
                    with open(file_path, 'r', encoding='utf-8') as file:
                        data = json.load(file)
                        if isinstance(data, dict):
                            processed_text = data.get('processed_text', '')
                            self.processed_text.setText(processed_text)
                        else:
                            self.processed_text.setText('')
                else:
                    self.processed_text.setText('')



    
 
    def export_table(self):
        self.output_folder = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
        if self.output_folder:
            selected_mode = self.view_mode.currentText()
            if selected_mode == "Tablo":
                self.save_table()
            elif selected_mode == "Processed Text":
                self.export_processed_text()
            self.file_path_label.setText(f"Output Folder: {self.output_folder}")
            self.status_bar.showMessage(f"{selected_mode} exported successfully.", 5000)

    def export_processed_text(self):
        if 0 <= self.current_index < len(self.files):
            current_file = self.files[self.current_index]
            file_path = os.path.join(self.input_folder, current_file)
            processed_content = self.processed_text.toPlainText()
            output_path = os.path.join(self.output_folder, f"{os.path.splitext(current_file)[0]}_processed.txt")
            with open(output_path, 'w', encoding='utf-8') as file:
                file.write(processed_content)
            self.status_bar.showMessage("Processed text exported successfully.", 5000)


    def create_read_button(self):
        button = QPushButton("Not Readed")
        button.setFixedSize(100, 40)
        button.setStyleSheet("""
            QPushButton {
                background-color: red;
                color: white;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: darkred;
            }
        """)
        button.clicked.connect(self.mark_as_read)
        return button
    
    def update_read_button(self):
        if 0 <= self.current_index < len(self.files):
            file_name = self.files[self.current_index]
            file_path = os.path.join(self.input_folder, file_name)
            if file_path in self.read_files:
                self.read_button.setText("Readed")
                self.read_button.setStyleSheet("""
                    QPushButton {
                        background-color: #4CAF50;
                        color: white;
                        border-radius: 8px;
                        padding: 10px;
                        font-family: 'Roboto', sans-serif;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #45a049;
                    }
                """)
                self.read_label.setText("Readed")
            else:
                self.read_button.setText("Not Readed")
                self.read_button.setStyleSheet("""
                    QPushButton {
                        background-color: #FF4C4C;
                        color: white;
                        border-radius: 8px;
                        padding: 10px;
                        font-family: 'Roboto', sans-serif;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #E63E3E;
                    }
                """)
                self.read_label.setText("Not Readed")
        else:
            self.read_button.setText("Not Readed")
            self.read_button.setStyleSheet("""
                QPushButton {
                    background-color: #FF4C4C;
                    color: white;
                    border-radius: 8px;
                    padding: 10px;
                    font-family: 'Roboto', sans-serif;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #E63E3E;
                }
            """)
            self.read_label.setText("Not Readed")





    def mark_as_read(self):
        if 0 <= self.current_index < len(self.files):
            file_name = self.files[self.current_index]
            file_path = os.path.join(self.input_folder, file_name)
            if file_path in self.read_files:
                self.read_files.remove(file_path)
                self.read_button.setText("Not Readed")
                self.read_button.setStyleSheet("""
                    QPushButton {
                        background-color: #FF4C4C;
                        color: white;
                        border-radius: 8px;
                        padding: 10px;
                        font-family: 'Roboto', sans-serif;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #E63E3E;
                    }
                """)
                self.read_label.setText("Not Readed")
                self.status_bar.showMessage(f"Marked {file_name} as unread.", 5000)
            else:
                self.read_files.append(file_path)
                self.read_button.setText("Readed")
                self.read_button.setStyleSheet("""
                    QPushButton {
                        background-color: #4CAF50;
                        color: white;
                        border-radius: 8px;
                        padding: 10px;
                        font-family: 'Roboto', sans-serif;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #45a049;
                    }
                """)
                self.read_label.setText("Readed")
                self.status_bar.showMessage(f"Marked {file_name} as read.", 5000)
            self.save_state()  # Save state after marking as read/unread



    def show_context_menu(self, pos):
        context_menu = QMenu(self)
        copy_action = QAction('Kopyala', self)
        copy_action.triggered.connect(self.copy_file_path)
        context_menu.addAction(copy_action)
        context_menu.exec_(self.file_path_label.mapToGlobal(pos))

    def copy_file_path(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.file_path_label.text())
        self.status_bar.showMessage("Dosya yolu kopyalandı.", 5000)



    def create_right_widget(self):
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(10)
        right_layout.setContentsMargins(10, 10, 10, 10)

        self.scroll_area = QScrollArea()
        self.scroll_widget = QWidget()
        self.table_layout = QGridLayout(self.scroll_widget)
        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setWidgetResizable(True)
        right_layout.addWidget(self.scroll_area)

        button_layout = QHBoxLayout()
        self.add_row_button = self.create_icon_button('icons/add_icon.png', self.add_table_row, 40)
        button_layout.addWidget(self.add_row_button)

        self.save_button = self.create_icon_button('icons/save_icon.png', self.save_state)
        button_layout.addWidget(self.save_button)

        self.export_button = self.create_icon_button('icons/export_icon.png', self.export_table)
        button_layout.addWidget(self.export_button)

        right_layout.addLayout(button_layout)

        return right_widget

       


    def create_icon_button(self, icon_path, callback, size=40):
        button = QPushButton(self)
        button.setIcon(QIcon(icon_path))
        button.setIconSize(QSize(size, size))
        button.setFixedSize(size + 20, size + 20)  # Add padding around the icon
        button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
                border-radius: 5px;
            }
        """)
        button.clicked.connect(callback)
        return button

    def select_input_folder(self):
        # Mevcut klasörün indeksini kaydet
        if self.input_folder:
            self.folder_indices[self.input_folder] = self.current_index

        self.input_folder = QFileDialog.getExistingDirectory(self, 'Select Input Folder')
        self.load_files()
        
        # Yeni klasörün mevcut dosya indeksini yükle
        if self.input_folder in self.folder_indices:
            self.current_index = self.folder_indices[self.input_folder]
        else:
            self.current_index = 0 if self.files else -1  # Yeni klasör için indeks ayarla

        self.show_message()
        self.save_state()  # Yeni klasör yüklendikten sonra durumu kaydet
        self.file_path_label.setText(f"Input Folder: {self.input_folder}")
        self.status_bar.showMessage("Input folder selected.", 5000)

    def go_to_selected_file(self):
        selected_index = self.file_selector.currentIndex()
        if 0 <= selected_index < len(self.files):
            self.current_index = selected_index
            self.show_message()
            self.save_state()
            self.status_bar.showMessage(f"Showing file {self.current_index + 1} of {len(self.files)}.", 5000)



    def load_files(self):
        if self.input_folder:
            all_files = sorted([f for f in os.listdir(self.input_folder) if f.endswith('.txt')])
            self.files = all_files  # Update to load all files, including read files


            # QComboBox'ı numaralarla doldurun
            self.file_selector.clear()
            for i in range(len(self.files)):
                self.file_selector.addItem(f"File {i+1}")

            # Set current_index to the first unread file
            self.current_index = 0
            for i, file in enumerate(self.files):
                if os.path.join(self.input_folder, file) not in self.read_files:
                    self.current_index = i
                    break

            if self.files:
                self.show_message()
                self.update_status()
                self.status_bar.showMessage(f"{len(self.files)} files loaded.", 5000)
            else:
                self.current_index = -1  # Dosya yoksa geçersiz index
                self.status_bar.showMessage("No txt files found in the selected folder.", 5000)
            self.update_read_button()  # Read button textini güncelle


    def show_message(self):
        if 0 <= self.current_index < len(self.files):
            file_path = os.path.join(self.input_folder, self.files[self.current_index])
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.message_text.setText(content)
                file_name = self.files[self.current_index]  # Sadece dosya adı
                self.file_path_label.setText(f"Current File: {file_name}")  # Update file name label
                self.update_status()
                self.update_read_button()  # Update read button text
                self.status_bar.showMessage(f"Showing file {self.current_index + 1} of {len(self.files)}.", 5000)
                self.file_selector.setCurrentIndex(self.current_index)  # QComboBox'ı güncelleyin

                # Load table data for the current file
                self.load_table_data_for_current_file()
            else:
                self.status_bar.showMessage(f"File not found: {file_path}", 5000)
        else:
            self.message_text.setText("")
            self.file_path_label.setText("")
            self.processed_text.setText("")  # Temizle
            self.update_status()
            self.update_read_button()
            self.status_bar.showMessage("No files to display.", 5000)









    def load_table_data_for_current_file(self):
        if 0 <= self.current_index < len(self.files):
            current_file = self.files[self.current_index]
            file_path = os.path.join(self.input_folder, current_file)
            table_data_path = file_path + '.json'
            if os.path.exists(table_data_path):
                with open(table_data_path, 'r', encoding='utf-8') as file:
                    data = json.load(file)
                    if isinstance(data, list):  # Eğer data bir liste ise eski formatı dönüştür
                        table_data = data
                        processed_text = ''
                    else:  # Eğer data bir sözlük ise yeni formatı kullan
                        table_data = data.get('table_data', [])
                        processed_text = data.get('processed_text', '')
                    self.set_table_data(table_data)
                    self.processed_text.setText(processed_text)  # Load the processed text for the current file
            else:
                self.set_table_data([])  # Clear the table if no data is found
                self.processed_text.setText("")  # Clear processed text if no data is found
        # Add default row if no data is present
        if not self.table_data:
            self.add_table_row()





                

    
    def get_table_data(self):
        data = []
        for row_data in self.table_data:
            row = []
            for item in row_data:
                if isinstance(item, QComboBox):
                    row.append(item.currentText())
                elif isinstance(item, QLineEdit):
                    row.append(item.text())
            data.append(row)
        return data

    def set_table_data(self, data):
        # Clear existing table data
        for i in reversed(range(len(self.table_data))):
            self.delete_table_row(i)

        self.table_data = []
        num_headers = len(self.headers)  # Number of headers

        for row_position, row_data in enumerate(data):
            # Ensure row_data has the same length as headers by truncating or padding with empty strings
            if len(row_data) > num_headers:
                row_data = row_data[:num_headers]  # Truncate extra columns
            else:
                row_data.extend([""] * (num_headers - len(row_data)))  # Pad with empty strings

            row = []
            for col in range(num_headers):
                value = row_data[col]

                if self.headers[col].lower() == "offer type":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(["WTB", "WTS"])
                    combo_box.setCurrentText(value)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "product name":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.product_name_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.product_name_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_product_name_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "brand":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.brand_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.brand_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_brand_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "category":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.category_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.category_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_category_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "color":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.color_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.color_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_color_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                else:
                    line_edit = QLineEdit()
                    line_edit.setFont(QFont('Arial', 10))
                    line_edit.setText(value)
                    line_edit.setFixedHeight(40)
                    line_edit.setFixedWidth(120)
                    line_edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    line_edit.installEventFilter(self)
                    self.table_layout.addWidget(line_edit, row_position + 1, col)
                    row.append(line_edit)

            delete_button = self.create_icon_button('icons/delete_icon.png', partial(self.delete_table_row, row_position), size=40)
            self.table_layout.addWidget(delete_button, row_position + 1, len(self.headers))

            self.table_data.append(row)
        self.status_bar.showMessage("State loaded.", 5000)





    def show_next_message(self):
        if self.current_index < len(self.files) - 1:
            self.current_index += 1
            while self.current_index < len(self.files) and self.files[self.current_index] in self.read_files:
                self.current_index += 1
            self.show_message()
            self.save_state()
            self.status_bar.showMessage(f"Showing file {self.current_index + 1} of {len(self.files)}.", 5000)
            self.file_selector.setCurrentIndex(self.current_index)  # QComboBox'ı güncelleyin


    def show_prev_message(self):
        if self.current_index > 0:
            self.current_index -= 1
            while self.current_index >= 0 and self.files[self.current_index] in self.read_files:
                self.current_index -= 1
            self.show_message()
            self.save_state()
            self.status_bar.showMessage(f"Showing file {self.current_index + 1} of {len(self.files)}.", 5000)
            self.file_selector.setCurrentIndex(self.current_index)  # QComboBox'ı güncelleyin

    def load_table_headers(self):
        config_file = 'config.conf'
        if os.path.exists(config_file):
            with open(config_file, 'r') as file:
                headers = file.read().strip().split(',')
                self.headers = headers
                self.create_table()
                if not self.table_data:  # Add an initial empty row if table_data is empty
                    self.add_table_row()

    def create_table(self):
        self.table_layout.setSpacing(0)
        self.table_layout.setContentsMargins(0, 0, 0, 0)

        for col, header in enumerate(self.headers):
            if header.lower() == "offer type":
                header_widget = QWidget()
                header_layout = QHBoxLayout(header_widget)
                header_layout.setContentsMargins(0, 0, 0, 0)
                
                header_label = QLabel(header)
                header_label.setFont(QFont('Arial', 10, QFont.Bold))
                header_layout.addWidget(header_label)
                header_label.setStyleSheet("""
                    background-color: #C8E6C9;  
                    color: black;  
                    padding: 10px;
                    border: 1px solid #BDBDBD;  
                    border-radius: 5px;
                    text-align: center;
                    margin: 0;  
                """)

                header_combo = QComboBox()
                header_combo.addItems(["WTB", "WTS"])
                header_combo.currentTextChanged.connect(lambda: self.update_all_offer_types(header_combo.currentText()))
                header_layout.addWidget(header_combo)

                self.table_layout.addWidget(header_widget, 0, col)
            else:
                header_label = QLabel(header)
                header_label.setFont(QFont('Arial', 10, QFont.Bold))
                header_label.setStyleSheet("""
                    background-color: #C8E6C9;  
                    color: black;  
                    padding: 10px;
                    border: 1px solid #BDBDBD;  
                    border-radius: 5px;
                    text-align: center;
                    margin: 0;  
                """)
                header_label.setFixedHeight(40)
                header_label.setFixedWidth(120)
                header_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                self.table_layout.addWidget(header_label, 0, col)

        self.table_data = []


            

    def update_all_offer_types(self, selected_value):
        for row in self.table_data:
            for col, item in enumerate(row):
                if self.headers[col].lower() == "offer type":
                    item.setCurrentText(selected_value)

    def add_table_row(self):
        row_position = len(self.table_data)
        row_data = []

        current_offer_type = None
        for row in self.table_data:
            for col, item in enumerate(row):
                if self.headers[col].lower() == "offer type":
                    current_offer_type = item.currentText()
                    break
            if current_offer_type:
                break

        if current_offer_type is None:
            current_offer_type = "WTB"

        for col in range(len(self.headers)):
            color = "#FFFFFF" if row_position % 2 == 0 else "#F9F9F9"
            if self.headers[col].lower() == "offer type":
                combo_box = QComboBox()
                combo_box.setFont(QFont('Arial', 10))
                combo_box.addItems(["WTB", "WTS"])
                combo_box.setCurrentText(current_offer_type)
                combo_box.setFixedHeight(40)
                combo_box.setFixedWidth(120)
                combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                combo_box.setStyleSheet(f"""
                    padding: 0px;
                    margin: 0px;
                    border: 1px solid #BDBDBD;
                    background-color: {color};
                """)
                combo_box.installEventFilter(self)
                self.table_layout.addWidget(combo_box, row_position + 1, col)
                row_data.append(combo_box)
            elif self.headers[col].lower() == "product name":
                combo_box = QComboBox()
                combo_box.setFont(QFont('Arial', 10))
                combo_box.addItems(self.product_name_list)
                combo_box.setEditable(True)
                combo_box.setFixedHeight(40)
                combo_box.setFixedWidth(120)
                combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                combo_box.setStyleSheet(f"""
                    padding: 0px;
                    border: 1px solid #BDBDBD;
                    margin: 0px;
                    background-color: {color};
                """)
                combo_box.installEventFilter(self)
                completer = QCompleter(self.product_name_list, self)
                combo_box.setCompleter(completer)
                combo_box.lineEdit().editingFinished.connect(self.update_product_name_list)
                self.table_layout.addWidget(combo_box, row_position + 1, col)
                row_data.append(combo_box)
            elif self.headers[col].lower() == "brand":
                combo_box = QComboBox()
                combo_box.setFont(QFont('Arial', 10))
                combo_box.addItems(self.brand_list)
                combo_box.setEditable(True)
                combo_box.setFixedHeight(40)
                combo_box.setFixedWidth(120)
                combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                combo_box.setStyleSheet(f"""
                    padding: 0px;
                    border: 1px solid #BDBDBD;
                    margin: 0px;
                    background-color: {color};
                """)
                combo_box.installEventFilter(self)
                completer = QCompleter(self.brand_list, self)
                combo_box.setCompleter(completer)
                combo_box.lineEdit().editingFinished.connect(self.update_brand_list)
                self.table_layout.addWidget(combo_box, row_position + 1, col)
                row_data.append(combo_box)
            elif self.headers[col].lower() == "category":
                combo_box = QComboBox()
                combo_box.setFont(QFont('Arial', 10))
                combo_box.addItems(self.category_list)
                combo_box.setEditable(True)
                combo_box.setFixedHeight(40)
                combo_box.setFixedWidth(120)
                combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                combo_box.setStyleSheet(f"""
                    padding: 0px;
                    border: 1px solid #BDBDBD;
                    margin: 0px;
                    background-color: {color};
                """)
                combo_box.installEventFilter(self)
                completer = QCompleter(self.category_list, self)
                combo_box.setCompleter(completer)
                combo_box.lineEdit().editingFinished.connect(self.update_category_list)
                self.table_layout.addWidget(combo_box, row_position + 1, col)
                row_data.append(combo_box)
            elif self.headers[col].lower() == "color":
                combo_box = QComboBox()
                combo_box.setFont(QFont('Arial', 10))
                combo_box.addItems(self.color_list)
                combo_box.setEditable(True)
                combo_box.setFixedHeight(40)
                combo_box.setFixedWidth(120)
                combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                combo_box.setStyleSheet(f"""
                    padding: 0px;
                    border: 1px solid #BDBDBD;
                    margin: 0px;
                    background-color: {color};
                """)
                combo_box.installEventFilter(self)
                completer = QCompleter(self.color_list, self)
                combo_box.setCompleter(completer)
                combo_box.lineEdit().editingFinished.connect(self.update_color_list)
                self.table_layout.addWidget(combo_box, row_position + 1, col)
                row_data.append(combo_box)
            else:
                line_edit = QLineEdit()
                line_edit.setFont(QFont('Arial', 10))
                line_edit.setFixedHeight(40)
                line_edit.setFixedWidth(120)
                line_edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                line_edit.setStyleSheet(f"""
                    padding: 0px;
                    border: 1px solid #BDBDBD;
                    margin: 0px;
                    background-color: {color};
                """)
                line_edit.installEventFilter(self)
                self.table_layout.addWidget(line_edit, row_position + 1, col)
                row_data.append(line_edit)

        delete_button = QPushButton()
        delete_button.setIcon(QIcon('icons/delete_icon.png'))
        delete_button.setIconSize(QSize(40, 40))
        delete_button.setFixedSize(60, 60)
        delete_button.setStyleSheet("""
            border: 1px solid #BDBDBD;
            margin: 0;
            border-radius: 5px;
        """)
        delete_button.clicked.connect(partial(self.delete_table_row, row_position))
        self.table_layout.addWidget(delete_button, row_position + 1, len(self.headers))

        self.table_data.append(row_data)
        self.update_delete_buttons()
        self.save_state()
        self.status_bar.showMessage("Row added.", 5000)







    def update_product_name_list(self):
        sender = self.sender()
        if isinstance(sender, QLineEdit):
            text = sender.text()
            if text and text not in self.product_name_list:
                self.product_name_list.append(text)
                for row in self.table_data:
                    for col, item in enumerate(row):
                        if self.headers[col].lower() == "product name":
                            completer = QCompleter(self.product_name_list, self)
                            item.setCompleter(completer)


    def update_offer_type(self, combo_box):
        selected_value = combo_box.currentText()
        for row in self.table_data:
            for col, item in enumerate(row):
                if self.headers[col].lower() == "offer type":
                    item.setCurrentText(selected_value)

    def delete_table_row(self, row):
        if 0 <= row < len(self.table_data):
            # Remove all widgets in this row from the layout and delete them
            for col in range(len(self.headers) + 1):  # Include the delete button column
                item = self.table_layout.itemAtPosition(row + 1, col)
                if item is not None:
                    widget = item.widget()
                    if widget:
                        self.table_layout.removeWidget(widget)
                        widget.deleteLater()

            # Remove the row data from the list
            self.table_data.pop(row)

            # Move up all widgets in rows below the deleted row
            for subsequent_row in range(row, len(self.table_data)):
                for col in range(len(self.headers) + 1):
                    item = self.table_layout.itemAtPosition(subsequent_row + 2, col)  # original position
                    if item is not None:
                        widget = item.widget()
                        if widget:
                            self.table_layout.removeWidget(widget)
                            self.table_layout.addWidget(widget, subsequent_row + 1, col)

            self.update_delete_buttons()
            self.save_state()
            self.status_bar.showMessage("Row deleted.", 5000)

    def update_delete_buttons(self):
        for i in range(len(self.table_data)):
            item = self.table_layout.itemAtPosition(i + 1, len(self.headers))
            if item is not None:
                delete_button = item.widget()
                if delete_button:
                    delete_button.clicked.disconnect()
                    delete_button.clicked.connect(partial(self.delete_table_row, i))



    def save_table(self):
        table_data = self.get_table_data()
        headers = self.headers

        # Ensure each row has the correct number of columns
        formatted_table_data = []
        for row in table_data:
            if len(row) == len(headers):
                formatted_table_data.append({headers[i]: row[i] for i in range(len(headers))})
            else:
                print(f"Skipping row due to length mismatch: {row}")

        export_data = {
            "table_data": formatted_table_data
        }

        if self.output_folder:
            file_path = os.path.join(self.output_folder, f"{os.path.splitext(self.files[self.current_index])[0]}.json")
            with open(file_path, 'w', encoding='utf-8') as file:
                json.dump(export_data, file, indent=4, ensure_ascii=False)
            self.status_bar.showMessage("Table saved successfully.", 5000)
        else:
            self.status_bar.showMessage("No output folder selected.", 5000)


    def save_state(self):
        if self.initial_load_done:  # Only save state if the initial load is done
            state = {
                'current_index': self.current_index,
                'table_data': self.get_table_data(),
                'input_folder': self.input_folder,
                'output_folder': self.output_folder,
                'files': self.files,
                'folder_indices': self.folder_indices,  # Add folder_indices to save state
                'product_name_list': self.product_name_list,  # Save product name list
                'brand_list': self.brand_list,  # Save brand list
                'category_list': self.category_list,  # Save category list
                'color_list': self.color_list,  # Save color list
                'read_files': self.read_files,  # Save read files list
                'processed_text': self.processed_text.toPlainText()  # Save processed text regardless of view mode
            }
            print("Saving state:", state)  # Debug print
            with open('state.json', 'w') as file:
                json.dump(state, file, ensure_ascii=False, indent=4)

            # Save table data for the current file
            self.save_table_data_for_current_file()
            self.status_bar.showMessage("State saved.", 5000)








    def save_table_data_for_current_file(self):
        if 0 <= self.current_index < len(self.files):
            current_file = self.files[self.current_index]
            file_path = os.path.join(self.input_folder, current_file)
            table_data = self.get_table_data()
            table_data_path = file_path + '.json'
            with open(table_data_path, 'w', encoding='utf-8') as file:
                json.dump({'table_data': table_data, 'processed_text': self.processed_text.toPlainText()}, file, ensure_ascii=False, indent=4)




    def load_external_data(self):
        try:
            with open('brand_category_color.json', 'r') as file:
                data = json.load(file)
                self.product_name_list = data.get('product_names', [])
                self.brand_list = data.get('brands', [])
                self.category_list = data.get('categories', [])
                self.color_list = data.get('colors', [])
        except Exception as e:
            print(f"Error loading external data: {e}")
            self.product_name_list = []
            self.brand_list = []
            self.category_list = []
            self.color_list = []



    def load_state(self):
        self.load_external_data()  # Load external data first
        if os.path.exists('state.json'):
            with open('state.json', 'r') as file:
                state = json.load(file)
                print("Loaded state:", state)  # Debug print
                self.current_index = state.get('current_index', -1)
                self.input_folder = state.get('input_folder', '')
                self.output_folder = state.get('output_folder', '')
                self.files = state.get('files', [])
                self.folder_indices = state.get('folder_indices', {})  # Load folder_indices
                self.product_name_list = state.get('product_name_list', [])
                self.brand_list = state.get('brand_list', [])  # Load brand list
                self.category_list = state.get('category_list', [])  # Load category list
                self.color_list = state.get('color_list', [])  # Load color list
                self.read_files = state.get('read_files', [])  # Load read files list
                processed_text = state.get('processed_text', '')  # Load processed text

                # Load table headers
                self.load_table_headers()

                # Set table data
                self.set_table_data(state.get('table_data', []))

                if self.input_folder:
                    self.load_files()  # Load files after setting the state

                self.processed_text.setText(processed_text)  # Load the processed text
                
                # Set the file selector to the current index
                if 0 <= self.current_index < len(self.files):
                    self.file_selector.setCurrentIndex(self.current_index)
                    self.show_message()

            self.start_auto_save()  # Start auto save after loading state



    def get_table_data(self):
        data = []
        for row_data in self.table_data:
            row = []
            for item in row_data:
                if isinstance(item, QComboBox):
                    row.append(item.currentText())
                elif isinstance(item, QLineEdit):
                    row.append(item.text())
            data.append(row)
        return data


    def set_table_data(self, data):
        # Clear existing table data
        for i in reversed(range(len(self.table_data))):
            self.delete_table_row(i)

        self.table_data = []
        num_headers = len(self.headers)  # Number of headers

        for row_position, row_data in enumerate(data):
            # Ensure row_data has the same length as headers by truncating extra columns
            row_data = row_data[:num_headers]
            # Ensure row_data has the same length as headers by padding with empty strings
            row_data += [""] * (num_headers - len(row_data))

            row = []
            for col in range(num_headers):
                value = row_data[col]

                if self.headers[col].lower() == "offer type":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(["WTB", "WTS"])
                    combo_box.setCurrentText(value)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "product name":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.product_name_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.product_name_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_product_name_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "brand":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.brand_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.brand_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_brand_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "category":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.category_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.category_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_category_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                elif self.headers[col].lower() == "color":
                    combo_box = QComboBox()
                    combo_box.setFont(QFont('Arial', 10))
                    combo_box.addItems(self.color_list)
                    combo_box.setCurrentText(value)
                    combo_box.setEditable(True)
                    combo_box.setFixedHeight(40)
                    combo_box.setFixedWidth(120)
                    combo_box.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    combo_box.installEventFilter(self)
                    completer = QCompleter(self.color_list, self)
                    combo_box.setCompleter(completer)
                    combo_box.lineEdit().editingFinished.connect(self.update_color_list)
                    self.table_layout.addWidget(combo_box, row_position + 1, col)
                    row.append(combo_box)
                else:
                    line_edit = QLineEdit()
                    line_edit.setFont(QFont('Arial', 10))
                    line_edit.setText(value)
                    line_edit.setFixedHeight(40)
                    line_edit.setFixedWidth(120)
                    line_edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    line_edit.installEventFilter(self)
                    self.table_layout.addWidget(line_edit, row_position + 1, col)
                    row.append(line_edit)

            delete_button = self.create_icon_button('icons/delete_icon.png', partial(self.delete_table_row, row_position), size=40)
            self.table_layout.addWidget(delete_button, row_position + 1, len(self.headers))

            self.table_data.append(row)
        self.status_bar.showMessage("State loaded.", 5000)



    def save_external_data(self):
        try:
            data = {
                'brands': self.brand_list,
                'categories': self.category_list,
                'colors': self.color_list,
                'product_names': self.product_name_list
            }
            with open('brand_category_color.json', 'w') as file:
                json.dump(data, file, indent=4)
            print("External data saved.")
        except Exception as e:
            print(f"Error saving external data: {e}")
            
    def update_product_name_list(self):
        sender = self.sender()
        if isinstance(sender, QLineEdit):
            text = sender.text()
            if text and text not in self.product_name_list:
                self.product_name_list.append(text)
                self.save_external_data()
                self.save_state()
                # Update all existing combo boxes for product name
                for row in self.table_data:
                    for col, item in enumerate(row):
                        if self.headers[col].lower() == "product name" and isinstance(item, QComboBox):
                            item.addItem(text)

    def update_brand_list(self):
        sender = self.sender()
        if isinstance(sender, QLineEdit):
            text = sender.text()
            if text and text not in self.brand_list:
                self.brand_list.append(text)
                self.save_external_data() 
                self.save_state()
                # Update all existing combo boxes for brand_list
                for row in self.table_data:
                    for col, item in enumerate(row):
                        if self.headers[col].lower() == "brand_list" and isinstance(item, QComboBox):
                            item.addItem(text)

    def update_category_list(self):
        sender = self.sender()
        if isinstance(sender, QLineEdit):
            text = sender.text()
            if text and text not in self.category_list:
                self.category_list.append(text)
                self.save_external_data()
                self.save_state()
                # Update all existing combo boxes for category_list
                for row in self.table_data:
                    for col, item in enumerate(row):
                        if self.headers[col].lower() == "category_list" and isinstance(item, QComboBox):
                            item.addItem(text)

    def update_color_list(self):
        sender = self.sender()
        if isinstance(sender, QLineEdit):
            text = sender.text()
            if text and text not in self.color_list:
                self.color_list.append(text)
                self.save_external_data()
                self.save_state()
                # Update all existing combo boxes for color_list
                for row in self.table_data:
                    for col, item in enumerate(row):
                        if self.headers[col].lower() == "color_list" and isinstance(item, QComboBox):
                            item.addItem(text)


    def eventFilter(self, source, event):
        if event.type() == QEvent.FocusIn:
            if isinstance(source, QLineEdit) or isinstance(source, QComboBox):
                source.setStyleSheet("""
                    padding: 5px;
                    border: 2px solid #4A90E2;
                    border-radius: 2px;
                    text-align: center;
                    margin: 0;
                """)
        elif event.type() == QEvent.FocusOut:
            if isinstance(source, QLineEdit) or isinstance(source, QComboBox):
                source.setStyleSheet("""
                    padding: 5px;
                    border: 1px solid #BDBDBD;
                    border-radius: 2px;
                    text-align: center;
                    margin: 0;
                """)
        return super(MessageTool, self).eventFilter(source, event)


    def update_status(self):
        if self.files:
            self.status_bar.showMessage(f"File {self.current_index + 1} of {len(self.files)}", 5000)
        else:
            self.status_bar.showMessage("No files loaded", 5000)

    def auto_save(self):
        print("Auto saving state")  # Debug print
        self.save_state()

    def start_auto_save(self):
     self.timer.start(7000)  # Auto save every 7 seconds

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MessageTool()
    ex.show()
    sys.exit(app.exec_())
