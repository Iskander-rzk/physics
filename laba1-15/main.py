from PyQt6 import QtWidgets, QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QDialog, QLabel, QApplication, QWidget, QPushButton, QMainWindow, QVBoxLayout, QFileDialog, \
    QTableWidget, QTableWidgetItem, QMessageBox, QHBoxLayout, QFileDialog
from PyQt6.QtGui import QFont
import os
import csv
import sys
from datetime import datetime
import matplotlib.pyplot as plt
from PyQt6.QtGui import QPixmap
from PIL import Image
import pikepdf
import win32com.client

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Лабораторная работа №1-15")
        self.setGeometry(500, 150, 500, 600)
        
        
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        label=QLabel("Лабораторная работа №1-15")
        label.setFont(QFont("Arial",18,QFont.Weight.Bold))
        main_layout.addWidget(label, alignment=Qt.AlignmentFlag.AlignHCenter)



        # Кнопки
        self.button1 = QPushButton("Методические указания")
        self.button2 = QPushButton("Создать таблицу")
        self.button3 = QPushButton("Построить график")
        self.button4 = QPushButton("Изменить таблицу")
        self.button5 = QPushButton("Сохранить таблицу как PNG")
        self.button6 = QPushButton("Объединить PDF")  # Новая кнопка для объединения PDF
        self.button_word = QPushButton("Редактировать титульный лист")

        font = QFont("Arial", 13)
        self.button1.setFont(font)
        self.button2.setFont(font)
        self.button3.setFont(font)
        self.button4.setFont(font)
        self.button5.setFont(font)
        self.button6.setFont(font)
        self.button_word.setFont(font)
        
        

        self.button1.setFixedSize(308, 60)
        self.button2.setFixedSize(150, 60)
        self.button3.setFixedSize(308, 60)
        self.button4.setFixedSize(150, 60)
        self.button5.setFixedSize(308, 60)
        self.button6.setFixedSize(308, 60)  # Размер новой кнопки
        self.button_word.setFixedSize(308, 60)

        main_layout.addWidget(self.button1, alignment=Qt.AlignmentFlag.AlignHCenter)
        main_layout.addWidget(self.button_word, alignment=Qt.AlignmentFlag.AlignHCenter)

        # Горизонтальный макет для кнопок 2 и 4
        horizontal_layout = QHBoxLayout()
        horizontal_layout.addWidget(self.button2, alignment=Qt.AlignmentFlag.AlignRight)
        horizontal_layout.addWidget(self.button4, alignment=Qt.AlignmentFlag.AlignLeft)
        main_layout.addLayout(horizontal_layout)

        main_layout.addWidget(self.button3, alignment=Qt.AlignmentFlag.AlignHCenter)
        main_layout.addWidget(self.button5, alignment=Qt.AlignmentFlag.AlignHCenter)

        label2 = QLabel()
        text = """
        <span style="font-family: Wingdings; font-size: 20pt;">FF</span>
        <span style="font-family: Arial; font-size: 18pt; font-weight: bold;"> ВНИМАНИЕ </span>
        <span style="font-family: Wingdings; font-size: 20pt;">EE</span>
        """


        label2.setText(text)
        main_layout.addWidget(label2, alignment=Qt.AlignmentFlag.AlignHCenter)
        label3=QLabel()
        text2 = """
                <span style="font-family: TimesNewRoman; font-size: 12pt;"> Перед экспортом не забудьте сохранить таблицу как PNG</span>
                """
        label3.setText(text2)
        main_layout.addWidget(label3, alignment=Qt.AlignmentFlag.AlignHCenter)


        main_layout.addWidget(self.button6, alignment=Qt.AlignmentFlag.AlignHCenter)
        self.button1.clicked.connect(self.on_button1_click)
        self.button2.clicked.connect(self.on_button2_click)
        self.button3.clicked.connect(self.plot_graph)
        self.button4.clicked.connect(self.on_button4_click)
        self.button5.clicked.connect(self.save_table_as_image)
        self.button6.clicked.connect(self.merge_pdfs)  # Подключаем функцию для объединения PDF
        self.button_word.clicked.connect(self.edit_titul_docx)

        # Инициализируем таблицу
        self.table = None
        self.create_table_dialog = None
        
    
    
    def edit_titul_docx(self):
        """Метод для открытия документа titul.docx в MS Word и сохранения его как изображение."""
        try:
            # Путь к файлу titul.docx
            doc_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'titul.docx')
            if not os.path.exists(doc_path):
                QMessageBox.warning(self, "Ошибка", f"Файл {doc_path} не найден.")
                return

            # Открытие Word и документа
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = True  # Делаем приложение видимым
            doc = word_app.Documents.Open(doc_path)

            # Сохранение документа как изображения после закрытия Word
            def save_as_pdf():
                pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'titul.pdf')
                doc.SaveAs(pdf_path, FileFormat=17)  # Формат 17 — это PDF
                QMessageBox.information(self, "Успех", f"Документ сохранен как PDF: {pdf_path}")
                doc.Close(False)  # Закрываем документ без сохранения изменений
                word_app.Quit()
                save_dialog.accept()


            # Создаем диалог с кнопкой сохранения
            save_dialog = QDialog(self)
            save_dialog.setWindowTitle("Сохранить как PDF")
            save_layout = QVBoxLayout(save_dialog)

            save_label = QLabel("После редактирования нажмите Сохранить, не закрывая Word")
            save_label.setFont(QFont("Arial", 12))
            save_layout.addWidget(save_label, alignment=Qt.AlignmentFlag.AlignCenter)

            save_button = QPushButton("Сохранить")
            save_button.clicked.connect(save_as_pdf)
            save_layout.addWidget(save_button, alignment=Qt.AlignmentFlag.AlignCenter)

            save_dialog.exec()

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")
                
    def get_resource_path(self, relative_path):
        """Функция для получения правильного пути к ресурсам (для скомпилированного приложения и режима разработки)"""
        if getattr(sys, 'frozen', False):
            # Если приложение запущено как скомпилированный файл
            app_path = sys._MEIPASS
        else:
            # Если приложение запущено в среде разработки
            app_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(app_path, relative_path)

    def on_button1_click(self):
        pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'metod.pdf')

        if os.path.exists(pdf_path):
            print(f"Открытие файла: {pdf_path}")  # Отладочный вывод
            os.startfile(pdf_path)
        else:
            print(f"Файл {pdf_path} не найден.")

    def on_button2_click(self):
        # Создаем окно для таблицы только при первом нажатии
        if self.create_table_dialog is None:
            self.create_table_dialog = QDialog(self)
            self.create_table_dialog.setWindowTitle("Таблица данных")
            self.create_table_dialog.setGeometry(300, 200, 1200, 600)

            layout = QVBoxLayout()

            self.table = QTableWidget(7, 9)
            self.table.setHorizontalHeaderLabels(
                ['F=mq, Н', 'N₀, мм', 'N₁, мм', 'N₁-N₀, мм', 'l₀, мм', 'D, мм', 'Δl, мм', 'E, Н/м²',
                'E из графика, Н/м²'])
            layout.addWidget(self.table)
            self.create_table_dialog.setFixedSize(1000, 400) 
            # Кнопка для сохранения данных таблицы
            save_button = QPushButton('Сохранить таблицу')
            save_button.clicked.connect(self.save_table)
            layout.addWidget(save_button)
            self.table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                color: black;
                border: 1px solid black;
            }
            QTableWidget::item {
                padding: 5px;
                border: 1px solid #ddd;
            }
            QTableWidget::horizontalHeader, QTableWidget::verticalHeader {
                background-color: #f4f4f4;
                border: 1px solid #ddd;
            }
        """)
            self.create_table_dialog.setLayout(layout)

        # Показываем уже созданный диалог для редактирования
        self.create_table_dialog.show()

    def save_table(self):
        # Проверяем или создаем директорию "tables"
        project_dir = os.path.dirname(os.path.abspath(__file__))
        tables_dir = os.path.join(project_dir, 'tables')

        if not os.path.exists(tables_dir):
            os.makedirs(tables_dir)

        # Генерируем имя файла с использованием времени
        file_name = f"table.csv"
        file_path = os.path.join(tables_dir, file_name)

        # Сохраняем данные таблицы в CSV
        if self.table:
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)
                for row in range(self.table.rowCount()):
                    row_data = []
                    for column in range(self.table.columnCount()):
                        item = self.table.item(row, column)
                        if item:
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    writer.writerow(row_data)

            # Показываем сообщение об успешном сохранении
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setText(f"Таблица сохранена в {file_path}")
            msg_box.setWindowTitle("Сохранение")
            msg_box.exec()

    def save_table_as_image(self):
    # Проверяем или создаем директорию "tables"
        project_dir = os.path.dirname(os.path.abspath(__file__))
        tables_dir = os.path.join(project_dir, 'tables')

        if not os.path.exists(tables_dir):
            os.makedirs(tables_dir)

    # Генерируем имя файла с использованием времени
        file_name = f"table.png"
        file_path = os.path.join(tables_dir, file_name)

    # Делаем скриншот только таблицы и сохраняем как изображение
        if self.table:
        # Создаем QPixmap из таблицы
            pixmap = self.table.grab()

        # Сохраняем QPixmap как PNG
            pixmap.save(file_path, "PNG")

        # Показываем сообщение об успешном сохранении
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setText(f"Таблица сохранена в виде изображения в {file_path}")
            msg_box.setWindowTitle("Сохранение изображения таблицы")
            msg_box.exec()

    def plot_graph(self):
        # Извлекаем данные из таблицы
        if not self.table:
            return

        tension = []
        elongation = []

        for row in range(self.table.rowCount()):
            f_item = self.table.item(row, 0)  # Натяжение T = F
            delta_l_item = self.table.item(row, 6)  # Удлинение ΔL

            if f_item and delta_l_item:
                try:
                    tension.append(float(f_item.text()))
                    elongation.append(float(delta_l_item.text()))
                except ValueError:
                    # Игнорируем строки с некорректными данными
                    continue

        if tension and elongation:
            fig, ax = plt.subplots()
            ax.plot(tension, elongation, marker='o')
            ax.set_title('Зависимость удлинения ΔL от натяжения T')
            ax.set_xlabel('Натяжение T (Н)')
            ax.set_ylabel('Удлинение ΔL (мм)')
            ax.grid(True)

            # Указываем директорию для сохранения графиков
            project_dir = os.path.dirname(os.path.abspath(__file__))
            plots_dir = os.path.join(project_dir, 'plots')

            if not os.path.exists(plots_dir):
                os.makedirs(plots_dir)

            # Генерируем имя файла с использованием времени
            file_name = f"plot.png"
            file_path = os.path.join(plots_dir, file_name)

            # Определяем функцию сохранения при закрытии
            def on_close(event):
                fig.savefig(file_path)  # Сохраняем график
                plt.close(fig)  # Закрываем окно графика

                # Показываем сообщение об успешном сохранении
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Icon.Information)
                msg_box.setText(f"График сохранен в {file_path}")
                msg_box.setWindowTitle("Сохранение графика")
                msg_box.exec()

            # Привязываем событие закрытия окна к функции сохранения
            fig.canvas.mpl_connect('close_event', on_close)

            # Показываем график
            plt.show()
        else:
            QMessageBox.warning(self, "Ошибка", "Недостаточно данных для построения графика.")

    def on_button4_click(self):
        if self.table is not None:
            self.create_table_dialog.show()
        else:
            QMessageBox.warning(self, "Ошибка", "Сначала создайте таблицу.")

    def merge_pdfs(self):
        # Открываем диалог для выбора места сохранения файла
        # Открываем диалог для выбора места сохранения файла
        options = QFileDialog.Option.DontUseNativeDialog  # Используем флаг напрямую

    # Открываем диалог для выбора места сохранения файла
        output_pdf_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Сохранить объединенный PDF", 
            "", 
            "PDF Files (*.pdf);;All Files (*)", 
            options=options
        )

        if not output_pdf_path:
            return  # Пользователь отменил выбор файла

        # Добавляем .pdf, если оно не указано
        if not output_pdf_path.lower().endswith('.pdf'):
            output_pdf_path += '.pdf'
        
        
    # Директория для временных файлов
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_files')
        os.makedirs(temp_dir, exist_ok=True)  # Создаем папку, если её нет

        # Filenames to be merged
        pdf_files = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'titul.pdf'),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'metod_for_titul.pdf'),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tables', 'table.png'),  # Path to PNG
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plots', 'plot.png')  # Path to PNG
        ]

        # Create a new empty PDF for merging
        with pikepdf.Pdf.new() as pdf_output:
            for pdf_file in pdf_files:
                if os.path.exists(pdf_file):
                    if pdf_file.endswith('.pdf'):
                        # Open existing PDF and add its pages to the new file
                        with pikepdf.open(pdf_file) as pdf_input:
                            pdf_output.pages.extend(pdf_input.pages)
                    elif pdf_file.endswith('.png'):
                        # Convert the PNG image to PDF
                        with Image.open(pdf_file) as img:
                            img = img.convert('RGB')  # Ensure image is in RGB mode
                            temp_pdf_path = os.path.join(temp_dir, 'temp.pdf')  # Path to save temporary PDF
                            img.save(temp_pdf_path, 'PDF')  # Save the image as a PDF
                            with pikepdf.open(temp_pdf_path) as img_pdf:
                                pdf_output.pages.extend(img_pdf.pages)
                else:
                    QMessageBox.warning(self, "Ошибка", f"Файл не найден: {pdf_file}")
                    return

            # Save the merged PDF
            pdf_output.save(output_pdf_path)

        # Show a success message
        QMessageBox.information(self, "Успех", f"Файлы объединены в {output_pdf_path}")

def application():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    application()