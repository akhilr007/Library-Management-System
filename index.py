from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import datetime
import sys
import MySQLdb
from xlrd import *
from xlsxwriter import *

from PyQt5.uic import loadUiType

ui, _ = loadUiType('library_management_system.ui')
login, _ = loadUiType('login.ui')

# loading our login UI


class Login(QWidget, login):
    def __init__(self):
        QWidget.__init__(self)
        self.setupUi(self)
        #self.handle_login()
        # to connect login push button
        self.pushButton.clicked.connect(self.handle_login)

    def handle_login(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        sql = '''SELECT * FROM users'''

        self.cur.execute(sql)
        data = self.cur.fetchall()
        print(data)

        username = self.lineEdit.text()
        password = self.lineEdit_2.text()

        for info in data:
            if username == info[1] and password == info[3]:
                print('user match')
                self.window2 = MainApp()
                self.close()
                self.window2.show()

            else:
                print('user not match')
                self.label.setText('Enter username or password correctly')


class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handle_ui_changes()  # to disable the hiding themes when we run the app
        self.handle_buttons()

        self.show_category()
        self.show_author()
        self.show_publisher()

        self.show_category_ui()
        self.show_author_ui()
        self.show_publisher_ui()

        self.dark_orange_themes()

        self.show_client()
        self.show_books()
        self.show_operations()

    # method to handle UI changes
    def handle_ui_changes(self):
        self.hide_themes()
        # to hide the tab bar in the user interface design
        self.tabWidget.tabBar().setVisible(False)

    # method to handle buttons
    def handle_buttons(self):
        # when we click on themes button to show the show themes
        self.pushButton_5.clicked.connect(self.show_themes)

        # when we click on this button the theme bar closes
        self.pushButton_22.clicked.connect(self.hide_themes)

        # to connect operations button button
        self.pushButton.clicked.connect(self.open_operation_tab)

        # to connect books button tab
        self.pushButton_2.clicked.connect(self.open_book_tab)

        # to connect users button tab
        self.pushButton_3.clicked.connect(self.open_user_tab)

        # to connect clients button tab
        self.pushButton_26.clicked.connect(self.open_client_tab)

        # to connect settings buttons tab
        self.pushButton_4.clicked.connect(self.open_setting_tab)

        # to connect add new book button
        self.pushButton_7.clicked.connect(self.add_new_books)

        # to connect add category button
        self.pushButton_15.clicked.connect(self.add_category)

        # to connect add author button
        self.pushButton_16.clicked.connect(self.add_author)

        # to connect add publisher button
        self.pushButton_17.clicked.connect(self.add_publisher)

        # to connect book search button
        self.pushButton_10.clicked.connect(self.search_books)

        # to connect edit book button
        self.pushButton_8.clicked.connect(self.edit_books)

        # to connect delete book button
        self.pushButton_9.clicked.connect(self.delete_book)

        # to connect add user button
        self.pushButton_11.clicked.connect(self.add_new_user)

        # to connect login button
        self.pushButton_12.clicked.connect(self.login)

        # to connect edit user button
        self.pushButton_14.clicked.connect(self.edit_user)

        # to connect themes button
        self.pushButton_18.clicked.connect(self.classic_themes)
        self.pushButton_19.clicked.connect(self.dark_blue_themes)
        self.pushButton_20.clicked.connect(self.dark_gray_themes)
        self.pushButton_21.clicked.connect(self.dark_orange_themes)

        # to connect add client button
        self.pushButton_13.clicked.connect(self.add_new_client)

        # to connect search client button
        self.pushButton_25.clicked.connect(self.search_client)

        # to connect delete client button
        self.pushButton_24.clicked.connect(self.delete_client)

        # to connect edit client button
        self.pushButton_23.clicked.connect(self.edit_client)

        # to connect operations button
        self.pushButton_6.clicked.connect(self.operations)

        # to connect export data buttons excel file
        self.pushButton_28.clicked.connect(self.export_operations)
        self.pushButton_29.clicked.connect(self.export_clients)
        self.pushButton_27.clicked.connect(self.export_books)

    # method to show themes
    def show_themes(self):
        self.groupBox_3.show()

    # method to hide themes
    def hide_themes(self):
        self.groupBox_3.hide()

    # #############################
    # ###### opening tab ##########

    # method to open operations tab
    def open_operation_tab(self):
        # connect tab 1 with the operations button
        # so that we can move from one tab to another tab by clicking buttons
        self.tabWidget.setCurrentIndex(0)

    # method to open books tab
    def open_book_tab(self):
        # connect tab 2 with the books button
        self.tabWidget.setCurrentIndex(1)

    # method to open users tab
    def open_client_tab(self):
        # connect tab 3 with the users button
        self.tabWidget.setCurrentIndex(2)

    # method to open clients tab
    def open_user_tab(self):
        # connect tab 3 with the users button
        self.tabWidget.setCurrentIndex(3)

    # method to open settings tab
    def open_setting_tab(self):
        # connect tab 4 with the settings button
        self.tabWidget.setCurrentIndex(4)

    # ######## operations ############
    def operations(self):
        book_title = self.lineEdit.text()
        client_name = self.lineEdit_29.text()

        type = self.comboBox.currentText()
        days = self.comboBox_2.currentIndex()+1
        date_from = datetime.date.today()
        date_to = date_from + datetime.timedelta(days=int(days))

        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        self.cur.execute('''
        INSERT INTO operations(book_name, client_name, type, days, date_from, date_to)
        VALUES(%s, %s, %s, %s, %s, %s)
        ''', (book_title,  client_name, type, days, date_from, date_to))

        self.db.commit()
        self.statusBar().showMessage('New Operations Added')

        self.show_operations()

    # function to show all operations in our UI
    def show_operations(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()
        self.cur.execute('''
        SELECT book_name, client_name, type, days, date_from, date_to FROM operations
        ''')

        data = self.cur.fetchall()

        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.insertRow(0)

        for row, info in enumerate(data):
            for column, item in enumerate(info):
                self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_pos = self.tableWidget_2.rowCount()
            self.tableWidget_2.insertRow(row_pos)

        self.db.close()


    # #############################
    # ###### operations under books section ##########

    # to show all the books in our database
    def show_books(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        self.cur.execute('''
        SELECT book_code, book_title, book_description, book_category,
                            book_author, book_publisher, book_price FROM book
        ''')

        data = self.cur.fetchall()

        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)

        for row, info in enumerate(data):
            for column, item in enumerate(info):
                self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_pos = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_pos)

        self.db.close()

    # for adding new books to our database
    def add_new_books(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        book_title = self.lineEdit_2.text()
        book_description = self.textEdit.toPlainText()
        book_code = self.lineEdit_3.text()
        book_category = self.comboBox_3.currentText()
        book_author = self.comboBox_4.currentText()
        book_publisher = self.comboBox_5.currentText()
        book_price = self.lineEdit_4.text()

        self.cur.execute('''
            INSERT INTO book (book_title, book_description, book_code, book_category,
                            book_author, book_publisher, book_price)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''', (book_title, book_description, book_code, book_category, book_author,
              book_publisher, book_price))

        self.db.commit()
        self.statusBar().showMessage('New Book Added')
        # to start from new after inserting all the data
        self.lineEdit_2.setText('')
        self.textEdit.setPlainText('')
        self.lineEdit_3.setText('')
        self.comboBox_3.setCurrentText('Gaming')
        self.comboBox_4.setCurrentText('Robin Sharma')
        self.comboBox_5.setCurrentText('Penguin International')
        self.lineEdit_4.setText('')

        self.show_books()

    # for searching of a particular book
    def search_books(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        book_title = self.lineEdit_8.text()

        sql = '''SELECT * FROM book WHERE book_title = %s'''
        self.cur.execute(sql, [(book_title)])

        data = self.cur.fetchone()

        print(data)

        self.lineEdit_5.setText(data[1])
        self.lineEdit_7.setText(data[3])
        self.comboBox_7.setCurrentText(data[4])
        self.comboBox_8.setCurrentText(data[5])
        self.comboBox_6.setCurrentText(data[6])
        self.lineEdit_6.setText(str(data[7]))
        self.textEdit_2.setPlainText(data[2])

    # for editing a book
    def edit_books(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()


        book_title = self.lineEdit_5.text()
        book_description = self.textEdit_2.toPlainText()
        book_code = self.lineEdit_7.text()
        book_category = self.comboBox_7.currentText()
        book_author = self.comboBox_8.currentText()
        book_publisher = self.comboBox_6.currentText()
        book_price = self.lineEdit_6.text()

        search_book = self.lineEdit_5.text()

        self.cur.execute('''
        UPDATE book SET book_title = %s, book_description = %s, book_code = %s, book_category = %s, 
        book_author = %s, book_publisher = %s, book_price = %s WHERE book_title = %s
        ''', (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price,
              search_book))

        self.db.commit()
        self.statusBar().showMessage('Book Updated')

        self.show_books()

    # for deleting a book
    def delete_book(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        book_title = self.lineEdit_8.text()

        warning = QMessageBox.warning(self, 'Delete Book', "Are you sure you want to delete this book ?",
                            QMessageBox.Yes | QMessageBox.No)

        if warning == QMessageBox.Yes:
            sql = '''DELETE FROM book WHERE book_title = %s'''
            self.cur.execute(sql, [book_title])
            self.db.commit()
            self.statusBar().showMessage('Book Deleted')

            self.show_books()


    # #############################
    # ###### operations under user section ##########

    def add_new_client(self):
        client_name = self.lineEdit_22.text()
        client_email = self.lineEdit_23.text()
        client_number = self.lineEdit_24.text()

        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        self.cur.execute('''
        INSERT INTO clients (client_name, client_email, client_number)
        VALUES (%s, %s, %s)
        ''', (client_name, client_email, client_number))

        self.db.commit()
        self.db.close()
        self.statusBar().showMessage('New Client Added')
        self.show_client()

    def edit_client(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        original_number = self.lineEdit_28.text()

        client_name = self.lineEdit_25.text()
        client_email = self.lineEdit_26.text()
        client_number = self.lineEdit_27.text()

        self.cur.execute('''
        UPDATE clients SET client_name = %s, client_email = %s, client_number = %s WHERE client_number = %s
        ''',(client_name, client_email, client_number, original_number))

        self.db.commit()
        self.db.close()
        self.statusBar().showMessage('Client Data Updated')
        self.show_client()

    def show_client(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        self.cur.execute('''
        SELECT client_name, client_email, client_number FROM clients
        ''')

        data = self.cur.fetchall()

        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)

        for row, info in enumerate(data):
            for column, item in enumerate(info):
                self.tableWidget_6.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_pos = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_pos)

        self.db.close()

    def search_client(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        client_number = self.lineEdit_28.text()

        sql = '''SELECT * FROM clients WHERE client_number = %s'''
        self.cur.execute(sql, [client_number])
        data = self.cur.fetchone()

        self.lineEdit_25.setText(data[1])
        self.lineEdit_26.setText(data[2])
        self.lineEdit_27.setText(data[3])

    def delete_client(self):

        warning_message = QMessageBox.warning(self, 'Delete Client',
                                              'Are you sure you want to delete this client',
                                              QMessageBox.Yes | QMessageBox.No)

        if warning_message == QMessageBox.Yes:
            self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                      db='library_management_system')
            self.cur = self.db.cursor()

            client_number = self.lineEdit_28.text()

            sql = ''' DELETE FROM clients WHERE client_number = %s '''
            self.cur.execute(sql, [client_number])

            self.db.commit()
            self.db.close()
            self.statusBar().showMessage('Client Data Deleted')

    # #############################
    # ###### operations under user section ##########

    # for adding a new user
    def add_new_user(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        username = self.lineEdit_9.text()
        email = self.lineEdit_10.text()
        password = self.lineEdit_11.text()
        confirmPassword = self.lineEdit_12.text()

        # checking the password with confirm password
        if password == confirmPassword:
            self.cur.execute('''
            INSERT INTO users(user_name, user_email, user_password)
            VALUES (%s, %s, %s)
            ''', (username, email, password))

            self.db.commit()
            self.statusBar().showMessage('New User Added')
        else:
            self.label_30.setText("Your password & confirm password doesn't match")

    # for login of a user
    def login(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        username = self.lineEdit_13.text()
        password = self.lineEdit_14.text()

        sql = '''SELECT * FROM users'''

        self.cur.execute(sql)
        data = self.cur.fetchall()

        for info in data:
            if username == info[1] and password == info[3]:
                self.statusBar().showMessage('Login Successful')
                self.groupBox_4.setEnabled(True)

                self.lineEdit_18.setText(info[1])
                self.lineEdit_15.setText(info[2])
                self.lineEdit_17.setText(info[3])

            else:
                self.statusBar().showMessage('Login Unsuccessful')

    # for editing the user data
    def edit_user(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add books tab
        self.cur = self.db.cursor()

        original_username = self.lineEdit_13.text()
        username = self.lineEdit_18.text()
        email = self.lineEdit_15.text()
        password = self.lineEdit_17.text()
        confirmPassword = self.lineEdit_16.text()

        if password == confirmPassword:
            self.cur.execute('''
            UPDATE users SET user_name = %s, user_email = %s, user_password = %s WHERE user_name = %s
            ''', (username, email, password, original_username))

            self.db.commit()
            self.statusBar().showMessage('User Data Updated Successfully')


        else:
            self.statusBar().showMessaage("Password and Confirm Password doesn't match")

    # #############################
    # ###### operations under settings section ##########

    def add_category(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # taking input to the database
        category_name = self.lineEdit_19.text()

        # adding category names to database
        self.cur.execute('''
            INSERT INTO category (category_name) VALUES (%s)
        ''', (category_name,))

        # commit this change to database
        self.db.commit()
        # to display that the new category is added
        self.statusBar().showMessage('New Category Added')

        self.lineEdit_19.setText('') # to clear the text editor after anyone type the information

        self.show_category()
        self.show_category_ui()

    # function to show names in the category tab
    def show_category(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # fetching category names from database and showing it to our table
        self.cur.execute('''SELECT category_name FROM category''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget_3.setRowCount(0)  # to show new category added after anyone inserts it
            self.tableWidget_3.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                    row_pos = self.tableWidget_3.rowCount()
                    self.tableWidget_3.insertRow(row_pos)

    # function to insert author names into the database
    def add_author(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add author tab
        self.cur = self.db.cursor()

        author_name = self.lineEdit_20.text()

        # adding author names to database
        self.cur.execute('''
            INSERT INTO author (author_name) VALUES (%s)
        ''', (author_name,))

        # commit this change to database
        self.db.commit()
        self.lineEdit_20.setText('')  # to clean the text editor
        # to display that the new author is added
        self.statusBar().showMessage('New Author Added')
        self.show_author()
        self.show_author_ui()

    # function to show author names in the settings tab
    def show_author(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # fetching category names from database and showing it to our table
        self.cur.execute('''SELECT author_name FROM author''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget_4.setRowCount(0)  # to show new category added after anyone inserts it
            self.tableWidget_4.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                    row_pos = self.tableWidget_4.rowCount()
                    self.tableWidget_4.insertRow(row_pos)

    # function to add publisher names in the database
    def add_publisher(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        # execute new cursor for add publisher tab
        self.cur = self.db.cursor()

        publisher_name = self.lineEdit_21.text()

        # adding publisher names to database
        self.cur.execute('''
            INSERT INTO publisher (publisher_name) VALUES (%s)
        ''', (publisher_name,))

        # commit this change to database
        self.db.commit()
        self.lineEdit_21.setText('')
        # to display that the new publisher is added
        self.statusBar().showMessage('New Publisher Added')
        self.show_publisher()
        self.show_publisher_ui()

    # function to show publisher names in the settings tab
    def show_publisher(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # fetching category names from database and showing it to our table
        self.cur.execute('''SELECT publisher_name FROM publisher''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget_5.setRowCount(0)  # to show new category added after anyone inserts it
            self.tableWidget_5.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_5.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1

                    row_pos = self.tableWidget_5.rowCount()
                    self.tableWidget_5.insertRow(row_pos)

    # ######### show settings data in UI ############
    # ###############################################

    # to show our category data in our overall app
    def show_category_ui(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # fetching category names from database and showing it to our table
        self.cur.execute('''SELECT category_name FROM category''')

        data = self.cur.fetchall()

        self.comboBox_3.clear()
        for category in data:
            self.comboBox_3.addItem(category[0])
            self.comboBox_7.addItem(category[0])


    # to show our author data in our overall app
    def show_author_ui(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # fetching category names from database and showing it to our table
        self.cur.execute('''SELECT author_name FROM author''')

        data = self.cur.fetchall()

        self.comboBox_4.clear()
        for author in data:
            self.comboBox_4.addItem(author[0])
            self.comboBox_8.addItem(author[0])


    # to show our publisher data in our overall app
    def show_publisher_ui(self):
        # connecting to database
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')

        self.cur = self.db.cursor()

        # fetching category names from database and showing it to our table
        self.cur.execute('''SELECT publisher_name FROM publisher''')

        data = self.cur.fetchall()

        self.comboBox_5.clear()
        for publisher in data:
            self.comboBox_5.addItem(publisher[0])
            self.comboBox_6.addItem(publisher[0])

    # ######### show settings data in UI ############
    # ###############################################

    def classic_themes(self):
        style = open('themes/classic.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def dark_blue_themes(self):
        style = open('themes/dark_blue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def dark_gray_themes(self):
        style = open('themes/dark_gray.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def dark_orange_themes(self):
        style = open('themes/dark_orange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    # Export data to excel files ############
    # function to export operations data to excel file
    def export_operations(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()
        self.cur.execute('''
        SELECT book_name, client_name, type, days, date_from, date_to FROM operations
        ''')

        data = self.cur.fetchall()

        wb = Workbook('Operations_data.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0, 0, 'Book title')
        sheet1.write(0, 1, 'Client Name')
        sheet1.write(0, 2, 'Type')
        sheet1.write(0, 3, 'Days')
        sheet1.write(0, 4, 'Date_from')
        sheet1.write(0, 5, 'Date_to')

        row_num = 1
        for row in data:
            col_num = 0
            for item in row:
                sheet1.write(row_num, col_num, str(item))
                col_num += 1
            row_num += 1

        wb.close()
        self.statusBar().showMessage('Report created successfully')

# function to export books data to excel file
    def export_books(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        self.cur.execute('''
        SELECT book_code, book_title, book_description, book_category,
                            book_author, book_publisher, book_price FROM book
        ''')

        data = self.cur.fetchall()

        wb = Workbook('Books_data.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0, 0, 'Book_Code')
        sheet1.write(0, 1, 'Book_Title')
        sheet1.write(0, 2, 'Book_Description')
        sheet1.write(0, 3, 'Book_Category')
        sheet1.write(0, 4, 'Book_Author')
        sheet1.write(0, 5, 'Book_Publisher')
        sheet1.write(0, 6, 'Book_Price')

        row_num = 1
        for row in data:
            col_num = 0
            for item in row:
                sheet1.write(row_num, col_num, str(item))
                col_num += 1
            row_num += 1

        wb.close()
        self.statusBar().showMessage('Report created successfully')

# function to export clients data to excel file
    def export_clients(self):
        self.db = MySQLdb.connect(host='localhost', user='akhil', password='movies',
                                  db='library_management_system')
        self.cur = self.db.cursor()

        self.cur.execute('''
        SELECT client_name, client_email, client_number FROM clients
        ''')

        data = self.cur.fetchall()

        wb = Workbook('Clients_data.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0, 0, 'Client_Name')
        sheet1.write(0, 1, 'Client_Email')
        sheet1.write(0, 2, 'Client_Number')

        row_num = 1
        for row in data:
            col_num = 0
            for item in row:
                sheet1.write(row_num, col_num, str(item))
                col_num += 1
            row_num += 1

        wb.close()
        self.statusBar().showMessage('Report created successfully')



def main():
    app = QApplication(sys.argv)
    window = Login()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()

