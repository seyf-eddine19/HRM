#      Project        :   مسابقة تقافية 
#      Date           :   13-08-2025
#      Author         :   Seyf Eddine
#      Gmail          :   seyfeddine.freelance@gmail.com
#      WhatsApp       :   (+213) 794 87 85 08
#      Python         :   3.12.10


import os
import sys
import shlex
import shutil
import subprocess
import sqlite3
import logging
import pandas as pd
from datetime import datetime
from hijridate import Gregorian
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from MainWindow import Ui_MainWindow
from EditEmployeePage import Ui_EditEmployeeDialog


class DB_conn:
    def __init__(self, database='db\\employees.db'):
        self.database = database
        self.switch_cols = {
            'department_types': ['id', 'name'],
            'job_titles': ['id', 'name'],
            'passport_types': ['id', 'name'],
            'visa_types': ['id', 'name'],
            'employees': [
                'id', 'general_number', 'name_ar', 'name_en',
                'birth_date', 'national_id', 'id_issue_date', 'id_expiry_date',
                'department_id', 'job_title_id', 'phone', 'iban_number', 'role',
                'photo_path', 'docs_path'
            ],
            'passports': [
                'id', 'employee_id', 'passport_number', 'passport_type_id', 'issue_date',
                'expiry_date', 'issue_authority', 'delivered_by',
                'received_by', 'received_at', 'custodian', 'doc_path'
            ],
            'visas': [
                'id', 'passport_id', 'visa_number', 'visa_type_id', 'issue_date', 'expiry_date',
                'doc_path'
            ]
        }

        if not os.path.exists(self.database):
            os.makedirs(os.path.dirname(self.database), exist_ok=True)
            self.init_db()

    def _get_connection(self):
        cnx = sqlite3.connect(self.database)
        cnx.execute("PRAGMA foreign_keys = ON")
        return cnx

    def execute_query(self, query, data=None, fetch=False, return_id=False):
        cnx = self._get_connection()
        cur = cnx.cursor()
        try:
            if data:
                cur.execute(query, data)
            else:
                cur.execute(query)

            if fetch or query.strip().lower().startswith('select'):
                df = pd.read_sql_query(query, cnx, params=data)
                return df

            cnx.commit()
            if return_id:
                return cur.lastrowid
            return "تمت العملية بنجاح"
        except Exception as err:
            return f"حدث خطأ: {str(err)}"
        finally:
            cur.close()
            cnx.close()

    def insert(self, table, data):
        if table not in self.switch_cols:
            return f"جدول {table} غير موجود"
        cols = self.switch_cols[table][1:]  # بدون id
        placeholders = ', '.join(['?' for _ in cols])
        query = f'INSERT INTO "{table}" ({",".join(cols)}) VALUES ({placeholders})'
        return self.execute_query(query, data, return_id=True)

    def update(self, table, data, ids):
        if table not in self.switch_cols:
            return f"جدول {table} غير موجود"
        cols = self.switch_cols[table][1:]  # بدون id
        set_clause = ', '.join([f'{col}=?' for col in cols])
        query = f'UPDATE "{table}" SET {set_clause} WHERE id=?'
        vals = data + [ids[0]]
        return self.execute_query(query, vals)

    def select(self, table):
        if table not in self.switch_cols:
            return f"جدول {table} غير موجود"
        query = f'SELECT * FROM "{table}"'
        return self.execute_query(query, fetch=True)

    def delete(self, table, ids):
        if table not in self.switch_cols:
            return f"جدول {table} غير موجود"
        query = f'DELETE FROM "{table}" WHERE id=?'
        return self.execute_query(query, ids)

    def delete_all(self, table):
        if table not in self.switch_cols:
            return f"جدول {table} غير موجود"
        query = f'DELETE FROM "{table}"'
        return self.execute_query(query)


class ManageTypesDialog(QtWidgets.QDialog):
    def __init__(self, db_conn, table_name, parent=None):
        super().__init__(parent)
        self.db_conn = db_conn
        self.table_name = table_name

        # تحويل اسم الجدول إلى العربية للعرض
        self.ar_table_name = {
            'department_types': 'الأقسام',
            'job_titles': 'المسمى الوظيفي',
            'passport_types': 'أنواع الجوازات',
            'visa_types': 'أنواع التأشيرات'
        }.get(table_name, table_name)

        self.setWindowTitle(f"إدارة {self.ar_table_name}")
        self.resize(500, 350)
        self.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Label أعلى الجدول
        self.label_title = QtWidgets.QLabel(f"إدارة {self.ar_table_name}")
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        self.label_title.setFont(font)
        layout.addWidget(self.label_title)

        # Table
        self.table = QtWidgets.QTableWidget()
        layout.addWidget(self.table)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['ID', 'الاسم'])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)

        # Buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_save_row = QtWidgets.QPushButton("حفظ")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_save_row.setIcon(icon1)
        self.btn_save_row.setIconSize(QtCore.QSize(28, 28))
        
        self.btn_edit_row = QtWidgets.QPushButton("تعديل")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-compose.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_edit_row.setIcon(icon1)
        self.btn_edit_row.setIconSize(QtCore.QSize(28, 28))

        self.btn_delete = QtWidgets.QPushButton("حذف")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-delete.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_delete.setIcon(icon1)
        self.btn_delete.setIconSize(QtCore.QSize(28, 28))

        self.btn_close = QtWidgets.QPushButton("إغلاق")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_close.setIcon(icon1)
        self.btn_close.setIconSize(QtCore.QSize(28, 28))

        btn_layout.addWidget(self.btn_close)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_edit_row)
        btn_layout.addWidget(self.btn_save_row)
        layout.addLayout(btn_layout)

        # Connections
        self.btn_save_row.clicked.connect(self.save_row)
        self.btn_edit_row.clicked.connect(self.edit_selected_row)
        self.btn_delete.clicked.connect(self.delete_item)
        self.btn_close.clicked.connect(self.close)

    def load_data(self):
        df = self.db_conn.select(self.table_name)  # DataFrame
        self.table.setRowCount(len(df) + 1)
        # الصف الفارغ الأول للإضافة أو تعديل
        id_item = QtWidgets.QTableWidgetItem("")  
        id_item.setFlags(QtCore.Qt.ItemIsEnabled)  # غير قابل للتعديل
        self.table.setItem(0, 0, id_item)
    
        name_item = QtWidgets.QTableWidgetItem("")  
        name_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsEnabled)
        self.table.setItem(0, 1, name_item)
    
        # باقي الصفوف من قاعدة البيانات (غير قابلة للتعديل مباشرة)
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            id_val = str(row['id'])
            name_val = str(row['name'])
    
            id_item = QtWidgets.QTableWidgetItem(id_val)
            id_item.setFlags(QtCore.Qt.ItemIsEnabled)
            self.table.setItem(i, 0, id_item)
    
            name_item = QtWidgets.QTableWidgetItem(name_val)
            name_item.setFlags(QtCore.Qt.ItemIsEnabled)
            self.table.setItem(i, 1, name_item)

    def save_row(self):
        # الصف الأول دائمًا للإضافة أو تعديل
        id_item = self.table.item(0, 0)
        name_item = self.table.item(0, 1)
        if not name_item or not name_item.text().strip():
            QtWidgets.QMessageBox.warning(self, "خطأ", "الرجاء إدخال الاسم")
            return

        name = name_item.text().strip()

        if id_item and id_item.text().isdigit():
            # تعديل سجل موجود
            self.db_conn.update(self.table_name, [name], [int(id_item.text())])
        else:
            # إضافة سجل جديد
            self.db_conn.insert(self.table_name, [name])

        self.load_data()

    def edit_selected_row(self):
        row = self.table.currentRow()
        if row <= 0:
            QtWidgets.QMessageBox.warning(self, "خطأ", "الرجاء اختيار صف للتعديل")
            return

        id_item = self.table.item(row, 0)
        name_item = self.table.item(row, 1)
        if not id_item or not name_item:
            return

        # نسخ البيانات إلى الصف الأول ليتم تعديلها
        self.table.item(0, 0).setText(id_item.text())
        self.table.item(0, 1).setText(name_item.text())
        self.table.setCurrentCell(0, 1)

    def delete_item(self):
        row = self.table.currentRow()
        if row <= 0:
            QtWidgets.QMessageBox.warning(self, "خطأ", "لا يمكن حذف الصف الفارغ")
            return

        id_item = self.table.item(row, 0)
        if not id_item or not id_item.text().isdigit():
            return

        reply = QtWidgets.QMessageBox.question(self, "حذف", "هل أنت متأكد من الحذف؟",
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            self.db_conn.delete(self.table_name, [int(id_item.text())])

            self.load_data()


class PassportDialog(QtWidgets.QDialog):
    def __init__(self, db_conn, employee_id, passport_data=None, parent=None):
        self.root = parent
        super().__init__(parent)
        self.db_conn = db_conn
        self.employee_id = employee_id
        self.passport_data = passport_data or {}
        self.setWindowTitle("إضافة / تعديل جواز السفر")
        self.resize(600, 300)
        self.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QtWidgets.QFormLayout(self)

        # رقم الجواز
        self.passport_number = QtWidgets.QLineEdit()
        layout.addRow("رقم الجواز:", self.passport_number)

        # نوع الجواز + زر إدارة الأنواع
        df_types = self.db_conn.select('passport_types')
        self.passport_types = {row['id']: row['name'] for _, row in df_types.iterrows()}
        type_layout = QtWidgets.QHBoxLayout()
        self.passport_type = QtWidgets.QComboBox()
        for pid, pname in self.passport_types.items():
            self.passport_type.addItem(pname, pid)
        btn_manage = QtWidgets.QToolButton()
        btn_manage.setIcon(QtGui.QIcon(":/img/icon-add.png"))
        btn_manage.setIconSize(QtCore.QSize(20,20))
        btn_manage.clicked.connect(lambda _, t='passport_types': self.open_manage_types(t))
        type_layout.addWidget(self.passport_type)
        type_layout.addWidget(btn_manage)
        layout.addRow("نوع الجواز:", type_layout)

        # تاريخ الإصدار
        self.issue_date = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.issue_date.setCalendarPopup(True)
        layout.addRow("تاريخ الإصدار:", self.issue_date)

        # تاريخ الانتهاء
        self.expiry_date = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.expiry_date.setCalendarPopup(True)
        layout.addRow("تاريخ الانتهاء:", self.expiry_date)
        # جهة الإصدار
        self.issue_authority = QtWidgets.QLineEdit()
        layout.addRow("جهة الإصدار:", self.issue_authority)
        # المسلم
        self.delivered_by = QtWidgets.QLineEdit()
        layout.addRow("المسلّم:", self.delivered_by)
        # المستلم
        self.received_by = QtWidgets.QLineEdit()
        layout.addRow("المستلم:", self.received_by)
        # تاريخ الاستلام
        self.received_at = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.received_at.setCalendarPopup(True)
        layout.addRow("تاريخ الاستلام:", self.received_at)
        # على عهدة
        # self.custodian = QtWidgets.QComboBox()
        # self.custodian.addItems(['الشركة', 'الموظف'])
        # layout.addRow("على عهدة:", self.custodian)

        # صورة الجواز
        file_layout = QtWidgets.QHBoxLayout()
        self.doc_path = QtWidgets.QLineEdit()
        self.doc_path.setReadOnly(True)
        btn_upload = QtWidgets.QPushButton("اختر صورة")
        btn_upload.clicked.connect(self.upload_file)
        file_layout.addWidget(self.doc_path)
        file_layout.addWidget(btn_upload)
        layout.addRow("صورة الجواز:", file_layout)

        # أزرار حفظ وإلغاء
        btn_save = QtWidgets.QPushButton("حفظ")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        btn_save.setIcon(icon1)
        btn_save.setIconSize(QtCore.QSize(28, 28))
        
        btn_save.clicked.connect(self.save_passport)
        btn_cancel = QtWidgets.QPushButton("إلغاء")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        btn_cancel.setIcon(icon1)
        btn_cancel.setIconSize(QtCore.QSize(28, 28))
        btn_cancel.clicked.connect(self.reject)

        hlayout_btn = QtWidgets.QHBoxLayout()
        hlayout_btn.addWidget(btn_save)
        hlayout_btn.addWidget(btn_cancel)
        layout.addRow(hlayout_btn)

    def open_manage_types(self, table_name):
        dlg = ManageTypesDialog(self.db_conn, table_name, self)
        dlg.exec_()
        # تحديث الأنواع بعد الإضافة
        self.passport_type.clear()
        df_types = self.db_conn.select('passport_types')
        for _, row in df_types.iterrows():
            self.passport_type.addItem(row['name'], row['id'])

    def load_data(self):
        if not self.passport_data:
            return
        self.passport_number.setText(self.passport_data.get("passport_number", ""))
        pid = self.passport_data.get("passport_type_id")
        if pid:
            idx = self.passport_type.findData(pid)
            if idx >= 0:
                self.passport_type.setCurrentIndex(idx)
        self.issue_date.setDate(QtCore.QDate.fromString(
            self.passport_data.get("issue_date","1999-01-01"), "yyyy-MM-dd"
        ))
        self.expiry_date.setDate(QtCore.QDate.fromString(
            self.passport_data.get("expiry_date","1999-01-01"), "yyyy-MM-dd"
        ))
        self.issue_authority.setText(self.passport_data.get("issue_authority",""))
        self.delivered_by.setText(self.passport_data.get("delivered_by",""))
        self.received_by.setText(self.passport_data.get("received_by",""))
        self.received_at.setDate(QtCore.QDate.fromString(
            self.passport_data.get("received_at","1999-01-01"), "yyyy-MM-dd"
        ))
        # custodian = self.passport_data.get("custodian","الشركة")
        # idx = self.custodian.findText(custodian)
        # if idx >= 0:
        #     self.custodian.setCurrentIndex(idx)
        self.doc_path.setText(self.passport_data.get("doc_path",""))

    def upload_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "اختر صورة الجواز", "", "Images (*.png *.jpg *.jpeg *.bmp)")
        if path:
            self.doc_path.setText(path)

    def save_passport(self):
        # التحقق من إدخال البيانات الأساسية
        if not self.passport_number.text().strip():
            QtWidgets.QMessageBox.warning(self, "تنبيه", "يرجى إدخال رقم الجواز.")
            return

        doc_path = self.doc_path.text()
        if not os.path.exists(self.root.docs_folder):
            os.makedirs(self.root.docs_folder)
        
        doc_path = self.copy_file(doc_path, self.root.docs_folder)

        data = [
            self.employee_id,
            self.passport_number.text(),
            self.passport_type.currentData(),
            self.issue_date.date().toString("yyyy-MM-dd"),
            self.expiry_date.date().toString("yyyy-MM-dd"),
            self.issue_authority.text(),
            self.delivered_by.text(),
            self.received_by.text(),
            self.received_at.text(),
            self.passport_data.get("custodian", "الشركة"),# self.custodian.currentText(),
            doc_path
        ]

        try:
            if self.passport_data.get("id"):  # تعديل
                self.db_conn.update('passports', data, [self.passport_data["id"]])
                QtWidgets.QMessageBox.information(self, "تم التعديل", "تم تعديل بيانات الجواز بنجاح.")
            else:
                result = self.db_conn.insert('passports', data)
                if isinstance(result, str) and result.startswith("حدث خطأ"):
                    QtWidgets.QMessageBox.critical(self, "خطأ", result)
                else:
                    QtWidgets.QMessageBox.information(self, "تم الحفظ", "تمت إضافة الجواز بنجاح.")

            self.accept()

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء الحفظ:\n{str(e)}")

    @staticmethod
    def copy_file(src, dest_folder):
        if not src or not os.path.exists(src):
            return "" 
       
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
    
        filename = os.path.basename(src)
        name, ext = os.path.splitext(filename)
        dst = os.path.join(dest_folder, filename)
    
        counter = 1
        while os.path.exists(dst) or os.path.abspath(src) == os.path.abspath(dst):
            dst = os.path.join(dest_folder, f"{name} - copie{counter}{ext}")
            counter += 1
    
        shutil.copy(src, dst)
        return dst


class VisaDialog(QtWidgets.QDialog):
    def __init__(self, db_conn, passport_id, visa_data=None, parent=None):
        super().__init__(parent)
        self.root = parent
        self.db_conn = db_conn
        self.passport_id = passport_id
        self.visa_data = visa_data or {}
        self.setWindowTitle("إضافة / تعديل التأشيرة")
        self.resize(600, 300)
        self.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QtWidgets.QFormLayout(self)

        # رقم التأشيرة
        self.visa_number = QtWidgets.QLineEdit()
        layout.addRow("رقم التأشيرة:", self.visa_number)

        # نوع التأشيرة + زر إدارة الأنواع
        df_types = self.db_conn.select('visa_types')
        self.visa_types = {row['id']: row['name'] for _, row in df_types.iterrows()}
        type_layout = QtWidgets.QHBoxLayout()
        self.visa_type = QtWidgets.QComboBox()
        for vid, vname in self.visa_types.items():
            self.visa_type.addItem(vname, vid)
        btn_manage = QtWidgets.QToolButton()
        btn_manage.setIcon(QtGui.QIcon(":/img/icon-add.png"))
        btn_manage.setIconSize(QtCore.QSize(20, 20))
        btn_manage.clicked.connect(lambda _, t='visa_types': self.open_manage_types(t))
        type_layout.addWidget(self.visa_type)
        type_layout.addWidget(btn_manage)
        layout.addRow("نوع التأشيرة:", type_layout)

        # تاريخ الإصدار
        self.issue_date = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.issue_date.setCalendarPopup(True)
        layout.addRow("تاريخ الإصدار:", self.issue_date)

        # تاريخ الانتهاء
        self.expiry_date = QtWidgets.QDateEdit(QtCore.QDate.currentDate())
        self.expiry_date.setCalendarPopup(True)
        layout.addRow("تاريخ الانتهاء:", self.expiry_date)

        # صورة التأشيرة
        file_layout = QtWidgets.QHBoxLayout()
        self.doc_path = QtWidgets.QLineEdit()
        self.doc_path.setReadOnly(True)
        btn_upload = QtWidgets.QPushButton("اختر صورة")
        btn_upload.clicked.connect(self.upload_file)
        file_layout.addWidget(self.doc_path)
        file_layout.addWidget(btn_upload)
        layout.addRow("صورة التأشيرة:", file_layout)

        # أزرار
        btn_save = QtWidgets.QPushButton("حفظ")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        btn_save.setIcon(icon1)
        btn_save.setIconSize(QtCore.QSize(28, 28))
        
        btn_save.clicked.connect(self.save_visa)
        btn_cancel = QtWidgets.QPushButton("إلغاء")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        btn_cancel.setIcon(icon1)
        btn_cancel.setIconSize(QtCore.QSize(28, 28))
        btn_cancel.clicked.connect(self.reject)
        
        hlayout_btn = QtWidgets.QHBoxLayout()
        hlayout_btn.addWidget(btn_save)
        hlayout_btn.addWidget(btn_cancel)
        layout.addRow(hlayout_btn)

    def open_manage_types(self, table_name):
        dlg = ManageTypesDialog(self.db_conn, table_name, self)
        dlg.exec_()
        self.visa_type.clear()
        df_types = self.db_conn.select('visa_types')
        for _, row in df_types.iterrows():
            self.visa_type.addItem(row['name'], row['id'])

    def load_data(self):
        if not self.visa_data:
            return
        self.visa_number.setText(self.visa_data.get("visa_number", ""))
        vtid = self.visa_data.get("visa_type_id")
        if vtid:
            idx = self.visa_type.findData(vtid)
            if idx >= 0:
                self.visa_type.setCurrentIndex(idx)
        self.issue_date.setDate(QtCore.QDate.fromString(
            self.visa_data.get("issue_date", "1999-01-01"), "yyyy-MM-dd"
        ))
        self.expiry_date.setDate(QtCore.QDate.fromString(
            self.visa_data.get("expiry_date", "1999-01-01"), "yyyy-MM-dd"
        ))
        self.doc_path.setText(self.visa_data.get("doc_path", ""))

    def upload_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "اختر صورة التأشيرة", "", "Images (*.png *.jpg *.jpeg *.bmp)")
        if path:
            self.doc_path.setText(path)

    def save_visa(self):
        if not self.visa_number.text().strip():
            QtWidgets.QMessageBox.warning(self, "تنبيه", "يرجى إدخال رقم التأشيرة.")
            return

        doc_path = self.doc_path.text()
        if not os.path.exists(self.root.docs_folder):
            os.makedirs(self.root.docs_folder)
        
        doc_path = self.copy_file(doc_path, self.root.docs_folder)

        data = [
            self.passport_id,
            self.visa_number.text(),
            self.visa_type.currentData(),
            self.issue_date.date().toString("yyyy-MM-dd"),
            self.expiry_date.date().toString("yyyy-MM-dd"),
            doc_path
        ]

        try:
            if self.visa_data.get("id"):  # تعديل
                self.db_conn.update('visas', data, [self.visa_data["id"]])
                QtWidgets.QMessageBox.information(self, "تم التعديل", "تم تعديل بيانات التأشيرة بنجاح.")
            else:
                result = self.db_conn.insert('visas', data)
                if isinstance(result, str) and result.startswith("حدث خطأ"):
                    QtWidgets.QMessageBox.critical(self, "خطأ", result)
                else:
                    QtWidgets.QMessageBox.information(self, "تم الحفظ", "تمت إضافة التأشيرة بنجاح.")

            self.accept()

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء الحفظ:\n{str(e)}")

    @staticmethod
    def copy_file(src, dest_folder):
        if not src or not os.path.exists(src):
            return ""  
        
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
    
        filename = os.path.basename(src)
        name, ext = os.path.splitext(filename)
        dst = os.path.join(dest_folder, filename)
    
        counter = 1
        while os.path.exists(dst) or os.path.abspath(src) == os.path.abspath(dst):
            dst = os.path.join(dest_folder, f"{name} - copie{counter}{ext}")
            counter += 1
    
        shutil.copy(src, dst)
        return dst


class EditEmployeeDialog(Ui_EditEmployeeDialog, QtWidgets.QDialog):
    def __init__(self, root, employee__id=None, action='create'):
        super().__init__()
        self.root = root
        self.db_conn = self.root.db_conn
        self.employee__id = employee__id
        self.employee_data = {}
        self.setup()

    def setup(self):
        self.setupUi(self)
        self.setup_connections()
        self.load_departments()
        self.load_job_titles()
        self.setup_passport_tab()
        self.setup_visa_tab()
        self.tabWidget.setTabEnabled(1, False)
        self.tabWidget.setTabEnabled(2, False)
        if self.employee__id:
            self.tabWidget.setTabEnabled(1, True)
            self.tabWidget.setTabEnabled(2, True)
            self.load_employee_data()
            self.load_passports()
            self.docs_folder = "documents/" + str(self.name_ar.text())
            self.load_documents()

    def setup_connections(self):
        self.saveBtn.clicked.connect(self.save_employee)
        self.cancelBtn.clicked.connect(self.exit)
        self.btn_upload_photo.clicked.connect(self.upload_photo)
        self.btn_manage_departments.clicked.connect(lambda: self.open_manage_types('department_types'))
        self.btn_manage_jobs.clicked.connect(lambda: self.open_manage_types('job_titles'))

        # Passport buttons
        self.btn_add_passport.clicked.connect(self.add_passport)
        self.btn_edit_passport.clicked.connect(self.edit_passport)
        self.btn_delete_passport.clicked.connect(self.delete_passport)
        self.table_passport.itemSelectionChanged.connect(self.on_passport_selected)

        # Visa buttons
        self.btn_add_visa.clicked.connect(self.add_visa)
        self.btn_edit_visa.clicked.connect(self.edit_visa)
        self.btn_delete_visa.clicked.connect(self.delete_visa)

        self.btnAddFile.clicked.connect(self.add_document)
        self.btnDeleteFile.clicked.connect(self.delete_document)
        self.btnOpen.clicked.connect(self.open_document)
        self.btnRefresh.clicked.connect(self.load_documents)

    def open_manage_types(self, table_name):
        dlg = ManageTypesDialog(self.db_conn, table_name, self)
        dlg.exec_()
        # Reload comboboxes or tables
        if table_name == 'department_types':
            self.load_departments()
        elif table_name == 'job_titles':
            self.load_job_titles()
        elif table_name == 'passport_types':
            self.setup_passport_tab()
            self.load_passports()
        elif table_name == 'visa_types':
            self.setup_visa_tab()

    # Departments & Jobs
    def load_departments(self):
        df_dept = self.db_conn.select('department_types')
        self.department.clear()
        for _, row in df_dept.iterrows():
            self.department.addItem(row['name'], row['id'])

        if self.employee_data:
            dept_id = self.employee_data.get("department_id")
            if dept_id:
                idx = self.department.findData(dept_id)
                if idx >= 0: self.department.setCurrentIndex(idx)

    def load_job_titles(self):
        df_job = self.db_conn.select('job_titles')
        self.job_title.clear()
        for _, row in df_job.iterrows():
            self.job_title.addItem(row['name'], row['id'])

        if self.employee_data:
            job_id = self.employee_data.get("job_title_id")
            if job_id:
                idx = self.job_title.findData(job_id)
                if idx >= 0: self.job_title.setCurrentIndex(idx)

    # Employee data
    def load_employee_data(self):
        if not self.employee__id: return
        df = self.db_conn.execute_query(
            "SELECT * FROM employees WHERE id=?", [self.employee__id], fetch=True
        )
        if df.empty: return
        self.employee_data = df.iloc[0].to_dict()
        self.general_number.setText(str(self.employee_data.get("general_number", "")))
        self.name_ar.setText(self.employee_data.get("name_ar", ""))
        self.name_en.setText(self.employee_data.get("name_en", ""))
        self.national_id.setText(self.employee_data.get("national_id", ""))
        self.phone.setText(self.employee_data.get("phone", ""))
        self.iban_number.setText(self.employee_data.get("iban_number", ""))
        self.birth_date.setDate(QtCore.QDate.fromString(self.employee_data.get("birth_date","1999-01-01"), "yyyy-MM-dd"))
        self.id_issue_date.setDate(QtCore.QDate.fromString(self.employee_data.get("id_issue_date","1999-01-01"), "yyyy-MM-dd"))
        self.id_expiry_date.setDate(QtCore.QDate.fromString(self.employee_data.get("id_expiry_date","1999-01-01"), "yyyy-MM-dd"))

        role = self.employee_data.get("role", "فرد")
        self.role_admin.setChecked(role=="مسؤول")
        self.role_user.setChecked(role=="فرد")

        photo_path = self.employee_data.get("photo_path","")
        if photo_path and os.path.exists(photo_path):
            pixmap = QtGui.QPixmap(photo_path).scaled(self.label_photo.width(), self.label_photo.height())
            self.label_photo.setPixmap(pixmap)

    def save_employee(self):
        required_fields = {
            "الرقم_العام": self.general_number.text(),
            "رقم الهوية": self.national_id.text(),
            "الاسم بالعربية": self.name_ar.text(),
            "الاسم بالإنجليزية": self.name_en.text(),
        }
        missing_fields = [name for name, value in required_fields.items() if not value.strip()]

        if missing_fields:
            QtWidgets.QMessageBox.warning(
                self,
                "بيانات ناقصة",
                "الحقول التالية مطلوبة:\n- " + "\n- ".join(missing_fields)
            )
            return None 
        
        self.docs_folder = "documents/" + str(self.name_ar.text())
        photo_path = getattr(self, "photo_path", "")
        if not os.path.exists(self.docs_folder):
            os.makedirs(self.docs_folder)
        
        photo_path = self.copy_file(photo_path, self.docs_folder)

        data = [
            self.general_number.text(),
            self.name_ar.text(),
            self.name_en.text(),
            self.birth_date.date().toString("yyyy-MM-dd"),
            self.national_id.text(),
            self.id_issue_date.date().toString("yyyy-MM-dd"),
            self.id_expiry_date.date().toString("yyyy-MM-dd"),
            self.department.currentData(),
            self.job_title.currentData(),
            self.phone.text(),
            self.iban_number.text(),
            "مسؤول" if self.role_admin.isChecked() else "فرد",
            photo_path,
            self.docs_folder
        ]
  
        if self.employee__id:
            self.db_conn.update('employees', data, [self.employee__id])
            QMessageBox.information(self, "نجاح", "تم تعديل الموظف بنجاح")
        else:
            new_id = self.db_conn.insert('employees', data)
            if isinstance(new_id, int):
                self.employee__id = new_id
                self.tabWidget.setTabEnabled(1, True)
                self.tabWidget.setTabEnabled(2, True)
                QMessageBox.information(self, "نجاح", f"تم إضافة الموظف بنجاح، رقم التعريف: {new_id}")
            else:
                QMessageBox.warning(self, "خطأ", str(new_id))

    @staticmethod
    def copy_file(src, dest_folder):
        if not src or not os.path.exists(src):
            return "" 
        
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
    
        filename = os.path.basename(src)
        name, ext = os.path.splitext(filename)
        dst = os.path.join(dest_folder, filename)
    
        counter = 1
        while os.path.exists(dst) or os.path.abspath(src) == os.path.abspath(dst):
            dst = os.path.join(dest_folder, f"{name} - copie{counter}{ext}")
            counter += 1
    
        shutil.copy(src, dst)
        return dst

    def upload_photo(self):
        path, _ = QFileDialog.getOpenFileName(self, "اختر صورة", "", "Images (*.png *.jpg *.jpeg)")
        if path:
            self.photo_path = path
            pixmap = QtGui.QPixmap(path).scaled(self.label_photo.width(), self.label_photo.height())
            self.label_photo.setPixmap(pixmap)

    # Passport Tab
    def setup_passport_tab(self):
        self.table_passport.setColumnCount(8)
        self.table_passport.setHorizontalHeaderLabels([
            'id', 'رقم الجواز', 'نوع الجواز', 'تاريخ الإصدار', 'تاريخ الانتهاء',
            'جهة الإصدار', 'الحالة', 'صورة الجواز'
        ])

        header = self.table_passport.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        # إخفاء عمود ID عن المستخدم
        self.table_passport.setColumnHidden(0, True)

        # جعل الاختيار على الصف بالكامل
        self.table_passport.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table_passport.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)

    def load_passports(self):
        if not self.employee__id:
            return
        
        df = self.db_conn.execute_query(
            """
            SELECT 
                p.id, 
                p.passport_number, 
                t.name AS passport_type, 
                p.issue_date, 
                p.expiry_date, 
                p.issue_authority, 
                p.delivered_by,
                p.received_by,
                p.received_at,
                p.custodian,
                p.doc_path
            FROM passports p
            LEFT JOIN passport_types t ON p.passport_type_id = t.id
            WHERE p.employee_id = ?
            """, 
            [self.employee__id], 
            fetch=True
        )
    
        if not hasattr(df, "iterrows"):  # يعني مش DataFrame
            QtWidgets.QMessageBox.warning(self, "خطأ", f"فشل تحميل الجوازات: {df}")
            return
    
        self.table_passport.setRowCount(len(df))
    
        for i, row in df.iterrows():
            # for j, col in enumerate(df.columns): dont use loop
            #     if col == 'custodian':
            #         text = "مستلم" if row['custodian'] == "الموظف" else "غير مستلم"
            #     else:
            #         text = str(row[col]) if row[col] else ""
            #     item = QtWidgets.QTableWidgetItem(text)
            #     self.table_passport.setItem(i, j, item)
            self.table_passport.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row["id"])))
            self.table_passport.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row["passport_number"])))
            self.table_passport.setItem(i, 2, QtWidgets.QTableWidgetItem(str(row["passport_type"])))
            self.table_passport.setItem(i, 3, QtWidgets.QTableWidgetItem(str(row["issue_date"])))
            self.table_passport.setItem(i, 4, QtWidgets.QTableWidgetItem(str(row["expiry_date"])))
            self.table_passport.setItem(i, 5, QtWidgets.QTableWidgetItem(str(row["issue_authority"])))
            custodian_text = "مستلم" if row["custodian"] == "الموظف" else "غير مستلم"
            self.table_passport.setItem(i, 6, QtWidgets.QTableWidgetItem(custodian_text))

            doc_path = str(row["doc_path"]) if row["doc_path"] else ""
            if doc_path:
                btn_open = QtWidgets.QPushButton("فتح")
                icon = QtGui.QIcon()
                icon.addPixmap(QtGui.QPixmap(":/img/icon-open.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
                btn_open.setIcon(icon)
                btn_open.setIconSize(QtCore.QSize(32, 32))
                btn_open.clicked.connect(lambda _, path=doc_path: self.open_doc(path))
                self.table_passport.setCellWidget(i, 7, btn_open)
            else:
                self.table_passport.setItem(i, 7, QtWidgets.QTableWidgetItem("لا يوجد"))

    def on_passport_selected(self):
        row = self.table_passport.currentRow()
        if row < 0:
            self.label_visa.setText("بيانات التأشيرات")
            self.table_visa.setRowCount(0)
            return

        passport_id_item = self.table_passport.item(row, 0)
        passport_number_item = self.table_passport.item(row, 1)

        if not passport_id_item:
            return

        passport_id = int(passport_id_item.text())
        passport_number = passport_number_item.text() if passport_number_item else ""

        # تحديث النص في label_visa
        self.label_visa.setText(f"بيانات تأشيرات الجواز رقم: {passport_number}")

        # تحميل التأشيرات الخاصة بهذا الجواز
        self.load_visas_for_passport(passport_id)

    def add_passport(self):
        dlg = PassportDialog(self.db_conn, self.employee__id, parent=self)
        if dlg.exec_():
            self.load_passports()

    def edit_passport(self):
        row = self.table_passport.currentRow()
        if row < 0:
            return
        passport_id_item = self.table_passport.item(row, 0)
        if not passport_id_item:
            return

        passport_id = int(passport_id_item.text())
        df = self.db_conn.execute_query("SELECT * FROM passports WHERE id=?", [passport_id], fetch=True)
        if df.empty:
            return

        passport_data = df.iloc[0].to_dict()
        dlg = PassportDialog(self.db_conn, self.employee__id, passport_data, parent=self)
        if dlg.exec_():
            self.load_passports()

    def delete_passport(self):
        row = self.table_passport.currentRow()
        if row < 0:
            return
        passport_id_item = self.table_passport.item(row, 0)
        if not passport_id_item:
            return

        passport_id = int(passport_id_item.text())
        reply = QMessageBox.question(
            self,
            "تأكيد الحذف",
            f"هل أنت متأكد أنك تريد حذف الجواز ({passport_id})؟",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.db_conn.delete('passports', [passport_id])
            self.load_passports()
            self.table_visa.setRowCount(0)

    # Visa Tab
    def setup_visa_tab(self):
        cols = ['id','رقم التأشيرة', 'نوع التأشيرة', 'تاريخ الإصدار', 'تاريخ الانتهاء', 'صورة التأشيرة']
        self.table_visa.setColumnCount(len(cols))
        self.table_visa.setHorizontalHeaderLabels(cols)
        self.table_visa.setColumnHidden(0, True)  # إخفاء ID

        header = self.table_visa.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        df_types = self.db_conn.select('visa_types')
        self.visa_types = {row['id']: row['name'] for _, row in df_types.iterrows()}

    def load_visas_for_passport(self, passport_id):
        df = self.db_conn.execute_query(
            """
            SELECT 
                v.id, 
                v.visa_number, 
                t.name AS visa_type, 
                v.issue_date, 
                v.expiry_date, 
                v.doc_path
            FROM visas v
            LEFT JOIN visa_types t ON v.visa_type_id = t.id
            WHERE v.passport_id = ?
            """, 
            [passport_id], 
            fetch=True
        )

        self.table_visa.setRowCount(len(df))

        for i, row in df.iterrows():
            for j, col in enumerate(df.columns[:-1]):
                self.table_visa.setItem(i, j, QtWidgets.QTableWidgetItem(str(row[col])))

            doc_path = str(row["doc_path"]) if row["doc_path"] else ""
            if doc_path:
                btn_open = QtWidgets.QPushButton("فتح")
                icon = QtGui.QIcon()
                icon.addPixmap(QtGui.QPixmap(":/img/icon-open.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
                btn_open.setIcon(icon)
                btn_open.setIconSize(QtCore.QSize(32, 32))
                btn_open.clicked.connect(lambda _, path=doc_path: self.open_doc(path))
                self.table_visa.setCellWidget(i, len(df.columns), btn_open)
            else:
                self.table_visa.setItem(i, len(df.columns), QtWidgets.QTableWidgetItem("لا يوجد"))

    def add_visa(self):
        row = self.table_passport.currentRow()
        if row < 0:
            return
        passport_id_item = self.table_passport.item(row, 0)
        if not passport_id_item:
            return

        passport_id = int(passport_id_item.text())
        dlg = VisaDialog(self.db_conn, passport_id, parent=self)
        if dlg.exec_():
            self.load_visas_for_passport(passport_id)

    def edit_visa(self):
        row = self.table_visa.currentRow()
        if row < 0:
            return
        visa_id_item = self.table_visa.item(row, 0)
        if not visa_id_item:
            return

        visa_id = int(visa_id_item.text())
        df = self.db_conn.execute_query("SELECT * FROM visas WHERE id=?", [visa_id], fetch=True)
        if df.empty:
            return

        visa_data = df.iloc[0].to_dict()
        dlg = VisaDialog(self.db_conn, visa_data['passport_id'], visa_data, parent=self)
        if dlg.exec_():
            self.load_visas_for_passport(visa_data['passport_id'])

    def delete_visa(self):
        row = self.table_visa.currentRow()
        if row < 0:
            return
        visa_id_item = self.table_visa.item(row, 0)
        if not visa_id_item:
            return

        visa_id = int(visa_id_item.text())
        reply = QMessageBox.question(
            self,
            "تأكيد الحذف",
            f"هل أنت متأكد أنك تريد حذف التاشيرة({visa_id})؟",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.db_conn.delete('visas', [visa_id])

            passport_row = self.table_passport.currentRow()
            if passport_row >= 0:
                passport_id = int(self.table_passport.item(passport_row, 0).text())
                self.load_visas_for_passport(passport_id)

    # Exit
    def exit(self):
        reply = QtWidgets.QMessageBox.question(
            self,'الغاء','لن يتم حفظ البيانات الغير محفوظة, هل أنت متأكد من أنك تريد الخروج ؟',
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No
        )
        if reply == QtWidgets.QMessageBox.Yes:
            self.reject()

    def load_documents(self):
        """تحميل المستندات في القائمة"""
        if not os.path.exists(self.docs_folder):
            os.makedirs(self.docs_folder)

        self.documentsTree.clear()
        for file in os.listdir(self.docs_folder):
            self.documentsTree.addItem(file)

    def add_document(self):
        """إضافة مستند عبر نسخ ملف للمجلد"""
        filePath, _ = QtWidgets.QFileDialog.getOpenFileName(self, "اختر ملف")
        if filePath:
            self.copy_file(filePath, self.docs_folder)
            self.load_documents()

    def delete_document(self):
        """حذف المستند المحدد"""
        selected = self.documentsTree.currentItem()
        if selected:
            fileName = selected.text()
            filePath = os.path.join(self.docs_folder, fileName)
            if os.path.exists(filePath):
                os.remove(filePath)
            self.load_documents()

    def open_document(self):
        """فتح المستند المحدد بالبرنامج الافتراضي"""
        selected = self.documentsTree.currentItem()
        if selected:
            filePath = os.path.join(self.docs_folder, selected.text())
            if sys.platform.startswith("win"):  # Windows
                os.startfile(filePath)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", filePath])
            else:  # Linux
                subprocess.run(["xdg-open", filePath])
            
    def open_doc(self, path):
        if os.path.exists(path):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
        else:
            QtWidgets.QMessageBox.warning(self, "خطأ", f"الملف غير موجود:\n{path}")


class CustodyViewer:
    def __init__(self, root):
        self.root = root
        self.db_conn = root.db_conn

        self.table_employee_custody = root.table_employee_custody
        self.table_company_custody = root.table_company_custody
        self.dateEdit_from = root.dateEdit_from
        self.dateEdit_to = root.dateEdit_to
        self.btn_filter_employee_custody = root.btn_filter_employee_custody

        today = QtCore.QDate.currentDate()
        last_month = today.addMonths(-1)
        self.dateEdit_from.setDate(last_month)
        self.dateEdit_to.setDate(today)

        self.setup_employee_table()
        self.setup_company_table()
        self.load_custody_data()

    def setup_employee_table(self):
        self.table_employee_custody.setColumnCount(6)
        self.table_employee_custody.setHorizontalHeaderLabels([
            "ID", "emp_ID", "اسم الموظف", "رقم الجواز", "نوع الجواز", "تاريخ التسليم"
        ])
        header = self.table_employee_custody.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table_employee_custody.setColumnHidden(0, True)
        self.table_employee_custody.setColumnHidden(1, True)
        self.table_employee_custody.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)

    def setup_company_table(self):
        self.table_company_custody.setColumnCount(5)
        self.table_company_custody.setHorizontalHeaderLabels([
            "ID", "emp_ID", "اسم الموظف", "رقم الجواز", "نوع الجواز" #, "تاريخ الإستلام"
        ])
        header = self.table_company_custody.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table_company_custody.setColumnHidden(0, True)
        self.table_company_custody.setColumnHidden(1, True)
        self.table_company_custody.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)

    def load_custody_data(self):
        start_date = self.dateEdit_from.date().toString("yyyy-MM-dd") + " 00:00:00"
        end_date = self.dateEdit_to.date().toString("yyyy-MM-dd") + " 23:59:59"
        query = """
            SELECT p.id, p.employee_id, e.name_ar AS employee, 
                   p.passport_number, pt.name AS passport_type,
                   p.custodian, p.received_at, p.delivered_by, p.received_by
            FROM passports p
            LEFT JOIN employees e ON p.employee_id = e.id
            LEFT JOIN passport_types pt ON p.passport_type_id = pt.id
            ORDER BY p.received_at DESC
        """
        
        df = self.db_conn.execute_query(query, fetch=True)
        df = df.fillna("غير موجود")

        if df is None or df.empty:
            self.table_employee_custody.setRowCount(0)
            self.table_company_custody.setRowCount(0)
            return

        df_employee = df[df['custodian'] == "الموظف"].reset_index(drop=True)
        df_company = df[df['custodian'] == "الشركة"].reset_index(drop=True)

        df_employee["received_at"] = pd.to_datetime(df_employee["received_at"], errors="coerce")
        
        df_employee = df_employee[
            (df_employee["received_at"] >= start_date) & (df_employee["received_at"] <= end_date)
        ].reset_index(drop=True)

        self.populate_table(self.table_employee_custody, df_employee, "received_at")
        self.populate_table(self.table_company_custody, df_company)

    @staticmethod
    def populate_table(table, df, date_field=None):
        table.setRowCount(len(df))
        for i, row in df.iterrows():
            table.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row['id'])))
            table.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row['employee_id'])))
            table.setItem(i, 2, QtWidgets.QTableWidgetItem(str(row['employee'])))
            table.setItem(i, 3, QtWidgets.QTableWidgetItem(str(row['passport_number'])))
            table.setItem(i, 4, QtWidgets.QTableWidgetItem(str(row['passport_type'])))
            if date_field:
                formatted_date = row[date_field].strftime("%Y-%m-%d") if pd.notna(row[date_field]) else ""
                table.setItem(i, 5, QtWidgets.QTableWidgetItem(formatted_date))


class EmployeeDataHandler:
    def __init__(self, db_conn):
        self.db_conn = db_conn

    def _get_or_create(self, table, name):
        """Ensure foreign key exists and return its ID."""
        if pd.isna(name):
            return None
        name = str(name).strip()
        # Check if exists
        res = self.db_conn.execute_query(
            f"SELECT id FROM {table} WHERE name = ?", (name,), fetch=True
        )
        if not res.empty:
            return int(res.iloc[0]['id'])
        # Insert new row
        try:
            new_id = int(self.db_conn.execute_query(
                f"INSERT INTO {table} (name) VALUES (?)", (name,), return_id=True
            ))
            return new_id
        except Exception as e:
            logging.error("Failed to insert into %s: %s", table, e)
            return None

    def export_selected_data(self, selected_ids):
        """Returns a DataFrame of selected employees."""
        if not selected_ids:
            return None

        placeholders = ",".join(["?"] * len(selected_ids))
        query = f"""
        SELECT 
            e.general_number AS الرقم_العام, e.name_ar AS الاسم_بالعربي,
            e.name_en AS الاسم_بالانجليزي, j.name AS المسمى_الوظيفي,
            d.name AS القسم, e.iban_number AS رقم_الايبان, e.phone AS رقم_الهاتف, e.birth_date AS تاريخ_الميلاد,
            e.national_id AS رقم_بطاقة_الهوية, e.id_issue_date AS تاريخ_بداية_الهوية,
            e.id_expiry_date AS تاريخ_نهاية_الهوية,
            p.passport_number AS رقم_الجواز, p.issue_date AS تاريخ_بداية_الجواز,
            p.expiry_date AS تاريخ_نهاية_الجواز, pt.name AS نوع_الجواز,
            v.visa_number AS رقم_التأشيرة, vt.name AS نوع_التأشيرة,
            v.issue_date AS تاريخ_بداية_التأشيرة, v.expiry_date AS تاريخ_نهاية_التأشيرة
        FROM employees e
        LEFT JOIN department_types d ON e.department_id = d.id
        LEFT JOIN job_titles j ON e.job_title_id = j.id
        LEFT JOIN passports p ON e.id = p.employee_id
        LEFT JOIN passport_types pt ON p.passport_type_id = pt.id
        LEFT JOIN visas v ON p.id = v.passport_id
        LEFT JOIN visa_types vt ON v.visa_type_id = vt.id
        WHERE e.id IN ({placeholders})
        """
        return self.db_conn.execute_query(query, data=selected_ids, fetch=True)

    def import_data(self, file_path):
        """Imports data from Excel and returns a dict with status info."""
        df = pd.read_excel(file_path)
        obj_cols = df.select_dtypes(include="object").columns
        df[obj_cols] = df[obj_cols].apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))
        df = df.dropna(how="all")

        not_created, errors = [], []

        for idx, row in df.iterrows():
            try:
                dept_id = self._get_or_create("department_types", row.get("القسم"))
                job_id = self._get_or_create("job_titles", row.get("المسمى_الوظيفي"))
                passport_type_id = self._get_or_create("passport_types", row.get("نوع_الجواز"))
                visa_type_id = self._get_or_create("visa_types", row.get("نوع_التأشيرة"))

                if dept_id is None or job_id is None:
                    not_created.append({"الرقم_العام": row.get("الرقم_العام"), "الاسم_بالعربي": row.get("الاسم_بالعربي")})
                    continue

                general_number = int(row.get("الرقم_العام"))
                emp = self.db_conn.execute_query(
                    "SELECT id FROM employees WHERE general_number = ?",
                    (general_number,), fetch=True
                )

                if not emp.empty:
                    not_created.append({"الرقم_العام": general_number, "الاسم_بالعربي": row.get("الاسم_بالعربي")})
                    employee_id = emp.iloc[0]["id"]
                else:
                    employee_id = self.db_conn.execute_query("""
                        INSERT INTO employees (general_number, name_ar, name_en, birth_date,
                                               national_id, id_issue_date, id_expiry_date,
                                               department_id, job_title_id, phone, iban_number, role)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        general_number, row.get("الاسم_بالعربي"), row.get("الاسم_بالانجليزي"),
                        row.get("تاريخ_الميلاد"), row.get("رقم_بطاقة_الهوية"),
                        row.get("تاريخ_بداية_الهوية"), row.get("تاريخ_نهاية_الهوية"),
                        dept_id, job_id, row.get("رقم_الهاتف"), row.get("رقم_الايبان", ""), None
                    ), return_id=True)

                passport_id = None
                if pd.notna(row.get("رقم_الجواز")):
                    passport_number = str(row.get("رقم_الجواز")).split(".")[0]  # handle floats
                    r = self.db_conn.execute_query(
                        "SELECT id FROM passports WHERE passport_number = ?",
                        (passport_number,), fetch=True
                    )
                    if not r.empty:
                        passport_id = int(r.iloc[0]["id"])
                    elif passport_type_id:
                        passport_id = self.db_conn.execute_query("""
                            INSERT INTO passports (employee_id, passport_number, passport_type_id,
                                                   issue_date, expiry_date, custodian)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (
                            int(employee_id), passport_number, int(passport_type_id),
                            row.get("تاريخ_بداية_الجواز"), row.get("تاريخ_نهاية_الجواز"), "الشركة"
                        ), return_id=True)
                        # print(f"[INFO] Created passport {passport_number} with ID {passport_id}")
                    # else:
                        # print(f"[WARN] Skipping passport for employee {general_number}: missing passport type")

                if pd.notna(row.get("رقم_التأشيرة")) and passport_id and visa_type_id:
                    visa_number = str(row.get("رقم_التأشيرة")).split(".")[0]
                    r = self.db_conn.execute_query(
                        "SELECT id FROM visas WHERE visa_number = ?",
                        (visa_number,), fetch=True
                    )
                    if r.empty:
                        self.db_conn.execute_query("""
                            INSERT INTO visas (passport_id, visa_number, visa_type_id, issue_date, expiry_date) VALUES (?, ?, ?, ?, ?)
                        """, (
                            passport_id, visa_number, visa_type_id,
                            row.get("تاريخ_بداية_التأشيرة"), row.get("تاريخ_نهاية_التأشيرة")
                        ))
                        # print(f"[INFO] Created visa {visa_number} for passport ID {passport_id}")
                    # else:
                        # print(f"[INFO] Visa {visa_number} already exists")

            except Exception as row_err:
                error_msg = f"صف {idx+2}: {row_err}"
                errors.append(error_msg)

        return {"not_created": not_created, "errors": errors}


class MainWindow(Ui_MainWindow, QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        # بيانات التطبيق
        self.app_title = 'الخطط والمهام - تصميم / عبدالحكيم الشمري'
        self.app_icon = QtGui.QIcon(":/img/logo.png")
        self.db_conn = DB_conn()
        self.setMinimumSize(1786, 924)
        self.showMaximized()

        self.employees_data = pd.DataFrame()
        self.employees_data_filtered = pd.DataFrame()
        self.current_page = 1
        self.rows_per_page = 30
        self.is_filtered = False  
        self.selected_ids = []  
        self.employee_data_handler = EmployeeDataHandler(self.db_conn)
        self.setup()

    def setup(self):
        """إعداد الواجهة والأزرار"""
        self.setupUi(self)
        self.setWindowTitle(self.app_title)
        self.setWindowIcon(self.app_icon)

        # التنقل بين الصفحات
        self.NavBtn1.clicked.connect(lambda: self.Change_Page(0))
        self.NavBtn2.clicked.connect(lambda: self.Change_Page(1))
        self.NavBtn3.clicked.connect(lambda: self.Change_Page(2))
        self.NavBtn4.clicked.connect(lambda: self.Change_Page(3))
        self.SettingButton.clicked.connect(lambda: self.Change_Page(4))

        # أزرار النظام
        self.LoginButton.clicked.connect(self.login)
        self.LogoutButton.clicked.connect(self.logout)
        self.btnSave.clicked.connect(self.change_credentials)
        self.ExitButton.clicked.connect(self.exit)

        # التاريخ والوقت
        self.update_datetime()
        timer = QtCore.QTimer(self)
        timer.timeout.connect(self.update_datetime)
        timer.start(30000)

        # الموظفين
        self.AddEmployeeButton.clicked.connect(self.open_add_employee_dialog)

        headers = [
            "الرقم", "الاسم عربي/النوع", "الاسم انجليزي/تاريخ الإصدار",
            "رقم الهوية/تاريخ الانتهاء", "الهاتف/جهة الإصدار",
            "القسم", "المسمى_الوظيفي", "الدور", ""
        ]
        self.EmployeesList.setColumnCount(len(headers))
        self.EmployeesList.setHeaderLabels(headers)
        header = self.EmployeesList.header()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.EmployeesList.setColumnWidth(0, 320)
        self.EmployeesList.itemClicked.connect(self.toggle_selection)

        # التحكم في الصفحات
        self.PrevPageButton.clicked.connect(self.prev_page)
        self.NextPageButton.clicked.connect(self.next_page)
        self.RefrechEmpolyeeButton.clicked.connect(self.refresh_emloyees)

        # البحث
        self.SearchButton.clicked.connect(self.search_employees)
        self.SearchEntry.returnPressed.connect(self.search_employees)
        self.ExportExcelButton.clicked.connect(self.export_selected_to_excel)
        self.ImportButton.clicked.connect(self.import_from_file)

        # تحميل البيانات الأولية
        self.load_combobox_data()
        self.refresh_emloyees()

        # جداول إضافية
        self.setup_custody_passports_table()
        self.setup_notifications()

        # العهدة
        self.custody_viewer = CustodyViewer(self)
        self.btn_filter_employee_custody.clicked.connect(self.custody_viewer.load_custody_data)
        self.btn_refresh_custody.clicked.connect(self.custody_viewer.load_custody_data)

    def open_add_employee_dialog(self):
        dialog = EditEmployeeDialog(self)
        dialog.exec_()
        self.refresh_emloyees()

    def open_edit_employee_dialog(self, item):
        emp_id = item.data(0, QtCore.Qt.UserRole)
        if not emp_id:
            QtWidgets.QMessageBox.warning(self, "خطأ", "لم يتم العثور على رقم الموظف!")
            return
        
        dialog = EditEmployeeDialog(self, employee__id=emp_id)
        dialog.exec_()
        self.refresh_emloyees()

    @staticmethod
    def set_item_bg(item, color_hex):
        color = QtGui.QBrush(QtGui.QColor(color_hex))
        for col in range(item.columnCount()):
            item.setBackground(col, color)

    @staticmethod
    def add_action_buttons(tree_widget, item, update_callback, delete_callback):
        # إنشاء حاوية أفقية
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setObjectName("horizontalLayout")
        layout.setSpacing(3)

        btn_update = QtWidgets.QToolButton(widget)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/img/icon-compose.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        btn_update.setIcon(icon)
        btn_update.setIconSize(QtCore.QSize(28, 28))
        btn_update.setObjectName("btn_update")
        btn_update.clicked.connect(lambda: update_callback(item))
        
        btn_delete = QtWidgets.QToolButton(widget)
        btn_delete.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/img/icon-delete.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        btn_delete.setIcon(icon1)
        btn_delete.setIconSize(QtCore.QSize(28, 28))
        btn_delete.setObjectName("btn_delete")
        btn_delete.clicked.connect(lambda: delete_callback(item))
        
        layout.addWidget(btn_update)
        layout.addWidget(btn_delete)

        layout.addStretch()

        # وضع الحاوية في آخر عمود
        tree_widget.setItemWidget(item, item.columnCount()-1, widget)

    def delete_employee(self, item):
        emp_id = item.data(0, QtCore.Qt.UserRole)
        if not emp_id:
            QtWidgets.QMessageBox.warning(self, "خطأ", "لم يتم العثور على رقم الموظف!")
            return

        reply = QtWidgets.QMessageBox.question(
            self,
            "تأكيد الحذف",
            "هل أنت متأكد أنك تريد حذف هذا الموظف؟\nسيتم حذف كل بياناته المرتبطة أيضًا.",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if reply == QtWidgets.QMessageBox.Yes:
            # حذف التأشيرات والجوازات أولاً إذا كانت مربوطة
            self.db_conn.execute_query("DELETE FROM visas WHERE passport_id IN (SELECT id FROM passports WHERE employee_id=?)", [emp_id])
            self.db_conn.execute_query("DELETE FROM passports WHERE employee_id=?", [emp_id])
            self.db_conn.execute_query("DELETE FROM employees WHERE id=?", [emp_id])
            docs_folder = "documents/" + str(item.text(1))
            if os.path.exists(docs_folder):
                def remove_readonly(func, path, _excinfo):
                    # Make file/folder writable and retry
                    os.chmod(path, 0o700)
                    func(path)
                try:
                    shutil.rmtree(docs_folder, onerror=remove_readonly)
                    # print(f"تم حذف المجلد: {docs_folder}")
                except Exception as e:
                    print(f"خطأ أثناء حذف المجلد: {e}")
            else:
                print(f"المجلد غير موجود: {docs_folder}")
            
            QtWidgets.QMessageBox.information(self, "نجاح", "تم حذف الموظف بنجاح")
            self.refresh_emloyees()

    def make_combobox_multiselect(self, combo: QtWidgets.QComboBox):
        combo_model = QtGui.QStandardItemModel(combo)
        combo.setCurrentText("")
        combo.setModel(combo_model)
        def on_item_pressed(index):
            item = combo_model.itemFromIndex(index)
            item.setCheckState(QtCore.Qt.Unchecked if item.checkState() == QtCore.Qt.Checked else QtCore.Qt.Checked)
        combo.view().pressed.connect(on_item_pressed)
    
        # دالة لإضافة العناصر
        def add_items(items):
            combo_model.clear()
            item = QtGui.QStandardItem("")
            item.setData("", QtCore.Qt.UserRole)
            combo_model.appendRow(item)
            for text, data in items:
                item = QtGui.QStandardItem(text)
                item.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                item.setData(QtCore.Qt.Unchecked, QtCore.Qt.CheckStateRole)
                item.setData(data, QtCore.Qt.UserRole)  # حفظ الـ id
                combo_model.appendRow(item)
        combo.add_items = add_items  # ربط دالة الإضافة بالـ combo
    
        # دالة للحصول على العناصر المحددة
        def get_selected_ids():
            return [combo_model.item(i).data(QtCore.Qt.UserRole) for i in range(combo_model.rowCount()) if combo_model.item(i).checkState() == QtCore.Qt.Checked]
        combo.get_selected_ids = get_selected_ids
    
        def clear_all():
            for i in range(1, self.VisaTypeEntry.model().rowCount()):
                item = self.VisaTypeEntry.model().item(i)
                item.setCheckState(QtCore.Qt.Unchecked)
        combo.clear_all = clear_all

    def load_combobox_data(self):
        """تحميل بيانات الأقسام والوظائف والتأشيرات"""
        def fill_combo(combo, df, placeholder="— الكل —"):
            combo.clear()
            combo.addItem(placeholder, None)
            if isinstance(df, pd.DataFrame) and not df.empty:
                for _, row in df.iterrows():
                    combo.addItem(row['name'], row['id'])

        df_dept = self.db_conn.select('department_types')
        df_job = self.db_conn.select('job_titles')

        fill_combo(self.DepartmentEntry, df_dept, "")
        fill_combo(self.JopEntry, df_job, "")
        
        self.make_combobox_multiselect(self.VisaTypeEntry)
        df_visa = self.db_conn.select('visa_types')

        if isinstance(df_visa, pd.DataFrame) and not df_visa.empty:
            items = [(row['name'], row['id']) for _, row in df_visa.iterrows() ]
            self.VisaTypeEntry.add_items(items)

    def reset_search_fields(self):
        self.SearchEntry.clear()
        self.DepartmentEntry.setCurrentIndex(0)
        self.JopEntry.setCurrentIndex(0)
        self.VisaTypeEntry.clear_all()
        self.rolecheckBox1.setChecked(False)
        self.rolecheckBox2.setChecked(False)

    def get_page_data(self):
        """إرجاع بيانات الصفحة الحالية"""
        start_idx = (self.current_page - 1) * self.rows_per_page
        end_idx = start_idx + self.rows_per_page
        return self.employees_data.iloc[start_idx:end_idx]

    def next_page(self):
        data = self.employees_data_filtered if self.is_filtered else self.employees_data
        total_pages = max(1, -(-len(data) // self.rows_per_page))
        if self.current_page < total_pages:
            self.current_page += 1
            self.render_employees(filtered=self.is_filtered)
    
    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.render_employees(filtered=self.is_filtered)

    def update_page_label(self):
        """تحديث نص عدد الصفحات والموظفين"""
        data = self.employees_data_filtered if not self.employees_data_filtered.empty else self.employees_data
        total_employees = len(data)
        total_pages = max(1, -(-total_employees // self.rows_per_page))

        showing_start = (self.current_page - 1) * self.rows_per_page + 1
        showing_end = min(self.current_page * self.rows_per_page, total_employees)

        self.PageLabel.setText(
            f"الموظفون: {showing_start}-{showing_end} من {total_employees} | صفحة {self.current_page} / {total_pages}"
        )
    

    def load_all_employees(self):
        """تحميل كل بيانات الموظفين مرة واحدة فقط"""
        self.employees_data = self.db_conn.execute_query("""
            SELECT 
                e.general_number,
                e.name_ar,
                e.name_en,
                e.national_id,
                e.phone,
                d.name AS department,
                j.name AS job_title,
                e.role,
                e.id
            FROM employees e
            LEFT JOIN department_types d ON e.department_id = d.id
            LEFT JOIN job_titles j ON e.job_title_id = j.id
        """, fetch=True)

        if not isinstance(self.employees_data, pd.DataFrame):
            QtWidgets.QMessageBox.critical(self, "خطأ", f"فشل جلب بيانات الموظفين:\n{self.employees_data}")
            self.employees_data = pd.DataFrame()

    @staticmethod
    def filter_by_text(data, search_text):
        if not search_text:
            return data
        return data[
            data.apply(
                lambda row: any(
                    search_text in str(row[col]).lower()
                    for col in ["general_number", "name_ar", "name_en", "national_id", "phone"]
                ),
                axis=1
            )
        ]

    @staticmethod
    def filter_by_department(data, department_id):
        if not department_id:
            return data
        return data[data["department_id"] == department_id]

    @staticmethod
    def filter_by_job_title(data, job_title_id):
        if not job_title_id:
            return data
        return data[data["job_title_id"] == job_title_id]
    
    @staticmethod
    def filter_by_role(data, role_fard, role_masoul):
        if role_fard and not role_masoul:
            return data[data["role"] == "فرد"]
        elif role_masoul and not role_fard:
            return data[data["role"] == "مسؤول"]
        return data

    def filter_by_visa(self, data, selected_visa_ids):
        if not selected_visa_ids:
            return data

        valid_ids = set()
        for emp_id in data["id"]:
            visas = self.db_conn.execute_query(
                """
                SELECT v.visa_type_id
                FROM visas v
                JOIN passports p ON v.passport_id = p.id
                WHERE p.employee_id = ?
                """,
                data=[emp_id],
                fetch=True
            )
            if isinstance(visas, pd.DataFrame) and not visas.empty and any(v_id in selected_visa_ids for v_id in visas["visa_type_id"]):
                valid_ids.add(emp_id)

        return data[data["id"].isin(valid_ids)]

    def search_employees(self):
        self.SearchButton.setEnabled(False)
        self.NextPageButton.setEnabled(False)
        self.PrevPageButton.setEnabled(False)

        if self.employees_data is None or self.employees_data.empty:
            return

        search_text = self.SearchEntry.text().strip().lower()
        department_id = self.DepartmentEntry.currentData()
        job_title_id = self.JopEntry.currentData()
        role_fard = self.rolecheckBox1.isChecked()
        role_masoul = self.rolecheckBox2.isChecked()
        selected_visa_ids = self.VisaTypeEntry.get_selected_ids()

        filtered_data = self.employees_data.copy()

        filtered_data = self.filter_by_text(filtered_data, search_text)
        filtered_data = self.filter_by_department(filtered_data, department_id)
        filtered_data = self.filter_by_job_title(filtered_data, job_title_id)
        filtered_data = self.filter_by_role(filtered_data, role_fard, role_masoul)
        filtered_data = self.filter_by_visa(filtered_data, selected_visa_ids)

        self.employees_data_filtered = filtered_data.reset_index(drop=True)
        self.current_page = 1
        self.is_filtered = True
        self.render_employees(filtered=True)

        self.SearchButton.setEnabled(True)
        self.NextPageButton.setEnabled(True)
        self.PrevPageButton.setEnabled(True)

    def render_employees(self, filtered=False):
        """عرض بيانات الموظفين (كاملة أو مفلترة)"""
        self.EmployeesList.clear()
        data = self.employees_data_filtered if filtered else self.employees_data

        if data is None or data.empty:
            return

        start_idx = (self.current_page - 1) * self.rows_per_page
        end_idx = start_idx + self.rows_per_page
        page_data = data.iloc[start_idx:end_idx]

        for _, emp in page_data.iterrows():
            emp_item = QtWidgets.QTreeWidgetItem([
                str(emp.get('general_number', "")),
                emp.get('name_ar', ""),
                emp.get('name_en', ""),
                emp.get('national_id', ""),
                emp.get('phone', ""),
                emp.get('department', ""),
                emp.get('job_title', ""),
                emp.get('role', ""),
                ""
            ])
            emp_item.setData(0, QtCore.Qt.UserRole, emp.get('id'))
            self.set_item_bg(emp_item, "#e8f4ff")
            emp_item.setFlags(emp_item.flags() | QtCore.Qt.ItemIsUserCheckable)
            emp_item.setCheckState(0, QtCore.Qt.Unchecked)
            self.EmployeesList.addTopLevelItem(emp_item)

            self.add_action_buttons(
                self.EmployeesList,
                emp_item,
                update_callback=self.open_edit_employee_dialog,
                delete_callback=self.delete_employee
            )

            # جلب الجوازات والتأشيرات فقط في حالة عرض البيانات الكاملة
            passports = self.db_conn.execute_query("""
                SELECT 
                    p.passport_number,
                    t.name AS passport_type,
                    p.issue_date,
                    p.expiry_date,
                    p.issue_authority,
                    p.id
                FROM passports p
                LEFT JOIN passport_types t ON p.passport_type_id = t.id
                WHERE p.employee_id = ?
            """, data=[emp['id']], fetch=True)

            if not isinstance(passports, pd.DataFrame) or passports.empty:
                emp_item.addChild(QtWidgets.QTreeWidgetItem(["❌ لا يوجد جواز"]))
                continue

            for _, pp in passports.iterrows():
                pp_item = QtWidgets.QTreeWidgetItem([
                    pp.get('passport_number', ""),
                    str(pp.get('passport_type', "")),
                    pp.get('issue_date', ""),
                    pp.get('expiry_date', ""),
                    pp.get('issue_authority', ""),
                    "", "", "", ""
                ])
                self.set_item_bg(pp_item, "#e0ffd6")
                emp_item.addChild(pp_item)

                visas = self.db_conn.execute_query("""
                    SELECT 
                        v.visa_number,
                        t.name AS visa_type,
                        v.issue_date,
                        v.expiry_date,
                        v.id
                    FROM visas v
                    LEFT JOIN visa_types t ON v.visa_type_id = t.id
                    WHERE v.passport_id = ?
                """, data=[pp['id']], fetch=True)

                if not isinstance(visas, pd.DataFrame) or visas.empty:
                    pp_item.addChild(QtWidgets.QTreeWidgetItem(["❌ لا يوجد تأشيرات"]))
                    continue

                for _, vs in visas.iterrows():
                    vs_item = QtWidgets.QTreeWidgetItem([
                        vs.get('visa_number', ""),
                        str(vs.get('visa_type', "")),
                        vs.get('issue_date', ""),
                        vs.get('expiry_date', ""),
                        "", "", "", "", ""
                    ])
                    self.set_item_bg(vs_item, "#f9fbe7")
                    pp_item.addChild(vs_item)

        self.update_page_label()

    def refresh_emloyees(self):
        """تحديث بيانات الموظفين"""
        self.current_page = 1
        self.is_filtered = False
        self.selected_ids = []
        self.reset_search_fields()
        self.load_all_employees()
        self.render_employees(filtered=False)

    def toggle_selection(self, item, column):
        """Store selected employee IDs across pages (only top-level rows)."""    
        if item.parent() is None:
            current_state = item.checkState(0)
            new_state = QtCore.Qt.Unchecked if current_state == QtCore.Qt.Checked else QtCore.Qt.Checked
            item.setCheckState(0, new_state)
        else:
            return


        emp_id = item.data(0, QtCore.Qt.UserRole) 
    
        if item.checkState(0) == QtCore.Qt.Checked:
            if emp_id not in self.selected_ids:
                self.selected_ids.append(emp_id)
        else:
            if emp_id in self.selected_ids:
                self.selected_ids.remove(emp_id)
    
    def export_selected_to_excel(self, filename="selected_employees"):
        if not self.selected_ids:
            QtWidgets.QMessageBox.warning(self, "تنبيه", "لم يتم تحديد أي موظف للتصدير.")
            return
    
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "حفظ ملف Excel", "selected_employees", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return
        if not file_path.endswith(".xlsx"):
            file_path += ".xlsx"
    
        df = self.employee_data_handler.export_selected_data(self.selected_ids)
        if df is None or df.empty:
            QtWidgets.QMessageBox.warning(self, "تنبيه", "لم يتم العثور على بيانات للتصدير")
            return
    
        try:
            df.to_excel(file_path, index=False)
            QtWidgets.QMessageBox.information(self, "نجاح", f"تم تصدير البيانات إلى:\n{file_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء التصدير:\n{str(e)}")
    
    def import_from_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "استيراد بيانات الموظفين", "", "ملفات Excel (*.xlsx *.xls)"
        )
        if not file_path:
            return
    
        result = self.employee_data_handler.import_data(file_path)
    
        if result["errors"]:
            msg = "❌ صفوف تم تخطيها بسبب أخطاء:\n" + "\n".join(result["errors"])
            QtWidgets.QMessageBox.warning(self, "تنبيه", msg)
        else:
            QtWidgets.QMessageBox.information(self, "نجاح", "تم استيراد البيانات بنجاح ✅")
            self.refresh_emloyees()

    def toggle_select_all(self, state):
        if state == QtCore.Qt.Checked:
            self.table_passport_custody.selectAll()
        else:
            self.table_passport_custody.clearSelection()
        self.table_passport_custody.setFocus()

    def setup_custody_passports_table(self):
        self.table_passport_custody.setColumnCount(9)
        self.table_passport_custody.setHorizontalHeaderLabels([
            "ID", "emp_ID", "اسم الموظف", "رقم الجواز", "نوع الجواز", "الحالة", "المسلّم", "المستلم", "تاريخ الاستلام"
        ])
        header = self.table_passport_custody.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
    
        # إخفاء العمود ID و emp_ID
        self.table_passport_custody.setColumnHidden(0, True)
        self.table_passport_custody.setColumnHidden(1, True)
        self.table_passport_custody.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)
        self.table_passport_custody.setSelectionMode(QtWidgets.QTableWidget.MultiSelection)
        self.table_passport_custody.itemDoubleClicked.connect(self.open_employee_from_passport)
    
        # أزرار العهدة
        self.btn_deliver_custody.clicked.connect(lambda: self.update_custody_status("الموظف"))
        self.btn_receive_custody.clicked.connect(lambda: self.update_custody_status("الشركة"))
        self.refrech_passport_custody.clicked.connect(self.refresh_custody)
    
        # زر البحث
        self.searchButton.clicked.connect(self.search_custody_passports)
        self.selectAllCheckBox.stateChanged.connect(self.toggle_select_all)
    
        self.load_custody_passports()
    
    def load_custody_passports(self, where_clause="", params=None):
        query = """
            SELECT p.id, p.employee_id, e.name_ar AS employee, p.passport_number, 
                   pt.name AS passport_type, p.custodian,
                   p.delivered_by, p.received_by, p.received_at
            FROM passports p
            LEFT JOIN employees e ON p.employee_id = e.id
            LEFT JOIN passport_types pt ON p.passport_type_id = pt.id
        """
        if where_clause:
            query += " WHERE " + where_clause
    
        df = self.db_conn.execute_query(query, params or [], fetch=True)
    
        if not hasattr(df, "iterrows"):
            QtWidgets.QMessageBox.warning(self, "خطأ", f"فشل تحميل الجوازات: {df}")
            return
    
        self.table_passport_custody.clearSelection()
        self.table_passport_custody.setRowCount(len(df))
        for i, row in df.iterrows():
            color = QtGui.QColor("#e3f2e3") if row['custodian'] == "الموظف" else QtGui.QColor("#f7f5e3")
            for j, col in enumerate(df.columns):
                if col == 'custodian':
                    text = "مستلم" if row['custodian'] == "الموظف" else "غير مستلم"
                elif col == 'received_at':
                    text = row['received_at'][:10] if pd.notna(row['received_at']) else ""
                else:
                    text = str(row[col]) if row[col] else ""
                item = QtWidgets.QTableWidgetItem(text)
                item.setBackground(color)
                self.table_passport_custody.setItem(i, j, item)
    
    def search_custody_passports(self):
        conditions, params = [], []
    
        emp_name = self.employeeNameLineEdit.text().strip()
        delivered_by = self.deliveredByLineEdit.text().strip()
        received_by = self.receivedByLineEdit.text().strip()
        is_received = self.receivedCheckBox.isChecked()
        is_not_received = self.notReceivedCheckBox.isChecked()
    
        if emp_name:
            conditions.append("e.name_ar LIKE ?")
            params.append(f"%{emp_name}%")
        if delivered_by:
            conditions.append("p.delivered_by LIKE ?")
            params.append(f"%{delivered_by}%")
        if received_by:
            conditions.append("p.received_by LIKE ?")
            params.append(f"%{received_by}%")
        if is_received and not is_not_received:
            conditions.append("p.custodian = ?")
            params.append("الموظف")
        elif is_not_received and not is_received:
            conditions.append("p.custodian = ?")
            params.append("الشركة")
    
        where_clause = " AND ".join(conditions) if conditions else ""
        self.load_custody_passports(where_clause, params)

    def open_employee_from_passport(self, item):
        row = item.row()
        emp_id = self.table_passport_custody.item(row, 1).text()
        emp_name = self.table_passport_custody.item(row, 2).text()
        emp_id_str = self.table_passport_custody.item(row, 1).text()
        emp_name = self.table_passport_custody.item(row, 2).text()

        # تحقق أن emp_id صالح
        if not emp_id_str or not emp_id_str.isdigit():
            QtWidgets.QMessageBox.warning(self, "خطأ", f"لم يتم العثور على الموظف {emp_name}")
            return

        emp_id = int(emp_id_str)

        dialog = EditEmployeeDialog(self, employee__id=emp_id, action='update')
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.load_custody_passports()

    def update_custody_status(self, new_status):
        selected_rows = self.table_passport_custody.selectionModel().selectedRows()
        if not selected_rows:
            QtWidgets.QMessageBox.warning(self, "تحذير", "لم يتم اختيار أي جوازات.")
            return
    
        action_text = "تسليم" if new_status == "الموظف" else "استلام"
        reply = QtWidgets.QMessageBox.question(
            self, self.app_title,
            f"هل أنت متأكد من {action_text} الجوازات المحددة؟",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.Yes
        )
        if reply == QtWidgets.QMessageBox.No:
            return
    
        ids_to_update = []
        already_status = []
        already_numbers = []
    
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for row in selected_rows:
            passport_id = int(self.table_passport_custody.item(row.row(), 0).text())
            employee_id = int(self.table_passport_custody.item(row.row(), 1).text())
            passport_number = self.table_passport_custody.item(row.row(), 3).text()
            status = self.table_passport_custody.item(row.row(), 5).text()
            status = "الموظف" if status == "مستلم" else "الشركة"
                
            if status == new_status:
                already_status.append(passport_id)
                already_numbers.append(passport_number)
            else:
                ids_to_update.append((passport_id, employee_id))
    
        if already_status:
            QtWidgets.QMessageBox.warning(
                self, "تنبيه",
                f"الجوازات التالية في عهدة {new_status} بالفعل:\n{', '.join(already_numbers)}"
            )
    
        if ids_to_update:
            # تحديث جدول passports + الأعمدة الجديدة
            passport_ids = [pid for pid, _ in ids_to_update]
            placeholders = ','.join(['?'] * len(passport_ids))
            self.db_conn.execute_query(
                f"""UPDATE passports SET custodian=?, received_at=CURRENT_TIMESTAMP  WHERE id IN ({placeholders})""",
                [new_status] + passport_ids
            )
    
            # تسجيل handover
            for passport_id, employee_id in ids_to_update:
                self.db_conn.execute_query(
                    """INSERT INTO handover (passport_id, employee_id, action_type, action_at)
                       VALUES (?, ?, ?, ?)""",
                    (passport_id, employee_id, action_text, now)
                )
    
            QtWidgets.QMessageBox.information(self, "نجاح", f"تم {action_text} الجوازات المحددة.")
            self.load_custody_passports()

    def refresh_custody(self):
        self.employeeNameLineEdit.clear()
        self.deliveredByLineEdit.clear()
        self.receivedByLineEdit.clear()
        self.receivedCheckBox.setChecked(False)
        self.notReceivedCheckBox.setChecked(False)
        self.selectAllCheckBox.setChecked(False)
        self.load_custody_passports()


    def setup_notifications(self):
        # ربط الراديو بوتنز بالجوازات
        self.rb_pass_expired.toggled.connect(lambda: self.filter_passports(0))
        self.rb_pass_15.toggled.connect(lambda: self.filter_passports(15))
        self.rb_pass_30.toggled.connect(lambda: self.filter_passports(30))
        self.rb_pass_45.toggled.connect(lambda: self.filter_passports(45))
        self.rb_pass_60.toggled.connect(lambda: self.filter_passports(60))
        self.rb_pass_90.toggled.connect(lambda: self.filter_passports(90))
        self.rb_pass_180.toggled.connect(lambda: self.filter_passports(180))
        self.rb_pass_expired.setChecked(True)

        # ربط الراديو بوتنز بالتأشيرات
        self.rb_visa_expired.toggled.connect(lambda: self.filter_visas(0))
        self.rb_visa_15.toggled.connect(lambda: self.filter_visas(15))
        self.rb_visa_30.toggled.connect(lambda: self.filter_visas(30))
        self.rb_visa_45.toggled.connect(lambda: self.filter_visas(45))
        self.rb_visa_60.toggled.connect(lambda: self.filter_visas(60))
        self.rb_visa_90.toggled.connect(lambda: self.filter_visas(90))
        self.rb_visa_180.toggled.connect(lambda: self.filter_visas(180))
        self.rb_visa_expired.setChecked(True)

        self.table_passports.setColumnCount(3)
        self.table_passports.setHorizontalHeaderLabels(["اسم الموظف", "رقم الجواز", "تاريخ الانتهاء"])
        header = self.table_passports.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table_passports.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table_passports.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)

        self.table_visas.setColumnCount(4)
        self.table_visas.setHorizontalHeaderLabels(["اسم الموظف", "رقم التأشيرة", "رقم الجواز", "تاريخ الانتهاء"])
        header = self.table_visas.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table_visas.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table_visas.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.filter_passports(0)
        self.filter_visas(0)

    def filter_passports(self, days):
        query = """
        SELECT p.id, e.name_ar AS employee_name, p.passport_number, p.expiry_date
        FROM passports p
        LEFT JOIN employees e ON e.id = p.employee_id
        """
        df = self.db_conn.execute_query(query, fetch=True)
        self.show_passports(self._filter_by_days(df, days, "expiry_date"))

    def filter_visas(self, days):
        query = """
        SELECT v.id, e.name_ar AS employee_name, p.passport_number, v.visa_number, v.expiry_date
        FROM visas v
        LEFT JOIN passports p ON p.id = v.passport_id
        LEFT JOIN employees e ON e.id = p.employee_id
        """
        df = self.db_conn.execute_query(query, fetch=True)
        self.show_visas(self._filter_by_days(df, days, "expiry_date"))

    @staticmethod
    def _filter_by_days(df, days, date_col):
        today = datetime.today().date()
        filtered = []
        
        if not isinstance(df, pd.DataFrame):
            return

        for _, row in df.iterrows():
            try:
                exp_date = datetime.strptime(row[date_col], "%Y-%m-%d").date()
            except Exception:
                continue
            diff_days = (exp_date - today).days
            if days == 0:
                if diff_days <= 0:  # منتهية فعليًا
                    filtered.append(row)
            else:
                if 0 < diff_days <= days:
                    filtered.append(row)
        return filtered

    def show_passports(self, rows):
        self.table_passports.setRowCount(0)
        for i, row in enumerate(rows):
            self.table_passports.insertRow(i)
            self.table_passports.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row["employee_name"])))
            self.table_passports.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row["passport_number"])))
            self.table_passports.setItem(i, 2, QtWidgets.QTableWidgetItem(str(row["expiry_date"])))

    def show_visas(self, rows):
        self.table_visas.setRowCount(0)
        for i, row in enumerate(rows):
            self.table_visas.insertRow(i)
            self.table_visas.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row["employee_name"])))
            self.table_visas.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row["passport_number"])))
            self.table_visas.setItem(i, 2, QtWidgets.QTableWidgetItem(str(row["visa_number"])))
            self.table_visas.setItem(i, 3, QtWidgets.QTableWidgetItem(str(row["expiry_date"])))


    def logout(self):
        self.Root.setCurrentIndex(0)
  
    def login(self):
        username = self.UserEntry.text().strip()
        password = self.PassEntry.text()
    
        if not username or not password:
            QtWidgets.QMessageBox.warning(self, "خطأ", "يرجى إدخال اسم المستخدم وكلمة المرور!")
            return
    
        query = "SELECT username, password FROM user WHERE id = 1"
        result = self.db_conn.execute_query(query, fetch=True)
    
        if isinstance(result, pd.DataFrame) and not result.empty:
            db_username = str(result.at[0, 'username']).strip()
            db_password = str(result.at[0, 'password'])
    
            if username == db_username and password == db_password:
                self.Main.setCurrentIndex(0)
                self.Root.setCurrentIndex(1)
            else:
                QtWidgets.QMessageBox.warning(self, "خطأ", "اسم المستخدم أو كلمة المرور غير صحيحة!")
        else:
            QtWidgets.QMessageBox.warning(self, "خطأ", "المستخدم (ID=1) غير موجود!")
    
    def change_credentials(self):
        username = self.lineEditUsername.text().strip()
        password = self.lineEditPassword.text().strip()
    
        if not username or not password:
            QtWidgets.QMessageBox.warning(self, "خطأ", "يرجى إدخال اسم المستخدم وكلمة المرور!")
            return
    
        try:
            check_query = "SELECT * FROM user WHERE id=1"
            result = self.db_conn.execute_query(check_query, fetch=True)
    
            if isinstance(result, pd.DataFrame) and not result.empty:
                update_query = "UPDATE user SET username=?, password=? WHERE id=1"
                self.db_conn.execute_query(update_query, [username, password])
            else:
                insert_query = "INSERT INTO user (username, password, id) VALUES (?, ?, ?)"
                self.db_conn.execute_query(insert_query, [username, password, 1])
    
            QtWidgets.QMessageBox.information(self, "نجاح", "تم تحديث بيانات الدخول بنجاح!")
    
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء التحديث:\n{e}")

    def Change_Page(self, index):
        self.Main.setCurrentIndex(index)

    def update_datetime(self):
        today = datetime.today()
        MyTime = datetime.now().strftime("%H:%M")
        HijriDate = Gregorian(today.year, today.month, today.day).to_hijri()        

        # تعريف أسماء الأيام والأشهر بالعربي
        arabic_days = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
        arabic_hijri_months = ["محرم", "صفر", "ربيع الأول", "ربيع الآخر", "جمادى الأولى", "جمادى الآخرة", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة"]
        arabic_gregorian_months = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
        
        HijriDate = Gregorian(today.year, today.month, today.day).to_hijri()
        
        MyDate = f"{arabic_days[HijriDate.weekday()]} {HijriDate.day} {arabic_hijri_months[HijriDate.month - 1]} {HijriDate.year} هـ / {arabic_gregorian_months[today.month - 1]} {today.year} {today.day} م"
        MyDate = MyDate + '  \t  \t  ' + MyTime 
        self.DateLabel.setText(MyDate)        

    def exit(self):
        reply = QtWidgets.QMessageBox.question(
            self, 
            self.app_title, 
            'هل أنت متأكد من أنك تريد الخروج ؟', 
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No
        )
        if reply == QtWidgets.QMessageBox.Yes:
            self.close()


if __name__ == '__main__':
    try:
        app = QtWidgets.QApplication(sys.argv)
        window = MainWindow()
        window.show()
        app.exec_()
    except Exception as e:
        QtWidgets.QMessageBox.information(
            None, 'Error',
            f'Hi, you can\'t access this app, please contact the developer\n\n{e}'
        )
        raise e

# pyuic5 ui/MainWindow.ui -o MainWindow.py
# pyuic5 ui/EditEmployeePage.ui -o EditEmployeePage.py
# pyrcc5 ui/img/img.qrc -o img_rc.py
# pyinstaller --windowed --icon=ui\img\logo.ico --add-data="ui\img\logo.png;." --name "HRM" main.py
