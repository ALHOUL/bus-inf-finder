import os
import sys
import csv
import openpyxl

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from kivy.core.window import Window

# ضبط خلفية التطبيق
Window.clearcolor = (0.95, 0.95, 0.96, 1)

class BusApp(App):
    def build(self):
        self.title = "قسم وسائل نقل الطلبة - عرض البيانات"
        self.data_list = []  # قائمة لتخزين الصفوف بدلاً من DataFrame
        self.data_row = {}

        # التصميم الرئيسي
        main_layout = BoxLayout(orientation='vertical', padding=15, spacing=10)

        # عنوان الواجهة
        title_label = Label(
            text="قسم وسائل نقل الطلبة",
            font_size='22sp',
            bold=True,
            color=(0.17, 0.24, 0.31, 1),
            size_hint_y=None,
            height=40
        )
        main_layout.add_widget(title_label)

        # أزرار التحميل والتحديث
        btn_box = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=45)
        
        btn_load = Button(
            text="1. تحميل ملف الإكسل / CSV",
            background_color=(0.2, 0.6, 0.86, 1),
            color=(1, 1, 1, 1),
            bold=True
        )
        btn_load.bind(on_release=self.open_file_chooser)
        
        btn_reset = Button(
            text="تحديث ↻",
            background_color=(0.95, 0.61, 0.07, 1),
            color=(1, 1, 1, 1),
            bold=True
        )
        btn_reset.bind(on_release=self.reset_filters)

        btn_box.add_widget(btn_load)
        btn_box.add_widget(btn_reset)
        main_layout.add_widget(btn_box)

        # قسم القوائم المنسدلة للبحث
        search_layout = GridLayout(cols=2, spacing=8, size_hint_y=None, height=150)

        search_layout.add_widget(Label(text=":المدرسة", color=(0.1, 0.1, 0.1, 1), size_hint_x=0.3))
        self.school_spinner = Spinner(text='اختر المدرسة', values=['الجميع'], size_hint_x=0.7)
        self.school_spinner.bind(text=self.on_school_change)
        search_layout.add_widget(self.school_spinner)

        search_layout.add_widget(Label(text=":اسم المالك", color=(0.1, 0.1, 0.1, 1), size_hint_x=0.3))
        self.owner_spinner = Spinner(text='اختر المالك', values=[], size_hint_x=0.7)
        self.owner_spinner.bind(text=self.on_owner_change)
        search_layout.add_widget(self.owner_spinner)

        search_layout.add_widget(Label(text=":رقم اللوحة", color=(0.1, 0.1, 0.1, 1), size_hint_x=0.3))
        self.plate_spinner = Spinner(text='اختر اللوحة', values=[], size_hint_x=0.7)
        self.plate_spinner.bind(text=self.on_plate_change)
        search_layout.add_widget(self.plate_spinner)

        main_layout.add_widget(search_layout)

        # عنوان قسم عرض البيانات
        info_header = Label(
            text="--- معلومات المركبة والعقد المحددة ---",
            font_size='16sp',
            bold=True,
            color=(0.2, 0.4, 0.6, 1),
            size_hint_y=None,
            height=30
        )
        main_layout.add_widget(info_header)

        # قائمة البيانات المطلوبة (16 حقل)
        self.fields_to_show = [
            ("المدرسة:", "المدرسة"),
            ("اسم المالك:", "المالك"),
            ("رقم اللوحة:", "رقم اللوحة"),
            ("رقم هاتف المالك:", "هاتف المالك"),
            ("نوع الوسيلة:", "نوع الوسيلة"),
            ("نوع العقد:", "نوع العقد"),
            ("المبلغ:", "المبلغ"),
            ("السائق:", "السائق"),
            ("الصنع:", "الصنع"),
            ("عدد المقاعد:", "عدد المقاعد"),
            ("رقم العقد:", "رقم العقد"),
            ("ماركة الوسيلة:", "ماركة الوسيلة"),
            ("اللون:", "اللون"),
            ("نوع اللوحة2:", "نوع اللوحة2"),
            ("الرمز:", "الرمز"),
            ("رقم المالك المدني:", "رقم المالك المدني")
        ]

        scroll = ScrollView()
        self.info_grid = GridLayout(cols=2, spacing=10, size_hint_y=None, padding=10)
        self.info_grid.bind(minimum_height=self.info_grid.setter('height'))

        self.info_value_labels = {}

        for label_text, key in self.fields_to_show:
            val_lbl = Label(
                text="-",
                color=(0.16, 0.5, 0.73, 1),
                bold=True,
                size_hint_y=None,
                height=35,
                halign='center'
            )
            title_lbl = Label(
                text=label_text,
                color=(0.2, 0.2, 0.2, 1),
                bold=True,
                size_hint_y=None,
                height=35,
                halign='right'
            )
            
            self.info_grid.add_widget(val_lbl)
            self.info_grid.add_widget(title_lbl)
            self.info_value_labels[key] = val_lbl

        scroll.add_widget(self.info_grid)
        main_layout.add_widget(scroll)

        return main_layout

    def open_file_chooser(self, instance):
        content = BoxLayout(orientation='vertical')
        file_chooser = FileChooserListView(path=os.getcwd(), filters=['*.xlsx', '*.xls', '*.csv'])
        content.add_widget(file_chooser)

        btn_layout = BoxLayout(size_hint_y=None, height=40, spacing=10)
        btn_select = Button(text="اختيار الملف", background_color=(0.18, 0.8, 0.44, 1))
        btn_cancel = Button(text="إلغاء", background_color=(0.9, 0.3, 0.23, 1))
        
        btn_layout.add_widget(btn_select)
        btn_layout.add_widget(btn_cancel)
        content.add_widget(btn_layout)

        popup = Popup(title="اختر ملف الإكسل أو CSV", content=content, size_hint=(0.9, 0.9))

        def load_selected_file(btn):
            if file_chooser.selection:
                path = file_chooser.selection[0]
                try:
                    if path.endswith('.csv'):
                        self.data_list = self.read_csv_file(path)
                    else:
                        self.data_list = self.read_excel_file(path)
                    self.reset_filters(None)
                    popup.dismiss()
                except Exception as e:
                    print("Error loading file:", e)
            else:
                popup.dismiss()

        btn_select.bind(on_release=load_selected_file)
        btn_cancel.bind(on_release=lambda x: popup.dismiss())
        popup.open()

    def read_csv_file(self, path):
        for enc in ['utf-8-sig', 'utf-8', 'cp1256', 'latin-1']:
            try:
                with open(path, mode='r', encoding=enc) as f:
                    reader = csv.DictReader(f)
                    rows = []
                    for row in reader:
                        cleaned = {str(k).strip(): (str(v).strip() if v is not None else "") for k, v in row.items() if k}
                        rows.append(cleaned)
                    return rows
            except Exception:
                continue
        return []

    def read_excel_file(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = wb.active
        rows_gen = sheet.iter_rows(values_only=True)
        try:
            headers = [str(cell).strip() if cell is not None else "" for cell in next(rows_gen)]
        except StopIteration:
            return []
        
        data = []
        for row in rows_gen:
            if not any(row):
                continue
            row_dict = {}
            for h, cell in zip(headers, row):
                if h:
                    row_dict[h] = str(cell).strip() if cell is not None and str(cell) != 'None' else ""
            data.append(row_dict)
        return data

    def reset_filters(self, instance):
        if self.data_list:
            schools = sorted(list({str(r.get('المدرسة', '')).strip() for r in self.data_list if r.get('المدرسة')}))
            self.school_spinner.values = ["الجميع"] + schools
            self.school_spinner.text = "الجميع"

            owners = sorted(list({str(r.get('المالك', '')).strip() for r in self.data_list if r.get('المالك')}))
            self.owner_spinner.values = owners
            self.owner_spinner.text = "اختر المالك"

            plates = sorted(list({str(r.get('رقم اللوحة', '')).strip() for r in self.data_list if r.get('رقم اللوحة')}))
            self.plate_spinner.values = plates
            self.plate_spinner.text = "اختر اللوحة"

            self.data_row = {}
            self.update_info_display()

    def get_filtered_data(self):
        if not self.data_list:
            return []
        school_val = self.school_spinner.text
        if school_val in ["", "الجميع", "اختر المدرسة"]:
            return self.data_list
        return [r for r in self.data_list if str(r.get('المدرسة', '')) == school_val]

    def on_school_change(self, spinner, text):
        if not self.data_list or text == "اختر المدرسة":
            return
        filtered = self.get_filtered_data()
        owners = sorted(list({str(r.get('المالك', '')).strip() for r in filtered if r.get('المالك')}))
        self.owner_spinner.values = owners
        self.owner_spinner.text = "اختر المالك"
        
        plates = sorted(list({str(r.get('رقم اللوحة', '')).strip() for r in filtered if r.get('رقم اللوحة')}))
        self.plate_spinner.values = plates
        self.plate_spinner.text = "اختر اللوحة"
        
        self.data_row = {}
        self.update_info_display()

    def on_owner_change(self, spinner, text):
        if not self.data_list or text in ["اختر المالك", ""]:
            return
        filtered = self.get_filtered_data()
        for r in filtered:
            if str(r.get('المالك', '')) == text:
                self.data_row = r
                self.plate_spinner.text = str(r.get('رقم اللوحة', ''))
                self.school_spinner.text = str(r.get('المدرسة', ''))
                self.update_info_display()
                break

    def on_plate_change(self, spinner, text):
        if not self.data_list or text in ["اختر اللوحة", ""]:
            return
        filtered = self.get_filtered_data()
        for r in filtered:
            if str(r.get('رقم اللوحة', '')) == text:
                self.data_row = r
                self.owner_spinner.text = str(r.get('المالك', ''))
                self.school_spinner.text = str(r.get('المدرسة', ''))
                self.update_info_display()
                break

    def update_info_display(self):
        for key, label_widget in self.info_value_labels.items():
            if self.data_row:
                val = self.data_row.get(key, "-")
                label_widget.text = str(val) if val != "" else "-"
            else:
                label_widget.text = "-"

if __name__ == "__main__":
    BusApp().run()