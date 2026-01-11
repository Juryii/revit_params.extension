# -*- coding: utf-8 -*-

__title__ = "Изменить параметры"
__author__ = "Yura Polyanskii"
__doc__ = """Version = 0.1
Description:
Применяет значения выбранных параметров INK ко всем элементам проекта.
Изменяются только те параметры, у которых включён соответствующий чекбокс.
Если отмечен чекбокс INK_ID_Element, формируется уникальный параметр INK_ID_Element на основе № позиции по генплану, раздела проекта и уникального ID элемента.
"""


__highlight__ = "new"
__min_revit_ver__ = 2019
__max_revit_ver = 2025


# ==================================================
# IMPORTS
# ==================================================
#.NET Imports
import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
# clr.AddReference('IronPython.WPF')

from pyrevit import script, UI
from Autodesk.Revit.DB import *
import wpf
from System import Windows

# ==================================================
# VARIABLES
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

xamlfile = script.get_bundle_file('ui.xaml')


def get_elements_on_pararam(param):
    result = []
    elements = (FilteredElementCollector(doc)
        .WhereElementIsNotElementType())
    
    for el in elements:
        p = el.LookupParameter(param)
        if p is not None:
            result.append(el)
    return result

def generate_ink_id(el):
    """
    Формирует значение параметра INK_ID_Element:
    [№ГП]-[Раздел]_ElementId
    Пустые значения пропускаются.
    """
    parts = []

    # № позиции по ГП
    p_gp = el.LookupParameter("INK_№ поз. по ГП")
    if p_gp and p_gp.AsString():
        parts.append(p_gp.AsString().strip())

    # Раздел проекта РД
    p_section = el.LookupParameter("INK_Раздел проекта РД")
    if p_section and p_section.AsString():
        parts.append(p_section.AsString().strip())

    # уникальный номер элемента (обязателен)
    parts.append(str(el.Id.IntegerValue))

    # соединяем:
    # между ГП и разделом → "-"
    # между разделом и ID → "_"
    if len(parts) == 3:
        return "{}-{}_{}".format(parts[0], parts[1], parts[2])
    elif len(parts) == 2:
        return "{}_{}".format(parts[0], parts[1])
    else:
        return parts[0]

def set_param_value(p, value):
    st = p.StorageType

    if st == StorageType.String:
        p.Set(str(value))

    elif st == StorageType.Integer:
        p.Set(int(value))

    elif st == StorageType.Double:
        p.Set(float(value))

    elif st == StorageType.ElementId:
        # пока не поддерживаем
        pass

class MyWindow(Windows.Window):
    def __init__(self):
        wpf.LoadComponent(self, xamlfile)

    def apply_params(self, sender, args):
        data = self.collect_ui_data()
        applied = 0
        skipped = 0
        errors = 0

        t = Transaction(doc, "Apply INK parameters")
        t.Start()

        for param_name, info in data.items():

            # если чекбокс не включён — пропускаем
            if info["enabled"] is not True:
                continue

            value = info["value"]

            # получаем элементы с таким параметром
            elements = get_elements_on_pararam(param_name)

            for el in elements:
                p = el.LookupParameter(param_name)

                if p is None or p.IsReadOnly:
                    skipped += 1
                    continue

                if not value.strip():
                    skipped += 1
                    continue

                try:
                    set_param_value(p, value)
                    applied += 1
                except:
                    errors += 1


        if self.cb_ID.IsChecked is True:
            elements = get_elements_on_pararam("INK_ID_Element")

            for el in elements:
                p_id = el.LookupParameter("INK_ID_Element")

                if p_id is None or p_id.IsReadOnly:
                    skipped += 1
                    continue

                ink_id = generate_ink_id(el)

                if ink_id is None:
                    skipped += 1
                    continue

                try:
                    p_id.Set(ink_id)
                    applied += 1
                except:
                    errors += 1

        t.Commit()
        
        self.tbStatus.Text = (
            "Готово. Применено: {} | Пропущено: {} | Ошибки: {}"
            ).format(applied, skipped, errors)

    def clear_params(self, sender, args):
        data = self.collect_ui_data()

        t = Transaction(doc, "Clear INK parameters")
        t.Start()

        cleared = 0
        skipped = 0
        errors = 0

        # обычные INK параметры
        for param_name, info in data.items():

            if info["enabled"] is not True:
                skipped += 1
                continue

            elements = get_elements_on_pararam(param_name)

            for el in elements:
                p = el.LookupParameter(param_name)

                if not p or p.IsReadOnly:
                    skipped += 1
                    continue

                try:
                    if p.AsString():
                        p.Set("")
                        cleared += 1
                except:
                    errors += 1

        # INK_ID_Element (отдельно)
        if self.cb_ID.IsChecked:
            elements = get_elements_on_pararam("INK_ID_Element")

            for el in elements:
                p = el.LookupParameter("INK_ID_Element")

                if not p or p.IsReadOnly:
                    continue

                try:
                    if p.AsString():
                        p.Set("")
                        cleared += 1
                except:
                    pass

        t.Commit()

        self.tbStatus.Text = (
            "Готово. Очищено значений: {} | Пропущено: {} | Ошибки: {}"
            ).format(cleared, skipped, errors)


    def collect_ui_data(self):
        data = {
            "INK_Объект": {
                "enabled": self.cb_INK_Object.IsChecked,
                "value": self.tb_INK_Object.Text
            },
            "INK_Подобъект": {
                "enabled": self.cb_INK_SubObject.IsChecked,
                "value": self.tb_INK_SubObject.Text
            },
            "INK_№ поз. по ГП": {
                "enabled": self.cb_INK_GP.IsChecked,
                "value": self.tb_INK_GP.Text
            },
            "INK_Проектная организация": {
                "enabled": self.cb_INK_ProjectOrg.IsChecked,
                "value": self.tb_INK_ProjectOrg.Text
            },
            "INK_Раздел проекта РД": {
                "enabled": self.cb_INK_RD_Section.IsChecked,
                "value": self.tb_INK_RD_Section.Text
            },
            "INK_Шифр комплекта РД": {
                "enabled": self.cb_INK_RD_Code.IsChecked,
                "value": self.tb_INK_RD_Code.Text
            },
            "INK_Изменение": {
                "enabled": self.cb_INK_Revision.IsChecked,
                "value": self.tb_INK_Revision.Text
            },
            "INK_Статус": {
                "enabled": self.cb_INK_Status.IsChecked,
                "value": self.tb_INK_Status.Text
            }
        }
        return data

    def say_hello(self, sender, args):
        data = self.collect_ui_data()

        text = ""
        for k, v in data.items():
            text += "{} | enabled: {} | value: {}\n".format(
                k, v["enabled"], v["value"]
            )
        UI.TaskDialog.Show("UI data", text)



# ==================================================
# MAIN
# ==================================================
MyWindow().ShowDialog()
    
