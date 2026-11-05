"""Разделяет прямоугольные воздуховоды без производителя на 1250 мм и остаток."""
from collections import Counter
from pyrevit import DB, UI, forms, revit, script
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.Exceptions import OperationCanceledException
from duct_split import plan_duct, split_duct, verify_split


class Cancelled(Exception):
    pass


def main():
    doc, uidoc = revit.doc, revit.uidoc
    selected = list(uidoc.Selection.GetElementIds())
    scope = u'Выбранные элементы'
    if selected:
        elements = [doc.GetElement(eid) for eid in selected]
    else:
        scope = forms.CommandSwitchWindow.show(
            [u'Выбрать элементы', u'Текущий вид', u'Весь проект'],
            message=u'Где разделить воздуховоды без производителя?')
        if not scope:
            return
        if scope == u'Выбрать элементы':
            try:
                with forms.WarningBar(title=u'Выберите воздуховоды и нажмите «Готово»'):
                    refs = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element,
                                                       u'Выберите воздуховоды')
                elements = [doc.GetElement(ref.ElementId) for ref in refs]
            except OperationCanceledException:
                return
        else:
            collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id) if scope == u'Текущий вид' else DB.FilteredElementCollector(doc)
            elements = list(collector.OfClass(Duct))
    plans, skipped = [], Counter()
    for element in elements:
        plan, reason = plan_duct(element)
        if plan:
            plans.append(plan)
        else:
            skipped[reason] += 1
    skip_text = u'\n'.join(u'{}: {}'.format(reason, count) for reason, count in sorted(skipped.items()))
    if not plans:
        forms.alert(u'Нет воздуховодов для разделения.\n\n' + skip_text,
                    title=u'Разбить по 1250')
        return
    cuts = sum(len(plan['lengths_mm']) - 1 for plan in plans)
    message = u'{}\nРазделить воздуховодов: {}. Разрезов: {}.'.format(
        scope, len(plans), cuts)
    if skipped:
        message += u'\nПропущено: {}.'.format(sum(skipped.values()))
    if not forms.alert(message, title=u'Разбить по 1250', yes=True, no=True):
        return
    group = DB.TransactionGroup(doc, u'Разбить воздуховоды по 1250 мм')
    transaction = None
    failed_id = None
    group.Start()
    try:
        with forms.ProgressBar(title=u'Разделение: {value}/{max_value}', cancellable=True) as progress:
            for index, plan in enumerate(plans):
                if progress.cancelled:
                    raise Cancelled()
                failed_id = plan['id']
                transaction = DB.Transaction(doc, u'Разделить воздуховод')
                transaction.Start()
                state = split_duct(doc, plan)
                if transaction.Commit() != DB.TransactionStatus.Committed:
                    raise ValueError(u'Revit отменил разделение.')
                verify_split(doc, state)
                progress.update_progress(index + 1, len(plans))
            if progress.cancelled:
                raise Cancelled()
        group.Assimilate()
    except Exception as error:
        if transaction and transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        if group.GetStatus() == DB.TransactionStatus.Started:
            group.RollBack()
        if isinstance(error, Cancelled):
            forms.alert(u'Разделение отменено. Все разрезы этого запуска отменены.')
        else:
            forms.alert(u'Разделение отменено целиком.\nВоздуховод ID: {}.\n{}'.format(
                failed_id, error), title=u'Не удалось разделить воздуховоды')
        return
    output = script.get_output()
    output.set_title(u'Разделение воздуховодов по 1250 мм')
    output.print_md(u'Разделено **{}** воздуховодов. Выполнено **{}** разрезов. Область: {}.'.format(
        len(plans), cuts, scope))
    output.print_table(table_data=[
        [output.linkify(DB.ElementId(plan['id'])),
         u' + '.join(u'{:g}'.format(length) for length in plan['lengths_mm'])] for plan in plans],
        columns=[u'Исходный воздуховод', u'Отрезки, мм'])
    if skipped:
        output.print_table(table_data=[[reason, count] for reason, count in sorted(skipped.items())],
                           columns=[u'Причина пропуска', u'Количество'])


main()
