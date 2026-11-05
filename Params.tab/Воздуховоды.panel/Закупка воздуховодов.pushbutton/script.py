"""Рассчитывает закупку воздуховодов 1250 мм, дополнительной шинорейки и уголков."""
import math
from collections import Counter
from xml.sax.saxutils import escape
import clr
clr.AddReference('PresentationFramework')
from System.Windows import SystemParameters
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB.Mechanical import Duct
from duct_flanges import collect, calculate
from support_cutting import SearchBudget, SearchCancelled

MODES = ('stock', 'one_factory_end', 'no_factory_ends')
MODE_NAMES = (u'Новый раскрой заготовок 1250 мм', u'Один заводской торец у каждого короткого отрезка',
              u'У коротких отрезков нет заводских фланцев')


def mm(value):
    return ('%.3f' % value).rstrip('0').rstrip('.').replace('.', ',')


class FlangeWindow(forms.WPFWindow):
    def __init__(self, has_selection):
        forms.WPFWindow.__init__(self, script.get_bundle_file('ui.xaml'))
        self.MaxHeight = SystemParameters.WorkArea.Height - 40
        self.values = None
        self.selected_scope.IsEnabled = has_selection
        if has_selection:
            self.scope.SelectedIndex = 2

    def calculate_click(self, sender, args):
        try:
            values = [float(box.Text.strip().replace(',', '.')) for box in
                      (self.stock20, self.stock30, self.deduction20, self.deduction30, self.duct_kerf, self.rail_kerf)]
            if any(math.isnan(v) or math.isinf(v) for v in values) or min(values[:2]) <= 0 or min(values[2:]) < 0:
                raise ValueError()
            if values[-1] >= min(values[:2]) or values[-2] >= 1250:
                raise ValueError()
            self.values = {'stock20_mm': values[0], 'stock30_mm': values[1],
                           'deduction20_mm': values[2], 'deduction30_mm': values[3],
                           'duct_kerf_mm': values[4], 'rail_kerf_mm': values[5],
                           'mode': MODES[self.mode.SelectedIndex],
                           'corner20': u'УГФ-0' if self.corner20.SelectedIndex == 0 else u'УГФ-1'}
            self.Close()
        except ValueError:
            self.error.Text = u'Длины заготовок должны быть больше нуля; вычеты и потери — неотрицательными. Потеря на рез должна быть меньше заготовки.'


def report(output, result, skipped, problems, settings, scope, details, title):
    output.set_title(u'Закупка воздуховодов, шин и уголков')
    output.print_md(u'# Закупка воздуховодов, шин и уголков')
    output.print_md(u'Проект: **{}**. Область: **{}**. {}'.format(escape(title), scope, MODE_NAMES[MODES.index(settings['mode'])]))
    raw_rejected = result['duct_rejected']
    rail_rejected = [p for rail in result['rails'] for p in rail['rejected']]
    if problems or raw_rejected or rail_rejected:
        output.print_md(u'**НЕПОЛНЫЙ РАСЧЁТ: есть исключенные детали. Закупка ниже покрывает только учтенные отрезки. См. список в конце.**')
    purchase = result['purchase']
    if settings['mode'] == 'stock':
        output.print_md(u'**Купить воздуховодов 1250 мм: {} шт.** Из них **{} шт. установить целыми**, '
                        u'**{} шт. пустить на раскрой** → **{} коротких отрезков в монтаж**.'.format(
                            purchase['stock_quantity'], purchase['whole_installed'],
                            purchase['cut_stock_quantity'], purchase['cut_installed']))
        output.print_md(u'В монтаж всего: **{} участков / {} м**. Остатки раскроя: **{} м**; потеря на рез: **{} м**.'.format(
            purchase['installed_parts'], mm(purchase['installed_length_mm'] / 1000),
            mm(purchase['offcut_mm'] / 1000), mm(purchase['kerf_loss_mm'] / 1000)))
    else:
        output.print_md(u'В монтаж: **{} целых участков 1250 мм и {} коротких отрезков**. '
                        u'Закупка воздуховодов в режиме готовых деталей не рассчитывается; '
                        u'для неё выберите новый раскрой заготовок 1250 мм.'.format(
                            purchase['whole_installed'], purchase['cut_installed']))
    output.print_md(u'## Воздуховоды по сечениям')
    duct_rows, specifications = [], []
    for index, r in enumerate(result['groups']):
        group, counts = r['group'], r['purchase']
        description = [group['type_name']]
        for name, value in group.get('signature', []):
            label = u'{}: {}'.format(name, value)
            if value and label not in description:
                description.append(label)
        specification = escape(u'; '.join(description))
        if specification not in specifications:
            specifications.append(specification)
        duct_rows.append([
            u'{}: {}×{}'.format(index + 1, mm(group['width_mm']), mm(group['height_mm'])),
            specifications.index(specification) + 1,
            counts['stock_quantity'] if counts['stock_quantity'] is not None else u'—',
            counts['whole_installed'],
            counts['cut_stock_quantity'] if counts['cut_stock_quantity'] is not None else u'—',
            counts['cut_installed'],
            mm(counts['offcut_mm'] / 1000) if counts['offcut_mm'] is not None else u'—'])
    output.print_table(table_data=duct_rows,
                       columns=[u'Группа: сечение, мм', u'Исполнение', u'Купить 1250, шт.',
                                u'Целыми в монтаж, шт.', u'Заготовок в раскрой, шт.',
                                u'Коротких в монтаж, шт.', u'Остаток, м'])
    for index, specification in enumerate(specifications):
        output.print_md(u'**Исполнение {}:** {}'.format(index + 1, specification))
    if settings['mode'] == 'stock' and any(not r['packing']['optimal'] for r in result['groups']):
        output.print_md(u'Количество воздуховодов подобрано по выполнимой схеме; минимум заготовок подтвержден не для всех сечений.')
    output.print_md(u'## Дополнительные шины и уголки')
    output.print_table(table_data=[
        [u'Шина %s' % rail['size'], rail['flanges'], mm(rail['cut_rail_mm'] / 1000),
         mm(rail['packing']['stock_mm']), rail['packing']['quantity'], rail['corner'], rail['corners'],
         u'Минимум заготовок подтвержден' if rail['packing']['optimal'] else u'Подбор заготовок']
        for rail in result['rails']], columns=[u'Шина', u'Новых фланцев', u'Расход, м', u'Заготовка, мм',
                                              u'Заготовок, шт.', u'Уголок', u'Уголков, шт.', u'Раскрой шин'])
    output.print_md(u'**Новых фланцев: {}. Дополнительных уголков: {} шт.** Заводские шинорейки и уголки в эти количества не включены.'.format(result['new_flanges'], result['corners']))
    output.print_md(u'Вычет на каждую шину 20/30: **{} / {} мм**; потери на рез воздуховода/шины: **{} / {} мм**. '
                    u'При вычете 0 длины шин считаются по сторонам сечения.'.format(
                        mm(settings['deduction20_mm']), mm(settings['deduction30_mm']),
                        mm(settings['duct_kerf_mm']), mm(settings['rail_kerf_mm'])))
    output.print_md(u'Фасонные детали и внешний складской запас не учитываются. '
                    u'Размер шины по большей стороне: до 1000 мм — 20, больше — 30.')
    if details:
        if settings['mode'] == 'stock':
            output.print_md(u'## Раскрой воздуховодов')
            output.print_md(u'Для раскроя резервируется по одному резу на короткую деталь; целая заготовка — без реза. '
                            u'Крайние детали получают заводские торцы, средние — два новых фланца. Остатки этой карты уже учтены.')
            output.print_md(u'Для каждой детали: ссылка на ID, длина, число сохраненных заводских торцов. Детали перечислены от одного заводского края к другому; промежуточный остаток располагается перед последней деталью. Для одиночной короткой детали остаток остается со вторым заводским торцом.')
            for index, r in enumerate(result['groups']):
                output.print_md(u'**Группа {}: {}×{} мм.** {}'.format(index + 1, mm(r['group']['width_mm']), mm(r['group']['height_mm']),
                    u'Минимум заготовок по принятому резерву на рез подтвержден.' if r['packing']['optimal'] else u'Подбор заготовок; минимум не доказан.'))
                output.print_table(table_data=[
                    [n + 1, u'; '.join(u'{}: {} мм, заводских торцов {}'.format(
                        output.linkify(DB.ElementId(p['id'])), mm(p['length_mm']), p['factory_ends']) for p in layout['parts']),
                     layout['new_flanges'], mm(layout['offcut_mm']), layout['offcut_factory_ends']]
                    for n, layout in enumerate(r['layouts'])],
                    columns=[u'Заготовка', u'Детали', u'Новых фланцев', u'Остаток, мм', u'Зав. торцов у остатка'])
        else:
            output.print_md(u'## Торцы учтенных деталей')
            output.print_table(table_data=[
                [output.linkify(DB.ElementId(p['id'])), mm(p['length_mm']), p['factory_ends'], p['new_flanges']]
                for r in result['groups'] for p in r['parts']],
                columns=[u'Воздуховод', u'Длина, мм', u'Заводских торцов', u'Новых фланцев'])
        output.print_md(u'## Раскрой шинорейки')
        for rail in result['rails']:
            output.print_md(u'**Шина {}: заготовка {} мм.**'.format(rail['size'], mm(rail['packing']['stock_mm'])))
            output.print_table(table_data=[
                [u' + '.join(u'{} × {}'.format(mm(length), count) for length, count in sorted(Counter(p['lengths_mm']).items(), reverse=True)),
                 p['repeat'], mm(p['offcut_mm'])] for p in rail['packing']['patterns']],
                columns=[u'Отрезки в одной заготовке, мм × шт.', u'Повторить, раз', u'Остаток, мм'])
    exclusions = list(problems) + [(p['id'], u'Отрезок воздуховода не помещается с учетом реза') for p in raw_rejected]
    exclusions += [(p['duct_id'], u'Шина длиной {} мм не помещается в заготовку'.format(mm(p['length_mm']))) for p in rail_rejected]
    if exclusions:
        output.print_md(u'## Не учтено в закупке')
        output.print_table(table_data=[[output.linkify(DB.ElementId(eid)), escape(reason)] for eid, reason in exclusions],
                           columns=[u'Воздуховод', u'Причина'])
    if skipped:
        output.print_table(table_data=[[reason, len(ids)] for reason, ids in sorted(skipped.items())],
                           columns=[u'Пропущено', u'Количество'])


def main():
    doc = revit.doc
    if doc is None or doc.IsFamilyDocument:
        forms.alert(u'Откройте проект Revit для расчёта закупки воздуховодов.')
        return
    selected = list(revit.uidoc.Selection.GetElementIds())
    window = FlangeWindow(bool(selected))
    window.ShowDialog()
    if window.values is None:
        return
    scope_index = window.scope.SelectedIndex
    scope = (u'Весь проект', u'Текущий вид', u'Выбранные элементы')[scope_index]
    elements = None
    if scope_index == 1:
        if not DB.FilteredElementCollector.IsViewValidForElementIteration(doc, doc.ActiveView.Id):
            forms.alert(u'На этом виде нельзя собрать воздуховоды. Выберите план, 3D-вид или область «Весь проект».')
            return
        elements = list(DB.FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(Duct))
    elif scope_index == 2:
        elements = [doc.GetElement(eid) for eid in selected]
    groups, skipped, problems = collect(doc, elements)
    if not groups:
        output = script.get_output()
        output.set_title(u'Закупка воздуховодов, шин и уголков')
        output.print_md(u'Нет подходящих воздуховодов для расчёта.')
        output.print_table(table_data=[[reason, len(ids)] for reason, ids in sorted(skipped.items())],
                           columns=[u'Пропущено', u'Количество'])
        if problems:
            output.print_table(table_data=[[output.linkify(DB.ElementId(eid)), reason] for eid, reason in problems],
                               columns=[u'Воздуховод', u'Причина'])
        return
    try:
        with forms.ProgressBar(title=u'Расчёт закупки: {value}/{max_value}', cancellable=True) as progress:
            def advance(value, maximum):
                progress.update_progress(value, maximum)
                return progress.cancelled
            def refresh():
                return advance(0, len(groups) + 2)
            budget = SearchBudget(2.0, 200000, refresh)
            result = calculate(groups, budget=budget, on_progress=advance, **window.values)
            if progress.cancelled:
                return
    except SearchCancelled:
        return
    except ValueError as error:
        forms.alert(u'Не удалось рассчитать: {}'.format(error), title=u'Закупка воздуховодов')
        return
    report(script.get_output(), result, skipped, problems, window.values, scope, bool(window.details.IsChecked), doc.Title)


main()
