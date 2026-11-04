"""Количество закупаемых шпилек и траверс с картой раскроя. Модель не изменяется."""
import math
from collections import Counter
from System.Windows import SystemParameters
from pyrevit import DB, forms, revit, script
from support_cutting import collect, pack, SearchBudget, SearchCancelled


class PurchaseWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, script.get_bundle_file('ui.xaml'))
        self.MaxHeight = SystemParameters.WorkArea.Height - 40
        self.values = None

    def calculate(self, sender, args):
        try:
            vals = [float(box.Text.strip().replace(',', '.'))
                    for box in [self.rod_length, self.channel_length, self.kerf]]
            if not all(not math.isnan(v) and not math.isinf(v) for v in vals):
                raise ValueError()
            if min(vals[:2]) <= 0 or vals[2] < 0 or vals[2] >= min(vals[:2]):
                raise ValueError()
            self.allowances = [float(box.Text.strip().replace(',', '.'))
                               for box in [self.rod_allowance, self.channel_allowance]]
            if any(math.isnan(v) or math.isinf(v) or v < 0 for v in self.allowances):
                raise ValueError()
            self.search_limits = (.4, 30000)
            if self.deep_search.IsChecked:
                seconds = float(self.search_seconds.Text.strip().replace(',', '.'))
                steps = int(self.search_steps.Text.strip())
                if math.isnan(seconds) or math.isinf(seconds) or seconds <= 0 or steps <= 0:
                    raise ValueError()
                self.search_limits = (seconds, steps)
            self.values = vals
            self.Close()
        except ValueError:
            self.error.Text = u'Проверьте ввод: длины и время больше нуля; шаги — целое число больше нуля; припуски неотрицательны; рез — от нуля до длины заготовки.'


def mm(value):
    return ('%.3f' % value).rstrip('0').rstrip('.').replace('.', ',')


def main():
    window = PurchaseWindow()
    window.ShowDialog()
    if window.values is None:
        return
    rod, channel, kerf = window.values
    groups, problems = collect(revit.doc)
    if not groups:
        forms.alert(u'В текущем проекте нет подходящих деталей подвесов РП.'
                    + (u' Неверная длина у элементов: ' + ', '.join(map(str, problems)) if problems else ''))
        return
    output = script.get_output()
    output.set_title(u'Закупка креплений — раскрой')
    output.print_md(u'# Закупка креплений')
    output.print_md(u'Проект: **{}**. Шпильки: {} мм; траверсы: {} мм; рез: {} мм.'.format(
        revit.doc.Title, mm(rod), mm(channel), mm(kerf)))
    output.print_md(u'Весь текущий проект; связанные модели исключены. Системы объединены для совместного раскроя. '
                    u'Один рез на отрезок, целая заготовка без реза.')
    output.print_md(u'Припуск на каждую шпильку: **{} мм**, на каждую траверсу: **{} мм**. '
                    u'Длины в карте раскроя включают припуск один раз на деталь. Модель и спецификация не изменены.'.format(
        mm(window.allowances[0]), mm(window.allowances[1])))
    rows, results, rejected_count = [], [], 0
    with forms.ProgressBar(title=u'Расчёт раскроя: {value}/{max_value}', cancellable=True) as progress:
        current_group = [0]

        def refresh_search():
            progress.title = u'Поиск {value}/{max_value}: %d с, %d шагов' % (budget.elapsed(), budget.nodes)
            progress.update_progress(current_group[0], len(groups))
            return progress.cancelled

        budget = SearchBudget(window.search_limits[0], window.search_limits[1], refresh_search)
        for i, g in enumerate(groups):
            current_group[0] = i
            if progress.cancelled:
                forms.alert(u'Расчёт отменён. Отчёт неполный, для заказа не использовать.')
                return
            allowance = window.allowances[0] if g['kind'] == 'rod' else window.allowances[1]
            r = pack(g['lengths_mm'], rod if g['kind'] == 'rod' else channel, kerf,
                     budget=budget, allowance_mm=allowance)
            rejected_count += len(r['rejected_mm'])
            results.append((g, r))
            status = u'Минимум подтверждён' if r['optimal'] else u'Подбор; минимум не доказан'
            if r['stop_reason']:
                status += u' (лимит времени)' if r['stop_reason'] == 'time' else u' (лимит шагов)'
            rows.append([g['name'], mm(r['stock_mm']), r['quantity'], r['parts'],
                         mm(r['required_mm']/1000), mm(r['offcut_mm']/1000),
                         mm(r['kerf_loss_mm']/1000),
                         status,
                         len(r['rejected_mm'])])
            progress.update_progress(i+1, len(groups))
    output.print_md(u'Общие лимиты поиска: {} с / {} шагов. Использовано {} шагов; время расчёта {} с. '
                    u'Остановка по первому лимиту; при доказанном минимуме — раньше.'.format(
        mm(window.search_limits[0]), window.search_limits[1], budget.nodes, mm(budget.elapsed())))
    output.print_table(table_data=rows, columns=[u'Наименование', u'Заготовка, мм', u'Заказать, шт.',
        u'Отрезков учтено', u'С припуском, м', u'Остаток, м', u'Рез, м', u'Точность', u'Не помещаются'])
    for g, r in results:
        output.print_md(u'## ' + g['name'])
        output.print_md(u'Припуск {} мм на деталь; суммарно в учтённых отрезках: {} м.'.format(
            mm(r['allowance_mm']), mm(r['allowance_total_mm']/1000)))
        patterns = []
        for pat in r['patterns']:
            counts = Counter(pat['lengths_mm'])
            formula = u' + '.join(u'{} × {}'.format(n, mm(length)) for length, n in sorted(counts.items(), reverse=True))
            patterns.append([pat['repeat'], formula, mm(pat['offcut_mm'])])
        if patterns:
            output.print_table(table_data=patterns, columns=[u'Заготовок, шт.', u'Раскрой, мм', u'Остаток каждой, мм'])
        if r['rejected_mm']:
            bad = [(g['ids'][j], g['lengths_mm'][j]+r['allowance_mm']) for j in r['rejected_indices']]
            output.print_md(u'**Не включены в заказ: {} отрезков — превышают заготовку или не хватает места на рез.**'.format(len(r['rejected_mm'])))
            for eid, length in bad:
                element = revit.doc.GetElement(DB.ElementId(eid))
                parent = element.SuperComponent if element else None
                detail = output.linkify(DB.ElementId(eid), title=u'Деталь {}'.format(eid))
                support = output.linkify(parent.Id, title=u'Подвес') if parent else ''
                reason = u'Длиннее заготовки' if length > r['stock_mm'] else u'Не хватает места на рез'
                output.print_md(u'{} мм — {}. {} {}'.format(mm(length), reason, detail, support))
    if problems:
        output.print_md(u'**Элементы с неверной длиной, не учтены:** ' + ', '.join(map(str, problems)))
    forms.alert(u'Расчёт готов: {} заготовок. Отчёт и карта раскроя открыты в окне pyRevit.{}'.format(
        sum(r['quantity'] for g, r in results),
        u' Есть неучтённые детали — см. предупреждения в отчёте.' if rejected_count or problems else ''))


try:
    main()
except SearchCancelled:
    forms.alert(u'Расчёт отменён. Отчёт неполный, для заказа не использовать.')
except Exception as exc:
    forms.alert(u'Не удалось завершить расчёт закупки: {}'.format(exc))
