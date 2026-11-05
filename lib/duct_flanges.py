"""Read rectangular ducts and plan additional flanges from auditable stock cuts."""
import math
from collections import defaultdict, deque
from support_cutting import pack, ticks, SearchBudget, SearchCancelled
from duct_split import manufacturer_values, flange_size, TOLERANCE_MM

STOCK_MM = 1250.0


def _finite(value):
    return not math.isnan(float(value)) and not math.isinf(float(value))


def collect(doc, elements=None):
    """Group compatible straight ducts; manufacturer and unsupported elements are excluded."""
    from pyrevit.api import DB
    from Autodesk.Revit.DB.Mechanical import Duct
    from pyrevit.revit.db.query import get_param
    groups, skipped, problems = {}, defaultdict(list), []
    elements = elements if elements is not None else DB.FilteredElementCollector(doc).OfClass(Duct)
    for duct in elements:
        eid = int(duct.Id.Value)
        if not isinstance(duct, Duct) or duct.IsPlaceholder or duct.DuctType.Shape != DB.ConnectorProfileType.Rectangular:
            skipped[u'Не прямоугольный жесткий воздуховод'].append(eid)
            continue
        if manufacturer_values(duct):
            skipped[u'Заполнен производитель'].append(eid)
            continue
        if not isinstance(duct.Location, DB.LocationCurve) or not isinstance(duct.Location.Curve, DB.Line):
            skipped[u'Непрямолинейная ось'].append(eid)
            continue
        length = duct.Location.Curve.Length * 304.8
        width, height = duct.Width * 304.8, duct.Height * 304.8
        if not all(_finite(v) and v > 0 for v in (length, width, height)):
            problems.append((eid, u'Неверные размеры'))
            continue
        if length > STOCK_MM + TOLERANCE_MM:
            problems.append((eid, u'Длина больше 1250 мм — сначала разделите воздуховод'))
            continue
        if abs(length - STOCK_MM) <= TOLERANCE_MM:
            length = STOCK_MM
        signatures = []
        for obj in (duct, doc.GetElement(duct.GetTypeId())):
            for name in (u'ADSK_Толщина стенки', u'Толщина стенки', u'ADSK_Материал', u'Материал', u'ADSK_Покрытие'):
                parameter = get_param(obj, name)
                if parameter:
                    signatures.append((name, parameter.AsValueString() or parameter.AsString() or u''))
        key = (int(duct.GetTypeId().Value), round(width, 3), round(height, 3), tuple(signatures))
        if key not in groups:
            typ = doc.GetElement(duct.GetTypeId())
            groups[key] = {'type_id': key[0], 'type_name': typ.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString(),
                           'width_mm': round(width, 3), 'height_mm': round(height, 3),
                           'flange': flange_size(width, height), 'signature': signatures, 'parts': []}
        system = duct.get_Parameter(DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
        groups[key]['parts'].append({'id': eid, 'length_mm': length,
                                     'system': system.AsString() if system else u''})
    ordered = sorted(groups.values(), key=lambda g: (g['flange'], g['width_mm'], g['height_mm'], g['type_id'], repr(g['signature'])))
    return ordered, dict(skipped), problems


def calculate_group(group, mode='stock', duct_kerf_mm=0, budget=None):
    """Assign factory ends to an explicit new-stock layout, or use known ends per short part."""
    if mode not in ('stock', 'one_factory_end', 'no_factory_ends'):
        raise ValueError('Unknown stock mode')
    if not _finite(duct_kerf_mm) or duct_kerf_mm < 0:
        raise ValueError('Invalid kerf')
    parts = group['parts']
    if any(not _finite(p['length_mm']) or p['length_mm'] <= 0 or p['length_mm'] > STOCK_MM for p in parts):
        raise ValueError('Split ducts to at most 1250 mm before calculating')
    rows, layouts, rejected = [], [], []
    packing = None
    if mode == 'stock':
        packing = pack([p['length_mm'] for p in parts], STOCK_MM, duct_kerf_mm, budget=budget)
        rejected_indices = set(packing['rejected_indices'])
        queues = defaultdict(deque)
        for index, part in enumerate(parts):
            if index in rejected_indices:
                rejected.append(dict(part))
            else:
                queues[ticks(part['length_mm'])].append(part)
        for pattern in packing['patterns']:
            for repeat in range(pattern['repeat']):
                lengths = pattern['lengths_mm']
                layout = {'parts': [], 'offcut_mm': pattern['offcut_mm'], 'new_flanges': 0}
                for index, length in enumerate(lengths):
                    part = dict(queues[ticks(length)].popleft())
                    factory = 2 if len(lengths) == 1 and length == STOCK_MM else (1 if index == 0 or index == len(lengths) - 1 else 0)
                    part.update({'factory_ends': factory, 'new_flanges': 2 - factory})
                    rows.append(part)
                    layout['parts'].append(part)
                    layout['new_flanges'] += 2 - factory
                layout['offcut_factory_ends'] = 2 - sum(p['factory_ends'] for p in layout['parts'])
                layouts.append(layout)
        if any(queues.values()):
            raise ValueError('Incomplete stock assignment')
    else:
        for original in parts:
            part = dict(original)
            factory = 2 if part['length_mm'] == STOCK_MM else (1 if mode == 'one_factory_end' else 0)
            part.update({'factory_ends': factory, 'new_flanges': 2 - factory})
            rows.append(part)
    total = sum(p['new_flanges'] for p in rows)
    whole = sum(1 for p in rows if p['length_mm'] == STOCK_MM)
    purchase = {'stock_quantity': len(layouts) if mode == 'stock' else None,
                'whole_installed': whole, 'cut_installed': len(rows) - whole,
                'cut_stock_quantity': len(layouts) - whole if mode == 'stock' else None,
                'installed_parts': len(rows),
                'installed_length_mm': sum(p['length_mm'] for p in rows),
                'purchased_length_mm': len(layouts) * STOCK_MM if mode == 'stock' else None,
                'offcut_mm': packing['offcut_mm'] if packing is not None else None,
                'kerf_loss_mm': packing['kerf_loss_mm'] if packing is not None else None}
    return {'group': group, 'parts': rows, 'layouts': layouts, 'packing': packing,
            'purchase': purchase,
            'rejected': rejected, 'new_flanges': total, 'corners': total * 4,
            'gross_rail_mm': total * 2 * (group['width_mm'] + group['height_mm'])}


def calculate(groups, mode='stock', corner20=u'УГФ-0', stock20_mm=3000, stock30_mm=3000,
              deduction20_mm=0, deduction30_mm=0, duct_kerf_mm=0, rail_kerf_mm=0, budget=None, on_progress=None):
    """Return duct layouts, individual rail cuts and a purchase plan; never edit the model."""
    if corner20 not in (u'УГФ-0', u'УГФ-1'):
        raise ValueError('Invalid P20 corner')
    for value in (stock20_mm, stock30_mm):
        if not _finite(value) or value <= 0:
            raise ValueError('Stock length must be positive')
    for value in (deduction20_mm, deduction30_mm, duct_kerf_mm, rail_kerf_mm):
        if not _finite(value) or value < 0:
            raise ValueError('Losses and deductions must be nonnegative')
    budget = budget or SearchBudget(1.0, 100000)
    results, rails = [], {}
    for index, group in enumerate(groups):
        if on_progress and on_progress(index, len(groups) + 2):
            raise SearchCancelled()
        result = calculate_group(group, mode, duct_kerf_mm, budget)
        size = group['flange']
        deduction = deduction20_mm if size == 20 else deduction30_mm
        lengths = [group['width_mm'] - deduction, group['height_mm'] - deduction]
        if result['new_flanges'] and min(lengths) <= 0:
            raise ValueError(u'Вычет для шины %s не меньше стороны сечения %s x %s мм.' % (
                size, group['width_mm'], group['height_mm']))
        if size not in rails:
            rails[size] = {'size': size, 'corner': corner20 if size == 20 else u'УГФ-2',
                           'corners': 0, 'flanges': 0, 'gross_rail_mm': 0, 'cuts': []}
        rail = rails[size]
        rail['corners'] += result['corners']
        rail['flanges'] += result['new_flanges']
        rail['gross_rail_mm'] += result['gross_rail_mm']
        for part in result['parts']:
            for unused in range(part['new_flanges']):
                for length in (lengths[0], lengths[0], lengths[1], lengths[1]):
                    rail['cuts'].append({'length_mm': length, 'duct_id': part['id']})
        results.append(result)
    for index, size in enumerate(sorted(rails)):
        rail = rails[size]
        if on_progress and on_progress(len(groups) + index, len(groups) + 2):
            raise SearchCancelled()
        rail['packing'] = pack([p['length_mm'] for p in rail['cuts']],
                               stock20_mm if size == 20 else stock30_mm, rail_kerf_mm, budget=budget)
        rail['cut_rail_mm'] = sum(p['length_mm'] for p in rail['cuts'])
        rail['rejected'] = [rail['cuts'][i] for i in rail['packing']['rejected_indices']]
    purchase = {}
    for key in ('stock_quantity', 'whole_installed', 'cut_installed', 'cut_stock_quantity',
                'installed_parts', 'installed_length_mm', 'purchased_length_mm', 'offcut_mm', 'kerf_loss_mm'):
        if mode != 'stock' and key not in ('whole_installed', 'cut_installed', 'installed_parts', 'installed_length_mm'):
            purchase[key] = None
        else:
            purchase[key] = sum(r['purchase'][key] for r in results)
    return {'groups': results, 'rails': [rails[size] for size in sorted(rails)], 'purchase': purchase,
            'duct_rejected': [part for result in results for part in result['rejected']],
            'new_flanges': sum(r['new_flanges'] for r in results),
            'corners': sum(r['corners'] for r in results), 'mode': mode}
