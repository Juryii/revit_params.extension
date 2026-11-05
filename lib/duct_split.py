"""Split straight rectangular ducts without changing their route or end connections."""
import math

STANDARD_MM = 1250.0
TOLERANCE_MM = 0.01
MANUFACTURERS = (u'ADSK_Завод-изготовитель', u'ADSK_Изготовитель',
                 u'Производитель', u'Изготовитель', 'Manufacturer')


def segment_lengths(length_mm, minimum_mm=2.55):
    """Return full 1250 mm sections and the remainder; refuse an unbuildable sliver."""
    if math.isnan(length_mm) or math.isinf(length_mm):
        raise ValueError(u'Некорректная длина.')
    if length_mm <= 0:
        raise ValueError(u'Длина должна быть больше нуля.')
    if length_mm <= STANDARD_MM + TOLERANCE_MM:
        return [length_mm]
    count = int(length_mm // STANDARD_MM)
    remainder = length_mm - count * STANDARD_MM
    if remainder <= TOLERANCE_MM:
        return [STANDARD_MM] * (count - 1) + [STANDARD_MM + remainder]
    if remainder < minimum_mm:
        raise ValueError(u'Остаток %.3f мм слишком мал для Revit; участок пропущен.' % remainder)
    return [STANDARD_MM] * count + [remainder]


def flange_size(width_mm, height_mm):
    """User's rule: select by the largest cross-section side, including 1000 in P20."""
    return 20 if max(width_mm, height_mm) <= 1000.0 + TOLERANCE_MM else 30


def _api():
    from pyrevit.api import DB
    from Autodesk.Revit.DB.Mechanical import Duct, MechanicalUtils
    return DB, Duct, MechanicalUtils


def element_id(element):
    return int(element.Id.Value)


def manufacturer_values(duct):
    DB, _, _ = _api()
    values = []
    for obj in (duct, duct.Document.GetElement(duct.GetTypeId())):
        if obj is None:
            continue
        for name in MANUFACTURERS:
            for parameter in obj.GetParameters(name):
                if parameter.StorageType == DB.StorageType.String:
                    value = parameter.AsString()
                    if value and value.strip():
                        values.append(value.strip())
        parameter = obj.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MANUFACTURER)
        if parameter and parameter.StorageType == DB.StorageType.String:
            value = parameter.AsString()
            if value and value.strip():
                values.append(value.strip())
    return values


def plan_duct(duct):
    """Validate eligibility without editing; return (plan, reason)."""
    DB, Duct, _ = _api()
    if not isinstance(duct, Duct):
        return None, u'Не жесткий воздуховод'
    if duct.IsPlaceholder or duct.DuctType.Shape != DB.ConnectorProfileType.Rectangular:
        return None, u'Не прямоугольный воздуховод'
    if manufacturer_values(duct):
        return None, u'Заполнен производитель'
    if duct.GroupId != DB.ElementId.InvalidElementId:
        return None, u'Элемент в группе'
    if duct.Pinned:
        return None, u'Элемент закреплен'
    if not isinstance(duct.Location, DB.LocationCurve) or not isinstance(duct.Location.Curve, DB.Line):
        return None, u'Непрямолинейная ось'
    length = duct.Location.Curve.Length * 304.8
    try:
        lengths = segment_lengths(length, max(2.55, duct.Document.Application.ShortCurveTolerance * 304.8 + .01))
    except ValueError as error:
        return None, u'%s' % error
    if len(lengths) < 2:
        return None, u'Длина не больше 1250 мм'
    return {'id': element_id(duct), 'length_mm': length, 'lengths_mm': lengths,
            'flange': flange_size(duct.Width * 304.8, duct.Height * 304.8)}, None


def _connectors(duct):
    DB, _, _ = _api()
    return [c for c in duct.ConnectorManager.Connectors
            if c.ConnectorType in (DB.ConnectorType.End, DB.ConnectorType.Curve)]


def _layers(doc, duct):
    DB, _, _ = _api()
    result = []
    for kind, ids in [('insulation', DB.InsulationLiningBase.GetInsulationIds(doc, duct.Id)),
                      ('lining', DB.InsulationLiningBase.GetLiningIds(doc, duct.Id))]:
        for eid in ids:
            layer = doc.GetElement(eid)
            result.append((kind, int(layer.GetTypeId().Value), round(layer.Thickness, 9)))
    return sorted(result)


def _end_at(duct, point):
    DB, _, _ = _api()
    return next(c for c in _connectors(duct)
                if c.ConnectorType == DB.ConnectorType.End and c.Origin.DistanceTo(point) < 1e-6)


def split_duct(doc, plan):
    """Inside an open transaction: split, reconnect coincident ends, keep original external refs."""
    DB, _, MechanicalUtils = _api()
    duct = doc.GetElement(DB.ElementId(plan['id']))
    current_plan, reason = plan_duct(duct)
    if current_plan is None or abs(current_plan['length_mm'] - plan['length_mm']) > TOLERANCE_MM:
        raise ValueError(reason or u'Воздуховод изменился после предварительного расчета.')
    curve = duct.Location.Curve
    start, direction = curve.GetEndPoint(0), curve.Direction
    external = []
    for connector in _connectors(duct):
        for ref in connector.AllRefs:
            if ref.Owner.Id != duct.Id and ref.ConnectorType in (DB.ConnectorType.End, DB.ConnectorType.Curve):
                external.append((int(ref.Owner.Id.Value), ref.Id, connector.Origin))
    state = {'type_id': duct.GetTypeId(), 'width': duct.Width, 'height': duct.Height,
             'system_id': duct.MEPSystem.Id if duct.MEPSystem else None,
             'layers': _layers(doc, duct), 'external': external,
             'start': start, 'end': curve.GetEndPoint(1), 'joints': []}
    pieces = [duct]
    offset = 0.0
    for length in plan['lengths_mm'][:-1]:
        offset += length
        point = start + direction * (offset / 304.8)
        target = next(piece for piece in pieces
                      if piece.Location.Curve.Distance(point) < 1e-6
                      and all(piece.Location.Curve.GetEndPoint(i).DistanceTo(point) > 1e-6 for i in (0, 1)))
        new_id = MechanicalUtils.BreakCurve(doc, target.Id, point)
        if new_id == DB.ElementId.InvalidElementId:
            raise ValueError(u'Revit не смог разделить воздуховод.')
        doc.Regenerate()
        new_piece = doc.GetElement(new_id)
        a, b = _end_at(target, point), _end_at(new_piece, point)
        if not a.IsConnectedTo(b):
            a.ConnectTo(b)
        pieces.append(new_piece)
        state['joints'].append((element_id(target), element_id(new_piece), point))
        doc.Regenerate()
    state['ids'] = [element_id(piece) for piece in pieces]
    state['plan'] = plan
    return state


def verify_split(doc, state):
    """After transaction commit: validate lengths, systems, layers and every previous connection."""
    DB, _, _ = _api()
    pieces = [doc.GetElement(DB.ElementId(i)) for i in state['ids']]
    actual = sorted(piece.Location.Curve.Length * 304.8 for piece in pieces)
    expected = sorted(state['plan']['lengths_mm'])
    if len(actual) != len(expected) or any(abs(a - b) > TOLERANCE_MM for a, b in zip(actual, expected)):
        raise ValueError(u'После соединения изменились длины отрезков.')
    for piece in pieces:
        if piece.GetTypeId() != state['type_id'] or abs(piece.Width - state['width']) > 1e-7 or abs(piece.Height - state['height']) > 1e-7:
            raise ValueError(u'Изменился тип или сечение воздуховода.')
        if state['system_id'] and (not piece.MEPSystem or piece.MEPSystem.Id != state['system_id']):
            raise ValueError(u'Не сохранилась принадлежность к системе.')
        if _layers(doc, piece) != state['layers']:
            raise ValueError(u'Не сохранилась изоляция или внутренняя облицовка.')
    for owner_id, connector_id, point in state['external']:
        owner = doc.GetElement(DB.ElementId(owner_id))
        manager = owner.ConnectorManager if hasattr(owner, 'ConnectorManager') else owner.MEPModel.ConnectorManager
        ref = manager.Lookup(connector_id)
        if not any(c.Origin.DistanceTo(point) < 1e-6 and c.IsConnectedTo(ref)
                   for piece in pieces for c in _connectors(piece)):
            raise ValueError(u'Потеряно существующее соединение воздуховода.')
    for a_id, b_id, point in state['joints']:
        at_joint = [_end_at(piece, point) for piece in pieces
                    if any(piece.Location.Curve.GetEndPoint(i).DistanceTo(point) < 1e-6 for i in (0, 1))]
        if len(at_joint) != 2 or not at_joint[0].IsConnectedTo(at_joint[1]):
            raise ValueError(u'Новые торцы не соединены.')
    for point in (state['start'], state['end']):
        if not any(any(piece.Location.Curve.GetEndPoint(i).DistanceTo(point) < 1e-6 for i in (0, 1)) for piece in pieces):
            raise ValueError(u'Изменилась исходная трасса.')
    return actual
