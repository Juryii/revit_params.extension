"""Read individual support parts and pack them into purchased blanks. No model edits."""
import math
import time
from collections import Counter


FAMILIES = {
    u'РП_Шпилька_М8': 'rod',
    u'РП_Шпилька_М10': 'rod',
    u'РП_Шпилька_подвеса_выборная': 'rod',
    u'РП_Траверса_монтажная_отрезок': 'channel',
}


def ticks(mm):
    return int(math.ceil(float(mm) * 1000.0 - 0.00001))


class SearchStopped(Exception):
    pass


class SearchCancelled(Exception):
    pass


class SearchBudget(object):
    """One time/step budget shared by all material groups; callback pumps the UI."""
    def __init__(self, seconds, node_limit, callback=None):
        self.clock = getattr(time, 'monotonic', time.time)
        self.started = self.clock()
        self.seconds = seconds
        self.node_limit = node_limit
        self.nodes = 0
        self.callback = callback
        self.last_update = self.started - 1
        self.reason = None

    def step(self):
        now = self.clock()
        if self.callback and now-self.last_update >= .1:
            self.last_update = now
            if self.callback():
                raise SearchCancelled()
        if self.nodes >= self.node_limit:
            self.reason = 'steps'
            raise SearchStopped()
        if now-self.started >= self.seconds:
            self.reason = 'time'
            raise SearchStopped()
        self.nodes += 1

    def elapsed(self):
        return self.clock()-self.started


def pack(lengths_mm, stock_mm, kerf_mm=0, seconds=.4, node_limit=30000, budget=None, allowance_mm=0):
    """Bounded exact search after best-fit. Status never claims an unproven optimum.

    Integer micrometres round each required length upward. Each cut part consumes
    one kerf; an uncut whole blank consumes none. No splicing or reusable inventory.
    """
    lengths_mm = list(lengths_mm)
    if not all(not math.isnan(float(v)) and not math.isinf(float(v)) for v in [stock_mm, kerf_mm, allowance_mm] + lengths_mm):
        raise ValueError('Non-finite length')
    stock = int(math.floor(float(stock_mm) * 1000.0 + 0.00001))
    kerf = ticks(kerf_mm)
    if stock <= 0 or kerf_mm < 0 or allowance_mm < 0 or any(v <= 0 for v in lengths_mm):
        raise ValueError('Lengths must be positive; kerf cannot be negative')
    valid, rejected, full, rejected_indices = [], [], [], []
    for index, value in enumerate(lengths_mm):
        value += allowance_mm
        n = ticks(value)
        if n == stock:
            full.append([n])
        elif n + kerf > stock:
            rejected.append(value)
            rejected_indices.append(index)
        else:
            valid.append(n)
    parts = sorted(valid, reverse=True)
    bins, used = [], []
    for n in parts:
        cost = n + kerf
        candidates = [j for j in range(len(bins)) if used[j] + cost <= stock]
        if candidates:
            j = min(candidates, key=lambda j: stock - used[j] - cost)
            bins[j].append(n)
            used[j] += cost
        else:
            bins.append([n])
            used.append(cost)
    lower = (sum(parts) + len(parts)*kerf + stock - 1)//stock
    optimal = len(bins) == lower
    budget = budget or SearchBudget(seconds, node_limit)
    if not optimal:
        costs = [n + kerf for n in parts]
        for count in range(lower, len(bins)):
            trial, loads = [[] for _ in range(count)], [0]*count
            try:
                options, cursor, assignment = [None]*len(parts), [0]*len(parts), [-1]*len(parts)
                i = 0
                while i >= 0 and i < len(parts):
                    budget.step()
                    if options[i] is None:
                        seen, choices = set(), []
                        for j in sorted(range(count), key=lambda j: -loads[j]):
                            load = loads[j]
                            if load in seen or load+costs[i] > stock:
                                continue
                            seen.add(load)
                            choices.append(j)
                            if load == 0:
                                break
                        options[i] = choices
                    if cursor[i] == len(options[i]):
                        options[i], cursor[i] = None, 0
                        i -= 1
                        if i >= 0:
                            j = assignment[i]
                            loads[j] -= costs[i]
                            trial[j].pop()
                    else:
                        j = options[i][cursor[i]]
                        cursor[i] += 1
                        assignment[i] = j
                        loads[j] += costs[i]
                        trial[j].append(parts[i])
                        i += 1
                if i == len(parts):
                    bins = trial
                    optimal = True
                    break
            except SearchStopped:
                break
        else:
            optimal = True
    bins = full + bins
    patterns = Counter(tuple(sorted(b, reverse=True)) for b in bins)
    cutting_loss = sum(len(b)*kerf for b in bins if b != [stock])
    return {
        'stock_mm': stock/1000.0, 'quantity': len(bins),
        'stop_reason': None if optimal else budget.reason,
        'allowance_mm': allowance_mm,
        'allowance_total_mm': sum(len(b) for b in bins)*allowance_mm,
        'rejected_indices': rejected_indices,
        'parts': sum(len(b) for b in bins), 'optimal': optimal,
        'lower_bound': len(full)+lower, 'rejected_mm': rejected,
        'required_mm': sum(sum(b) for b in bins)/1000.0,
        'kerf_loss_mm': cutting_loss/1000.0,
        'offcut_mm': (len(bins)*stock-sum(sum(b) for b in bins)-cutting_loss)/1000.0,
        'patterns': [{'lengths_mm': [n/1000.0 for n in pat], 'repeat': repeat,
                      'offcut_mm': (stock-sum(pat)-(0 if pat == (stock,) else len(pat)*kerf))/1000.0}
                     for pat, repeat in sorted(patterns.items(), reverse=True)],
    }


def collect(doc):
    """Read the current document once, including shared nested instances exactly once."""
    from pyrevit import DB
    groups, problems = {}, []

    def text(element, name):
        p = element.LookupParameter(name)
        return p.AsString() if p and p.StorageType == DB.StorageType.String else ''

    for e in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance):
        kind = FAMILIES.get(e.Symbol.FamilyName)
        if not kind:
            continue
        eid = e.Id.Value if hasattr(e.Id, 'Value') else e.Id.IntegerValue
        p = next((e.LookupParameter(n) for n in
                  [u'ADSK_Размер_Длина', u'Длина элемента', u'Длина шпильки']
                  if e.LookupParameter(n) is not None), None)
        if (not p or p.StorageType != DB.StorageType.Double or p.AsDouble() <= 0
                or math.isnan(p.AsDouble()) or math.isinf(p.AsDouble())):
            problems.append(eid)
            continue
        name = text(e, u'ADSK_Наименование') or e.Symbol.FamilyName + ': ' + e.Name
        key = (kind, name, text(e, u'ADSK_Обозначение'), text(e, u'ADSK_Завод-изготовитель'))
        item = groups.setdefault(key, {'kind': kind, 'name': name, 'lengths_mm': [], 'ids': []})
        item['lengths_mm'].append(p.AsDouble()*304.8)
        item['ids'].append(eid)
    return sorted(groups.values(), key=lambda g: (g['kind'], g['name'])), problems
