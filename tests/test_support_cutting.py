import itertools
import random
import sys
import unittest
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'lib'))
from support_cutting import pack, SearchBudget, SearchCancelled


def oracle(items, stock):
    items = sorted(items, reverse=True)
    for count in range(1, len(items)+1):
        loads = [0]*count

        def fit(i):
            if i == len(items):
                return True
            seen = set()
            for j in range(count):
                if loads[j] in seen or loads[j]+items[i] > stock:
                    continue
                seen.add(loads[j])
                loads[j] += items[i]
                if fit(i+1):
                    return True
                loads[j] -= items[i]
            return False

        if fit(0):
            return count
    return 0


class CuttingTests(unittest.TestCase):
    def assert_conservation(self, lengths, stock, kerf=0, **kwargs):
        r = pack(lengths, stock, kerf, **kwargs)
        reconstructed = []
        for p in r['patterns']:
            self.assertGreaterEqual(p['offcut_mm'], 0)
            reconstructed.extend(p['lengths_mm']*p['repeat'])
        rejected = r['rejected_mm']
        self.assertEqual(len(reconstructed)+len(rejected), len(lengths))
        self.assertEqual(Counter(reconstructed+rejected), Counter(lengths))
        self.assertAlmostEqual(r['quantity']*stock,
                               r['required_mm']+r['kerf_loss_mm']+r['offcut_mm'])
        return r

    def test_cannot_divide_total_metres(self):
        r = self.assert_conservation([600]*3, 1000)
        self.assertEqual(r['quantity'], 3)
        self.assertTrue(r['optimal'])

    def test_exhaustive_small_against_independent_oracle(self):
        for n in range(1, 7):
            for parts in itertools.combinations_with_replacement([2, 3, 4, 6, 7, 8], n):
                r = self.assert_conservation(list(parts), 10, seconds=10)
                self.assertTrue(r['optimal'])
                self.assertEqual(r['quantity'], oracle(parts, 10))

    def test_loss_and_whole_blank(self):
        r = self.assert_conservation([3000, 1499, 1499, 3001], 3000, 2)
        self.assertEqual(r['quantity'], 3)
        self.assertEqual(r['kerf_loss_mm'], 4)
        self.assertEqual(r['rejected_mm'], [3001])
        self.assertEqual(pack([1000], 1000, 5)['quantity'], 1)
        self.assertEqual(pack([999], 1000, 5)['rejected_mm'], [999])

    def test_random_kerf_conservation(self):
        rng = random.Random(183)
        for _ in range(100):
            parts = [rng.randrange(10, 3010) for _ in range(rng.randrange(1, 60))]
            self.assert_conservation(parts, 3000, 3, seconds=.01)

    def test_budget_does_not_claim_optimality(self):
        r = self.assert_conservation([600]*3, 1000, node_limit=0)
        self.assertFalse(r['optimal'])

    def test_global_steps_and_more_than_100_parts(self):
        budget = SearchBudget(300, 20)
        first = pack([600]*101, 1000, budget=budget)
        self.assertFalse(first['optimal'])
        self.assertEqual(first['stop_reason'], 'steps')
        self.assertEqual(budget.nodes, 20)
        second = pack([600]*3, 1000, budget=budget)
        self.assertEqual(budget.nodes, 20)
        self.assertFalse(second['optimal'])

    def test_global_time(self):
        budget = SearchBudget(1, 100000)
        now = [0]
        budget.clock = lambda: now[0]
        budget.started = 0
        budget.step()
        now[0] = 2
        result = pack([600]*3, 1000, budget=budget)
        self.assertEqual(result['stop_reason'], 'time')
        self.assertFalse(result['optimal'])

    def test_cancel_interrupts_inside_search(self):
        calls = []
        def cancel():
            calls.append(True)
            return True
        budget = SearchBudget(300, 1000000, cancel)
        with self.assertRaises(SearchCancelled):
            pack([600]*101, 1000, budget=budget)
        self.assertEqual(len(calls), 1)

    def test_allowance_applies_once_to_every_part(self):
        lengths = [500]*6
        r = pack(lengths, 3000, allowance_mm=50)
        self.assertEqual(lengths, [500]*6)
        self.assertEqual(pack(lengths, 3000)['quantity'], 1)
        self.assertEqual(r['quantity'], 2)
        self.assertEqual(r['required_mm'], 3300)
        self.assertEqual(r['allowance_total_mm'], 300)
        self.assertTrue(all(p['lengths_mm'] == [550]*len(p['lengths_mm']) for p in r['patterns']))

    def test_allowance_rejection_and_validation(self):
        r = pack([2950, 2960, 920], 3000, allowance_mm=50)
        self.assertEqual(r['rejected_indices'], [1])
        self.assertEqual(r['rejected_mm'], [3010])
        self.assertEqual(r['required_mm'], 3970)
        for v in [-1, float('nan'), float('inf')]:
            with self.assertRaises(ValueError):
                pack([500], 3000, allowance_mm=v)

    def test_invalid_inputs(self):
        for parts, stock, kerf in [([0], 3000, 0), ([1], 0, 0), ([1], 3, -1),
                                   ([float('nan')], 3, 0)]:
            with self.assertRaises(ValueError):
                pack(parts, stock, kerf)


if __name__ == '__main__':
    unittest.main()
