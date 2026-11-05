"""Factory-end accounting, cut assignment and additional-only purchase quantities."""
import os
import sys
import unittest
from collections import Counter

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'lib'))
from duct_flanges import calculate, calculate_group
from support_cutting import SearchCancelled


def group(lengths, width=500, height=250, flange=20):
    return {'width_mm': width, 'height_mm': height, 'flange': flange,
            'parts': [{'id': i + 1, 'length_mm': length} for i, length in enumerate(lengths)]}


class FlangeTests(unittest.TestCase):
    def test_full_standard_includes_both_factory_ends(self):
        result = calculate([group([1250])])
        self.assertEqual(0, result['corners'])
        self.assertEqual(0, result['new_flanges'])
        self.assertEqual(0, result['rails'][0]['packing']['quantity'])

    def test_requested_1750_split(self):
        result = calculate([group([1250, 500])])
        self.assertEqual(1, result['new_flanges'])
        self.assertEqual(4, result['corners'])
        self.assertEqual(1500, result['rails'][0]['cut_rail_mm'])

    def test_two_shorts_from_opposite_ends(self):
        result = calculate_group(group([500, 500]))
        self.assertEqual(1, len(result['layouts']))
        self.assertEqual([1, 1], [p['factory_ends'] for p in result['parts']])
        self.assertEqual(2, result['new_flanges'])
        self.assertEqual(8, result['corners'])
        self.assertEqual(250, result['layouts'][0]['offcut_mm'])
        self.assertEqual(0, result['layouts'][0]['offcut_factory_ends'])

    def test_middle_piece_needs_two_new_flanges(self):
        result = calculate_group(group([400, 400, 400]))
        self.assertEqual([1, 0, 1], [p['factory_ends'] for p in result['parts']])
        self.assertEqual(4, result['new_flanges'])
        self.assertEqual(16, result['corners'])

    def test_full_stock_with_two_short_pieces(self):
        result = calculate_group(group([625, 625]))
        self.assertEqual(2, result['new_flanges'])
        self.assertEqual(0, result['layouts'][0]['offcut_mm'])

    def test_single_short_leaves_one_factory_end_on_offcut(self):
        result = calculate_group(group([300]))
        self.assertEqual(1, result['layouts'][0]['offcut_factory_ends'])
        self.assertEqual(1, result['new_flanges'])

    def test_every_source_accounted_for_once(self):
        original = group([1250, 300, 300, 300, 500, 500, 800])
        result = calculate_group(original)
        self.assertEqual(Counter(p['id'] for p in original['parts']), Counter(p['id'] for p in result['parts']))
        self.assertEqual(sum(p['length_mm'] for p in original['parts']), sum(p['length_mm'] for p in result['parts']))
        for layout in result['layouts']:
            self.assertLessEqual(sum(p['factory_ends'] for p in layout['parts']), 2)
            self.assertEqual(1250, sum(p['length_mm'] for p in layout['parts']) + layout['offcut_mm'])

    def test_known_end_modes_keep_whole_standard_free(self):
        source = group([1250, 500, 500])
        self.assertEqual(2, calculate_group(source, 'one_factory_end')['new_flanges'])
        self.assertEqual(4, calculate_group(source, 'no_factory_ends')['new_flanges'])

    def test_p20_selection_and_p30_fixed_corner(self):
        result = calculate([group([500]), group([500], 1200, 400, 30)], corner20=u'УГФ-1')
        self.assertEqual([u'УГФ-1', u'УГФ-2'], [r['corner'] for r in result['rails']])
        self.assertEqual([4, 4], [r['corners'] for r in result['rails']])

    def test_deduction_is_per_rail_not_per_flange(self):
        result = calculate([group([500])], deduction20_mm=20)
        self.assertEqual([480, 480, 230, 230], [p['length_mm'] for p in result['rails'][0]['cuts']])
        self.assertEqual(1420, result['rails'][0]['cut_rail_mm'])
        self.assertEqual(1500, result['rails'][0]['gross_rail_mm'])

    def test_incompatible_groups_do_not_share_duct_stock(self):
        result = calculate([group([500]), group([500], 600, 300)])
        self.assertEqual(2, sum(len(g['layouts']) for g in result['groups']))

    def test_rail_too_long_for_purchased_stock_is_reported(self):
        result = calculate([group([500], 1200, 400, 30)], stock30_mm=1000)
        self.assertEqual(2, len(result['rails'][0]['rejected']))
        self.assertEqual(4, result['corners'])

    def test_unusable_duct_due_to_kerf_is_not_silently_counted(self):
        result = calculate([group([1249])], duct_kerf_mm=3)
        self.assertEqual(1, len(result['duct_rejected']))
        self.assertEqual(0, result['new_flanges'])

    def test_bad_inputs(self):
        with self.assertRaises(ValueError):
            calculate([group([1251])])
        for kwargs in ({'stock20_mm': 0}, {'deduction20_mm': -1}, {'rail_kerf_mm': float('nan')},
                       {'deduction20_mm': 250}, {'corner20': u'УГФ-2'}):
            with self.assertRaises(ValueError):
                calculate([group([500])], **kwargs)

    def test_cancel_stops_before_producing_an_order(self):
        with self.assertRaises(SearchCancelled):
            calculate([group([500])], on_progress=lambda value, maximum: True)

    def test_no_demand_has_no_purchase(self):
        result = calculate([])
        self.assertEqual([], result['rails'])
        self.assertEqual(0, result['corners'])

    def test_duct_purchase_separates_whole_stocks_and_mounted_cut_parts(self):
        result = calculate([group([1250, 500, 500])])
        purchase = result['purchase']
        self.assertEqual(2, purchase['stock_quantity'])
        self.assertEqual(1, purchase['whole_installed'])
        self.assertEqual(1, purchase['cut_stock_quantity'])
        self.assertEqual(2, purchase['cut_installed'])
        self.assertEqual(3, purchase['installed_parts'])
        self.assertEqual(250, purchase['offcut_mm'])

    def test_middle_cuts_are_three_installed_parts_from_one_stock(self):
        purchase = calculate([group([400, 400, 400])])['purchase']
        self.assertEqual(1, purchase['stock_quantity'])
        self.assertEqual(1, purchase['cut_stock_quantity'])
        self.assertEqual(0, purchase['whole_installed'])
        self.assertEqual(3, purchase['cut_installed'])

    def test_whole_ducts_are_not_counted_as_cut_stock(self):
        purchase = calculate([group([1250, 1250])])['purchase']
        self.assertEqual(2, purchase['stock_quantity'])
        self.assertEqual(2, purchase['whole_installed'])
        self.assertEqual(0, purchase['cut_stock_quantity'])
        self.assertEqual(0, purchase['cut_installed'])

    def test_purchase_conserves_length_including_kerf_and_offcuts(self):
        result = calculate([group([1250, 400, 400, 400]), group([500, 500], 600, 300)], duct_kerf_mm=3)
        for purchase in [result['purchase']] + [g['purchase'] for g in result['groups']]:
            self.assertEqual(purchase['stock_quantity'], purchase['whole_installed'] + purchase['cut_stock_quantity'])
            self.assertEqual(purchase['installed_parts'], purchase['whole_installed'] + purchase['cut_installed'])
            self.assertAlmostEqual(purchase['purchased_length_mm'], purchase['installed_length_mm'] +
                                   purchase['offcut_mm'] + purchase['kerf_loss_mm'])

    def test_rejected_parts_are_excluded_from_duct_purchase(self):
        purchase = calculate([group([1250, 1249])], duct_kerf_mm=3)['purchase']
        self.assertEqual(1, purchase['stock_quantity'])
        self.assertEqual(1, purchase['whole_installed'])
        self.assertEqual(0, purchase['cut_stock_quantity'])
        self.assertEqual(0, purchase['cut_installed'])

    def test_known_ends_do_not_invent_a_stock_purchase(self):
        for mode in ('one_factory_end', 'no_factory_ends'):
            purchase = calculate([group([1250, 500, 500])], mode=mode)['purchase']
            self.assertIsNone(purchase['stock_quantity'])
            self.assertIsNone(purchase['cut_stock_quantity'])
            self.assertIsNone(purchase['offcut_mm'])
            self.assertEqual(1, purchase['whole_installed'])
            self.assertEqual(2, purchase['cut_installed'])


if __name__ == '__main__':
    unittest.main()
