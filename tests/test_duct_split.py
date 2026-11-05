"""Boundary and stock-length cases independent of Revit."""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'lib'))
from duct_split import segment_lengths, flange_size


class DuctSplitTests(unittest.TestCase):
    def test_requested_example(self):
        self.assertEqual([1250, 500], segment_lengths(1750))

    def test_exact_multiples_and_short_lengths(self):
        for length, expected in [(500, [500]), (1250, [1250]),
                                 (2500, [1250, 1250]), (5000, [1250] * 4),
                                 (4470, [1250, 1250, 1250, 720])]:
            self.assertEqual(expected, segment_lengths(length))

    def test_no_tiny_floating_point_remainder(self):
        lengths = segment_lengths(2500.0000001)
        self.assertEqual(2, len(lengths))
        self.assertAlmostEqual(2500.0000001, sum(lengths))

    def test_tiny_real_remainder_is_rejected(self):
        with self.assertRaises(ValueError):
            segment_lengths(2501)

    def test_minimum_revit_length(self):
        with self.assertRaises(ValueError):
            segment_lengths(1260, minimum_mm=20)

    def test_repeat_does_not_split_again(self):
        lengths = segment_lengths(4470)
        self.assertEqual(lengths, [part for length in lengths for part in segment_lengths(length)])

    def test_conservation_for_non_integer_lengths(self):
        for length in (1274.22815571, 2287.89321881, 10010.125, 5003):
            lengths = segment_lengths(length)
            self.assertAlmostEqual(length, sum(lengths))
            self.assertTrue(all(0 < part <= 1250 for part in lengths))

    def test_invalid_lengths(self):
        for value in (0, -1, float('inf'), float('nan')):
            with self.assertRaises(ValueError):
                segment_lengths(value)

    def test_flange_uses_largest_side_and_inclusive_boundary(self):
        self.assertEqual(20, flange_size(1000, 400))
        self.assertEqual(20, flange_size(400, 1000))
        self.assertEqual(30, flange_size(400, 1001))
        self.assertEqual(30, flange_size(1200, 300))


if __name__ == '__main__':
    unittest.main()
