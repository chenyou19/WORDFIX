from __future__ import annotations

import unittest

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.shading import get_shading_action, get_shading_decision
from docx_fixer.xml_utils import qn


def make_shd(**attrs):
    shd = etree.Element(qn("shd"), nsmap={"w": W_NS})
    for name, value in attrs.items():
        shd.set(qn(name), value)
    return shd


class ShadingTests(unittest.TestCase):
    def test_c0c0c0_maps_to_gray(self):
        decision = get_shading_decision(make_shd(fill="C0C0C0"))
        self.assertEqual(decision["action"], "gray")
        self.assertEqual(decision["reason"], "gray hex darker than default gray")

    def test_bfbfbf_maps_to_gray(self):
        decision = get_shading_decision(make_shd(fill="BFBFBF"))
        self.assertEqual(decision["action"], "gray")
        self.assertEqual(decision["reason"], "gray hex darker than default gray")

    def test_a6a6a6_maps_to_gray(self):
        decision = get_shading_decision(make_shd(fill="A6A6A6"))
        self.assertEqual(decision["action"], "gray")
        self.assertEqual(decision["reason"], "gray hex darker than default gray")

    def test_808080_maps_to_gray(self):
        decision = get_shading_decision(make_shd(fill="808080"))
        self.assertEqual(decision["action"], "gray")
        self.assertEqual(decision["reason"], "gray hex darker than default gray")

    def test_f2f2f2_is_kept(self):
        decision = get_shading_decision(make_shd(fill="F2F2F2"))
        self.assertEqual(decision["action"], "keep")
        self.assertEqual(decision["reason"], "gray hex lighter/equal default gray")

    def test_d9d9d9_is_kept(self):
        decision = get_shading_decision(make_shd(fill="D9D9D9"))
        self.assertEqual(decision["action"], "keep")
        self.assertEqual(decision["reason"], "gray hex lighter/equal default gray")

    def test_theme_fill_without_fill_is_kept(self):
        decision = get_shading_decision(make_shd(themeFill="accent1"))
        self.assertEqual(decision["action"], "keep")
        self.assertEqual(decision["reason"], "theme color unresolved")

    def test_theme_color_without_fill_is_kept(self):
        decision = get_shading_decision(make_shd(themeColor="accent1"))
        self.assertEqual(decision["action"], "keep")
        self.assertEqual(decision["reason"], "theme color unresolved")

    def test_explicit_red_is_cleared(self):
        decision = get_shading_decision(make_shd(fill="FF0000"))
        self.assertEqual(decision["action"], "clear")
        self.assertEqual(decision["normalized_fill_hex"], "FF0000")
        self.assertEqual(decision["reason"], "explicit non-gray hex color")

    def test_explicit_yellow_is_cleared(self):
        self.assertEqual(get_shading_action(make_shd(fill="FFFF00")), "clear")

    def test_explicit_light_yellow_is_cleared(self):
        self.assertEqual(get_shading_action(make_shd(fill="FFEB9C")), "clear")

    def test_explicit_light_red_is_cleared(self):
        self.assertEqual(get_shading_action(make_shd(fill="FFC7CE")), "clear")

    def test_explicit_light_blue_is_cleared(self):
        self.assertEqual(get_shading_action(make_shd(fill="DDEBF7")), "clear")


if __name__ == "__main__":
    unittest.main()
