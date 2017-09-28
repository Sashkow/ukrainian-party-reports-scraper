import unittest
from unittest import TestCase

from main import *


class TestGetYesNames(TestCase):

    def test_is_row_empty(self):
        xlsx = 'test_data/is_row_empty_test.xlsx'
        table = xlrd.open_workbook(xlsx).sheet_by_index(0)

        self.assertEqual(str(table.row(0)), "[text:'some value']" )
        self.assertFalse(is_row_empty(table.row(0)))

        self.assertEqual(str(table.row(1)),"[text:' ']")
        self.assertFalse(is_row_empty(table.row(1)))

        self.assertEqual(str(table.row(2)),"[empty:'']")
        self.assertTrue(is_row_empty(table.row(2)))

    def test_find_table_start(self):
        raw_xls = 'test_data/bpp_general.xlsx'
        title = '(1.1. Відомості про нерухоме майно){e<=10}'
        table = xlrd.open_workbook(raw_xls).sheet_by_index(0)
        self.assertEqual(find_table_start(table, title), 234)

    def test_find_table_end(self):
        raw_xls = 'test_data/bpp_general.xlsx'
        title = '(1.1. Відомості про нерухоме майно){e<=10}'
        table = xlrd.open_workbook(raw_xls).sheet_by_index(0)
        start = find_table_start(table, title)
        end_text = '(Загальна сума){e<=5}'
        self.assertEqual(find_table_end(table, start, end_text), 250)

    def test_get_table(self):
        raw_xls = 'test_data/bpp_general.xlsx'
        title = '(1.1. Відомості про нерухоме майно){e<=10}'
        end_text = '(Загальна сума){e<=5}'

        self.assertEqual(get_table_start_end(raw_xls, title, end_text)[1], 234)
        self.assertEqual(get_table_start_end(raw_xls, title, end_text)[2], 250)

    def test_table_to_answer_table(self):
        raw_xls = 'test_data/bpp_general.xlsx'
        table = xlrd.open_workbook(raw_xls).sheet_by_index(0)
        answer_workbook_path = 'test_data/answers/додаток I.1 - майно та активи (у власності)_IV-kv_parl.xlsx'
        sheet_name = 'I.1.1 - нерухомість'

        
        table_to_answer_table(table, 234, 250, answer_workbook_path, sheet_name)












if __name__ == '__main__':
    unittest.main()


self.assertEqual(find_table_end(table, start, end_text)[1], 234)