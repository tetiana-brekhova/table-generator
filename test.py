import unittest
import os
from main import TableCreator
from tkinter import Tk, Entry, Button


class TestTableCreator(unittest.TestCase):
    def setUp(self):
        self.app = TableCreator()

    def test_add_row(self):
        initial_rows = self.app.total_rows

        self.app.add_new_row()

        self.assertEqual(self.app.total_rows, initial_rows + 1)
        self.assertEqual(len(self.app.data_lst[-1]), self.app.total_columns)

    def test_add_column(self):
        initial_columns = self.app.total_columns

        self.app.add_new_column()

        self.assertEqual(self.app.total_columns, initial_columns + 1)
        self.assertEqual(len(self.app.data_lst), self.app.total_rows)
        self.assertEqual(len(self.app.data_lst[-1]), self.app.total_columns)

    def test_save_table(self):
        filename = "test_table.docx"

        self.app.save_table()
        self.assertTrue(os.path.exists(filename))

        os.remove(filename)

    def tearDown(self):
        self.app.destroy()


if __name__ == '__main__':
    unittest.main()
