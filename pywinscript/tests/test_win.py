"""
pywinscript.win test module

"""
import os
import sys
import unittest
module_path = os.path.dirname(os.path.dirname(__file__))
sys.path.append(module_path)
from win import create_folder, is_running, start


class TestCreateFolder(unittest.TestCase):

    EXISTING_DIR = os.path.join(os.path.dirname(__file__), 'testdir')
    NEW_DIR = os.path.join(EXISTING_DIR, 'new')

    def setUp(self):
        if not os.path.exists(self.EXISTING_DIR):
            raise WindowsError('"testdir" not found')

    def tearDown(self):
        try:
            os.rmdir(self.NEW_DIR)
        except WindowsError:
            pass

    def test_returns_existing_path(self):
        self.assertEqual(create_folder(self.EXISTING_DIR), self.EXISTING_DIR)

    def test_returns_new_path(self):
        self.assertEqual(create_folder(self.NEW_DIR), self.NEW_DIR)

    def test_raises_windowserror(self):
        with self.assertRaises(WindowsError):
            create_folder('link//path')


class TestIsRunning(unittest.TestCase):

    def test_returns_true(self):
        self.assertTrue(is_running('python.exe'))


class TestStart(unittest.TestCase):

    def test_raises_windowserror(self):
        with self.assertRaises(WindowsError):
            start('fakeprogram')

        with self.assertRaises(WindowsError):
            start('fakeprogram', 'fakepath')


if __name__ == '__main__':
    unittest.main(verbosity=2)
