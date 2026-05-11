import unittest
from unittest.mock import patch, MagicMock
import os
import sys

# Create a fake winreg module
fake_winreg = MagicMock()
fake_winreg.HKEY_LOCAL_MACHINE = 'HKEY_LOCAL_MACHINE'
sys.modules['winreg'] = fake_winreg

# Mock other problematic modules
class MockCTK:
    def CTkFont(*args, **kwargs): return MagicMock()
    def set_appearance_mode(*args, **kwargs): pass
    def set_default_color_theme(*args, **kwargs): pass
    class CTk: pass
    def CTkLabel(*args, **kwargs): return MagicMock()
    def CTkFrame(*args, **kwargs): return MagicMock()
    def CTkButton(*args, **kwargs): return MagicMock()
    def CTkImage(*args, **kwargs): return MagicMock()
    def CTkProgressBar(*args, **kwargs): return MagicMock()

sys.modules['customtkinter'] = MockCTK
sys.modules['PIL'] = MagicMock()
sys.modules['PIL.Image'] = MagicMock()

import main

class TestAppTecnico(unittest.TestCase):

    @patch('main.winreg.OpenKey')
    @patch('main.winreg.QueryInfoKey')
    @patch('main.winreg.EnumKey')
    @patch('main.winreg.QueryValueEx')
    def test_esta_instalado_true(self, mock_query_val, mock_enum, mock_info, mock_open_key):
        app = MagicMock()

        mock_info.return_value = [1, 0, 0] # One entry
        mock_enum.return_value = "ChromeSubKey"
        mock_query_val.return_value = ("Google Chrome", 1)

        result = main.AppTecnico.esta_instalado(app, "Chrome")

        self.assertTrue(result)

    @patch('main.winreg.OpenKey')
    def test_esta_instalado_registry_error(self, mock_open_key):
        app = MagicMock()
        mock_open_key.side_effect = OSError("Registry not accessible")

        result = main.AppTecnico.esta_instalado(app, "AnyApp")

        self.assertFalse(result)

    @patch('main.os.path.exists')
    def test_verificar_arquivos_hd(self, mock_exists):
        app = MagicMock()
        app.frame_checks = MagicMock()
        # Mocking the children destruction loop
        app.frame_checks.winfo_children.return_value = []
        app.fonte_negrito = MagicMock()

        mock_exists.side_effect = lambda path: "SDI" in path

        main.AppTecnico.verificar_arquivos_hd(app)

        # Verify that it checked for both
        self.assertTrue(mock_exists.called)

if __name__ == '__main__':
    unittest.main()
