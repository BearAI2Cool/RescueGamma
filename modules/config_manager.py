import os
import configparser
import shutil
from datetime import datetime
import json


class ConfigManager:
    def __init__(self):
        self.ini_path = os.path.join(os.getcwd(), 'setting.ini')
        self.config_files = [
            self.ini_path,
            os.path.join(os.getcwd(), 'config.json'),
            os.path.join(os.getcwd(), 'font_config.json')
        ]
        self.backup_dir = os.path.join(os.getcwd(), 'Backup')
        self.config = configparser.ConfigParser()
        self._initialize_ini()

    def _initialize_ini(self):
        if not os.path.exists(self.ini_path):
            self.config['Theme'] = {'dark_mode': 'False', 'sort_order': '0'}
            self.config['RecentFiles'] = {'last_ppt_path': ''}
            with open(self.ini_path, 'w', encoding='utf-8') as f:
                self.config.write(f)
        else:
            self.config.read(self.ini_path, encoding='utf-8')
            if not self.config.has_section('Theme'):
                self.config.add_section('Theme')
            if not self.config.has_option('Theme', 'sort_order'):
                self.config.set('Theme', 'sort_order', '0')
            if not self.config.has_option('Theme', 'dark_mode'):
                self.config.set('Theme', 'dark_mode', 'False')
            with open(self.ini_path, 'w', encoding='utf-8') as f:
                self.config.write(f)

    def get_dark_mode(self):
        return self.config.getboolean('Theme', 'dark_mode', fallback=False)

    def set_dark_mode(self, is_dark):
        self.config['Theme']['dark_mode'] = str(is_dark)
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get_last_ppt_path(self):
        return self.config.get('RecentFiles', 'last_ppt_path', fallback='')

    def set_last_ppt_path(self, path):
        self.config['RecentFiles']['last_ppt_path'] = path
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def restore_default_settings(self):
        os.makedirs(self.backup_dir, exist_ok=True)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        for file_path in self.config_files:
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                name, ext = os.path.splitext(file_name)
                backup_name = f'{name}_old_{timestamp}{ext}'
                backup_path = os.path.join(self.backup_dir, backup_name)
                shutil.copy2(file_path, backup_path)

        for file_path in self.config_files:
            if os.path.exists(file_path):
                if file_path.endswith('.ini'):
                    self.config['Theme'] = {'dark_mode': 'False'}
                    self.config['RecentFiles'] = {'last_ppt_path': ''}
                    with open(file_path, 'w', encoding='utf-8') as f:
                        self.config.write(f)
                else:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump({}, f)

        return True

    def get_sort_order(self):
        return self.config.getint('Theme', 'sort_order', fallback=0)

    def set_sort_order(self, order):
        if not self.config.has_section('Theme'):
            self.config.add_section('Theme')
        self.config.set('Theme', 'sort_order', str(order))
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)
