# -*- coding: utf-8 -*-
"""
Supabase Configuration Dialog
Configure Supabase connection settings
"""
__title__ = "Supabase Config"
__author__ = "T3Lab"

import os
import sys
import clr
import traceback

# .NET Imports
clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
from System.Windows import Window
from System.Windows.Markup import XamlReader

# pyRevit Imports
try:
    from pyrevit import script, forms
    logger = script.get_logger()
except:
    import logging
    logger = logging.getLogger(__name__)
    forms = None

# Import Supabase client
try:
    from lib.supabase_client import SupabaseClient
except:
    import supabase_client
    SupabaseClient = supabase_client.SupabaseClient


class SupabaseConfigDialog(Window):
    """Dialog for configuring Supabase connection"""

    def __init__(self):
        try:
            logger.info("Initializing SupabaseConfigDialog...")

            # Initialize the base Window class
            Window.__init__(self)

            # Load XAML
            xaml_path = os.path.join(os.path.dirname(__file__), 'SupabaseConfig.xaml')
            logger.info("Loading XAML from: {}".format(xaml_path))

            with open(xaml_path, 'r') as f:
                xaml_content = f.read()
            self.ui = XamlReader.Parse(xaml_content)

            # Set the parsed content
            self.Content = self.ui
            logger.info("XAML loaded successfully")

            # Get controls
            self.txt_url = self.ui.FindName('txt_url')
            self.txt_api_key = self.ui.FindName('txt_api_key')
            self.txt_bucket_name = self.ui.FindName('txt_bucket_name')
            self.txt_status = self.ui.FindName('txt_status')
            self.btn_test = self.ui.FindName('btn_test')
            self.btn_save = self.ui.FindName('btn_save')
            self.btn_cancel = self.ui.FindName('btn_cancel')

            # Wire up events
            self.btn_test.Click += self.test_clicked
            self.btn_save.Click += self.save_clicked
            self.btn_cancel.Click += self.cancel_clicked

            # Load existing config
            self.client = SupabaseClient()
            if self.client.url:
                self.txt_url.Text = self.client.url
            if self.client.api_key:
                self.txt_api_key.Text = self.client.api_key
            if self.client.bucket_name:
                self.txt_bucket_name.Text = self.client.bucket_name

            logger.info("SupabaseConfigDialog initialized successfully")

        except Exception as ex:
            logger.error("Critical error in SupabaseConfigDialog.__init__: {}".format(ex))
            logger.error(traceback.format_exc())
            raise

    def test_clicked(self, sender, e):
        """Test Supabase connection"""
        try:
            url = self.txt_url.Text.strip()
            api_key = self.txt_api_key.Text.strip()
            bucket_name = self.txt_bucket_name.Text.strip()

            if not url or not api_key:
                self.txt_status.Text = "Please enter both URL and API key"
                self.txt_status.Foreground = System.Windows.Media.Brushes.Red
                return

            self.txt_status.Text = "Testing connection..."
            self.txt_status.Foreground = System.Windows.Media.Brushes.Blue

            # Create temp client
            test_client = SupabaseClient(url=url, api_key=api_key, bucket_name=bucket_name)

            # Try to list storage files
            try:
                files = test_client.list_storage_files()
                self.txt_status.Text = "Connection successful! Found {} files in storage.".format(len(files))
                self.txt_status.Foreground = System.Windows.Media.Brushes.Green
                logger.info("Supabase connection test successful")
            except Exception as storage_ex:
                # If storage fails, try database query
                try:
                    families = test_client.get_all_families(limit=1)
                    self.txt_status.Text = "Database connection OK. Storage may need configuration."
                    self.txt_status.Foreground = System.Windows.Media.Brushes.Orange
                    logger.warning("Storage test failed but database OK: {}".format(storage_ex))
                except Exception as db_ex:
                    raise Exception("Database and storage tests failed")

        except Exception as ex:
            error_msg = str(ex)[:100]
            self.txt_status.Text = "Connection failed: {}".format(error_msg)
            self.txt_status.Foreground = System.Windows.Media.Brushes.Red
            logger.error("Supabase connection test failed: {}".format(ex))
            logger.error(traceback.format_exc())

    def save_clicked(self, sender, e):
        """Save configuration"""
        try:
            url = self.txt_url.Text.strip()
            api_key = self.txt_api_key.Text.strip()
            bucket_name = self.txt_bucket_name.Text.strip() or 'revit-family-free'

            if not url or not api_key:
                if forms:
                    forms.alert("Please enter both URL and API key", exitscript=False)
                self.txt_status.Text = "Please enter both URL and API key"
                self.txt_status.Foreground = System.Windows.Media.Brushes.Red
                return

            # Save config
            if self.client.save_config(url, api_key, bucket_name):
                self.txt_status.Text = "Configuration saved successfully!"
                self.txt_status.Foreground = System.Windows.Media.Brushes.Green
                logger.info("Supabase configuration saved")

                # Close dialog after short delay
                self.DialogResult = True
                self.Close()
            else:
                self.txt_status.Text = "Failed to save configuration"
                self.txt_status.Foreground = System.Windows.Media.Brushes.Red

        except Exception as ex:
            error_msg = str(ex)
            self.txt_status.Text = "Save failed: {}".format(error_msg[:50])
            self.txt_status.Foreground = System.Windows.Media.Brushes.Red
            logger.error("Failed to save Supabase config: {}".format(ex))
            logger.error(traceback.format_exc())

    def cancel_clicked(self, sender, e):
        """Cancel dialog"""
        self.DialogResult = False
        self.Close()


def show_supabase_config():
    """Show Supabase configuration dialog"""
    try:
        dialog = SupabaseConfigDialog()
        result = dialog.ShowDialog()
        return result
    except Exception as ex:
        logger.error("Error showing Supabase config dialog: {}".format(ex))
        logger.error(traceback.format_exc())
        if forms:
            forms.alert("Error: {}".format(ex))
        return False


if __name__ == '__main__':
    show_supabase_config()
