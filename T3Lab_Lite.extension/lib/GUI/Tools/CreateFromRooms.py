# -*- coding: utf-8 -*-
"""
Create From Rooms

Create elements based on room boundaries and properties.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Create From Rooms"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
import os, clr, traceback, sys

#>>>>>>>>>> pyRevit
from pyrevit import forms # Needed for wpf import to work.

# Custom Imports
from GUI.forms         import my_WPF
from Snippets._convert import convert_internal_units

#>>>>>>>>>> .NET IMPORTS
clr.AddReference("System")
from System.Collections.Generic import List
import wpf

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#====================================================================================================
PATH_SCRIPT = os.path.dirname(__file__)

uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application
doc     = __revit__.ActiveUIDocument.Document

active_view_id      = doc.ActiveView.Id
active_view         = doc.GetElement(active_view_id)
active_view_level   = active_view.GenLevel

class ListItem:
    """Helper Class for displaying selected sheets in my custom GUI."""
    def __init__(self,  Name='Unnamed', element = None, checked = False):
        self.Name       = Name
        self.IsChecked  = checked
        self.element    = element

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================
class CreateFromRooms(my_WPF):
    selected_type = []
    offset         = 0

    def __init__(self, items, title       = '__title',
                              label       = "Select Type:" ,
                              button_name = 'Create',
                              version     = 'version= 1.0'):
        """Function to show GUI to a user to Select Type and offset for creating elements from Rooms
            :param items:           Dictionary of items {e.Name:e}
            :param title:           Title of the window.
            :param label:           Label that is displayed above ListBox
            :param button_name:     Text in Button
            :param version:         Version of the script for footer."""
        self.items       = items        #type: dict
        self.title       = title        #type: str
        self.label       = label        #type: str
        self.button_name = button_name  #type: str
        self.version     = version      #type: str

        #>>>>>>>>>> SET RESOURCES AND LOAD WPF
        super(CreateFromRooms, self).add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'CreateFromRooms.xaml')
        wpf.LoadComponent(self, path_xaml_file )

        self.update_UI()
        self.ShowDialog()

    def update_UI(self):
        """Function to make updates to UI based on provided arguments."""
        # UPDATE GUI ELEMENTS
        self.main_title.Text          = self.title
        self.text_label.Content       = self.label
        self.button_main.Content      = self.button_name
        self.footer_version.Text      = self.version

        self.items                    = self.generate_list_items()
        self.main_ListBox.ItemsSource = self.items


    def generate_list_items(self):
        """Function to create a ICollection to pass to ListBox in GUI"""

        list_of_items = List[type(ListItem())]()
        first = True
        for type_name, typ in sorted(self.items.items()):
            checked = True if first else False
            first   = False
            list_of_items.Add(ListItem(type_name, typ, checked))
        return list_of_items


    # ╔═╗╦ ╦╦  ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║ ╦║ ║║  ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝╚═╝╩  ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ GUI EVENTS
    #==================================================
    def text_filter_updated(self, sender, e):
        """Function to filter items in the main_ListBox."""
        filtered_list_of_items = List[type(ListItem())]()
        filter_keyword = self.textbox_filter.Text

        #RESTORE ORIGINAL LIST
        if not filter_keyword:
            self.main_ListBox.ItemsSource = self.items
            return

        # FILTER ITEMS
        for item in self.items:
            if filter_keyword.lower() in item.Name.lower():
                filtered_list_of_items.Add(item)
            else:
                if item.IsChecked:
                    item.IsChecked = False

        # UPDATE LIST OF ITEMS
        self.main_ListBox.ItemsSource = filtered_list_of_items


    def UIe_ItemChecked(self, sender, e):
        """Single Selection in ListBox"""
        filtered_list_of_items = List[type(ListItem())]()
        for item in self.main_ListBox.Items:
            item.IsChecked = True if item.Name == sender.Content.Text else False
            filtered_list_of_items.Add(item)
        self.main_ListBox.ItemsSource = filtered_list_of_items

    def NumberValidationTextBox(self, sender, e):
        """Ensure that only Numbers and decimals can be inserted"""
        import re
        regex = re.compile('[^0-9.]+')
        e.Handled = regex.match(e.Text)


    # ╔╗ ╦ ╦╔╦╗╔╦╗╔═╗╔╗╔╔═╗
    # ╠╩╗║ ║ ║  ║ ║ ║║║║╚═╗
    # ╚═╝╚═╝ ╩  ╩ ╚═╝╝╚╝╚═╝ BUTTONS
    #==================================================
    def button_close(self, sender, e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.Close()
        sys.exit()

    def button_run(self, sender, e):
        """Button to finilize selection"""
        self.textbox_filter.Text = ''
        self.Close()

        try:
            selected_items = [item.element for item in self.main_ListBox.ItemsSource if item.IsChecked]
            self.offset    = convert_internal_units(float(self.UI_offset.Text),
                                                        get_internal=True,
                                                        units = 'cm')

            # if not selected_items:
            #     forms.alert('Type was not selected. Please Try Again', title='Create from Rooms',exitscript=True)

            self.selected_type = selected_items[0]

        except:
            # print(traceback.format_exc())

            self.offset = 0
