# -*- coding: utf-8 -*-
__title__ = "PAM Building\nConverter"
__author__ = "PAM Building UK"
__doc__ = """Version = 1.0
Date    = 01.11.2024
__________________________________________________________________
Compatibility:

Revit 2023+
__________________________________________________________________
Description:

Converts our generic Revit content to our full data families
__________________________________________________________________
How-to:

-> Click the button
-> Select coupling type and pipe type to convert to
-> Click 'Select Pipework' button
-> Select all elements that you wish to convert and press 'Finish'
__________________________________________________________________
"""

# Import required classes and add references to required libraries
import os
import clr
import webbrowser

# Add references to the necessary assemblies
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')

import System
from System.Windows.Controls import Button, ListBox, TextBox
from System.Windows.Input import MouseButtonState

from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit
from pyrevit.forms import WPFWindow

#import Autodesk
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB import *

from pyrevit import forms, script
#from rpw import revit

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference("RevitNodes")

import Revit
clr.ImportExtensions(Revit.Elements)

# Variables
doc = revit.doc
uidoc = revit.uidoc

output = script.get_output()

# Custom class to represent an item in the ListBox
class PipeTypeItem:
    """Class to represent a pipe type item in the ListBox."""
    def __init__(self, element, type_name):
        self.Element = element
        self.IsChecked = False
        self.Name = type_name

def get_dict_of_elements(built_in_category):
    elements = FilteredElementCollector(doc).OfCategory(built_in_category).WhereElementIsElementType().ToElements()
    return {Element.Name.GetValue(e): e for e in elements}

# Get pipe types
pipe_types_dict = get_dict_of_elements(BuiltInCategory.OST_PipeCurves)

# Function to get family type IDs with added support for single-type families
def get_family_type_ids(doc, family_name, type_names):
    type_ids = {}
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for symbol in collector:
        if symbol.FamilyName == family_name:
            symbol_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            # Add type if it matches one in type_names or if there’s only one type available
            if not type_names or symbol_name in type_names:
                type_ids[symbol_name] = symbol.Id
                # Handle single type families by setting the default ID
                if len(type_ids) == 1 and not type_names:
                    return symbol.Id
    return type_ids

# Convert the dictionary items into PipeTypeItem objects and sort them by Name
family_items = sorted(
    [PipeTypeItem(element, type_name) for type_name, element in pipe_types_dict.items()],
    key=lambda item: item.Name.lower()  # Sort case-insensitively by Name
)

# Sort the pipe type items by Name
#sorted_family_items = sorted(family_items, key=lambda item: item.Name)

def show_window():
    """Load and display the WPF window for user interaction."""
    # Path to your XAML file relative to the script directory
    script_dir = os.path.dirname(__file__)
    xaml_file_path = os.path.join(script_dir, 'UI.xaml')

    # Load the WPF window from the XAML file
    window = WPFWindow(xaml_file_path)

    selected_coupling = [None]  # To store the selected coupling type
    selected_pipe_type = []  # To store the selected pipe types

    def close_button_click(sender, args):
        """Handle the close button click event."""
        window.Close()

    def run_button_click(sender, args):
        """Handle the Place Couplings button click event."""
        list_box = window.FindName('list_pipetypes')
        if list_box:
            selected_pipe_type[:] = [item for item in list_box.Items if isinstance(item, PipeTypeItem) and item.IsChecked]
            if not selected_pipe_type:
                forms.alert('Please select pipe type', title='Select Pipe Type')
                return

        # Check which radio button is selected
        if window.EC002.IsChecked:
            selected_coupling[0] = "EC002 - Ductile Iron Coupling"
        elif window.EC002NG.IsChecked:
            selected_coupling[0] = "EC002NG - RAPID S NG Coupling"

        if selected_coupling[0] and selected_pipe_type:
            # Close the window and return the selected coupling and pipe types
            window.DialogResult = True
            window.Tag = (selected_coupling[0], selected_pipe_type) #It’s crucial to match the order of data when retrieving it later. So coupling first, then pipe type!!!
            window.Close()

        else:
            forms.alert('Please select a coupling type and pipe types.', title='Select Coupling and Pipe Types')

    def on_image_click(sender, event_args):
        """Handle the image click event to open a URL."""
        url = "https://www.pambuilding.co.uk/"
        webbrowser.open(url)

    def header_drag(sender, event_args):
        """Allow dragging the window by the header."""
        if event_args.LeftButton == MouseButtonState.Pressed:
            window.DragMove()
    
    def UI_text_filter_updated(sender, args):
        """Handle TextBox TextChanged event to filter ListBox items."""
        filter_text = sender.Text.lower()
        list_box = window.FindName('list_pipetypes')
        list_box.Items.Clear()
        for item in family_items:
            if filter_text in item.Name.lower():
                list_box.Items.Add(item)

    # Attach event handlers
    close_button = window.FindName('button_close')
    if close_button and isinstance(close_button, Button):
        close_button.Click += close_button_click

    run_button = window.FindName('button_run')
    if run_button and isinstance(run_button, Button):
        run_button.Click += run_button_click

    # Attach the image click event handler
    logo_image = window.FindName('logo')
    if logo_image:
        logo_image.MouseLeftButtonDown += on_image_click

    # Find the TitleBar and attach the drag event handler
    title_bar = window.FindName('TitleBar')
    if title_bar:
        title_bar.MouseLeftButtonDown += header_drag
    
    # Attach the TextChanged event handler to the TextBox in code
    text_box = window.FindName('textbox_filter')
    if text_box and isinstance(text_box, TextBox):
        text_box.TextChanged += UI_text_filter_updated
    
    # Populate ListBox with sorted pipe types
    list_box = window.FindName('list_pipetypes')
    if list_box and isinstance(list_box, ListBox):
        for family_item in family_items:
            list_box.Items.Add(family_item)

    # Show the window
    window.ShowDialog()

    # Retrieve and return selected coupling and pipe types
    if window.DialogResult:
        return window.Tag
    else:
        return None, None

# Show the custom window and get the selected coupling and pipe types
selected_coupling, selected_pipe_type = show_window()

if not selected_coupling or not selected_pipe_type:
    script.exit()

# Define a selection filter to allow the user to select only pipes, pipe fittings, pipe accessories, and plumbing fixtures
class CategorySelectionFilter(ISelectionFilter):
    def __init__(self, category_names):
        self.category_names = category_names

    def AllowElement(self, e):
        if e.Category.Name in self.category_names:
            return True
        else:
            return False

    def AllowReference(self, ref, point):
        return True

# Use the pipe type selected from the WPF window
pipe_type_name = selected_pipe_type[0].Name  # Because only one pipe type is selected
pipe_type = pipe_types_dict[pipe_type_name]
#rpm = pipe_type.RoutingPreferenceManager
#rc = RoutingConditions(RoutingPreferenceErrorLevel.None)

# Clear unused elements
del pipe_types_dict

# Main logic
try:
    # Picking elements
    with forms.WarningBar(title="Select pipework and press Finish when complete"):
        selected_elements = uidoc.Selection.PickObjects(
            ObjectType.Element, 
            CategorySelectionFilter(["Pipes", "Pipe Fittings", "Pipe Accessories"]), 
            'Select Pipework'
        )

    # Check if no elements were selected
    if not selected_elements:
        forms.alert('No elements have been selected', title='Select Pipework')
        script.exit()

    # Start the transaction for changing pipes and fittings
    transaction = Transaction(doc, 'Convert Pipework to PAM Building')
    transaction.Start()

    num_pipes_changed = 0  # Counter for pipes changed
    num_fittings_changed = 0  # Counter for fittings and accessories changed

    # Loop through picked elements to change pipe types
    for element in selected_elements:
        try:
            element = doc.GetElement(element)
            # Process pipes
            if element.Category.Name == "Pipes":
                element.ChangeTypeId(pipe_type.Id)
                num_pipes_changed += 1  # Increment the counter
            # Process pipe fittings and pipe accessories
            else:
                element_type = doc.GetElement(element.GetTypeId())
                description_param = element_type.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString()
                #family_name = element_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
                type_name = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()

                # Initialize family_name
                family_name = ""

                # Check if the description contains certain substrings
                #Pipe Fittings
                if "45° Single Long Arm Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_45° Single Long Arm Branch_EF008_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_45° Single Long Arm Branch_EF008_NG"
                elif "88° Long Radius Door Back Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_88° Long Radius Door Back Bend_EF05L_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_88° Long Radius Door Back Bend_EF05L_NG"
                elif "88° Medium Radius Door Back Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_88° Medium Radius Door Back Bend_EF05M_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_88° Medium Radius Door Back Bend_EF05M_NG"
                elif "88º Short Radius Door Back Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_88° Short Radius Door Back Bend_EF005_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_88° Short Radius Door Back Bend_EF005_NG"
                elif "88° Vented Bend Axial" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_88° Vented Bend_EF002AB_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_88° Vented Bend_EF002AB_NG"
                elif "88° Vented Bend Radial" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_88° Vented Bend_EF002RB_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_88° Vented Bend_EF002RB_NG"
                elif "Access Pipe Rectangular Door" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Access Pipe Rect Door_EF015_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Access Pipe Rect Door_EF015_NG" 
                elif "Access Pipe Round Door" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Access Pipe Round Door_AF014_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Access Pipe Round Door_AF014_NG"
                elif "Air~Wave Vent Cowl" in description_param:
                        family_name = "SGPAMUK_ES_Air Wave Vent Cowl_EF075"
                elif "Blank End Drilled and Taped" in description_param:
                        family_name = "SGPAMUK_ES_Blank End Drilled And Taped_EF071T"
                elif "Blank End Push-Fit Connection" in description_param:
                        family_name = "SGPAMUK_ES_Blank End Push Fit Connection_EF077"
                elif "Blank End Push-Fit" in description_param:
                        family_name = "SGPAMUK_ES_Blank End Push Fit_EF071"
                elif "Blank End" in description_param:
                    family_name = "SGPAMUK_ES_Blank End_EF070"
                elif "Corner Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Corner Branch_EF035_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Corner Branch_EF035_NG"
                elif "Corner Radius Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Corner Radius Branch_EF035R_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Corner Radius Branch_EF035R_NG"
                elif "Long Tail Double Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Bend_EF054_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Bend_EF054_NG"
                elif "Double Boss with Bosses Opposed at 88º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_AF091_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_AF091_NG"
                elif "Double Boss with Bosses at 90º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_AF092_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_AF092_NG"
                elif "Double Boss with Drilled/Tapped 50mm Bosses Opposed at 88º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_EF091T_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_EF091T_NG"
                elif "Double Boss with Drilled/Tapped 50mm Bosses at 90º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_EF092T_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Boss_EF092T_NG"
                elif "Double Branch Long Tail Radius Curve" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Branch Long Tail Radius Curve_EF097_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Branch Long Tail Radius Curve_EF097_NG"
                elif "Double Radius Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Branch Radius Curve_AF010R_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Branch Radius Curve_AF010R_NG"
                elif "Double Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Double Branch_AF010_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Double Branch_AF010_NG"
                elif "Expansion Plug" in description_param:
                    family_name = "SGPAMUK_ES_Expansion Plug_EF074"
                elif "Long Radius Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Radius Bend_EF02L_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Radius Bend_EF02L_NG"
                elif "Long Tail Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Bend_EF055_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Bend_EF055_NG"
                elif "Long Tail Corner Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Corner Branch_EF036_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Corner Branch_EF036_NG"
                elif "Long Tail Double Boss with Bosses Opposed at 88º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Double Boss_EF091LT_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Double Boss_EF091LT_NG"
                elif "Long Tail Double Boss with Bosses at 90º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Double Boss_EF092LT_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Double Boss_EF092LT_NG"
                elif "Long Tail Single Boss at 88º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Single Boss_EF090LT_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Single Boss_EF090LT_NG"
                elif "Long Tail Single Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Single Branch_EF056_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Long Tail Single Branch_EF056_NG"
                elif "Manifold Connector" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Manifold Connector_EF094_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Manifold Connector_EF094_NG"
                elif "Corner Multi-Waste Manifold Connector" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Manifold Connector Corner_EF099_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Manifold Connector Corner_EF099_NG"
                elif "Multi-Waste Manifold Connector" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Manifold Connector_EF095_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Manifold Connector_EF095_NG"
                elif "Push-Fit Movement Connector" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Movement Connector_EF058_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Movement Connector_EF058_NG"

                elif "Rodding Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Rodding Branch_EF009_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Rodding Branch_EF009_NG"

                elif "Short Radius Bend" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Bend_AF002_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Bend_AF002_NG"

                elif "Single Boss at 88º" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Single Boss_AF090_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Single Boss_AF090_NG"
                elif "Single Boss with Drilled/Tapped 50mm Boss Connection" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Single Boss_EF090T_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Single Boss_EF090T_NG"
                elif "Single Branch Long Tail Radius Curve" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Single Branch Long Tail Radius Curve_EF096_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Single Branch Long Tail Radius Curve_EF096_NG"
                elif "Single Branch with Radius Curve" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Single Branch Radius Curve_AF06R_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Single Branch Radius Curve_AF06R_NG"
                elif "Single Branch with Access Radius Curve" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Single Branch With Access Radius Curve_EF07R_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Single Branch With Access Radius Curve_EF07R_NG"

                elif "Single Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Single Branch_AF006_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Single Branch_AF006_NG"

                elif "Stack Support Pipe" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Stack Support Pipe_EF050 & EF051_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Stack Support Pipe_EF050 & EF051_NG"
                elif "Stench Trap" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Stench Trap_EF081_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Stench Trap_EF081_NG"
                elif "Taper Pipe" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Taper Pipe_EF028_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Taper Pipe_EF028_NG"
                elif "Transitional Connector" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Transitional Connector_EF059_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Transitional Connector_EF059_NG"
                elif "Universal Connector" in description_param:
                    family_name = "SGPAMUK_ES_Universal Connector_EF071R"
                elif "Entry/Terminal Venting Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Venting Branch Entry-Terminal_EF013_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Venting Branch Entry-Terminal_EF013_NG"
                elif "Interconnecting Venting Branch" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Venting Branch Interconnecting_EF013_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Venting Branch Interconnecting_EF013_NG"                        

                #elif "Metallic Coupling" in description_param:
                    #if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        #family_name = "SGPAMUK_ES_Two-Piece Ductile Iron Coupling_EC002_Union"
                    #elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        #family_name = "SGPAMUK_ES_RAPID S NG Coupling_EC002NG_Union"    
                
                elif "EN 12056 Calculation Connector" in description_param:
                    family_name = "SGPAMUK_ES_EN 12056 Calculation Connector"
                #Pipe Accessories
                elif "Offset" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Offset_EF024_DI"                                               
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Offset_EF024_NG"                        
                elif "Roof Connector for Asphalts" in description_param:
                    family_name = "SGPAMUK_ES_Roof Connector For Asphalts_EF073"
                elif "Roof Connector for Asphalts" in description_param:
                    family_name = "SGPAMUK_ES_Roof Penetration Flange with Gasket_EF079"
                elif "Strap-On Boss" in description_param:
                    family_name = "SGPAMUK_ES_Strap-On-Boss_EF133"
                elif "Branch Trap" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Trap Branch_EF080_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Trap Branch_EF080_NG" 
                elif "Trap Plain" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Trap Plain Branch_EF034_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Trap Plain Branch_EF034_NG" 
                elif "Trap Plain With Access Bottom" in description_param:
                    if selected_coupling == 'EC002 - Ductile Iron Coupling':
                        family_name = "SGPAMUK_ES_Trap Plain With Access Bottom_EF037_DI"
                    elif selected_coupling == 'EC002NG - RAPID S NG Coupling':
                        family_name = "SGPAMUK_ES_Trap Plain With Access Bottom_EF037_NG" 

                # Change the type 
                if family_name:
                    type_ids = get_family_type_ids(doc, family_name, [type_name])
                    if isinstance(type_ids, ElementId):
                        # Single type family case: Directly change to this type
                        element.ChangeTypeId(type_ids)
                    elif type_name in type_ids:
                        # Multiple type family case
                        element.ChangeTypeId(type_ids[type_name])

                    num_fittings_changed += 1 # Increment the counter

        except Exception as e:
            # Handle any exceptions that occur during the loop
            print("Error processing element: {} - {}".format(element.Id, str(e)))
            continue  

    doc.Regenerate()

    # Commit the transaction for changing pipes and fittings
    transaction.Commit()

except OperationCanceledException:
    # Show alert if the user cancels the selection or presses ESC
    forms.alert('User cancelled selection', title='Select Pipework')
    script.exit()  # Exit the script to prevent further processing

# Output message with the count of pipes and fittings changed
output_message = "Congratulations, you changed {} pipes.".format(num_pipes_changed)
output_message += " You also changed {} pipe fittings and accessories.".format(num_fittings_changed)
print(output_message)
