import sys, os

from PySide6.QtCore import QSize, Qt, Slot, QThreadPool
from PySide6.QtGui import QPixmap, QTextCursor, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QLabel,
    QLineEdit,
    QMainWindow,
    QVBoxLayout,
    QGridLayout,
    QWidget,
    QPushButton,
    QDialog,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QRadioButton,
    QTextEdit,
    QLayout
)

from PIL import Image, ImageGrab, ImageDraw
from time import sleep
# pip install pywin32 to get win32gui
import win32gui
import win32.lib.win32con as win32con

import pyautogui

import pytesseract

import numpy
# pip install opencv-python
import cv2

import random

# Regex
import re

# Version 1.2.2, as newer version didn't work
# pip install playsound==1.2.2
from playsound import playsound

import keyboard
from datetime import datetime
import traceback

# Get the containing folder of __file__ which holds the full path of the current Python file
basedir = os.path.dirname(__file__)

# Bundling data folders
try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'mycompany.myproduct.subproduct.version'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

# TODO: Look into Tesserocr as it is potentially faster?
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# pytesseract.pytesseract.tesseract_cmd = r'Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = os.path.join(basedir, "Tesseract-OCR", "tesseract.exe")

# CONSTANTS
EQUIPMENT_TYPES = ["---", "Weapon", "Headwear", "Bodywear", "Legwear"]
WEAPON_HIDDEN_POWERS = ["ATK", 
                        "HP", 
                        "DEF", 
                        "CRIT Rate", 
                        "CRIT Damage", 
                        "Fire Elemental ATK", 
                        "Water Elemental ATK", 
                        "Wind Elemental ATK", 
                        "Earth Elemental ATK", 
                        "Light Elemental ATK", 
                        "Dark Elemental ATK"
                        ]
HEADWEAR_HIDDEN_POWERS = ["ATK", 
                        "HP", 
                        "DEF", 
                        "CRIT Rate", 
                        "CRIT Damage", 
                        "Fire DMG", 
                        "Water DMG", 
                        "Wind DMG", 
                        "Earth DMG", 
                        "Light DMG", 
                        "Dark DMG"
                        ]
BODYWEAR_HIDDEN_POWERS = ["ATK", 
                        "HP", 
                        "DEF", 
                        "CRIT Rate", 
                        "CRIT Damage", 
                        "Buff AMP", 
                        "Debuff AMP", 
                        "Healing", 
                        "Shield", 
                        "DEF Penetration", 
                        "Damage Reduction"
                        ]
LEGWEAR_HIDDEN_POWERS = ["ATK", 
                        "HP", 
                        "DEF", 
                        "CRIT Rate", 
                        "CRIT Damage", 
                        "Buff AMP", 
                        "Debuff AMP", 
                        "Healing", 
                        "Shield", 
                        "DEF Penetration", 
                        "Damage Reduction"
                        ]

#
RARITY_VALUES_BASE = {"Common": (11, 12), "Uncommon": (13, 14), "Rare": (15, 16), "Epic": (17, 18), "Super Epic": (19, 20)}
RARITY_VALUES_CRIT_RATE = {"Common": (8.25, 9), "Uncommon": (9.75, 10.5), "Rare": (11.25, 12), "Epic": (12.75, 13.5), "Super Epic": (14.25, 15)}
RARITY_VALUES_CRIT_DAMAGE = {"Common": (13.75, 15), "Uncommon": (16.25, 17.5), "Rare": (18.75, 20), "Epic": (21.25, 22.5), "Super Epic": (23.75, 25)}
RARITY_VALUES_ELE_ATK = {"Common": (35, 40), "Uncommon": (45, 50), "Rare": (55, 60), "Epic": (65, 70), "Super Epic": (75, 80)}
RARITY_VALUES_AMP = {"Common": (5.5, 6), "Uncommon": (6.5, 7), "Rare": (7.5, 8), "Epic": (8.5, 9), "Super Epic": (9.5, 10)}
RARITY_VALUES_HEAL_SHIELD = {"Common": (8.8, 9.6), "Uncommon": (10.4, 11.2), "Rare": (12, 12.8), "Epic": (13.6, 14.4), "Super Epic": (15.2, 16)}
RARITY_VALUES_PEN_RED = {"Common": (4.4, 4.8), "Uncommon": (5.2, 5.6), "Rare": (6, 6.4), "Epic": (6.8, 7.2), "Super Epic": (7.6, 8)}

HIDDEN_POWER_RARITY_VALUES = {
    "ATK": RARITY_VALUES_BASE,
    "HP": RARITY_VALUES_BASE,
    "DEF": RARITY_VALUES_BASE,
    "CRIT Rate": RARITY_VALUES_CRIT_RATE,
    "CRIT Damage": RARITY_VALUES_CRIT_DAMAGE,
    "Fire Elemental ATK": RARITY_VALUES_ELE_ATK, 
    "Water Elemental ATK": RARITY_VALUES_ELE_ATK, 
    "Wind Elemental ATK": RARITY_VALUES_ELE_ATK, 
    "Earth Elemental ATK": RARITY_VALUES_ELE_ATK, 
    "Light Elemental ATK": RARITY_VALUES_ELE_ATK, 
    "Dark Elemental ATK": RARITY_VALUES_ELE_ATK,
    "Fire DMG": RARITY_VALUES_CRIT_RATE, 
    "Water DMG": RARITY_VALUES_CRIT_RATE, 
    "Wind DMG": RARITY_VALUES_CRIT_RATE, 
    "Earth DMG": RARITY_VALUES_CRIT_RATE, 
    "Light DMG": RARITY_VALUES_CRIT_RATE, 
    "Dark DMG": RARITY_VALUES_CRIT_RATE,
    "Buff AMP": RARITY_VALUES_AMP,
    "Debuff AMP": RARITY_VALUES_AMP, 
    "Healing": RARITY_VALUES_HEAL_SHIELD, 
    "Shield": RARITY_VALUES_HEAL_SHIELD, 
    "DEF Penetration": RARITY_VALUES_PEN_RED, 
    "Damage Reduction": RARITY_VALUES_PEN_RED
}

# ENTITY DIMENSIONS
# Given standard fullscreen dimensions of 1920 x 1080:

BBOX_BASE = (0, 0, 1920, 1080)

# Screen 1 -
# Change Power Button - (1215, 899, 390, 135)
S1_CHANGE_POWER_BUTTON = (1215, 899, 1215 + 390, 899 + 135)
S1_HIDDEN_POWER_LEVEL = (1112, 653, 1112 + 38, 653 + 48)

# These are currently not used for anything
S2_OLD_RARITY_1 = (285, 509, 285 + 198, 509 + 78)
S2_OLD_STAT_1 = (480, 509, 408 + 400, 509 + 78)
S2_OLD_RARITY_2 = (285, 602, 285 + 198, 602 + 78)
S2_OLD_STAT_2 = (480, 602, 480 + 440, 602 + 78)

S2_NEW_RARITY_1 = (1040, 509, 1040 + 198, 509 + 78)
S2_NEW_STAT_1 = (1236, 509, 1236 + 400, 509 + 78)
S2_NEW_RARITY_2 = (1040, 603, 1040 + 198, 603 + 78)
S2_NEW_STAT_2 = (1236, 603, 1236 + 400, 603 + 78)

S2_NEW_LOCK_BUTTON_1 = (1569, 519, 1569 + 63, 519 + 63)
S2_NEW_LOCK_BUTTON_2 = (1569, 609, 1569 + 63, 609 + 63)

S2_NEW_STAT_BOTH = (
    S2_NEW_STAT_1[0], S2_NEW_STAT_1[1], S2_NEW_STAT_2[2], S2_NEW_STAT_2[3])
S2_NEW_STAT_BOTH_WITH_LOCK = (
    S2_NEW_STAT_1[0], S2_NEW_STAT_1[1], S2_NEW_STAT_2[2] - 68, S2_NEW_STAT_2[3])

S2_KEEP_BUTTON = (386, 717, 386 + 390, 717 + 120)
S2_APPLY_NEW_BUTTON = (1142, 717, 386 + 390, 717 + 120)

# Going to offset this to the left half of the button to reduce chances of clicking the "Apply New" button which is below it
S3_CHANGE_POWER_KEEP = (973, 726, 973 + (362/2), 726 + 124)
    
# METHODS
# Given (x1, y1, x2, y2), return (x, y) that is the center
def get_center_of_coords(coords):
    return (coords[0] + (coords[2] - coords[0]) / 2), (coords[1] + (coords[3] - coords[1]) / 2)

# For entities that need to be clicked, get new point to click based on bbox scalings relative to base (1920 x 1080)
def get_center_of_entity_coords_based_on_window_offset(coords, bbox, randOffset = False):
    # With the two points, get the center for clicking purposes
    center_of_coords = get_center_of_coords(coords)
    
    # Random offset for pseudo-anti-detection purposes
    if randOffset:
        center_of_coords = (center_of_coords[0] + random.randint(-20, 20), center_of_coords[1] + random.randint(-20, 20))
    
    coord_to_bbox_ratios = (center_of_coords[0] / BBOX_BASE[2], center_of_coords[1] / BBOX_BASE[3])
    
    updated_x = bbox[0] + (coord_to_bbox_ratios[0] * (bbox[2] - bbox[0]))
    updated_y = bbox[1] + (coord_to_bbox_ratios[1] * (bbox[3] - bbox[1]))
    
    return (updated_x, updated_y)

def get_modified_entity_dimensions(coords, bbox):
    coord_to_bbox_ratios_1 = (coords[0] / BBOX_BASE[2], coords[1] / BBOX_BASE[3])
    coord_to_bbox_ratios_2 = (coords[2] / BBOX_BASE[2], coords[3] / BBOX_BASE[3])
    
    updated_x1 = bbox[0] + (coord_to_bbox_ratios_1[0] * (bbox[2] - bbox[0]))
    updated_y1 = bbox[1] + (coord_to_bbox_ratios_1[1] * (bbox[3] - bbox[1]))
    
    updated_x2 = bbox[0] + (coord_to_bbox_ratios_2[0] * (bbox[2] - bbox[0]))
    updated_y2 = bbox[1] + (coord_to_bbox_ratios_2[1] * (bbox[3] - bbox[1]))
    
    return (updated_x1, updated_y1, updated_x2, updated_y2)
    
# A pyautogui.click() but with a slight delay to increase chance for game to register it
def click(position):
    pyautogui.moveTo(position)
    sleep(0.1)
    pyautogui.mouseDown(x = position[0], y = position[1])
    sleep(0.1)
    pyautogui.mouseUp(x = position[0], y = position[1])
    
# https://stackoverflow.com/a/61730849
def get_dominant_color(img, palette_size=16):
    # Reduce colors (uses k-means internally)
    paletted = img.convert('P', palette=Image.ADAPTIVE, colors=palette_size)

    # Find the color that occurs most often
    palette = paletted.getpalette()
    color_counts = sorted(paletted.getcolors(), reverse=True)
    palette_index = color_counts[0][1]
    dominant_color = palette[palette_index*3:palette_index*3+3]

    return dominant_color

# Use this to check for lock icon (of which active icon is green) for hidden power 5 detection with lock (rather than OCR)
# https://stackoverflow.com/a/4846754
def check_if_dominant_color_is_green(img, palette_size=16):
    dominant_color = get_dominant_color(img, palette_size)

    return dominant_color[0] < dominant_color[1] / 2 and dominant_color[2] < dominant_color[1] / 2

# Use this to check for Keep button for existence of change power confirmation window (rather than OCR)
def check_if_dominant_color_is_yellow(img, palette_size=16):
    dominant_color = get_dominant_color(img, palette_size)
    
    return dominant_color[2] < dominant_color[0] / 2 and dominant_color[2] < dominant_color[1] / 2
###

class DisclaimerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("About this application")

        layout = QVBoxLayout()
        layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)
        message = QLabel("""
        <ul>
            <li>This application is meant to be used with the "CookieRun: Tower of Adventures" application, and serves as an automatic hidden power roller.</li>
            <li>It works by using Optical Character Recognition (OCR) and determining whether to reroll or to stop and let the user review the stats.
                <ul>
                    <li>The application essentially takes control of your mouse and simulates a user rolling the hidden powers.</li>
                    <li>As a consequence, it is not faster than manually rolling due to OCR processing times, but is meant to free yourself from the computer from time to time.</li>
                </ul>
            </li>
            <li>Important caveats:
                <ul>
                    <!--<li style="color: red">Please do a few manual rolls first to get rid of the 'Change Power' confirmation screen for not keeping higher rarities for the day, as this will not work if that is present.</li>-->
                    <li>Only works on the Google Play Games PC version of the game. Don't expect this to work on mobile.</li>
                    <li>Does not keep track of your scroll or gold levels and will click aimlessly if either of these run out.</li>
                    <li>Use at your own risk. I do not take responsibility for accidentally missed good rolls, but I will do my best to minimize this.</li>
                </ul>
            </li>
        </ul>

        <p>Application created by AltiV. Version 1.0.0 (January 2025).</p>
                         """
                         )
        message.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(message)
        self.setLayout(layout)

class InstructionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("How to use")

        layout = QVBoxLayout()
        layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)
        
        message = QLabel("""
        <ol>
            <li>Click the "Detect ToA application" and check to make sure the application has been detected before continuing.</li>
            <li>Select the equipment type that you are rolling (Weapon/Headwear/Bodywear/Legwear), as the types of hidden powers available vary depending on this.</li>
            <li>In the "Hidden Powers to Target" section, select any of the hidden powers that you are targeting, along with the minimum numerical value you desire.
                <ul>
                    <li>The application will essentially stop rolling when it sees new hidden powers that match the targeted selections.</li>
                </ul>
            </li>
            <li>For the criteria radio buttons section, select 1 or 2 depending on whether you want one or both of the stats to match the chosen hidden powers.</li>
            <li>The "Require mutually exclusive..." button makes it so getting two of the same hidden power type does not count as a satisfactory roll, and will continue rolling.</li>
            <li>From there, click "Start Rolling" to start the application alternating between clicking the "Change Power" in-game, checking the new powers, and either continuing to roll if the stats don't match or stopping if they do.
                <ul>
                    <li>Otherwise, you can manually stop rolling by clicking the "Stop Rolling" button, 'q' on your keyboard, or slamming your mouse cursor to the corner of your screen as a failsafe.</li>
                    <li>As this application uses OCR to parse the hidden power values, do not block the text with your cursor or other objects to minimize erroneous readings.</li>
                    <li>It is recommended to try the application with some lenient hidden power rolls first to make sure it works as expected with your system.</li>
                </ul>
            </li>
        </ol>
        <p style='color: red'>Make sure you have properly set up your rolling environment before starting the rolls (see below screenshot for example, including the Scroll of Potential being selected):</p>
        """)
        message.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(message)
        
        example_screenshot = QLabel(self)
        example_screenshot.setFixedSize(576, 324)
        
        pixmap = QPixmap(os.path.join(basedir, "images", "screen_1_sample.png"))
        example_screenshot.setPixmap(pixmap)
        example_screenshot.setScaledContents(True)
        layout.addWidget(example_screenshot, alignment=Qt.AlignmentFlag.AlignCenter)
        
        self.setLayout(layout)
        
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.threadpool = QThreadPool()
        # print("Multithreading with maximum %d threads" % self.threadpool.maxThreadCount())
        
        self.setWindowTitle("CookieRun: ToA Hidden Power Auto-Roller")
        self.setFixedSize(QSize(400, 760))
        
        layout = QVBoxLayout()
        
        # Set title
        label = QLabel("CookieRun: ToA Hidden Power Auto-Roller")
        font = label.font()
        font.setPointSize(14)
        label.setFont(font)
        label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(label)
        
        top_buttons = QGridLayout()
        
        # Add disclaimer button
        disclaimer_button = QPushButton("Read before use")
        disclaimer_button.clicked.connect(self.disclaimer_button_clicked)
        top_buttons.addWidget(disclaimer_button, 0, 0)
        
        usage_button = QPushButton("How to use this application")
        usage_button.clicked.connect(self.instructions_button_clicked)
        top_buttons.addWidget(usage_button, 0, 1)
        
        layout.addLayout(top_buttons)
        
        # Button to detect ToA window
        app_detect_button = QPushButton("Detect ToA application")
        app_detect_button.clicked.connect(self.app_detect_button_clicked)
        layout.addWidget(app_detect_button)
        
        # Application detected textbox
        self.app_window_info = QLineEdit()
        self.app_window_info.setReadOnly(True)
        self.app_window_info.setEnabled(False)
        self.app_window_info.setPlaceholderText("Information about the ToA application will be displayed here.")
        self.app_window_info.setStyleSheet("border: 2px solid red")
        layout.addWidget(self.app_window_info)
        
        # Equipment type dropdown (changing this will change what displays in the hidden power table)
        equip_label = QLabel("Equipment Type")
        layout.addWidget(equip_label)
        
        self.equip_combobox = QComboBox()
        self.equip_combobox.addItems(EQUIPMENT_TYPES)
        # The same signal can send a text hidden_power
        self.equip_combobox.currentTextChanged.connect(self.equip_dropdown_changed)
        layout.addWidget(self.equip_combobox)
        
        # Hidden Powers Selection Dropdown
        hidden_powers_label = QLabel("Hidden Powers to Target (Select one or more)")
        layout.addWidget(hidden_powers_label)
        
        # Hidden Powers Table (Dynamically changes based on what equipment type is selected)
        hidden_powers_table_headers = ["Hidden Power", "Minimum Target Value", "Minimum Rarity"]
        
        self.hidden_powers_table = QTableWidget()
        self.hidden_powers_table.setColumnCount(3)
        self.hidden_powers_table.setRowCount(11)
        self.hidden_powers_table.setHorizontalHeaderLabels(hidden_powers_table_headers)
        self.hidden_powers_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.hidden_powers_table.verticalHeader().setVisible(False)
        # While this applies to the whole table, this one is just to check the checkboxes in the first column
        self.hidden_powers_table.cellChanged.connect(self.hidden_power_checkbox_clicked)
        
        # Initialize a list of previous hidden power values for the usage of snapping between valid ranges
        self.previous_hidden_power_values = [None] * len(WEAPON_HIDDEN_POWERS)
        
        # Initialize table with empty widgets
        for row in range(len(WEAPON_HIDDEN_POWERS)):
            # First column: Stat name with checkbox
            hidden_power_checkbox = QTableWidgetItem()
            hidden_power_checkbox.setFlags(Qt.ItemIsEnabled)
            hidden_power_checkbox.setCheckState(Qt.CheckState.Unchecked)
            self.hidden_powers_table.setItem(row, 0, hidden_power_checkbox)
            
            # Second column: Stat value, of which range is dependent on stat name
            hidden_power_min_value = QDoubleSpinBox()
            # Start this as disabled, as it will be enabled if the user selects it for usage
            hidden_power_min_value.setEnabled(False)
            
            self.hidden_powers_table.setCellWidget(row, 1, hidden_power_min_value)
            
            # Third column: Hidden power rarity
            hidden_power_rarity = QTableWidgetItem()
            hidden_power_rarity.setFlags(Qt.ItemIsEnabled)
            hidden_power_rarity.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
            self.hidden_powers_table.setItem(row, 2, hidden_power_rarity)
        
        layout.addWidget(self.hidden_powers_table)
        
        number_of_stats_to_target_grid = QGridLayout()
        
        stat_criteria_label = QLabel("How many stats need to satisfy the above criteria?")
        number_of_stats_to_target_grid.addWidget(stat_criteria_label, 0, 0)
        
        self.stat_criteria_select_one = QRadioButton("1", self)
        number_of_stats_to_target_grid.addWidget(self.stat_criteria_select_one, 0, 1, Qt.AlignmentFlag.AlignRight)
        
        self.stat_criteria_select_two = QRadioButton("2", self)
        self.stat_criteria_select_two.setChecked(True)
        number_of_stats_to_target_grid.addWidget(self.stat_criteria_select_two, 0, 2, Qt.AlignmentFlag.AlignRight)
        
        layout.addLayout(number_of_stats_to_target_grid)
        
        # Checkbox for requiring mutually exclusive hidden powers 
        self.mutually_exclusive_hidden_powers_checkbox = QCheckBox("Require mutually exclusive hidden powers (no two same types)")
        layout.addWidget(self.mutually_exclusive_hidden_powers_checkbox)
        
        bottom_buttons = QGridLayout()
        
        # Start rolling button
        self.start_rolling_button = QPushButton("Start Rolling")
        self.start_rolling_button.setDisabled(True)
        self.start_rolling_button.clicked.connect(self.start_rolling_button_clicked)
        bottom_buttons.addWidget(self.start_rolling_button, 0, 0)
        
        # Stop rolling button
        self.stop_rolling_button = QPushButton("Stop Rolling (or hold 'q')")
        self.stop_rolling_button.setDisabled(True)
        self.stop_rolling_button.clicked.connect(self.stop_rolling_button_clicked)
        bottom_buttons.addWidget(self.stop_rolling_button, 0, 1)
        
        layout.addLayout(bottom_buttons)
        
        self.roll_data_info = QTextEdit()
        self.roll_data_info.setReadOnly(True)
        self.roll_data_info.setFixedHeight(100)
        self.roll_data_info.setPlaceholderText("Information about rolling results will be displayed here.")
        # self.roll_data_info.verticalScrollBar().setValue(self.roll_data_info.verticalScrollBar().maximum())
        self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
        layout.addWidget(self.roll_data_info)
        
        central_widget = QWidget()
        central_widget.setLayout(layout)

        self.setCentralWidget(central_widget)
    
    ###################
    # OTHER FUNCTIONS #
    ###################
    
    # Helper function to return list of selected hidden powers and their associated value(s)
    def get_hidden_powers_selected(self):
        if (self.equip_combobox.currentText() == "---"): return []
        
        # Get all the targeted stats and associated values to start scanning for
        targeted_stats = []
        
        for row, hidden_power in enumerate(eval(eval("self.equip_combobox.currentText().upper() + '_HIDDEN_POWERS'"))):
            if self.hidden_powers_table.item(row, 0).checkState() == Qt.CheckState.Checked:
                targeted_stats.append([hidden_power, round(self.hidden_powers_table.cellWidget(row, 1).value(), 2)])
                
        return targeted_stats
    
    def disclaimer_button_clicked(self, s):
        dlg = DisclaimerDialog(self)
        dlg.exec()
    
    def instructions_button_clicked(self, s):
        dlg = InstructionsDialog(self)
        dlg.exec()
    
    # Returns an array containing the window handle of the game and its bounding box 
    # Unminimize variable is when you need to make sure you need to start clicking in it
    def get_toa_app_window_handle_and_bounding_box(self, unminimize = False):
        # Get list of windows
        # hwnd = window handle
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                # print(f"window text: '{win32gui.GetWindowText(hwnd)}'")
                listOfWindows.append((hwnd, win32gui.GetWindowText(hwnd)))
        
        # Step 1 - Locate CookieRun TOA window
        topList, listOfWindows = [], []

        win32gui.EnumWindows(callback, None)

        # print(topList)
        # print(listOfWindows)

        toaWindow = [(hwnd, title) for hwnd, title in listOfWindows if title ==
                    'CookieRun: Tower of Adventures']

        # Get first instance
        if (toaWindow):            
            toaWindow = toaWindow[0]
            
            toaWindowHandle = toaWindow[0]
            
            if unminimize and win32gui.GetWindowPlacement(toaWindowHandle)[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(toaWindowHandle, win32con.SW_SHOWNORMAL)
            
            # Get the bounding box dimensions
            # bbox = win32gui.GetWindowRect(toaWindowHandle)
            bbox = win32gui.GetWindowPlacement(toaWindowHandle)[4]
            
            self.app_window_info.setStyleSheet("border: 2px solid green")
            
            return [toaWindowHandle, bbox]
        else:
            self.app_window_info.setStyleSheet("border: 2px solid red")
            
            return [None, None]
        
    def app_detect_button_clicked(self):
        self.bbox = self.get_toa_app_window_handle_and_bounding_box()[1]
        
        if (self.bbox):
            self.app_window_info.setText(f"Application detected! Window dimensions: {self.bbox}")
                    
            # Determine if the start rolling button should be disabled or not based on if there are hidden powers selected
            if (hasattr(self, 'start_rolling_button')):
                self.start_rolling_button.setDisabled(True if not self.get_hidden_powers_selected() or not hasattr(self, 'bbox') or not self.bbox else False)
        else:
            self.app_window_info.setText("Application not detected.")
    
    def hidden_power_checkbox_clicked(self, row, column):
        # Ignore changes from non-checkbox columns as they are handled by other functions
        if (column != 0): return
        
        # Disable and clear other columns if row is unchecked
        if (self.hidden_powers_table.item(row, column).checkState() == Qt.CheckState.Unchecked):
            if (self.hidden_powers_table.cellWidget(row, 1) is not None):
                self.hidden_powers_table.cellWidget(row, 1).setEnabled(False)
                
                self.hidden_powers_table.item(row, 2).setText("")
        else:
            self.hidden_powers_table.cellWidget(row, 1).setEnabled(True)
            
            hidden_power_stat = self.hidden_powers_table.item(row, column).text()
            hidden_power_value = round(self.hidden_powers_table.cellWidget(row, 1).value(), 2)
            
            if hidden_power_stat:
                value_to_set_to = list(HIDDEN_POWER_RARITY_VALUES[hidden_power_stat].items())[0]
                
                if (hidden_power_value <= list(HIDDEN_POWER_RARITY_VALUES[hidden_power_stat].items())[-1][1][1]):
                    value_to_set_to = next(x for x in HIDDEN_POWER_RARITY_VALUES[hidden_power_stat].items() if x[1][1] >= hidden_power_value)
                
                self.hidden_powers_table.item(row, 2).setText(value_to_set_to[0])
                self.previous_hidden_power_values[row] = value_to_set_to[1][0]
            else:
                self.hidden_powers_table.item(row, 2).setText("")
                self.previous_hidden_power_values[row] = None
        
        # Determine if the start rolling button should be disabled or not based on if there are hidden powers selected
        if (hasattr(self, 'start_rolling_button')):
            self.start_rolling_button.setDisabled(True if not self.get_hidden_powers_selected() or not hasattr(self, 'bbox') or not self.bbox else False)
    
    def hidden_power_numerical_value_changed(self, value, row, hidden_power):
        # Check if the new value is valid (between ranges) and modify if not
        modified_value = value
        value_in_range = value >= list(HIDDEN_POWER_RARITY_VALUES[hidden_power].values())[0][0] and value <= list(HIDDEN_POWER_RARITY_VALUES[hidden_power].values())[-1][1]
        value_is_valid = any(value >= x[1][0] and value <= x[1][1] for x in HIDDEN_POWER_RARITY_VALUES[hidden_power].items())
        
        if (value_in_range and not value_is_valid):
            # Snap to valid value if found to be in-between rarities
            # Snap upwards
            if (self.previous_hidden_power_values[row] and self.previous_hidden_power_values[row] < value):
                modified_value = next(x[1][0] for x in HIDDEN_POWER_RARITY_VALUES[hidden_power].items() if value < x[1][1])
                self.hidden_powers_table.cellWidget(row, 1).blockSignals(True)
                self.hidden_powers_table.cellWidget(row, 1).setValue(modified_value)
                self.hidden_powers_table.cellWidget(row, 1).blockSignals(False)
                
            # Snap downwards
            elif (self.previous_hidden_power_values[row] and self.previous_hidden_power_values[row] > value):
                modified_value = next(x[1][1] for x in reversed(HIDDEN_POWER_RARITY_VALUES[hidden_power].items()) if value > x[1][0])
                self.hidden_powers_table.cellWidget(row, 1).blockSignals(True)
                self.hidden_powers_table.cellWidget(row, 1).setValue(modified_value)
                self.hidden_powers_table.cellWidget(row, 1).blockSignals(False)
        
        self.previous_hidden_power_values[row] = modified_value
        
        # Determine what to show in the "Minimum Rarity" field        
        if (float(HIDDEN_POWER_RARITY_VALUES[hidden_power]["Super Epic"][1]) >= round(float(modified_value), 2)):
            self.hidden_powers_table.item(row, 2).setText(next(x[0] for x in HIDDEN_POWER_RARITY_VALUES[hidden_power].items() if x[1][1] >= round(float(modified_value), 2)))

    def equip_dropdown_changed(self, text):
        if text == "---":
            for row in range(self.hidden_powers_table.rowCount()):
                # First column: Stat name with checkbox
                self.hidden_powers_table.item(row, 0).setText("")
                self.hidden_powers_table.item(row, 0).setFlags(Qt.ItemIsEnabled)
                self.hidden_powers_table.item(row, 0).setCheckState(Qt.CheckState.Unchecked)
                
                # Second column: Stat value, of which range is dependent on stat name
                self.hidden_powers_table.cellWidget(row, 1).setEnabled(False)
                self.hidden_powers_table.cellWidget(row, 1).setMinimum(0)
                self.hidden_powers_table.cellWidget(row, 1).setMaximum(0)
                self.hidden_powers_table.cellWidget(row, 1).setValue(0)
                self.hidden_powers_table.cellWidget(row, 1).setSingleStep(0)
                self.hidden_powers_table.cellWidget(row, 1).setSuffix("")
                
                # Third column: Hidden power rarity
                self.hidden_powers_table.item(row, 2).setText("")
        else:
            for row, hidden_power in enumerate(eval(eval("text.upper() + '_HIDDEN_POWERS'"))):
                # First column: Stat name with checkbox
                self.hidden_powers_table.item(row, 0).setText(hidden_power)
                self.hidden_powers_table.item(row, 0).setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                self.hidden_powers_table.item(row, 0).setCheckState(Qt.CheckState.Unchecked)
                
                # Second column: Stat value, of which range is dependent on stat name          
                self.hidden_powers_table.cellWidget(row, 1).setEnabled(False)
                self.hidden_powers_table.cellWidget(row, 1).setMinimum(HIDDEN_POWER_RARITY_VALUES[hidden_power]["Common"][0])
                self.hidden_powers_table.cellWidget(row, 1).setMaximum(HIDDEN_POWER_RARITY_VALUES[hidden_power]["Super Epic"][1])
                self.hidden_powers_table.cellWidget(row, 1).setValue(HIDDEN_POWER_RARITY_VALUES[hidden_power]["Common"][0])
                # Each rarity can roll six different values between min and max (inclusive), so divide range by five to get step
                self.hidden_powers_table.cellWidget(row, 1).setSingleStep((HIDDEN_POWER_RARITY_VALUES[hidden_power]["Common"][1] - HIDDEN_POWER_RARITY_VALUES[hidden_power]["Common"][0]) * 0.2)
                self.hidden_powers_table.cellWidget(row, 1).setSuffix("%" if "Elemental ATK" not in self.hidden_powers_table.item(row, 0).text() else "")
                
                # This throws error on first pass but I am not sure of method to generically check for existing connections
                try: self.hidden_powers_table.cellWidget(row, 1).valueChanged.disconnect() 
                except Exception: pass
                self.hidden_powers_table.cellWidget(row, 1).valueChanged.connect(lambda value, row = row, hidden_power = hidden_power: 
                    self.hidden_power_numerical_value_changed(value, row, hidden_power))
                
                # Third column: Hidden power rarity
                self.hidden_powers_table.item(row, 2).setText("")

    def start_rolling_button_clicked(self):
        self.roll_data_info.setText("")
        
        # Let's disable stuff here to prevent user tampering
        self.equip_combobox.setDisabled(True)
        
        for row in range(self.hidden_powers_table.rowCount()):
            self.hidden_powers_table.item(row, 0).setFlags(Qt.ItemIsEnabled)
            self.hidden_powers_table.cellWidget(row, 1).setEnabled(False)
                
        self.stat_criteria_select_one.setDisabled(True)
        self.stat_criteria_select_two.setDisabled(True)
        self.mutually_exclusive_hidden_powers_checkbox.setDisabled(True)
        
        # Get the hidden powers the user wants to roll for
        self.targeted_stats = self.get_hidden_powers_selected()
        
        # Check that the application is (still) running
        self.toaWindowHandle = self.get_toa_app_window_handle_and_bounding_box(True)[0]
        
        if not self.targeted_stats or not self.toaWindowHandle: return
        
        self.start_rolling_button.setDisabled(True)
        self.stop_rolling_button.setDisabled(False)
        
        self.number_of_stats_to_check = 1 if self.stat_criteria_select_one.isChecked() else 2
        
        targeted_stats_str_formatted = " | ".join(list(map(lambda a: f"{a[0]}: {a[1]}{("%" if "Elemental ATK" not in a[0] else "")}+", self.targeted_stats)))
        self.roll_data_info.append(f"Targeted stats: {targeted_stats_str_formatted}")
        self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
        
        self.hidden_power_thread_active = True
        self.keyboard_thread_active = True
        
        # self.execute_hidden_power_thread = threading.Thread(target = hidden_power_rolling_logic_thread)
        # self.execute_hidden_power_thread.start()
        
        # self.execute_keyboard_thread = threading.Thread(target = keyboard_detecting_thread)
        # self.execute_keyboard_thread.start()
        
        self.threadpool.start(self.hidden_power_rolling_logic_thread)
        self.threadpool.start(self.keyboard_detecting_thread)
        
        # active_threads = self.threadpool.activeThreadCount()
        # max_thread_count = self.threadpool.maxThreadCount()
        # print(f"{active_threads} out of {max_thread_count} possible threads active")
            
    def stop_rolling_button_clicked(self):
        playsound(os.path.join(basedir, "sounds", "notification-cancel.mp3"), block = False)
        self.end_rolling_session()
        
    def end_rolling_session(self):
        self.roll_data_info.append("--------------------")
        
        self.start_rolling_button.setDisabled(False)
        self.stop_rolling_button.setDisabled(True)
        
        self.hidden_power_thread_active = False
        self.keyboard_thread_active = False
        
        # Let's re-enable the above controls
        self.equip_combobox.setDisabled(False)
        
        for row in range(self.hidden_powers_table.rowCount()):
            self.hidden_powers_table.item(row, 0).setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            self.hidden_powers_table.cellWidget(row, 1).setEnabled(True)
        
        self.stat_criteria_select_one.setDisabled(False)
        self.stat_criteria_select_two.setDisabled(False)
        self.mutually_exclusive_hidden_powers_checkbox.setDisabled(False)
    
    # Two separate threads that handle the main rolling logic, which is done to prevent application freezing
    @Slot()
    def hidden_power_rolling_logic_thread(self):
        try:
            while self.hidden_power_thread_active:
                self.bbox = self.get_toa_app_window_handle_and_bounding_box(True)[1]
                
                # "It will let you SetForegroundWindow, if you send an alt key first"
                # https://stackoverflow.com/a/15503675
                pyautogui.press("alt")
                
                # win32gui.SetForegroundWindow throws an exception if the start menu is open during its execution, so 
                try:
                    win32gui.SetForegroundWindow(self.toaWindowHandle)
                except Exception:
                    print("Set foreground window function does not work while the Start Menu is open, among other potential issues.")
                
                # # Separate OCR call just to get the hidden power level, since capture region changes at HP5 with the lock icon appearing...
                # img = ImageGrab.grab(get_modified_entity_dimensions(S1_HIDDEN_POWER_LEVEL, self.bbox), all_screens = True)
                # opencv_image = cv2.cvtColor(numpy.array(img), cv2.COLOR_RGB2BGR)
                # gry = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
                # thresh = cv2.threshold(gry, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
                # img = Image.fromarray(cv2.cvtColor(thresh, cv2.COLOR_BGR2RGB))
                # hidden_power_level = pytesseract.image_to_string(img, lang='eng', config='--psm 6')
                # # self.roll_data_info.append(f"Hidden Power Level: {hidden_power_level}")
                # # self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
                
                # Start by clicking Change Power
                click(get_center_of_entity_coords_based_on_window_offset(S1_CHANGE_POWER_BUTTON, self.bbox, True))
                
                # Wait for stuff to load
                sleep(1.4)
                
                # Midway check to break out of loop based on cancel
                if not self.hidden_power_thread_active: break
                
                # Check if lock is active which is non-OCR way to detect "relevant" Hidden Power 5
                img1 = ImageGrab.grab(get_modified_entity_dimensions(S2_NEW_LOCK_BUTTON_1, self.bbox), all_screens = True)
                img2 = ImageGrab.grab(get_modified_entity_dimensions(S2_NEW_LOCK_BUTTON_2, self.bbox), all_screens = True)
                
                hidden_power_5 = check_if_dominant_color_is_green(img1) or check_if_dominant_color_is_green(img2)
                
                # Delay a little longer if hidden power 5 due to the flying hidden power point coins possibly obscuring text as it flies upwards
                if hidden_power_5: sleep(0.6)
                                    
                # Midway check to break out of loop based on cancel
                if not self.hidden_power_thread_active: break
                
                # Let's check the new effects
                img = ImageGrab.grab(get_modified_entity_dimensions(S2_NEW_STAT_BOTH_WITH_LOCK if hidden_power_5 else S2_NEW_STAT_BOTH, self.bbox), all_screens = True)
                
                # Convert to greyscale and then binary inversion via array for easier text recognition
                opencv_image = cv2.cvtColor(numpy.array(img), cv2.COLOR_RGB2BGR)
                gry = cv2.cvtColor(opencv_image, cv2.COLOR_BGR2GRAY)
            
                # Convert back to image
                img = Image.fromarray(cv2.cvtColor(gry, cv2.COLOR_BGR2RGB))
                
                # img.show()
                
                # Let's check the text
                # https://pyimagesearch.com/2021/11/15/tesseract-page-segmentation-modes-psms-explained-how-to-improve-your-ocr-accuracy/
                new_effects_text = pytesseract.image_to_string(img, lang='eng', config='--psm 4')
                                                        
                # Midway check to break out of loop based on cancel
                if not self.hidden_power_thread_active: break
                
                # Due to some scaling configurations resulting in extrenous OCR data being sent over, filter it out by checking for min length
                new_effects_arr = filter(lambda a: len(a) >= 8, new_effects_text.split("\n"))
                # Instead of doing double OCR call just to check for hidden power 5 stuff, just assume the lock icon reads as an @ sign at times to speed things up
                # and hope for the best.....
                new_effects_arr = list(map(lambda a: a.replace("|", "").replace("@", "").strip(), new_effects_arr))

                # For saving screenshots with OCR text for testing purposes                
                # ImageDraw.Draw(
                #     img  # Image
                # ).text(
                #     (0, 0),  # Coordinates
                #     " | ".join(new_effects_arr),  # Text
                #     (0, 0, 0)  # Color
                # )
                # img.save(os.path.join(basedir, "images", "test-screenshots", f"{datetime.now().strftime('%Y-%m-%d %H_%M_%S')}.png"))
                
                self.roll_data_info.append(f"Hidden powers rolled: '{" | ".join(new_effects_arr)}'")
                self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
                
                # Now start calculating whether to stop the function or reroll based on if criteria are met
                matching_counter = 0
                
                # For each of the two new hidden powers rolled...
                for rolled_stat in new_effects_arr:
                    # Iterating through the targeted stats, check if the new hidden power matches both the type as well as the minimum value requirement
                    # The ".findall(stat[1]) or [0])[0]" logic is to default to sending a 0 through if OCR + regex didn't properly return a numerical value
                    if any((rolled_stat.upper().startswith(stat[0].upper()) and float((re.compile(r'[.\d]+').findall(rolled_stat) or [0])[0]) >= stat[1]) for stat in self.targeted_stats):
                        matching_counter = matching_counter + 1
                
                self.roll_data_info.append(f"Number of stats matching requirements: {matching_counter}")
                self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
                
                # Stop rolling if criteria is met
                if matching_counter >= self.number_of_stats_to_check:
                    # Scenario where the hidden powers meet standard criteria but not the mutually exclusive one when checked
                    new_effects_text_only = list(map(lambda a: " ".join(a.split(" ")[:-1]), new_effects_arr))
                    
                    if self.mutually_exclusive_hidden_powers_checkbox.isChecked() and all(x == new_effects_text_only[0] for x in new_effects_text_only):
                        self.roll_data_info.append("New hidden power rolls satisfy criteria but are not mutually exclusive; rerolling...")
                        self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
                        click(get_center_of_entity_coords_based_on_window_offset(S2_KEEP_BUTTON, self.bbox, True))
                        sleep(0.4)
                    else:
                        playsound(os.path.join(basedir, "sounds", "notification-complete.mp3"), block = False)
                        self.roll_data_info.append("New hidden power rolls satisfy criteria; stopping rolling.")
                        self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
                        self.end_rolling_session()
                        break
                else:
                    self.roll_data_info.append("New hidden power rolls do not satisfy criteria; rerolling...")
                    self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
                    click(get_center_of_entity_coords_based_on_window_offset(S2_KEEP_BUTTON, self.bbox, True))
                    sleep(0.4)
                
                # # This is in case the "Change Power" confirmation window has not been checked away (doesn't work)
                # pyautogui.keyDown('enter')
                # sleep(0.1)
                # pyautogui.keyUp('enter')
                # sleep(0.1)
                
                # In case the "Change Power" confirmation window has not been checked away (try to detect the yellow of the button...)
                image_potential_keep_button = ImageGrab.grab(get_modified_entity_dimensions(S3_CHANGE_POWER_KEEP, self.bbox), all_screens = True)
                
                if check_if_dominant_color_is_yellow(image_potential_keep_button):
                    click(get_center_of_entity_coords_based_on_window_offset(S3_CHANGE_POWER_KEEP, self.bbox, True))
        except pyautogui.FailSafeException:
            self.roll_data_info.append("PyAutoGUI failsafe triggered from mouse cursor being dragged to corner; stopping logic.")
            self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
            self.stop_rolling_button_clicked()
        except Exception as error:
            traceback.print_exc()
            self.roll_data_info.append("An error has occured in the rolling function; stopping logic.")
            self.roll_data_info.moveCursor(QTextCursor.MoveOperation.End)
            self.stop_rolling_button_clicked()
                
    # Pressing the designated key will exit both of the threads
    @Slot()
    def keyboard_detecting_thread(self):
        while self.keyboard_thread_active:
            if keyboard.is_pressed('q'):
                self.hidden_power_thread_active = False
                self.stop_rolling_button_clicked()
                break
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Build the relative paths for icons using os.path.join()
    app.setWindowIcon(QIcon(os.path.join(basedir, 'images/scroll_of_potential.ico')))
    window = MainWindow()
    window.show()
    app.exec()