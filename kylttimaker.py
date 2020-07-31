'''
KylttiMaker 1.0
TS 2020

This program is used automate product label/sign marking by creating 2D drawing
files in the .dxf format. The input data is read from a excel input file. The
data can be displayed as a text, a QR code or a combination of the options.

User can enter desired size of a single sign as well as the size of the sheets
the signs will be marked on and the program will automatically fit the data on
the sheets. Each sheet (=file) can have multiple layers of signs.
'''

from pathlib import Path
from typing import Union
from tkinter import Tk, LEFT, RIGHT, BOTH, END, StringVar, BooleanVar, DoubleVar, IntVar, Menu, Event
from tkinter.ttk import Treeview, Progressbar, Button, Checkbutton, Entry, Frame, Label, LabelFrame, Spinbox, OptionMenu
import tkinter.filedialog
import ezdxf
import xlrd
import pyqrcode


# QR code with relevant alignment options that can be drawn on a sheet.
class QR:
    SIDE_OPTIONS = ['LEFT', 'RIGHT']
    DEFAULT_MARGIN = 2
    DEFAULT_PADDING = 0

    # Initialize a GUI frame where user can enter the relevant options.
    def __init__(self, properties: LabelFrame) -> None:
        self.frame = Frame(properties)
        Label(self.frame, text='Side').grid(
            column=0, row=0, sticky='E', pady=App.PADDING)
        self.side = StringVar(self.frame)
        OptionMenu(self.frame, self.side,
                   QR.SIDE_OPTIONS[0], *QR.SIDE_OPTIONS).grid(column=1, row=0, sticky='W')
        Label(self.frame, text='Margin').grid(
            column=0, row=1, sticky='E', pady=App.PADDING)
        self.margin = StringVar(self.frame)
        self.margin.set(QR.DEFAULT_MARGIN)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.margin,
                width=App.SPINBOX_WIDTH).grid(column=1, row=1, sticky='W')
        Label(self.frame, text='Inverse').grid(
            column=0, row=2, sticky='E', pady=App.PADDING)
        self.inverse = BooleanVar(self.frame)
        self.inverse.set(False)
        Checkbutton(self.frame, variable=self.inverse).grid(
            column=1, row=2, sticky='W')
        Label(self.frame, text='Padding').grid(
            column=0, row=3, sticky='E', pady=App.PADDING)
        self.padding = StringVar(self.frame)
        self.padding.set(QR.DEFAULT_PADDING)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.padding,
                width=App.SPINBOX_WIDTH).grid(column=1, row=3, sticky='W')

    # Draws the QR code on the sheet
    def draw(self, value: str, sheet: ezdxf.drawing.Drawing, layer: int, sign_origin_x: float, sign_origin_y: float, sign_width: float, sign_height: float) -> None:
        margin = float(self.margin.get())
        padding = float(self.padding.get())
        assert margin >= 0, 'Margin must be positive'
        assert padding >= 0, 'Padding must be positive'
        side = self.side.get()
        inverse = self.inverse.get()
        size = sign_height - 2 * margin - 2 * padding
        hatch = sheet.modelspace().add_hatch(
            dxfattribs={'layer': str(layer), 'hatch_style': 1})

        if side == QR.SIDE_OPTIONS[1]:
            sign_origin_x += sign_width - sign_height

        if inverse:
            hatch.paths.add_polyline_path([
                (sign_origin_x + margin, sign_origin_y - margin),
                (sign_origin_x + sign_height - margin, sign_origin_y - margin),
                (sign_origin_x + sign_height - margin,
                 sign_origin_y + margin - sign_height),
                (sign_origin_x + margin, sign_origin_y + margin - sign_height)
            ])

        # Gets the QR code as multi line string of ones and zeros
        code = pyqrcode.create(value).text(quiet_zone=0)
        lines = code.splitlines()
        cell_size = size / len(lines)

        for y, line in enumerate(lines):
            for x, char in enumerate(line):
                if char == '1':
                    hatch.paths.add_polyline_path([
                        (sign_origin_x + margin + padding + x * cell_size,
                         sign_origin_y - margin - padding - y * cell_size),
                        (sign_origin_x + margin + padding + (x + 1) * cell_size,
                         sign_origin_y - margin - padding - y * cell_size),
                        (sign_origin_x + margin + padding + (x + 1) * cell_size,
                         sign_origin_y - margin - padding - (y + 1) * cell_size),
                        (sign_origin_x + margin + padding + x * cell_size,
                         sign_origin_y - margin - padding - (y + 1) * cell_size)
                    ])


# Text object with a position, size and alignment options that can be drawn on a sheet.
class Text:
    ALIGN_OPTIONS = [
        'TOP_LEFT', 'TOP_CENTER', 'TOP_RIGHT',
        'MIDDLE_LEFT', 'MIDDLE_CENTER', 'MIDDLE_RIGHT',
        'BOTTOM_LEFT', 'BOTTOM_CENTER', 'BOTTOM_RIGHT'
    ]

    # Initialize a GUI frame where user can enter the relevant options.
    def __init__(self, properties: LabelFrame) -> None:
        self.frame = Frame(properties)
        Label(self.frame, text='X').grid(
            column=0, row=0, sticky='E', pady=App.PADDING)
        self.position_x = StringVar(self.frame)
        self.position_x.set(App.DEFAULT_SIGN_WIDTH / 2)
        Spinbox(self.frame, to=App.MAX_SHEET_WIDTH, textvariable=self.position_x,
                width=App.SPINBOX_WIDTH).grid(column=1, row=0, sticky='W')
        Label(self.frame, text='Y').grid(
            column=0, row=1, sticky='E', pady=App.PADDING)
        self.position_y = StringVar(self.frame)
        self.position_y.set(App.DEFAULT_SIGN_HEIGHT / 2)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.position_y,
                width=App.SPINBOX_WIDTH).grid(column=1, row=1, sticky='W')
        Label(self.frame, text='Font size').grid(
            column=0, row=2, sticky='E', pady=App.PADDING)
        self.size = StringVar(self.frame)
        self.size.set(App.DEFAULT_SIGN_HEIGHT / 2)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.size,
                width=App.SPINBOX_WIDTH).grid(column=1, row=2, sticky='W')
        Label(self.frame, text='Align').grid(
            column=0, row=3, sticky='E', pady=App.PADDING)
        self.align = StringVar(self.frame)
        OptionMenu(self.frame, self.align,
                   Text.ALIGN_OPTIONS[4], *Text.ALIGN_OPTIONS).grid(column=1, row=3, sticky='W')

    def draw(self, value: str, sheet: ezdxf.drawing.Drawing, layer: int, sign_origin_x: float, sign_origin_y: float, sign_width: float, sign_height: float) -> None:
        size = float(self.size.get())
        assert size >= 0, 'Font size must be positive'
        position = (
            sign_origin_x + float(self.position_x.get()),
            sign_origin_y - float(self.position_y.get())
        )
        align = self.align.get()
        sheet.modelspace().add_text(value, dxfattribs={'layer': str(
            layer), 'height': size}).set_pos(position, align=align)


# A hole with a position and size that can be drawn on a sheet.
class Hole:
    DEFAULT_POSITION = 0.0
    DEFAULT_DIAMETER = 5.0

    # Initialize a GUI frame where user can enter the relevant options.
    def __init__(self, properties: LabelFrame) -> None:
        self.frame = Frame(properties)
        Label(self.frame, text='X').grid(
            column=0, row=0, sticky='E', pady=App.PADDING)
        self.position_x = StringVar(self.frame)
        self.position_x.set(Hole.DEFAULT_POSITION)
        Spinbox(self.frame, to=App.MAX_SHEET_WIDTH, textvariable=self.position_x,
                width=App.SPINBOX_WIDTH).grid(column=1, row=0, sticky='W')
        Label(self.frame, text='Y').grid(
            column=0, row=1, sticky='E', pady=App.PADDING)
        self.position_y = StringVar(self.frame)
        self.position_y.set(Hole.DEFAULT_POSITION)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.position_y,
                width=App.SPINBOX_WIDTH).grid(column=1, row=1, sticky='W')
        Label(self.frame, text='Diameter').grid(
            column=0, row=2, sticky='E', pady=App.PADDING)
        self.diameter = StringVar(self.frame)
        self.diameter.set(Hole.DEFAULT_DIAMETER)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.diameter,
                width=App.SPINBOX_WIDTH).grid(column=1, row=2, sticky='W')

    def draw(self, value: str, sheet: ezdxf.drawing.Drawing, layer: int, sign_origin_x: float, sign_origin_y: float, sign_width: float, sign_height: float) -> None:
        x = (sign_origin_x + float(self.position_x.get()))
        y = (sign_origin_y - float(self.position_y.get()))
        r = float(self.diameter.get()) / 2
        assert r >= 0, 'Diameter must be positive'
        sheet.modelspace().add_circle(
            (x, y), r, dxfattribs={'layer': str(layer)})

# A class that represents a collection of text data that get's read from a file and marked on a sign as for example a QR code or a simple text.
class Field:
    MAX_COLUMN = MAX_ROW = 100000

    # Initialize GUI for the user to enter field properties.
    def __init__(self, properties: LabelFrame) -> None:
        self.marks = {}
        self.data = []
        self.frame = Frame(properties)
        Label(self.frame, text='Column').grid(
            column=0, row=0, sticky='E', pady=App.PADDING)
        self.column = StringVar(self.frame)
        self.column.set(1)
        Spinbox(self.frame, from_=1, to=Field.MAX_COLUMN, textvariable=self.column,
                width=App.SPINBOX_WIDTH).grid(column=1, row=0, sticky='WE')
        Label(self.frame, text='Start row').grid(
            column=0, row=1, sticky='E', pady=App.PADDING)
        self.start_row = StringVar(self.frame)
        self.start_row.set(1)
        Spinbox(self.frame, from_=1, to=Field.MAX_ROW, textvariable=self.start_row,
                width=App.SPINBOX_WIDTH).grid(column=1, row=1, sticky='WE')
        Label(self.frame, text='End row').grid(
            column=0, row=2, sticky='E', pady=App.PADDING)
        self.end_row = StringVar(self.frame)
        self.end_row.set(0)
        Spinbox(self.frame, to=Field.MAX_ROW, textvariable=self.end_row,
                width=App.SPINBOX_WIDTH).grid(column=1, row=2, sticky='WE')
        Label(self.frame, text='(0 = No limit)').grid(
            column=2, row=2, columnspan=2, sticky='W')
        Label(self.frame, text='Source').grid(
            column=0, row=3, sticky='E', pady=App.PADDING)
        self.path = StringVar(self.frame)
        Entry(self.frame, textvariable=self.path).grid(
            column=1, row=3, columnspan=3, sticky='WE')
        Button(self.frame, text='Select', command=self.select,
               width=5).grid(column=4, row=3, sticky='W')
        Button(self.frame, text='Read', command=self.read).grid(
            column=1, row=4, sticky='WE')
        Label(self.frame, text='Values').grid(column=2, row=4, sticky='E')
        self.values_amount = IntVar(self.frame)
        self.values_amount.set(0)
        Label(self.frame, textvariable=self.values_amount).grid(
            column=3, row=4, sticky='W')

    # Get the field input path (excel file) from the user and displays it in the relevant entry box.
    def select(self) -> None:
        if dialog_path := tkinter.filedialog.askopenfilename(filetypes=(('Excel', '*.xlsx'), ('Excel', '*.xls'))):
            self.path.set(dialog_path)

    # Populate field's data variable by reading the file specified by path variable.
    def read(self) -> None:
        try:
            column = int(self.column.get())
            start_row = int(self.start_row.get())
            end_row = int(self.end_row.get())
            assert start_row > 0, 'Start row must be greater than 0'
            if end_row == 0:  # End row 0 = no limit.
                end_row = None
            else:
                assert end_row >= start_row, 'End row must be greater than or equal to start row'
            self.data = xlrd.open_workbook(self.path.get()).sheet_by_index(
                0).col_slice(column - 1, start_rowx=start_row - 1, end_rowx=end_row)
            # Show to the user how many values were read.
            self.values_amount.set(len(self.data))
        except Exception as e:
            print(e)

    # Draw field's marks (QR, Text and Hole objects).
    def draw(self, index: int, sheet: ezdxf.drawing.Drawing, layer: int, sign_origin_x: float, sign_origin_y: float, sign_width: float, sign_height: float) -> None:
        if len(self.data) > index:
            value = self.data[index].value
            for mark_iid in self.marks:
                self.marks[mark_iid].draw(
                    value, sheet, layer, sign_origin_x, sign_origin_y, sign_width, sign_height)


class App(Tk):
    DEFAULT_SIGN_WIDTH = 150
    DEFAULT_SIGN_HEIGHT = 22
    DEFAULT_SHEET_WIDHT = 300
    DEFAULT_SHEET_HEIGHT = 300
    DEFAULT_SHEETS_PER_FILE = 0
    MAX_SHEET_WIDTH = 470
    MAX_SHEET_HEIGHT = 310
    MAX_SHEETS_PER_FILE = 100
    SPINBOX_WIDTH = 8
    PADDING = 2
    DXF_VERSIONS = ('R12', 'R2000', 'R2004', 'R2007', 'R2010', 'R2013', 'R2018')

    # Initialize GUI layout.
    def __init__(self) -> None:
        super().__init__()
        self.title('KylttiMaker')
        self.minsize(640, 480)

        # Tree widget that displays fields and their relative marks in a hierarchy.
        self.tree = Treeview(self, selectmode='browse')
        self.tree.heading('#0', text='Fields', command=self.remove_selection)
        self.tree.bind('<Button-3>', self.tree_right_click)
        self.tree.bind('<<TreeviewSelect>>', self.tree_selection_changed)
        self.tree.bind('<Double-Button-1>', self.rename)
        self.bind('<Escape>', self.remove_selection)
        self.bind('<Delete>', self.remove)
        self.tree.pack(side=LEFT, fill=BOTH)
        self.properties = LabelFrame(self, text='Properties')
        self.properties.pack(side=RIGHT, fill=BOTH, expand=1)
        self.fields = {}
        self.selected_iid = None

        # Entry field that get's temporarily shown to the user whilst renaming a field or a mark.
        self.new_name = StringVar(self.tree)
        self.new_name_entry = Entry(self.tree, textvariable=self.new_name)
        self.new_name_entry.bind('<Key-Return>', self.new_name_entered)

        # Output options that get's shown to the user when nothing else is selected from the hierarchy.
        self.frame = Frame(self.properties)
        Label(self.frame, text='Sheet size').grid(
            column=0, row=0, sticky='E', pady=App.PADDING)
        self.sheet_width_var = StringVar(self.frame)
        self.sheet_width_var.set(App.DEFAULT_SHEET_WIDHT)
        Spinbox(self.frame, to=App.MAX_SHEET_WIDTH, textvariable=self.sheet_width_var,
                width=App.SPINBOX_WIDTH).grid(column=1, row=0, sticky='WE')
        Label(self.frame, text='x').grid(column=2, row=0)
        self.sheet_height_var = StringVar(self.frame)
        self.sheet_height_var.set(App.DEFAULT_SHEET_HEIGHT)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.sheet_height_var,
                width=App.SPINBOX_WIDTH).grid(column=3, row=0, sticky='WE')
        Label(self.frame, text='Sign size').grid(
            column=0, row=1, sticky='E', pady=App.PADDING)
        self.sign_width_var = StringVar(self.frame)
        self.sign_width_var.set(App.DEFAULT_SIGN_WIDTH)
        Spinbox(self.frame, to=App.MAX_SHEET_WIDTH, textvariable=self.sign_width_var,
                width=App.SPINBOX_WIDTH).grid(column=1, row=1, sticky='WE')
        Label(self.frame, text='x').grid(column=2, row=1)
        self.sign_height_var = StringVar(self.frame)
        self.sign_height_var.set(App.DEFAULT_SIGN_HEIGHT)
        Spinbox(self.frame, to=App.MAX_SHEET_HEIGHT, textvariable=self.sign_height_var,
                width=App.SPINBOX_WIDTH).grid(column=3, row=1, sticky='WE')
        Label(self.frame, text='Layers per sheet').grid(
            column=0, row=2, sticky='W', pady=App.PADDING)
        self.layers_per_sheet_var = StringVar(self.frame)
        self.layers_per_sheet_var.set(App.DEFAULT_SHEETS_PER_FILE)
        Spinbox(self.frame, to=App.MAX_SHEETS_PER_FILE, textvariable=self.layers_per_sheet_var,
                width=App.SPINBOX_WIDTH).grid(column=1, row=2, sticky='WE')
        Label(self.frame, text='(0 = No limit)').grid(
            column=2, row=2, columnspan=2, sticky='W')
        Label(self.frame, text='DXF version').grid(column=0, row=4, sticky='E', pady=App.PADDING)
        self.dxf_version = StringVar(self.frame)
        OptionMenu(self.frame, self.dxf_version,
                   App.DXF_VERSIONS[1], *App.DXF_VERSIONS).grid(column=1, row=4, sticky='W')
        Button(self.frame, text='Create',
               command=self.create).grid(column=2, row=4, columnspan=2)
        self.frame.pack()

    # Display a popup menu with relevant options when right clicking on the tree widget item.
    def tree_right_click(self, event: Event) -> None:
        menu = Menu(self, tearoff=0)
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid in self.fields:
                menu.add_command(
                    label='Add QR', command=lambda: self.add_mark(QR, iid))
                menu.add_command(label='Add Text',
                                 command=lambda: self.add_mark(Text, iid))
                menu.add_command(label='Add Hole',
                                 command=lambda: self.add_mark(Hole, iid))
            menu.add_command(
                label='Rename', command=lambda: self.rename(iid=iid))
            menu.add_command(
                label='Remove', command=lambda: self.remove(iid=iid))
        else:
            menu.add_command(label='Add field', command=self.add_field)
        menu.tk_popup(event.x_root, event.y_root)

    # Display the properties of the selected item.
    def tree_selection_changed(self, event: Event) -> None:
        # Hide the items previously shown in the properties pane.
        self.new_name_entry.place_forget()
        for child in self.properties.winfo_children():
            child.pack_forget()

        selected_items = self.tree.selection()
        if selected_items:
            self.selected_iid = selected_items[0]
            # Check if the selected item is a field or a mark object, in which case show its properties.
            if self.selected_iid in self.fields:
                self.fields[self.selected_iid].frame.pack()
            else:
                for field_iid in self.fields:
                    if self.selected_iid in self.fields[field_iid].marks:
                        self.fields[field_iid].marks[self.selected_iid].frame.pack()
        else:
            # Clear the properties pane.
            self.selected_iid = None
            self.frame.pack()

    # Create a new field object and add a corresponding node to the hierarchy.
    def add_field(self) -> None:
        iid = self.tree.insert('', END, text='Field')
        self.fields[iid] = Field(self.properties)

    # Display a entry for the user to input a new name for the item to be renamed.
    def rename(self, event: Event = None, iid: int = None) -> None:
        if not iid:
            if self.selected_iid:
                iid = self.selected_iid
            else:
                return
        self.editing_iid = iid
        self.new_name.set(self.tree.item(iid)['text'])
        self.new_name_entry.place(x=20, y=0)
        self.new_name_entry.focus_set()
        self.new_name_entry.select_range(0, END)

    # Display the renamed item in the hierarchy.
    def new_name_entered(self, event: Event) -> None:
        self.tree.item(self.editing_iid, text=self.new_name.get())
        self.new_name_entry.place_forget()

    # Link a new mark speciefied by mark_type parameter to the field speciefied by field_iid parameter.
    def add_mark(self, mark_type: Union[QR, Text, Hole], field_iid: int = None) -> None:
        if not field_iid:
            if self.selected_iid in self.fields:
                field_iid = self.selected_iid
            else:
                print('Select a field first')
                return
        iid = self.tree.insert(field_iid, END, text=mark_type.__name__)
        self.fields[field_iid].marks[iid] = mark_type(self.properties)
        self.tree.see(iid)

    # Remove a tree item speciefied by iid parameter, else removes the currently selected item.
    def remove(self, event: Event = None, iid: int = None) -> None:
        if not iid:
            if self.selected_iid:
                iid = self.selected_iid
            else:
                print('Select something first')
                return
        # Check if the item to be removed is a field item, else check if it is a mark item.
        if iid in self.fields:
            self.remove_selection()
            self.tree.delete(iid)
            del self.fields[iid]
        else:
            for field_iid in self.fields:
                if iid in self.fields[field_iid].marks:
                    self.remove_selection()
                    self.tree.delete(iid)
                    del self.fields[field_iid].marks[iid]

    # Clear the selection.
    def remove_selection(self, event: Event = None) -> None:
        for item in self.tree.selection():
            self.tree.selection_remove(item)

    # Create sheets according to entered settings.
    def create(self) -> None:
        if not self.fields:
            print('No fields')
            return

        # Calculate the length of the longest field (some fields can have less values than others).
        total_signs = 0
        for field_iid in self.fields:
            total_signs = max(total_signs, len(self.fields[field_iid].data))
        if total_signs == 0:
            print('No fields with data')
            return
        try:
            sheet_width = float(self.sheet_width_var.get())
            sheet_height = float(self.sheet_height_var.get())
            sign_width = float(self.sign_width_var.get())
            sign_height = float(self.sign_height_var.get())
            layers_per_sheet = int(self.layers_per_sheet_var.get())
            assert sign_width > 0, 'Sign width must be greater than 0'
            assert sign_height > 0, 'Sign height must be greater than 0'
            assert sheet_width >= sign_width, 'Sheet width must be greater than sign width'
            assert sheet_height >= sign_height, 'Sheet height must be greater than sign height'
        except ValueError:
            print('Invalid dimensions')
            return
        except AssertionError as e:
            print(e)
            return

        # Show progress bar.
        progress_bar = Progressbar(self.frame)
        progress_bar.grid(column=0, row=5, columnspan=4, sticky='WE')

        # Calculate the needed values to define sheet layout.
        signs_per_row = int(sheet_width // sign_width)
        signs_per_column = int(sheet_height // sign_height)
        signs_per_layer = signs_per_row * signs_per_column
        # Ceiling division.
        total_layers = -int(-total_signs // signs_per_layer)
        if layers_per_sheet > 0:
            total_sheets = -int(-total_layers // layers_per_sheet)
        else:
            total_sheets = 1

        # Create needed sheet objects.
        sheets = []
        for _ in range(total_sheets):
            sheets.append(ezdxf.new(self.dxf_version.get()))

        # Iterate over all layers and draw their outline based on how many signs that layer will have.
        for layer in range(total_layers):
            max_x = sign_width * signs_per_row
            max_y = -sign_height * signs_per_column
            if layer == total_layers - 1:  # If last layer.
                signs_in_last_sheet = total_signs - layer * signs_per_layer
                if signs_in_last_sheet < signs_per_row:
                    max_x = sign_width * signs_in_last_sheet
                max_y = sign_height * (-signs_in_last_sheet // signs_per_row)
            if layers_per_sheet > 0:
                sheet_index = layer // layers_per_sheet
            else:
                sheet_index = 0
            # Draw layer outline (left and top side bounds).
            sheets[sheet_index].modelspace().add_lwpolyline(
                [(0, max_y), (0, 0), (max_x, 0)], dxfattribs={'layer': str(layer)})

        # Iterate over each sign.
        for sign_index in range(total_signs):
            # Update progress bar value.
            progress_bar['value'] = (sign_index + 1) / total_signs * 100
            progress_bar.update()

            # Calculate in which position, in which layer of which sheet the current sign should be drawn.
            layer = sign_index // signs_per_layer
            layer_position = sign_index % signs_per_layer
            sign_origin_x = (layer_position % signs_per_row) * sign_width
            sign_origin_y = -(layer_position // signs_per_row) * sign_height
            if layers_per_sheet > 0:
                sheet_index = layer // layers_per_sheet
            else:
                sheet_index = 0
            sheet = sheets[sheet_index]

            # Draw marks (QR, Text and Hole objects).
            for field_iid in self.fields:
                try:
                    self.fields[field_iid].draw(
                        sign_index, sheet, layer, sign_origin_x, sign_origin_y, sign_width, sign_height)
                except Exception as e:
                    print(e)
                    progress_bar.grid_forget()
                    return

            # Draw sign outline (right and bottom side bounds).
            sign_outline = [
                (sign_origin_x, sign_origin_y - sign_height),
                (sign_origin_x + sign_width, sign_origin_y - sign_height),
                (sign_origin_x + sign_width, sign_origin_y)
            ]
            sheet.modelspace().add_lwpolyline(
                sign_outline, dxfattribs={'layer': str(layer)})

        # Save sheets.
        # Get a output directory if there are multiple sheets to be saved, otherwise get path for the single output (.dxf) file.
        if total_sheets > 1:
            if directory := tkinter.filedialog.askdirectory():
                for index, sheet in enumerate(sheets):
                    # Indicate save progress
                    progress_bar['value'] = (index + 1) / len(sheets) * 100
                    progress_bar.update()
                    sheet.saveas(Path(directory) / f'sheet{index}.dxf')
        elif path := tkinter.filedialog.asksaveasfilename(defaultextension='.dxf', filetypes=(('DXF', '*.dxf'), ('All files', '*.*'))):
            sheets[0].saveas(path)
        progress_bar.grid_forget()


if __name__ == '__main__':
    app = App()
    app.mainloop()
