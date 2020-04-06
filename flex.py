from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_cell_to_rowcol
import tkinter as tk
import tksheet
import parser
import re

INIT_ROWS = 100
INIT_COLS = 100

class flex(tk.Tk):
    formulas = [ ["=0"] * INIT_COLS for _ in range(INIT_ROWS)]
    selectionBuffer = None
    selectedCell = None
    selectedCellSumMean = None
    updateBinds = {}
    def __init__(self):
        tk.Tk.__init__(self)
        self.selectedCell = tk.StringVar()
        self.selectedCellSumMean = tk.StringVar()
        selectedCellLabel = tk.Label(self, textvariable=self.selectedCell)
        selectedCellLabel.grid(row=1, column=0, sticky="se")
        selectedCellSumMeanLabel = tk.Label(self, textvariable=self.selectedCellSumMean)
        selectedCellSumMeanLabel.grid(row=1, column=0, sticky="sw")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.sheet = tksheet.Sheet(self,
                                data=[],  # to set sheet data at startup
                                total_rows = INIT_ROWS, #if you want to set empty sheet dimensions at startup
                                total_columns = INIT_COLS, #if you want to set empty sheet dimensions at startup
                                height=500,  # height and width arguments are optional
                                width=1200)  # For full startup arguments see DOCUMENTATION.md
        self.sheet.enable_bindings(("single_select",  # "single_select" or "toggle_select"
                                         "drag_select",  # enables shift click selection as well
                                         "column_drag_and_drop",
                                         "row_drag_and_drop",
                                         "column_select",
                                         "row_select",
                                         "column_width_resize",
                                         "double_click_column_resize",
                                         # "row_width_resize",
                                         # "column_height_resize",
                                         "arrowkeys",
                                         "row_height_resize",
                                         "double_click_row_resize",
                                         "right_click_popup_menu",
                                         "rc_select",
                                         "rc_insert_column",
                                         "rc_delete_column",
                                         "rc_insert_row",
                                         "rc_delete_row",
                                         "copy",
                                         "cut",
                                         "paste",
                                         "delete",
                                         "undo",
                                         "edit_cell"))
        # self.sheet.enable_bindings("enable_all")
        # self.sheet.disable_bindings() #uses the same strings
        self.sheet.grid(row=0, column=0, sticky="nswe")

        # __________ HIGHLIGHT / DEHIGHLIGHT CELLS __________

        #self.sheet.highlight_cells(row=5, column=5, bg="#ed4337", fg="white")
        #self.sheet.highlight_cells(row=5, column=1, bg="#ed4337", fg="white")
        #self.sheet.highlight_cells(row=5, bg="#ed4337", fg="white", canvas="row_index")
        #self.sheet.highlight_cells(column=0, bg="#ed4337", fg="white", canvas="header")

        # __________ DISPLAY SUBSET OF COLUMNS __________

        # self.sheet.display_subset_of_columns(indexes = [3, 1, 2], enable = True) #any order

        # __________ DATA AND DISPLAY DIMENSIONS __________

        # self.sheet.total_rows(4) #will delete rows if set to less than current data rows
        # self.sheet.total_columns(2) #will delete columns if set to less than current data columns
        # self.sheet.sheet_data_dimensions(total_rows = 4, total_columns = 2)
        # self.sheet.sheet_display_dimensions(total_rows = 4, total_columns = 6) #currently resets widths and heights
        # self.sheet.set_sheet_data_and_display_dimensions(total_rows = 4, total_columns = 2) #currently resets widths and heights

        # __________ SETTING OR RESETTING TABLE DATA __________

        # .set_sheet_data() function returns the object you use as argument
        # verify checks if your data is a list of lists, raises error if not
        # self.data = self.sheet.set_sheet_data([[f"Row {r} Column {c}" for c in range(30)] for r in range(2000)], verify = False)

        # __________ SETTING ROW HEIGHTS AND COLUMN WIDTHS __________

        # self.sheet.set_cell_data(0, 0, "\n".join([f"Line {x}" for x in range(500)]))
        # self.sheet.set_column_data(1, ("" for i in range(2000)))
        # self.sheet.row_index((f"Row {r}" for r in range(2000))) #any iterable works
        # self.sheet.row_index("\n".join([f"Line {x}" for x in range(500)]), 2)
        # self.sheet.column_width(column = 0, width = 300)
        # self.sheet.row_height(row = 0, height = 60)
        # self.sheet.set_column_widths([120 for c in range(30)])
        # self.sheet.set_row_heights([30 for r in range(2000)])
        # self.sheet.set_all_column_widths()
        # self.sheet.set_all_row_heights()
        # self.sheet.set_all_cell_sizes_to_text()

        # __________ BINDING A FUNCTION TO USER SELECTS CELL __________

        self.sheet.extra_bindings([
            ("cell_select", self.cell_select),
            ("shift_cell_select", self.shift_select_cells),
            ("drag_select_cells", self.drag_select_cells),
            ("ctrl_a", self.ctrl_a),
            ("row_select", self.row_select),
            ("shift_row_select", self.shift_select_rows),
            ("drag_select_rows", self.drag_select_rows),
            ("column_select", self.column_select),
            ("shift_column_select", self.shift_select_columns),
            ("drag_select_columns", self.drag_select_columns),
            ("deselect", self.deselect),
            ("edit_cell", self.edit_cell),
            ("begin_edit_cell", self.edit_cell_begin)
        ])

        # self.sheet.extra_bindings([("cell_select", None)]) #unbind cell select
        # self.sheet.extra_bindings("unbind_all") #remove all functions set by extra_bindings()

        # __________ BINDING NEW RIGHT CLICK FUNCTION __________

        self.sheet.bind("<3>", self.rc)

        # __________ SETTING HEADERS __________

        # self.sheet.headers((f"Header {c}" for c in range(30))) #any iterable works
        # self.sheet.headers("Change header example", 2)
        # print (self.sheet.headers())
        # print (self.sheet.headers(index = 2))

        # __________ SETTING ROW INDEX __________

        # self.sheet.row_index((f"Row {r}" for r in range(2000))) #any iterable works
        # self.sheet.row_index("Change index example", 2)
        # print (self.sheet.row_index())
        # print (self.sheet.row_index(index = 2))

        # __________ INSERTING A ROW __________

        # self.sheet.insert_row(values = (f"my new row here {c}" for c in range(30)), idx = 0) # a filled row at the start
        # self.sheet.insert_row() # an empty row at the end

        # __________ INSERTING A COLUMN __________

        # self.sheet.insert_column(values = (f"my new col here {r}" for r in range(2050)), idx = 0) # a filled column at the start
        # self.sheet.insert_column() # an empty column at the end

        # __________ SETTING A COLUMNS DATA __________

        # any iterable works
        # self.sheet.set_column_data(0, values = (0 for i in range(2050)))

        # __________ SETTING A ROWS DATA __________

        # any iterable works
        # self.sheet.set_row_data(0, values = (0 for i in range(35)))

        # __________ SETTING A CELLS DATA __________

        # self.sheet.set_cell_data(1, 2, "NEW VALUE")

        # __________ GETTING FULL SHEET DATA __________

        # self.all_data = self.sheet.get_sheet_data()

        # __________ GETTING CELL DATA __________

        # print (self.sheet.get_cell_data(0, 0))

        # __________ GETTING ROW DATA __________

        # print (self.sheet.get_row_data(0)) # only accessible by index

        # __________ GETTING COLUMN DATA __________

        # print (self.sheet.get_column_data(0)) # only accessible by index

        # __________ GETTING SELECTED __________

        # print (self.sheet.get_currently_selected())
        # print (self.sheet.get_selected_cells())
        # print (self.sheet.get_selected_rows())
        # print (self.sheet.get_selected_columns())
        # print (self.sheet.get_selection_boxes())
        # print (self.sheet.get_selection_boxes_with_types())

        # __________ SETTING SELECTED __________

        # self.sheet.deselect("all")
        # self.sheet.create_selection_box(0, 0, 2, 2, type_ = "cells") #type here is "cells", "cols" or "rows"
        # self.sheet.set_currently_selected(0, 0)
        # self.sheet.set_currently_selected("row", 0)
        # self.sheet.set_currently_selected("column", 0)

        # __________ CHECKING SELECTED __________

        # print (self.sheet.is_cell_selected(0, 0))
        # print (self.sheet.is_row_selected(0))
        # print (self.sheet.is_column_selected(0))
        # print (self.sheet.anything_selected())

        # __________ HIDING THE ROW INDEX AND HEADERS __________

        # self.sheet.hide("row_index")
        # self.sheet.hide("top_left")
        # self.sheet.hide("header")

        # __________ ADDITIONAL BINDINGS __________

        # self.sheet.bind("<Motion>", self.mouse_motion)

    def mouse_motion(self, event):
        region = self.sheet.identify_region(event)
        row = self.sheet.identify_row(event, allow_end=False)
        column = self.sheet.identify_column(event, allow_end=False)
        print(region, row, column)

    def deselect(self, event):
        print("deselected")

    def rc(self, event):
        print(event)

    def edit_cell_begin(self, response):
        if(self.getFormulaForResponse(response)!="=0"):
            self.sheet.set_cell_data(response[0], response[1], self.getFormulaForResponse(response)) #update cell content with its formula
        self.sheet.refresh(False, False)

    def edit_cell(self, response):
        content = self.sheet.get_cell_data(response[0], response[1])
        if '\n' in content:
            newcells = content.splitlines()
            for x in range(len(newcells)):
                if '\t' in newcells[x]:
                    newcellsc = re.split(r'\t+', newcells[x])
                    for y in range(len(newcellsc)):
                        self.sheet.set_cell_data(response[0] + x, response[1] + y, newcellsc[y])
                        self.commitCellChanges([response[0] + x, response[1] + y])
                        if x == len(newcells)-1:
                            self.sheet.column_width(response[1]+y, 120)
                else:
                    self.sheet.set_cell_data(response[0]+x, response[1], newcells[x])
                    self.commitCellChanges([response[0]+x, response[1]])
                self.sheet.row_height(row=response[0], height=15)
        else:
            self.commitCellChanges(response)

    def commitCellChanges(self, response):
        content = self.sheet.get_cell_data(response[0], response[1])
        if (content == ""):
            self.formulas[int(response[0])][int(response[1])] = ""
            self.sheet.refresh(False, False)
        elif(content[0]!="="):
            self.formulas[int(response[0])][int(response[1])] = "=" + content
        else:
            self.formulas[int(response[0])][int(response[1])] = content
            self.updateCellFromFormulaResult(response)
        for c in self.updateBinds:
            if xl_rowcol_to_cell(response[0], response[1]) in self.updateBinds[c]:
                if c not in self.formulas[int(response[0])][int(response[1])]:
                    self.updateBinds[c].remove(xl_rowcol_to_cell(response[0], response[1]))

    def updateCellFromFormulaResult(self, response):
        if (xl_rowcol_to_cell(response[0], response[1]) in self.updateBinds):
            for updQE in self.updateBinds[xl_rowcol_to_cell(response[0], response[1])]:
                self.updateCellFromFormulaResult(xl_cell_to_rowcol(updQE))
        self.sheet.set_cell_data(response[0], response[1], self.interpret(self.getFormulaForResponse(response)[1:], response))
        self.sheet.refresh(False, False)

    def interpret(self, f, response):
        vinst = re.compile('[\$]?([aA-zZ]+)[\$]?(\d+)')
        itern = vinst.finditer(f)
        varsn = {}
        xln = xl_rowcol_to_cell(response[0], response[1])
        refs = []
        for match in itern:
            if match.group() == xln:
                return "RECURSION ERROR"
            else:
                refs.append(match.group())
                varsn[match.group()] = self.interpret(self.formulas[xl_cell_to_rowcol(match.group())[0]][xl_cell_to_rowcol(match.group())[1]][1:], xl_cell_to_rowcol(match.group()))
                if(match.group() not in self.updateBinds):
                    self.updateBinds[match.group()] = []
                self.updateBinds[match.group()].append(xln)
                #print("created a bind: when " + match.group() + " changes, "+xln+" is updated")
        for updc in self.updateBinds:
            self.updateBinds[updc] = list(dict.fromkeys(self.updateBinds[updc]))
        import math
        locals().update(varsn)
        return eval(parser.expr(f).compile())

    def getFormulaForResponse(self, response):
        return self.formulas[int(response[0])][int(response[1])]

    def cell_select(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(response[1], response[2]))

    def shift_select_cells(self, response):
        print(self.updateBinds)

    def drag_select_cells(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(response[1], response[2]) + ":" + xl_rowcol_to_cell(response[3]-1, response[4]-1))

    def ctrl_a(self, response):
        print(response)

    def row_select(self, response):
        print(response)

    def shift_select_rows(self, response):
        print(response)

    def drag_select_rows(self, response):
        pass
        # print (response)

    def column_select(self, response):
        print(response)

    def shift_select_columns(self, response):
        print(response)

    def drag_select_columns(self, response):
        pass
        # print (response)


app = flex()
app.mainloop()