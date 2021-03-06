from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_cell_to_rowcol
import xlsxwriter
import tkinter as tk
from tkinter import messagebox
import statistics
import datetime
import matplotlib.pyplot as plt
from collections import defaultdict
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter.filedialog import askopenfilename, asksaveasfilename
import tksheet
import parser
import math
import os
import csv
import sys
import re

INIT_ROWS = 1000
INIT_COLS = 1000

class flex(tk.Tk):
    formulas = [ ["=0"] * INIT_COLS for _ in range(INIT_ROWS)]
    selectionBuffer = None
    selectedCell = None
    selectedCellSumMean = None
    updateBinds = {}
    cellRefs = {}
    highlightedCells = []
    openfile = ""
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
        self.sheet.grid(row=0, column=0, sticky="nswe")
        self.buildMenu()

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
            ("begin_edit_cell", self.edit_cell_begin),
            ("delete_key", self.delk)
        ])

        # self.sheet.extra_bindings([("cell_select", None)]) #unbind cell select
        # self.sheet.extra_bindings("unbind_all") #remove all functions set by extra_bindings()

        self.sheet.bind("<3>", self.rc)
        self.sheet.bind("<BackSpace>", self.delk)

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

    def buildMenu(self):
        self.menubar = tk.Menu(self)
        importsubm = tk.Menu(self.menubar, tearoff=0)
        importsubm.add_command(label="Excel (.xlsx)")
        importsubm.add_command(label="CSV (.csv)", command=self.importCsv)
        exportsubm = tk.Menu(self.menubar, tearoff=0)
        exportsubm.add_command(label="Excel (.xlsx)", command=self.exportToExcel)
        exportsubm.add_command(label="CSV (.csv)", command=self.exportToCsv)
        filemenu = tk.Menu(self.menubar, tearoff=0)
        filemenu.add_command(label="New", command=self.restart, accelerator="Command-n")
        filemenu.add_command(label="Open...", command=self.open, accelerator="Command-o")
        filemenu.add_cascade(label='Import', menu=importsubm)
        filemenu.add_cascade(label='Export', menu=exportsubm)
        filemenu.add_command(label="Save", command=self.save, accelerator="Command-s")
        filemenu.add_command(label="Save as...", command=self.saveas, accelerator="Command-shift-s")
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.quit, accelerator="Command-q")
        self.menubar.add_cascade(label="File", menu=filemenu)
        graphmenu = tk.Menu(self.menubar, tearoff=0)
        graphmenu.add_command(label="Line plot", command=self.samplePlot)
        graphmenu.add_command(label="Scatter plot", command= lambda: self.samplePlot('ro'))
        helpmenu = tk.Menu(self.menubar, tearoff=0)
        helpmenu.add_command(label="Help Index", command=None)
        helpmenu.add_command(label="About...", command=None)
        funcmenu = tk.Menu(self.menubar, tearoff=0)
        self.generateMenuForModule(math, funcmenu, "Math")
        self.generateMenuForModule(statistics, funcmenu, "Statistics")
        self.generateMenuForModule(datetime, funcmenu, "Date/Time")
        self.menubar.add_cascade(label="Functions", menu=funcmenu)
        self.menubar.add_cascade(label="Graphing", menu=graphmenu)
        self.menubar.add_cascade(label="Help", menu=helpmenu)
        self.config(menu=self.menubar)

    def generatePlotValues(self):
        numbered = defaultdict(list)
        values = defaultdict(list)
        title= ""
        xlabel = ""
        ylabel = ""
        for i in self.sheet.get_selected_cells():
            numbered[i[1]].append(i)
        i=0
        for c in numbered:
            numbered[c] = sorted(numbered[c], key=lambda x: x[0])
            values[i] = []
            for el in numbered[c]:
                values[i].append(self.sheet.get_cell_data(el[0], el[1]))
            i+=1
        try:
            (float(values[0][0]))
        except ValueError:
            xlabel=values[1][0]
            ylabel = values[0][0]
            title = xlabel + " vs " + ylabel
            values[0].pop(0)
            values[1].pop(0)
        return values, title, xlabel, ylabel

    def samplePlot(self, options):
        figure = plt.Figure(figsize=(6, 5), dpi=100)
        ax = figure.add_subplot(111)
        new_window = tk.Toplevel(self)
        chart_type = FigureCanvasTkAgg(figure, new_window)
        chart_type.get_tk_widget().pack()
        values, title, xlabel, ylabel = self.generatePlotValues()
        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)
        ax.set_title(title)
        ax.plot(values[1], values[0], options)

    def generateMenuForModule(self, module, parent, label):
        menu = tk.Menu(self.menubar, tearoff=0)
        for i in dir(module):
            if(i[0]!="_"):
                menu.add_command(label=i)
        parent.add_cascade(label=label, menu=menu)


    def mouse_motion(self, event):
        region = self.sheet.identify_region(event)
        row = self.sheet.identify_row(event, allow_end=False)
        column = self.sheet.identify_column(event, allow_end=False)
        print(region, row, column)

    def deselect(self, event):
        print(event)

    def rc(self, event):
        print(event)

    def delk(self, event):
        for cell in self.sheet.get_selected_cells():
            for bnd in self.updateBinds:
                if xl_rowcol_to_cell(cell[0], cell[1]) in self.updateBinds[bnd]:
                    self.updateBinds[bnd].remove(xl_rowcol_to_cell(cell[0], cell[1]))
            self.formulas[cell[0]][cell[1]] = "=0"
            self.sheet.set_cell_data(cell[0], cell[1], "")
        self.sheet.refresh(False, False)

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
            self.formulas[int(response[0])][int(response[1])] = "=0"
        elif(content[0]!="="):
            self.formulas[int(response[0])][int(response[1])] = "=" + content
            self.updateCellFromFormulaResult(response)
        else:
            self.formulas[int(response[0])][int(response[1])] = content
            self.updateCellFromFormulaResult(response)
        for c in self.updateBinds:
            if xl_rowcol_to_cell(response[0], response[1]) in self.updateBinds[c]:
                if xl_rowcol_to_cell(response[0], response[1]) not in self.cellRefs:
                    self.updateBinds[c].remove(xl_rowcol_to_cell(response[0], response[1]))
                    pass
                elif c not in self.cellRefs[xl_rowcol_to_cell(response[0], response[1])]:
                    self.updateBinds[c].remove(xl_rowcol_to_cell(response[0], response[1]))
                    pass

    def updateCellFromFormulaResult(self, response):
        if (xl_rowcol_to_cell(response[0], response[1]) in self.updateBinds):
            for updQE in self.updateBinds[xl_rowcol_to_cell(response[0], response[1])]:
                self.updateCellFromFormulaResult(xl_cell_to_rowcol(updQE))
        self.sheet.set_cell_data(response[0], response[1], self.interpret(self.getFormulaForResponse(response)[1:], response))
        self.sheet.refresh(False, False)

    def interpret(self, f, response):
        vinst = re.compile('[\$]?([aA-zZ]+)[\$]?(\d+)')
        rinst = re.compile('([A-Z]{1,2}[0-9]{1,}:{1}[A-Z]{1,2}[0-9]{1,})|(^\$(([A-Z])|([a-z])){1,2}([0-9]){1,}:{1}\$(([A-Z])|([a-z])){1,2}([0-9]){1,}$)|(^\$(([A-Z])|([a-z])){1,2}(\$){1}([0-9]){1,}:{1}\$(([A-Z])|([a-z])){1,2}(\$){1}([0-9]){1,}$)')
        iterv = vinst.finditer(f)
        iterr = rinst.finditer(f)
        varsn = {}
        parspfr = []
        xln = xl_rowcol_to_cell(response[0], response[1])
        refs = []
        for match in iterr:
            cells = []
            values = []
            parspfr.append(match.span())
            c1 = xl_cell_to_rowcol(match.group().split(":")[0])
            c2 = xl_cell_to_rowcol(match.group().split(":")[1])
            if (c1[0]>c2[0] or c1[1]>c2[1]):
                return "RANGE ERROR"
            else:
                for x in range(c1[0], c2[0]+1):
                    for y in range(c1[1], c2[1]+1):
                        cells.append([x, y])
                        refs.append(xl_rowcol_to_cell(x, y))
                        varsn[xl_rowcol_to_cell(x, y)] = self.interpret(self.formulas[x][y][1:], [x, y])
                        values.append(varsn[xl_rowcol_to_cell(x, y)])
                        if(xl_rowcol_to_cell(x, y) not in self.updateBinds):
                            self.updateBinds[xl_rowcol_to_cell(x, y)] = []
                        self.updateBinds[xl_rowcol_to_cell(x, y)].append(xln)
                arrystr = "["
                for value in values:
                    arrystr += (str(value) + ",")
                arrystr = arrystr[:-1]
                arrystr += "]"
                f = f.replace(match.group(), arrystr)
        for match in iterv:
            if(match.group()[0].isalpha()):
                if match.group() == xln:
                    return "RECURSION ERROR"
                else:
                    if(self.checkAlreadyProcessed(parspfr, match)):
                        pass
                    else:
                        refs.append(match.group())
                        varsn[match.group()] = self.interpret(self.formulas[xl_cell_to_rowcol(match.group())[0]][xl_cell_to_rowcol(match.group())[1]][1:], xl_cell_to_rowcol(match.group()))
                        if(match.group() not in self.updateBinds):
                            self.updateBinds[match.group()] = []
                        self.updateBinds[match.group()].append(xln)
        for updc in self.updateBinds:
            self.updateBinds[updc] = list(dict.fromkeys(self.updateBinds[updc]))
        locals().update(varsn)
        if(refs!=[]):
            self.cellRefs[xln]=refs
        try:
            eval(parser.expr(f).compile())
        except:
            return f
        return eval(parser.expr(f).compile())

    def checkAlreadyProcessed(self, parspfr, match):
        for prl in parspfr:
            if (prl[0] <= match.start() <= prl[1]):
                return True
        return False

    def getFormulaForResponse(self, response):
        return self.formulas[int(response[0])][int(response[1])]

    def updateHighlightedCells(self, reset=False):
        for cell in self.highlightedCells:
            if reset:
                self.sheet.dehighlight_cells(row=cell[0], column=cell[1])
            else:
                self.sheet.highlight_cells(row=cell[0], column=cell[1], bg="#add8e6", fg="white")
        if reset:
            self.highlightedCells = []
        self.sheet.refresh(False, False)

    def cell_select(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(response[1], response[2]))
        self.selectedCellSumMean.set("")
        self.updateHighlightedCells(True)
        for bnd in self.updateBinds:
            if xl_rowcol_to_cell(response[1], response[2]) in self.updateBinds[bnd]:
                self.highlightedCells.append(xl_cell_to_rowcol(bnd))
        self.updateHighlightedCells()

    def shift_select_cells(self, response):
        print("update binds:" + str(self.updateBinds))
        print("cell refs:" + str(self.cellRefs))

    def drag_select_cells(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(response[1], response[2]) + ":" + xl_rowcol_to_cell(response[3]-1, response[4]-1))
        self.computeStatsForSelectedCells()

    def computeStatsForSelectedCells(self):
        cells=[]
        contents=[]
        contentsf=[]
        if(len(self.sheet.get_selected_columns())!=0):
            for col in self.sheet.get_selected_columns():
                for y in range(self.sheet.total_rows()):
                    cells.append([y, col])
        if (len(self.sheet.get_selected_rows()) != 0):
            for row in self.sheet.get_selected_rows():
                for x in range(self.sheet.total_columns()):
                    cells.append([row, x])
        for c in self.sheet.get_selected_cells():
            cells.append(c)
        for cell in cells:
            val = self.sheet.get_cell_data(cell[0], cell[1])
            contents.append(val)
            try:
                contentsf.append(float(val))
            except ValueError:
                pass
        if(len(contentsf)!=0):
            self.selectedCellSumMean.set("Sum: " + str(sum(contentsf)) + "\t Mean: " + str(statistics.mean(contentsf)) + "\t Median: " + str(statistics.median(contentsf)) +"\t Mode: " + str(statistics.mode(contents)))
        else:
            self.selectedCellSumMean.set("Mode: " + str(statistics.mode(contents)))

    def open(self):
        filename = askopenfilename(filetypes=[("Flex file","*.flx")])
        self.sheet.set_sheet_data([ [""] * INIT_COLS for _ in range(INIT_ROWS)])
        self.updateBinds={}
        self.openfile = filename
        self.formulas = list(csv.reader(open(filename)))
        for x in range(len(self.formulas)):
            for y in range(len(self.formulas[0])):
                if(self.formulas[x][y]!="=0"):
                    self.updateCellFromFormulaResult((x, y))

    def save(self):
        if(self.openfile==""):
            self.saveas()
        else:
            with open(self.openfile, "w+") as my_csv:
                csvWriter = csv.writer(my_csv, delimiter=',')
                csvWriter.writerows(self.formulas)

    def saveas(self):
        self.openfile = asksaveasfilename(filetypes=[("Flex file","*.flx")])
        with open(self.openfile, "w+") as my_csv:
            csvWriter = csv.writer(my_csv, delimiter=',')
            csvWriter.writerows(self.formulas)

    def importCsv(self):
        filename = askopenfilename(filetypes=[("Comma separated values", "*.csv")])
        with open(filename, "r") as f:
            reader = csv.reader(f)
            x=0
            for row in reader:
                y=0
                for e in row:
                    self.sheet.set_cell_data(x, y, e)
                    y+=1
                x+=1
        self.sheet.refresh(False, False)

    def exportToCsv(self):
        self.openfile = asksaveasfilename(filetypes=[("Comma separated values", "*.csv")])
        with open(self.openfile, "w+") as my_csv:
            csvWriter = csv.writer(my_csv, delimiter=',')
            csvWriter.writerows(self.sheet.get_sheet_data())

    def exportToExcel(self):
        messagebox.showinfo(message="Flex will not export cell formulas due to compatibility issues", title="Warning")
        writefile = asksaveasfilename(filetypes=[("Microsoft excel workbook","*.xlsx")])
        workbook = xlsxwriter.Workbook(writefile)
        worksheet = workbook.add_worksheet()
        for x in range(len(self.formulas)):
            for y in range(len(self.formulas[x])):
                try:
                    worksheet.write(x, y, float(self.sheet.get_cell_data(x, y)))
                except ValueError:
                    worksheet.write(x, y, self.sheet.get_cell_data(x, y))
        workbook.close()

    def ctrl_a(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(response[1], response[2]) + ":" + xl_rowcol_to_cell(response[3] - 1, response[4] - 1))

    def row_select(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(response[1], 0) + ":" + xl_rowcol_to_cell(response[1], INIT_COLS-1))
        self.computeStatsForSelectedCells()

    def shift_select_rows(self, response):
        print(response)

    def drag_select_rows(self, response):
        pass
        # print (response)

    def restart(self):
        os.execl(sys.executable, sys.executable, *sys.argv)

    def column_select(self, response):
        self.selectedCell.set(xl_rowcol_to_cell(0, response[1]) + ":" + xl_rowcol_to_cell(INIT_ROWS-1, response[1]))
        self.computeStatsForSelectedCells()

    def shift_select_columns(self, response):
        print(response)

    def drag_select_columns(self, response):
        pass
        # print (response)


app = flex()
app.mainloop()