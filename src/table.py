from tkinter import *
from tkinter.ttk import *
from copy import deepcopy

class Table:
    def __init__(self, window, scrollbar=True, height=20):
        """Contruct elements of the table."""
        self.__frame = Frame(window)    # create a frame to contain tableiew and scrollbar
        self.__table = Treeview(self.__frame, selectmode="browse", height=height) # create a treeview as a table
        self.__table.pack(side = LEFT)

        if scrollbar:   # create a scrollbar for the treeview
            vScrollbar = Scrollbar(self.__frame, command=self.__table.yview)
            vScrollbar.pack(side = RIGHT, fill = Y)
            self.__table.configure(yscrollcommand = vScrollbar.set)

    def construct(self, header, width, idcolumn=True, anchor="c"):
        """Contruct structure of table."""
        self.__idColumn = idcolumn
        if self.__idColumn:
            header.insert(0, "ID")  # insert id field into header
            width.insert(0, 25)     # id columm's width is 25
        index = list(map(str, range(1, len(header) + 1)))   # list indexes for header ("1" -> number of headers)
        self.__table["columns"] = index
        self.__table["show"] = "headings"

        for i in range(0, len(header)):     # available for use
            self.__table.column(index[i], width=width[i], anchor=anchor)
            self.__table.heading(index[i], text=header[i])

    def fill(self, datasource):
        """Fill data in table. Must construct structure of table first."""
        datasource_cpy = deepcopy(datasource)
        for i in range(0, len(datasource_cpy)):
            if self.__idColumn:
                datasource_cpy[i].insert(0, str(i + 1))     # id field - first column
            self.__table.insert("", END, values=(datasource_cpy[i]))     # insert each record into table

    def bind(self, event, function):
        """Bind an event to a function"""
        self.__table.bind(event, function)
    
    def focus(self):
        """Get current row"""
        return self.__table.focus()

    def item(self, object, key=None):
        """Analyze current row"""
        return self.__table.item(object, key)
    
    def delete(self, item):
        """Delete an item in table"""
        self.__table.delete(item)
    
    def set(self, item, column=None, value=None):
        """Set value for a cell in table"""
        self.__table.set(item, column=column, value=value)

    def get_children(self):
        """Get children from treeview"""
        return self.__table.get_children()

    def grid(self, column=0, columnspan=1, sticky=NW, padx=0, pady=0, row=0, rowspan=1):
        """Position a table by grid."""
        self.__frame.grid(column=column, columnspan=columnspan, sticky=sticky, padx=padx, pady=pady, row=row, rowspan=rowspan)

    def pack(self, expand=True, fill=NONE, side=TOP):
        """Position a table by pack."""
        self.__frame.pack(expand=expand, fill=fill, side=side)

    def place(self, anchor=NW, bordermode=INSIDE, x=0, y=0):
        """Position a table by place."""
        self.__frame.place(anchor=anchor, bordermode=bordermode, x=x, y=y)