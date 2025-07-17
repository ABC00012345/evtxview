import typing

import sys
from pathlib import Path

from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QBrush, QColor
from evtx import PyEvtxParser
from lxml import etree


class XmlElement:
    def __init__(self, root):
        self.__children = dict()
        for child in root:
            # remove namespace
            idx = child.tag.rfind('}')
            self.__children[child.tag[idx+1:]] = XmlElement(child)
        self.__node = root

    @property
    def text(self):
        return self.__node.text

    @property
    def attrib(self):
        return self.__node.attrib

    def __getitem__(self, item):
        return self.__children[item]

class EventRecord:
    def __init__(self, record):
        self.__id = record['event_record_id']
        self.__timestamp = record['timestamp']
        self.__data = record['data']
        self.__parsed = False
        self.__attrib = dict()

    @property
    def id(self):
        return self.__id

    @property
    def timestamp(self):
        return self.__timestamp

    @property
    def data(self):
        return self.__data

    @property
    def EventID(self):
        self.__parse_data()
        return self.__attrib["EventID"]

    @property
    def Provider(self):
        self.__parse_data()
        return self.__attrib["Provider"]

    @property
    def Level(self):
        levels = {
            0: "LogAlways",
            1: "Critical",
            2: "Error",
            3: "Warning",
            4: "Informational",
            5: "Verbose"
        }
        self.__parse_data()
        level_id = int(self.__attrib["Level"])
        return levels[level_id]

    def __parse_data(self):
        if self.__parsed:
            return
        # remove processing instruction
        idx = self.__data.find("?>")
        if idx != -1:
            self.__data = self.__data[idx+2:]
        root = XmlElement(etree.XML(self.__data))

        self.__attrib["EventID"] = root["System"]["EventID"].text
        self.__attrib["Level"] = root["System"]["Level"].text
        self.__attrib["Provider"] = root["System"]["Provider"].attrib["Name"]



class EvtxViewModel(QAbstractItemModel):
    def __init__(self, filename):
        super(QAbstractItemModel, self).__init__()
        self.__filename = filename
        self.__chunks = 0
        self.__columns = [
            ('ID', lambda r: r.id),
            ('TimeCreated', lambda r: r.timestamp),
            ('Provider', lambda r: r.Provider),
            ('EventID', lambda r: r.EventID),
            ('Level', lambda r: r.Level),
            ('Data', lambda r: r.data)
        ]

        #for idx in range(0, len(self.__columns)):
        #    self.setHeaderData(idx, Qt.Horizontal, self.__columns[0][0], Qt.DisplayRole)

        self.__records = dict()
        self.__record_ids = list()
        self.load_data()

        self._highlighted_row = None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.__columns[section][0]
        return super().headerData(section, orientation, role)


    def load_data(self):
        self.__parser = PyEvtxParser(self.__filename)
        for record in self.__parser.records():
            self.__records[record["event_record_id"]] = EventRecord(record)
            self.__record_ids.append(record["event_record_id"])
        self.__record_ids.sort()

    def rowCount(self, parent):
        return len(self.__record_ids)

    def columnCount(self, parent):
        return len(self.__columns)

    def index(self, row: int, column: int, parent: QModelIndex = ...) -> QModelIndex:
        return self.createIndex(row, column)

    def set_highlighted_row(self, row):
        old_row = self._highlighted_row
        self._highlighted_row = row
        # Emit dataChanged for old and new rows so the view updates
        if old_row is not None:
            topLeft_old = self.index(old_row, 0)
            bottomRight_old = self.index(old_row, self.columnCount(None) - 1)
            self.dataChanged.emit(topLeft_old, bottomRight_old, [Qt.BackgroundRole])

        if row is not None:
            topLeft_new = self.index(row, 0)
            bottomRight_new = self.index(row, self.columnCount(None) - 1)
            self.dataChanged.emit(topLeft_new, bottomRight_new, [Qt.BackgroundRole])

    def remove_highlight(self):
        self.set_highlighted_row(None)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return QVariant()

        if role == Qt.DisplayRole:
            row = index.row()
            record_id = self.__record_ids[row]
            record = self.__records[record_id]
            column = self.__columns[index.column()]
            return column[1](record)

        if role == Qt.BackgroundRole:
            if index.row() == self._highlighted_row:
                color = QColor("#FFFACD")
                color.setAlpha(128)  # 0 = fully transparent, 255 = fully opaque
                return QBrush(color)

        return QVariant()

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        return Qt.ItemIsEnabled | Qt.ItemNeverHasChildren

    def parent(self, qmodelindex=None):
        return QModelIndex()
    
    def get_records(self):
        return [self.__records[rid] for rid in self.__record_ids]


class EvtxView(QTableView):
    def __init__(self, filename):
        super(QTableView, self).__init__()

        self.__evtxViewModel = EvtxViewModel(filename)

        self.setModel(self.__evtxViewModel)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().hide()
        self.setShowGrid(False)

        self._highlighted_row = None

        # Connect the clicked signal to the handler method (opens then msg box woth the data)
        self.doubleClicked.connect(self.on_table_clicked)

    def get_evtxViewModel(self):
        return self.__evtxViewModel
    

    def scroll_to_record_id(self, record_id):
        model = self.get_evtxViewModel()
        try:
            row = model._EvtxViewModel__record_ids.index(record_id)
        except ValueError:
            print(f"Record ID {record_id} not found")
            return

        # Update highlighted row in the model
        model.set_highlighted_row(row)

        index = model.index(row, 0)
        if index.isValid():
            self.scrollTo(index, QAbstractItemView.PositionAtCenter)
            self.setCurrentIndex(index)


    def on_table_clicked(self, index: QModelIndex):
        # Check if the clicked column is the 'Data' column
        # Assuming 'Data' is last column (or find its index dynamically)
        data_column_index = None
        for i, (name, _) in enumerate(self.__evtxViewModel._EvtxViewModel__columns):
            if name == "Data":
                data_column_index = i
                break

        if index.column() == data_column_index:
            row = index.row()
            record_id = self.__evtxViewModel._EvtxViewModel__record_ids[row]
            record = self.__evtxViewModel._EvtxViewModel__records[record_id]
            # Now do something with the 'Data' cell clicked:
            print(f"Data cell clicked at row {row}: {record.data}")
            QMessageBox.information(self, f"Event Data for Event with Record id: {record_id}", record.data)



class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi("layout.ui", self)

        action = self.findChild(QAction, "actionExit")
        action.triggered.connect(self.action_exit)

        action = self.findChild(QAction, "actionOpen")
        action.triggered.connect(self.action_open)

        action = self.findChild(QAction, "actionSearch")
        action.triggered.connect(self.action_search)

        self.__tab_widget = self.findChild(QTabWidget, "tabWidget")
        self.__tab_widget.removeTab(0)
        self.__tab_widget.removeTab(0)
        self.__tab_widget.tabCloseRequested.connect(lambda idx: self.close_tab(idx))

        self.__files = dict()

        ## search in evt data toggle
        self.searchInEVTData = self.findChild(QAction, "actionSearchInEVTData")

        ## for search feature
        self.search_msgbox : QMessageBox
        self.cur_evtx_view : EvtxView
        self.cur_search_index = 0
        self.found = []

    def open_file(self, filename: str):
        index = -1
        if filename in self.__files.keys():
            index = self.__files[filename]
        else:
            new_tab = QWidget()
            layout = QVBoxLayout()

            evtx_view = EvtxView(filename)
            layout.addWidget(evtx_view)
            new_tab.setLayout(layout)

            # Store the EvtxView instance on the QWidget for later access
            new_tab.evtx_view = evtx_view

            index = self.__tab_widget.addTab(new_tab, Path(filename).name)
            self.__files[filename] = index

        assert index != -1
        self.__tab_widget.setCurrentIndex(index)


    def close_tab(self, index):
        self.__tab_widget.removeTab(index)
        tmp = dict()
        for filename, idx in self.__files.items():
            if idx != index:
                tmp[filename] = self.__files[filename]
        self.__files = tmp

    def action_exit(self):
        self.close()

    def action_open(self):
        dlg = QFileDialog()
        dlg.setAcceptMode(QFileDialog.AcceptOpen)
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setNameFilter("Windows event log files (*.evtx)")
        filenames = list()
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            assert len(filenames) == 1
            self.open_file(filenames[0])

    ## menu tab
    def action_search(self):
        if not self.__files:
            QMessageBox.information(self, "Search", f"No file opened!")
            return
        
        text, ok = QInputDialog.getText(self, "Search", "Enter search term:")
        if ok and text:
            #QMessageBox.information(self, "Search", f"You searched for: {text}")
            current_widget = self.__tab_widget.currentWidget()
            print(current_widget)
            print(current_widget.evtx_view)
            evtx_view : EvtxView = current_widget.evtx_view
            model : EvtxViewModel = evtx_view.get_evtxViewModel()

            ## search
            found = []
            for record in model.get_records():
                record : EventRecord
                # (text in record.data)
                if (text in record.Provider) or (text in record.EventID) or (self.searchInEVTData.isChecked() and (text in record.data)):
                    #found.append(f"Record ID: {record.id} Time: {record.timestamp}")
                    found.append([record.id, record.timestamp])

            if found:
                self.found = found
                self.cur_evtx_view = evtx_view

                #dialog = SearchResultsDialog(found, evtx_view)
                #dialog.show()  # Non-modal, main window stays usable
                self.display_search_results()
            else:
                QMessageBox.information(self, "Search Results", "Nothing found.")

    def display_search_results(self):
        self.cur_search_index = 0

        self.search_msgbox = QMessageBox(self)
        self.search_msgbox.setWindowTitle("Search Results")
        self.search_msgbox.setText(
            f"Found: {len(self.found)}, Current: {self.cur_search_index + 1}"
        )

        prev_button = self.search_msgbox.addButton("Previous", QMessageBox.ActionRole)
        next_button = self.search_msgbox.addButton("Next", QMessageBox.ActionRole)
        close_button = self.search_msgbox.addButton("Close", QMessageBox.ActionRole)

        # Disconnect any default signals (good practice)
        try:
            prev_button.clicked.disconnect()
            next_button.clicked.disconnect()
        except TypeError:
            pass  # Already disconnected

        def close_searchbox():
            cur_evtxViewModel : EvtxViewModel = self.cur_evtx_view.get_evtxViewModel()
            cur_evtxViewModel.remove_highlight()
            self.search_msgbox.close()

        # Reconnect with your desired behavior
        prev_button.clicked.connect(lambda: self.navigate_search(-1))
        next_button.clicked.connect(lambda: self.navigate_search(1))
        close_button.clicked.connect(close_searchbox)

        self.search_msgbox.open()


    def navigate_search(self, direction):
        if direction == -1:
            self.cur_search_index -= 1
            if self.cur_search_index<0:
                self.cur_search_index=len(self.found)-1
        elif direction == 1:
            self.cur_search_index += 1
            self.cur_search_index%=len(self.found)

        self.cur_evtx_view.scroll_to_record_id(self.found[self.cur_search_index][0])
        self.search_msgbox.setText(
            f"Found: {len(self.found)}, Current: {self.cur_search_index + 1}"
        )



def run_app():
    app = QApplication([])

    wnd_main = MainWindow()
    wnd_main.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    print("evtx viewer - mod by @abc00012345")
    run_app()
