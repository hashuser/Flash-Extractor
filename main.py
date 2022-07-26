import zipfile
import os
import sys
import shutil
from PyQt6 import QtCore, QtGui, QtWidgets
import multiprocessing
import threading


def unzip_file(input_path,output_path):
    filename = input_path[input_path.rfind("\\") + 1:]
    if '.pptx' not in filename:
        raise Exception('Please make sure file is in .pptx format')
    new_output_path = output_path+'\\temp_'+filename
    z = zipfile.ZipFile(input_path, 'r')
    os.makedirs(new_output_path, exist_ok=True)
    for file in z.namelist():
        z.extract(file, new_output_path)
    if not get_activex(new_output_path+'\\ppt\\activeX',output_path+'\\final_'+filename):
        shutil.rmtree(new_output_path)
        raise Exception('There is no Flash (SWF) in this PowerPoint')
    shutil.rmtree(new_output_path)
    return 1


def get_activex(input_path, output_path):
    try:
        files = os.listdir(input_path)
    except Exception as error:
        return 0
    has_activex = False
    for file in files:
        if '.xml' in file and 'active' in file:
            with open(input_path+'\\'+file, 'r') as f:
                content = f.read()
            if 'D27CDB6E-AE6D-11CF-96B8-444553540000' in content:
                decode_activex(input_path+'\\'+file[:file.rfind('.')] + '.bin',output_path,file[:file.rfind('.')])
                has_activex = True
    return has_activex


def decode_activex(input_path,output_path,filename):
    with open(input_path,'rb') as f:
        content = f.read()
    position = content.find(b'\x46\x57\x53')
    s = content[position+4:position+8].hex()
    t = ''
    for x in range(len(s)-1,0,-2):
        t += s[x-1] + s[x]
    length = int(t,16)
    os.makedirs(output_path,exist_ok=True)
    with open(output_path+'\\'+filename+'.swf','wb') as f:
        f.write(content[position:position+length])


def worker(pipe_r, pipe_s):
    while True:
        input_path, output_path = pipe_r.recv()
        input_path = input_path.replace('/','\\')
        if input_path == "kill":
            break
        try:
            unzip_file(input_path, output_path)
        except Exception as error:
            error = str(error)
            if 'Please make sure file is in .pptx format' in error:
                pipe_s.send((input_path, 'Fail', 'Not .pptx'))
            elif 'There is no Flash (SWF) in this PowerPoint' in error:
                pipe_s.send((input_path, 'Fail', 'No SWF'))
            else:
                pipe_s.send((input_path, 'Fail', 'Unknown'))
        else:
            pipe_s.send((input_path, 'Success', ''))
    pipe_s.send(("kill", "kill", "kill"))


class Personalized:
    def __init__(self):
        self._text = None
        self._text_font = None
        self._text_size = None
        self._text_color = None

    def init(self):
        if self._text:
            if not self._text_font:
                if self.is_contains_chinese(self._text):
                    self._text_font = "Microsoft JhengHei"
                else:
                    self._text_font = "Calibri"
            if not self._text_size:
                self._text_size = 10
            if not self._text_color:
                self._text_color = "black"
            self.setFont(QtGui.QFont(self._text_font, self._text_size))
            self.setStyleSheet("color:" + self._text_color)
            self.setText(self._text)

    def centre(self):
        if not self.isVisible():
            raise Exception('Please call show() before calling centre()')
        self.move(int((self.parent().geometry().width() - self.width()) / 2),
                  int((self.parent().geometry().height() - self.height()) / 2))

    def centre_x(self, y: int = None, p: float = None):
        if not self.isVisible():
            raise Exception('Please call show() before calling centre()')
        if not y:
            y = int(self.y())
        if not p:
            p = 2
        self.move(int((self.parent().geometry().width() - self.width()) / p), y)

    def centre_y(self, x: int = None, p: float = None):
        if not self.isVisible():
            raise Exception('Please call show() before calling centre()')
        if not x:
            x = int(self.x())
        if not p:
            p = 2
        self.move(x, int((self.parent().geometry().height() - self.height()) / p))

    @staticmethod
    def is_contains_chinese(s):
        if not s:
            return False
        for _char in s:
            if '\u4e00' <= _char <= '\u9fa5':
                return True
        return False


class MyQLabel(QtWidgets.QLabel, Personalized):
    def __init__(self,
                 _parent: QtWidgets.QWidget | None = None,
                 _text: str = None,
                 _text_font: str = None,
                 _text_size: int = None,
                 _text_color: str = None):
        super().__init__()
        self._text = _text
        self._text_font = _text_font
        self._text_size = _text_size
        self._text_color = _text_color
        self.setParent(_parent)
        self.init()


class MyPushbutton(QtWidgets.QPushButton, Personalized):
    def __init__(self,
                 _parent: QtWidgets.QWidget | None = None,
                 _text: str = None,
                 _text_font: str = None,
                 _text_size: int = None,
                 _text_color: str = None,
                 _button_size: tuple = None):
        super().__init__()
        self._text = _text
        self._text_font = _text_font
        self._text_size = _text_size
        self._text_color = _text_color
        self.setParent(_parent)
        self.setMinimumSize(_button_size[0],_button_size[1])
        self.init()


class MyWidgets(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.text = None
        self.table_items = set()

    def dropEvent(self, e):
        table = self.findChild(QtWidgets.QTableWidget, "main_table")
        pushbutton1 = self.findChild(QtWidgets.QPushButton, "main_pushbutton")
        pushbutton2 = self.findChild(QtWidgets.QPushButton, "clear_pushbutton")
        for x in e.mimeData().urls():
            url = str(x.url())[8:]
            if url in self.table_items:
                continue
            else:
                self.table_items.add(url)
                items = table.findItems(url, QtCore.Qt.MatchFlag.MatchExactly)
                for x in items:
                    table.removeRow(x.row())
            if table.rowCount() == 14:
                self.setMinimumSize(680, 580)
                self.setMinimumSize(680, 580)
                table.setMinimumSize(680, 480)
                table.setMinimumSize(680, 480)
            if table.rowCount() >= 14:
                shift = 250
            else:
                shift = 230
            table.setRowCount(table.rowCount()+1)
            table.setItem(table.rowCount() - 1, 0, QtWidgets.QTableWidgetItem(url))
            table.setItem(table.rowCount() - 1, 1, QtWidgets.QTableWidgetItem("Pending"))
            item_width = len(url)*6
            if table.width() - shift > item_width > table.columnWidth(0):
                table.setColumnWidth(0,item_width)
            elif item_width >= table.width() - shift:
                table.setColumnWidth(0, table.width() - shift)
            pushbutton1.centre_x(510, 4)
            pushbutton2.centre_x(510, 1.33)
        self.text.hide()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            if not self.text:
                self.text = MyQLabel(self.findChild(QtWidgets.QTableWidget,"main_table"),
                                     _text="Drop file here",
                                     _text_size=30,
                                     _text_color="rgb(192,192,192)")
            self.text.show()
            self.text.centre()
            e.accept()

    def dragLeaveEvent(self, e):
        self.text.hide()


class SWF_GUI(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.w = MyWidgets()
        self.w.setWindowTitle("SWF Extractor v3")
        self.w.setMinimumSize(640, 580)
        self.w.setMaximumSize(640, 580)
        self.table = QtWidgets.QTableWidget(self.w)
        self.table.setMinimumSize(640, 480)
        self.table.setMaximumSize(640, 480)
        self.table.setObjectName("main_table")
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Path', 'Status', 'Remark'])
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.editTriggers(self.table).NoEditTriggers)
        self.table.show()
        self.table.setColumnWidth(0, self.table.width() - 220)
        self.pushbutton1 = MyPushbutton(self.w, _text="Start", _text_size=15, _button_size=(200,30))
        self.pushbutton1.setObjectName("main_pushbutton")
        self.pushbutton1.clicked.connect(self.run)
        self.pushbutton1.setEnabled(True)
        self.pushbutton1.show()
        self.pushbutton2 = MyPushbutton(self.w, _text="Clear", _text_size=15, _button_size=(200,30))
        self.pushbutton2.setObjectName("clear_pushbutton")
        self.pushbutton2.clicked.connect(self.clear)
        self.pushbutton2.setEnabled(True)
        self.pushbutton2.show()
        self.w.show()
        self.pushbutton1.centre_x(510, 4)
        self.pushbutton2.centre_x(510, 1.33)

    def run(self):
        print('run')
        self.pushbutton1.setEnabled(False)
        pipe_r1, pipe_s1 = multiprocessing.Pipe(False)
        pipe_r2, pipe_s2 = multiprocessing.Pipe(False)
        P = multiprocessing.Process(target=worker,args=(pipe_r1, pipe_s2,))
        P.start()
        while self.w.table_items:
            pipe_s1.send((self.w.table_items.pop(), os.path.abspath(os.path.dirname(sys.argv[0]))))
        pipe_s1.send(("kill", "kill"))
        T = threading.Thread(target=self.accept, args=(pipe_r2,))
        T.start()

    def clear(self):
        while self.table.rowCount():
            self.table.removeRow(0)
        self.w.table_items.clear()

    def accept(self, pipe_r2):
        while True:
            path, status, remark = pipe_r2.recv()
            path = path.replace('\\','/')
            if path == "kill":
                self.pushbutton1.setEnabled(True)
                break
            row = self.table.findItems(path,QtCore.Qt.MatchFlag.MatchExactly)[0].row()
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(status))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(remark))


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('windowsvista')
    GUI = SWF_GUI()
    sys.exit(app.exec())


