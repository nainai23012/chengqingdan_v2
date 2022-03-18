from PyQt5.QtWidgets import QStyledItemDelegate, QDoubleSpinBox, QComboBox

from PyQt5.QtCore import Qt


##from PyQt5.QtWidgets import  QLabel, QAbstractItemView, QFileDialog

##from PyQt5.QtGui import QStandardItemModel, QStandardItem

# ==============基于QDoubleSpinbox的代理组件====================
class QmyFloatSpinDelegate(QStyledItemDelegate):
    def __init__(self, minV=0, maxV=10000, digi=2, parent=None):
        super().__init__(parent)
        self.__min = minV
        self.__max = maxV
        self.__decimals = digi

    ## 自定义代理组件必须继承以下4个函数
    def createEditor(self, parent, option, index):
        editor = QDoubleSpinBox(parent)
        editor.setFrame(False)
        editor.setRange(self.__min, self.__max)
        editor.setDecimals(self.__decimals)
        return editor

    def setEditorData(self, editor, index):
        model = index.model()  # 关联的数据模型
        text = model.data(index, Qt.EditRole)  # 单元格文字
        ##      value = float(index.model().data(index, Qt.EditRole))
        editor.setValue(float(text))

    def setModelData(self, editor, model, index):
        value = editor.value()
        model.setData(index, value, Qt.EditRole)

    ##        digi="{:.%df}"%self.__decimals  # 获得 "{:.2f}", 会改变小数位数
    ##        text=digi.format(value)     #相当于 "{:.2f}.format(value)"
    ##        model.setData(index, text, Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


# ==============基于QComboBox的代理组件====================
class QmyComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.__itemList = []
        self.__isEditable = False

    def setItems(self, itemList, isEditable=False):
        self.__itemList = itemList
        self.__isEditable = isEditable

    # 自定义代理组件必须继承以下4个函数
    def createEditor(self, parent, option, index):
        editor = QComboBox(parent)
        editor.setFrame(False)
        editor.setEditable(self.__isEditable)
        editor.addItems(self.__itemList)
        return editor

    def setEditorData(self, editor, index):
        model = index.model()
        text = model.data(index, Qt.EditRole)
        # text = str(index.model().data(index, Qt.EditRole))
        editor.setCurrentText(text)

    def setModelData(self, editor, model, index):
        text = editor.currentText()
        model.setData(index, text, Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)
