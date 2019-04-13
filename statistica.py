# -*- coding: utf-8 -*-

import sys, csv, os, time
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QMessageBox, QWidget, QPushButton, QLineEdit, QInputDialog, QApplication, QFileDialog
from PyQt5 import QtCore, QtGui, QtWidgets
from docxtpl import DocxTemplate
import json

############# открываем файл конфигурации JSON ##############
with open('config.txt') as json_file:  
    data = json.load(json_file)        

class Ui_Dialog(object):
    def __init__(self):
        self.name = []
        self.sequence = []
        self.context_item = ''
        self.context_item_3 = ''
        self.fname = []
        
      
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(606, 703)
        self.horizontalLayout = QtWidgets.QHBoxLayout(Dialog)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.comboBox = QtWidgets.QComboBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.comboBox.setFont(font)
        self.comboBox.setObjectName("comboBox")
        self.verticalLayout_3.addWidget(self.comboBox)
        self.listWidget = QtWidgets.QListWidget(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.listWidget.setFont(font)
        self.listWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout_3.addWidget(self.listWidget)
        self.line_3 = QtWidgets.QFrame(Dialog)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.verticalLayout_3.addWidget(self.line_3)
        spacerItem = QtWidgets.QSpacerItem(350, 1, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_3.addItem(spacerItem)
        self.horizontalLayout_2.addLayout(self.verticalLayout_3)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_2 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_3.addWidget(self.label_2)
        self.label = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label")
        self.horizontalLayout_3.addWidget(self.label)
        self.verticalLayout_4.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_4 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_5.addWidget(self.label_4)
        self.label_3 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_5.addWidget(self.label_3)
        self.verticalLayout_4.addLayout(self.horizontalLayout_5)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_6 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_6.addWidget(self.label_6)
        self.label_5 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_5.setFont(font)
        self.label_5.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_6.addWidget(self.label_5)
        self.verticalLayout_4.addLayout(self.horizontalLayout_6)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_8 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.horizontalLayout_8.addWidget(self.label_8)
        self.label_7 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_7.setFont(font)
        self.label_7.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_8.addWidget(self.label_7)
        self.verticalLayout_4.addLayout(self.horizontalLayout_8)
        self.line = QtWidgets.QFrame(Dialog)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout_4.addWidget(self.line)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_10 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.horizontalLayout_10.addWidget(self.label_10)
        self.label_9 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_9.setFont(font)
        self.label_9.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_9.setObjectName("label_9")
        self.horizontalLayout_10.addWidget(self.label_9)
        self.verticalLayout_4.addLayout(self.horizontalLayout_10)
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.label_12 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_12.setFont(font)
        self.label_12.setObjectName("label_12")
        self.horizontalLayout_11.addWidget(self.label_12)
        self.label_11 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_11.setFont(font)
        self.label_11.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout_11.addWidget(self.label_11)
        self.verticalLayout_4.addLayout(self.horizontalLayout_11)
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.label_14 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_14.setFont(font)
        self.label_14.setObjectName("label_14")
        self.horizontalLayout_12.addWidget(self.label_14)
        self.label_13 = QtWidgets.QLabel(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_13.setFont(font)
        self.label_13.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout_12.addWidget(self.label_13)
        self.verticalLayout_4.addLayout(self.horizontalLayout_12)
        self.line_2 = QtWidgets.QFrame(Dialog)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.verticalLayout_4.addWidget(self.line_2)
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        self.verticalLayout_4.addLayout(self.horizontalLayout_14)
        self.groupBox = QtWidgets.QGroupBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.groupBox.setFont(font)
        self.groupBox.setObjectName("groupBox")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout(self.groupBox)
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.radioButton = QtWidgets.QRadioButton(self.groupBox)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton.setFont(font)
        self.radioButton.setObjectName("radioButton")
        #self.radioButton.setChecked(True)
        self.horizontalLayout_13.addWidget(self.radioButton)
        self.radioButton_2 = QtWidgets.QRadioButton(self.groupBox)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_2.setFont(font)
        self.radioButton_2.setObjectName("radioButton_2")
        self.horizontalLayout_13.addWidget(self.radioButton_2)
        self.verticalLayout_4.addWidget(self.groupBox)
        self.groupBox_2 = QtWidgets.QGroupBox(Dialog)
        self.groupBox_2.setObjectName("groupBox_2")
        self.horizontalLayout_15 = QtWidgets.QHBoxLayout(self.groupBox_2)
        self.horizontalLayout_15.setObjectName("horizontalLayout_15")
        self.radioButton_3 = QtWidgets.QRadioButton(self.groupBox_2)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_3.setFont(font)
        self.radioButton_3.setObjectName("radioButton_3")
        #self.radioButton_3.setChecked(True)
        self.horizontalLayout_15.addWidget(self.radioButton_3)
        self.radioButton_4 = QtWidgets.QRadioButton(self.groupBox_2)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_4.setFont(font)
        self.radioButton_4.setObjectName("radioButton_4")
        self.horizontalLayout_15.addWidget(self.radioButton_4)
        self.verticalLayout_4.addWidget(self.groupBox_2)
        self.groupBox_3 = QtWidgets.QGroupBox(Dialog)
        self.groupBox_3.setObjectName("groupBox_3")
        self.formLayout = QtWidgets.QFormLayout(self.groupBox_3)
        self.formLayout.setObjectName("formLayout")
        self.radioButton_5 = QtWidgets.QRadioButton(self.groupBox_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_5.setFont(font)
        self.radioButton_5.setObjectName("radioButton_5")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.radioButton_5)
        self.radioButton_6 = QtWidgets.QRadioButton(self.groupBox_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_6.setFont(font)
        self.radioButton_6.setObjectName("radioButton_6")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.radioButton_6)
        self.radioButton_7 = QtWidgets.QRadioButton(self.groupBox_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_7.setFont(font)
        self.radioButton_7.setObjectName("radioButton_7")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.radioButton_7)
        self.radioButton_8 = QtWidgets.QRadioButton(self.groupBox_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.radioButton_8.setFont(font)
        self.radioButton_8.setObjectName("radioButton_8")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.radioButton_8)
        self.verticalLayout_4.addWidget(self.groupBox_3)
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_4.addItem(spacerItem1)
        spacerItem2 = QtWidgets.QSpacerItem(180, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_4.addItem(spacerItem2)
        self.pushButton_3 = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.verticalLayout_4.addWidget(self.pushButton_3)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout_4.addWidget(self.pushButton)
        self.pushButton_2 = QtWidgets.QPushButton(Dialog)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout_4.addWidget(self.pushButton_2)
        self.horizontalLayout_2.addLayout(self.verticalLayout_4)
        self.horizontalLayout.addLayout(self.horizontalLayout_2)
        self.pushButton_3.clicked.connect(self.openCSVfile)
        #self.pushButton_2.clicked.connect(self.close)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        
    
    def openCSVfile(self):
        self.fname = QFileDialog.getOpenFileName()[0]
        f = open(self.fname, 'r')
        self.comboBox.clear()
        self.name = []
        self.sequence = []
        readCSV = ''
        with f as csvfile:
            readCSV = csv.reader(csvfile, delimiter=';')
            for row in readCSV:
                self.name.append(row[data['config'][0]['name']])
                self.sequence.append(row[data['config'][0]['seqence']])
        for item in range(len(self.name)):
            self.comboBox.addItem(self.name[item])
        f.close()
        #self.comboBox.addItem('Select All'))
                       
################# основной функциональный блок ################
        self.pushButton.clicked.connect(self.totemplate)            # по нажатию кнопки Build формируем файл протокола синтеза
        self.comboBox.activated['QString'].connect(self.listWidget.addItem) # передаем текущий олиг из выпадающего списка в общий список олигов
        self.comboBox.activated['QString'].connect(self.getitemlist)  # передаем имя и последовательность олига из выпадающего списка в массивы олигов items и itemseq
        self.radioButton.toggled.connect(self.rbClicked)
        self.radioButton_3.toggled.connect(self.rb3Clicked)
        
        
    def getitemlist(self):
        items = []
        itemseq = []
        #print (self.listWidget.count())
        #print (items)
        #print(itemseq)
        Mon_dA = 0
        Mon_dC = 0
        Mon_dG = 0
        Mon_dT = 0
        Mon_All = 0
        # заполняем массивы имен "items" и последователностей "itemseq" олигов
        for index in range(self.listWidget.count()):
             print (index)
             items.append(str(self.listWidget.item(index).text())) 
             for i in range(len(self.sequence)):
                 if str(self.listWidget.item(index).text()) == str(self.name[i]):
                     itemseq.append(self.sequence[i])
                 '''if str(self.listWidget.item(index).text()) == 'Select All':
                     print (1)
                     itemseq = self.sequence'''
        #print (self.listWidget.count())
        #print (items)
        #print(itemseq)
        # подсчитываем количество оснований в каждом олиге в массиве "itemseq"
        for oligo in range(len(itemseq)):
             for mon in range(len(itemseq[oligo])):
                 Mon_All+=1
                 if str(itemseq[oligo][mon]) == 'a':
                     Mon_dA += 1
                 elif str(itemseq[oligo][mon]) == 'c':
                     Mon_dC += 1
                 elif str(itemseq[oligo][mon]) == 'g':
                     Mon_dG += 1
                 elif str(itemseq[oligo][mon]) == 't':
                     Mon_dT += 1
        # выводим значение количеств оснований на весь синтез 
        self.label.setText(str(Mon_dA))
        self.label_3.setText(str(Mon_dC))
        self.label_5.setText(str(Mon_dG))
        self.label_7.setText(str(Mon_dT))
        self.label_9.setText(str(self.listWidget.count()))
        self.label_11.setText(str(len(max(itemseq)))) # не работает как надо !!!
        self.label_13.setText(str((Mon_All/self.listWidget.count())*3.75)) #Est. time. WRONG!!!

    def rbClicked(self):
        if self.radioButton.isChecked():
            self.context_item = 'Universal'
        else:
            self.context_item = 'ACGT'

    def rb3Clicked(self):
        if self.radioButton.isChecked():
            self.context_item_3 = 'Сухой'
        else:
            self.context_item_3 = 'Суспензия'
           
#################### блок работы с шаблоном документа протокала на синтез ##########################
    def totemplate(self, Mon_All):  
        Mon_All = int(self.label.text())+int(self.label_3.text())+int(self.label_5.text())+int(self.label_7.text())
        context = { 'dA' : str(round((((float(self.label.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol'])*(data['config'][1]['Amd']/data['config'][1]['MeCN'])),5)) + ' г',
                    'MeCN_dA' : str(round(((float(self.label.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol']),4)) + ' мл',
                    'dC' : str(round((((float(self.label_3.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol'])*(data['config'][1]['Amd']/data['config'][1]['MeCN'])),5)) + ' г',
                    'MeCN_dC' : str(round(((float(self.label_3.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol']),4)) + ' мл',
                    'dG' : str(round((((float(self.label_5.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol'])*(data['config'][1]['Amd']/data['config'][1]['MeCN'])),5)) + ' г',
                    'MeCN_dG' : str(round(((float(self.label_5.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol']),4)) + ' мл',
                    'dT' : str(round((((float(self.label_7.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol'])*(data['config'][1]['Amd']/data['config'][1]['MeCN'])),5)) + ' г',
                    'MeCN_dT' : str(round(((float(self.label_7.text())*data['config'][1]['VolOneBase'])+data['config'][1]['DeadVol']),4)) + ' мл',
                    'THF_Ox' : str(round(float(((Mon_All*data['config'][2]['VolOneBase'])+data['config'][2]['DeadVol'])*((data['config'][2]['THF'])/(data['config'][2]['THF']+data['config'][2]['Py']+data['config'][2]['H2O']))),5)) + ' мл',
                    'Py_Ox' : str(round(float(((Mon_All*data['config'][2]['VolOneBase'])+data['config'][2]['DeadVol'])*((data['config'][2]['Py'])/(data['config'][2]['THF']+data['config'][2]['Py']+data['config'][2]['H2O']))),5)) + ' мл',
                    'H2O_Ox' : str(round(float(((Mon_All*data['config'][2]['VolOneBase'])+data['config'][2]['DeadVol'])*((data['config'][2]['H2O'])/(data['config'][2]['THF']+data['config'][2]['Py']+data['config'][2]['H2O']))),5)) + ' мл',
                    'I2_Ox' : str(round(float(((Mon_All*data['config'][2]['VolOneBase'])+data['config'][2]['DeadVol'])*((data['config'][2]['I2'])/(data['config'][2]['THF']+data['config'][2]['Py']+data['config'][2]['H2O']))),5)) + ' г',
                    'V_Ox' : str(round(float((Mon_All*data['config'][2]['VolOneBase'])+data['config'][2]['DeadVol']),5)) + ' мл',
                    'MeCN_Act' : str(round(float((Mon_All*data['config'][3]['VolOneBase'])+data['config'][3]['DeadVol']),5)) + ' мл',
                    'TET_Act' : str(round(float(((Mon_All*data['config'][3]['VolOneBase'])+data['config'][3]['DeadVol'])*(data['config'][3]['TET']/data['config'][3]['MeCN'])),5)) + ' г',
                    'DCE_Dbl' : str(round(float(((Mon_All*data['config'][4]['VolOneBase'])+data['config'][4]['DeadVol'])*(data['config'][4]['DCE']/(data['config'][4]['DCE']+data['config'][4]['DCA']))),5)) + ' мл',
                    'DCA_Dbl' : str(round(float(((Mon_All*data['config'][4]['VolOneBase'])+data['config'][4]['DeadVol'])*(data['config'][4]['DCA']/(data['config'][4]['DCE']+data['config'][4]['DCA']))),5)) + ' мл',
                    'V_Dbl' : str(round(float((Mon_All*data['config'][4]['VolOneBase'])+data['config'][4]['DeadVol']),5)) + ' мл',
                    'THF_CPA' : str(round(float(((Mon_All*data['config'][5]['VolOneBase'])+data['config'][5]['DeadVol'])*(data['config'][5]['THF']/(data['config'][5]['THF']+data['config'][5]['Anhydride']+data['config'][5]['Py']))),5)) + ' мл',
                    'ANH_CPA' : str(round(float(((Mon_All*data['config'][5]['VolOneBase'])+data['config'][5]['DeadVol'])*(data['config'][5]['Anhydride']/(data['config'][5]['THF']+data['config'][5]['Anhydride']+data['config'][5]['Py']))),5)) + ' мл',
                    'Py_CPA' : str(round(float(((Mon_All*data['config'][5]['VolOneBase'])+data['config'][5]['DeadVol'])*(data['config'][5]['Py']/(data['config'][5]['THF']+data['config'][5]['Anhydride']+data['config'][5]['Py']))),5)) + ' мл',
                    'V_CPA' : str(round(float((Mon_All*data['config'][5]['VolOneBase'])+data['config'][5]['DeadVol']),5)) + ' мл',
                    'THF_CPB' : str(round(float(((Mon_All*data['config'][6]['VolOneBase'])+data['config'][6]['DeadVol'])*(data['config'][6]['THF']/(data['config'][6]['THF']+data['config'][6]['MeIm']))),5)) + ' мл',
                    'MeIm_CPB' : str(round(float(((Mon_All*data['config'][6]['VolOneBase'])+data['config'][6]['DeadVol'])*(data['config'][6]['MeIm']/(data['config'][6]['THF']+data['config'][6]['MeIm']))),5)) + ' мл',
                    'V_CPB' : str(round(float((Mon_All*data['config'][6]['VolOneBase'])+data['config'][6]['DeadVol']),5)) + ' мл',
                    'NH3_AMA' : str(self.listWidget.count()*0.5) + ' мл',
                    'MeNH2_AMA' : str(self.listWidget.count()*0.5) + ' мл',
                    'V_AMA' : str(self.listWidget.count()) + ' мл',
                    'CPG' : self.context_item,
                    'CPG_state' : self.context_item_3}
        doc = DocxTemplate("template.docx")
        doc.render(context)
        doc.save("Протокол синтеза олигонуклеотидов.docx")
        # всплывающие окно информации о сборке файла протокола синтеза
        infoBox = QMessageBox()
        infoBox.setIcon(QMessageBox.Information)
        infoBox.setText("Протокол синтеза готов")
        infoBox.setWindowTitle("Информация")
        infoBox.setStandardButtons(QMessageBox.Ok)
        infoBox.exec_()

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label_2.setText(_translate("Dialog", "A:"))
        self.label.setText(_translate("Dialog", "0"))
        self.label_4.setText(_translate("Dialog", "C:"))
        self.label_3.setText(_translate("Dialog", "0"))
        self.label_6.setText(_translate("Dialog", "G:"))
        self.label_5.setText(_translate("Dialog", "0"))
        self.label_8.setText(_translate("Dialog", "T:"))
        self.label_7.setText(_translate("Dialog", "0"))
        self.label_10.setText(_translate("Dialog", "Count oligo:"))
        self.label_9.setText(_translate("Dialog", "0"))
        self.label_12.setText(_translate("Dialog", "Max base:"))
        self.label_11.setText(_translate("Dialog", "0"))
        self.label_14.setText(_translate("Dialog", "Est. time:"))
        self.label_13.setText(_translate("Dialog", "0"))
        self.groupBox.setTitle(_translate("Dialog", "CPG-type"))
        self.radioButton.setText(_translate("Dialog", "Unylink"))
        self.radioButton_2.setText(_translate("Dialog", "ACGT"))
        self.groupBox_2.setTitle(_translate("Dialog", "CPG-state"))
        self.radioButton_3.setText(_translate("Dialog", "Dry"))
        self.radioButton_4.setText(_translate("Dialog", "Suspension"))
        self.groupBox_3.setTitle(_translate("Dialog", "Deblock"))
        self.radioButton_5.setText(_translate("Dialog", "NH3"))
        self.radioButton_6.setText(_translate("Dialog", "NH3+T"))
        self.radioButton_7.setText(_translate("Dialog", "AMA"))
        self.radioButton_8.setText(_translate("Dialog", "AMA+T"))
        self.pushButton_3.setText(_translate("Dialog", "Open CSV"))
        self.pushButton.setText(_translate("Dialog", "Built DOC"))
        self.pushButton_2.setText(_translate("Dialog", "Exit"))




if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
