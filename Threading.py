#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Aug  29 17:40:44 2018

@author: gio
"""

import numpy as np

import serial
import glob
import datetime
#from serial import Serial
import sys

from PyQt5.QtWidgets import (QDialog, QApplication, QWidget,  QPushButton,
                             QComboBox, QVBoxLayout, QCheckBox, QLCDNumber,
                             QSlider, QProgressBar, QHBoxLayout, QLabel)
from PyQt5 import QtCore
from PyQt5.QtCore import QThread, pyqtSignal

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt

#import random, time

timespan=600
load=[0]*timespan  
production=[0]*timespan
pdc=[0]*timespan




class WorkerThread(QThread):
    #mysignal_i=pyqtSignal( int, name='Signal_i') ### 1) declare the signal
    measurements_signals=pyqtSignal(int, int, name='m_signals') ### 1) declare the signal

    def __init__(self, parent=None):
        QThread.__init__(self)
        #super(WorkerThread, self).__init__(parent)
   
    def run(self):
        print("reading")
        ser = serial.Serial(
            port='/dev/ttyUSB0', # use "dmesg | grep tty" to find out the port
            baudrate=9600,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE
        ) 
        
        datastring=""
        while ser.isOpen():
            c=ser.read(size=1).decode("ASCII") #better ASCII
            if(c=='\n'): #arduino terminate strings with \n\r
                print("received: "+datastring)
                self.writeData(datastring)
                values = datastring.split(",")
                print(values[1])
                self.measurements_signals.emit(int(values[3]), int(values[1]))
                datastring=""
                #update_chart(int(values[0]), int(values[1]), int(values[2]), int(values[3]))
            else:
                datastring=datastring+c#.decode('utf-8')
            #print datetime.utcnow().isoformat(), datastring   
            if self.isInterruptionRequested():
               print ("exit while loop of reading serial")
               ser.close()
               self.terminate() 
        ser.close()

    def writeData(self, value):
        # Get the current data
        now = datetime.datetime.now()
        today = datetime.date.today() 
        today=now.strftime("%Y-%m-%d")
        t=now.strftime("%Y-%m-%d %H:%M:%S")
        # Open log file 2012-6-23.log and append
        logline=t+","+value+'\n'
        print(logline) 
#        with open('/home/gio/python_prove/'+str(today)+'.csv', 'a') as f: 
#            f.write(logline) 
#            f.close()
        
    def stop(self):
        #self.ser.close()
        self.terminate()    
        print("stop")



class Window(QWidget):
    def __init__(self, parent=None):
        super(Window, self).__init__(parent)
        
        
        self.setGeometry(100, 100, 800, 600) #pos pos width height
                
                
        self.PortLabel=QLabel("Port:")
        self.LabelProd=QLabel("Production:")
        self.LabelLoad=QLabel("Load:")
        
        
        port_selectBox = QComboBox(self)
        ports=self.available_serial_ports()
        
        for port in ports:
            port_selectBox.addItem(port)
        
        
        self.buttonConnect = QPushButton('Connect')
        #self.button.clicked.connect(self.plot)
        
        
        self.b1 = QCheckBox("SerialRead")
        self.b1.setChecked(True)
        #self.b1.stateChanged.connect(self.myThread.btnstate(self.b1))
        
        self.b2 = QCheckBox("activateLog")
        self.b2.setChecked(True)
        #self.b2.stateChanged.connect(lambda:self.btnstate(self.b2))
        
        
        self.figure_bar        = plt.figure()
        self.figure_timeseries = plt.figure()
  
        self.canvas_bar        = FigureCanvas(self.figure_bar)
        self.canvas_timeseries = FigureCanvas(self.figure_timeseries)
        
        # this is the Navigation widget
        # it takes the Canvas widget and a parent
        #self.toolbar = NavigationToolbar(self.canvas, self)

        # set the layout
        
        #self.b1.stateChanged.connect(self.SerialRead(self,b1))
        self.b1.stateChanged.connect(lambda:self.SerialRead(self.b1))
        self.b2.stateChanged.connect(self.SerialLog)


        self.lcdProd = QLCDNumber(self)
        self.lcdLoad = QLCDNumber(self)
        self.lcdProd.setFixedHeight(100)
        self.lcdLoad.setFixedHeight(100)
        
       
        #--------------------------- Layout
        
        col1=QVBoxLayout()
        col1.addWidget(self.PortLabel)
        col1.addWidget(port_selectBox)
        col1.addWidget(self.buttonConnect)
        col1.addWidget(self.b1)
        col1.addWidget(self.b2)
        
        col2=QVBoxLayout()
        col2.addWidget(self.LabelProd)
        col2.addWidget(self.lcdProd)
        col2.addWidget(self.LabelLoad)
        col2.addWidget(self.lcdLoad)
        
        
        toprow=QHBoxLayout()
        toprow.addLayout(col1)
        toprow.addLayout(col2)
        toprow.addWidget(self.canvas_bar)
        
        layout = QVBoxLayout()
        layout.addLayout(toprow)
        layout.addWidget(self.canvas_timeseries)
        
        #layout.addWidget(self.toolbar)
        #layout.addWidget(self.button)
        self.setLayout(layout)
        
        #---------------------------------------------------
        
        
        self.wt=WorkerThread() # This is the thread object
        self.wt.start()
        # Connect the signal from the thread to the slot_method
        self.wt.measurements_signals.connect(self.slot_method)  ### 3) connect to the slot
        app.aboutToQuit.connect(self.wt.stop) #to stop the thread when closing the GUI
        
        timespan=600
        load=[0]*timespan  
        production=[550]*timespan
        pdc=[0]*timespan
        
    def slot_method(self, p,l):
        print("p=", p)
        print("l=", l)
        self.lcdProd.display(p)
        self.lcdLoad.display(l)
        self.update_chart_timeseries(p,l)
        self.update_chart_bar(p,l)
        
    def SerialRead(self,b):
        enable=b.isChecked()
        print("enable="+str(enable))
        #self.myThread.start()
        
    def SerialLog(self,b2):
        print("b2")
        
        
    def threadDone(self):
        print("Done")
        
        
#    def update_chart(self, produzione, carico):
#        load.pop(0)
#        load.append(carico)
#        production.pop(0)
#        production.append(produzione)
#        
#        self.figure.clear() #questo è importante
#        plt.plot(production, color="b")
#        plt.plot(load, color="r")
#        #plt.set_ylim([0,max(load, production)])
#        plt.ylim(ymin=0) 
#        plt.legend(['PV', 'Load'], loc='upper left')
#        self.canvas.draw()


    def update_chart_bar(self, produzione, carico): #bar
        objects = ('Carico','Produzione', 'Immissione')
        y_pos = np.arange(len(objects))
        immissione=produzione-carico
        performance = [carico, produzione, immissione]
        
        self.figure_bar.clear()
        plt.figure(num=self.figure_bar.number) #aggiunta da massimo maggi :)
        plt.bar(y_pos, performance, align='center', alpha=0.5)
        plt.xticks(y_pos, objects)
        plt.ylabel('Power [W]')
        plt.title('Power usage') #is printing this in the wrong canvas
        self.canvas_bar.draw()


    def update_chart_timeseries(self, produzione, carico): #time series
        load.pop(0)
        load.append(carico)
        production.pop(0)
        production.append(produzione)
        
        self.figure_timeseries.clear() #questo è importante
        plt.figure(num=self.figure_timeseries.number)#aggiunta da massimo maggi :)
        plt.plot(production, color="b")
        plt.plot(load, color="r")
        #plt.set_ylim([0,max(load, production)])
        plt.ylim(ymin=0) 
        plt.legend(['PV', 'Load'], loc='upper left')
        self.canvas_timeseries.draw()

            
        
    def available_serial_ports(self):
        """ Lists serial port names
    
            :raises EnvironmentError:
                On unsupported or unknown platforms
            :returns:
                A list of the serial ports available on the system
        """
        if sys.platform.startswith('win'):
            ports = ['COM%s' % (i + 1) for i in range(256)]
        elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
            # this excludes your current terminal "/dev/tty"
            ports = glob.glob('/dev/tty[A-Za-z]*')
        elif sys.platform.startswith('darwin'):
            ports = glob.glob('/dev/tty.*')
        else:
            raise EnvironmentError('Unsupported platform')
    
        result = []
        for port in ports:
            try:
                s = serial.Serial(port)
                s.close()
                result.append(port)
            except (OSError, serial.SerialException):
                pass
        return result

        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = Window()
    #main.minimumWidth(800) 
    main.show()
    sys.exit(app.exec_())
#

    


