# ------------------------------------------------------
# -------------------- mplwidget.py --------------------
# ------------------------------------------------------
from PyQt5.QtWidgets import*

from matplotlib.backends.backend_qt5agg import FigureCanvas

from matplotlib.figure import Figure
    
class MplWidget(QWidget):
    
    def __init__(self, parent = None):

        QWidget.__init__(self, parent)
        
        
        self.canvas = FigureCanvas(Figure(dpi=70))
        vertical_layout = QVBoxLayout()
        vertical_layout.addWidget(self.canvas)
        self.canvas.axes = self.canvas.figure.add_subplot(111)
        #self.canvas.figure.set_tight_layout(True)
        self.canvas.figure.tight_layout(pad=1.1, w_pad=0.1, h_pad=0.1)
        
        self.setLayout(vertical_layout)
        