import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

class GraphHelper:
    """グラフ描画ヘルパークラス"""
    
    def __init__(self, parent, figsize=(8, 6)):
        """
        parent: グラフを配置する親ウィジェット
        figsize: グラフのサイズ
        """
        self.figure = plt.Figure(figsize=figsize, dpi=100)
        self.axis = self.figure.add_subplot(111)
        
        self.canvas = FigureCanvasTkAgg(self.figure, parent)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
    
    def plot(self, x_data, y_data, xlabel="X", ylabel="Y", title="Graph", clear=True):
        """基本的な折れ線グラフを描画"""
        if clear:
            self.axis.clear()
        
        self.axis.plot(x_data, y_data, marker='o', linestyle='-')
        self.axis.set_xlabel(xlabel)
        self.axis.set_ylabel(ylabel)
        self.axis.set_title(title)
        self.axis.grid(True)
        
        self.canvas.draw()
    
    def scatter(self, x_data, y_data, xlabel="X", ylabel="Y", title="Scatter Plot", clear=True):
        """散布図を描画"""
        if clear:
            self.axis.clear()
        
        self.axis.scatter(x_data, y_data)
        self.axis.set_xlabel(xlabel)
        self.axis.set_ylabel(ylabel)
        self.axis.set_title(title)
        self.axis.grid(True)
        
        self.canvas.draw()
    
    def multi_plot(self, data_sets, labels, xlabel="X", ylabel="Y", title="Graph"):
        """複数のデータセットを同時にプロット"""
        self.axis.clear()
        
        for (x, y), label in zip(data_sets, labels):
            self.axis.plot(x, y, marker='o', label=label)
        
        self.axis.set_xlabel(xlabel)
        self.axis.set_ylabel(ylabel)
        self.axis.set_title(title)
        self.axis.legend()
        self.axis.grid(True)
        
        self.canvas.draw()
    
    def clear_plot(self):
        """グラフをクリア"""
        self.axis.clear()
        self.canvas.draw()
    
    def get_canvas(self):
        """キャンバスを取得"""
        return self.canvas