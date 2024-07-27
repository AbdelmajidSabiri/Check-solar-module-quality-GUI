from pathlib import Path
from tkinter import Tk, Canvas, Entry, Button, PhotoImage,ttk, DoubleVar,Label,StringVar, BooleanVar
from tkinter.font import Font
import customtkinter as ctk
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import sys,time
sys.path.append('C:\\Users\\dell\\Check-solar-module-quality-GUI')
from readData import Add_current, Bkp8600, Add_voltage, Add_serialNum, CollectData
from PIL import Image, ImageTk
from datetime import datetime

