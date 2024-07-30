from pathlib import Path
from tkinter import Tk, Canvas, Entry, Button, PhotoImage,ttk, DoubleVar,Label,StringVar, BooleanVar
from tkinter.font import Font
import customtkinter as ctk
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import sys,time
import math
import random
sys.path.append('C:\\Users\\dell\\GUI-read-Bk-precision-data')
from readData import Bkp8600, CollectData
from PIL import Image, ImageTk
from datetime import datetime



class GUI:

    def __init__(self):

        self.bk_device = Bkp8600()
        self.bk_device.initialize()
        self.data_list_current = []
        self.data_list_voltage = []
        self.data_list_power = []
        self.max_power = 0.0
        self.MPP = {"Vmpp" : 0 , "Impp" : 0}
        self.progress = 0
        self.running = False

        self.window = ctk.CTk()
        self.window.geometry("1420x800")
        self.window.title("Agamine Solar LAMINATE TESTING V 1.0")
        self.window.configure(fg_color = "#D7E1E7")

        # Create variables for displaying max power and other parameters
        self.Current_var = DoubleVar(value=0.00)
        self.Voltage_var =  DoubleVar(value=0.00)
        self.Power_var = DoubleVar(value=0.00)
        self.max_power_var = DoubleVar(value=0.00)
        self.Vmpp_var = DoubleVar(value= 0.00)
        self.Impp_var = DoubleVar(value= 0.00)
        self.Isc_var = DoubleVar(value= 0.00)
        self.Voc_var = DoubleVar(value= 0.00)
        self.FF_var = DoubleVar(value= 0.00)
        self.temp_var = DoubleVar(value=25)
        self.grade_var = StringVar(value="A")

        self.Current_formated = StringVar(value="0.00")
        self.Voltage_formated =  StringVar(value="0.00")
        self.Power_formated = StringVar(value="0.00")
        self.max_power_formated = StringVar(value="0.00")
        self.Vmpp_formated = StringVar(value="0.00")
        self.Impp_formated = StringVar(value="0.00")
        self.Isc_formated = StringVar(value="0.00")
        self.Voc_formated = StringVar(value="0.00")
        self.FF_formated = StringVar(value="0.00")




        self.left_frame = ctk.CTkFrame(master=self.window, width=200, corner_radius=0, fg_color='white')
        self.left_frame.pack(side="left", fill="y")

        # Load images
        self.image_dashboard_ON_img = PhotoImage(file="images\\dashboard_ON.png")
        self.image_dashboard_OFF_img = PhotoImage(file="images\\dashboard_OFF.png")

        self.image_bk_profiles_OFF_img = PhotoImage(file="images\\Bk_profiles_OFF.png")
        self.image_bk_profiles_ON_img = PhotoImage(file="images\\Bk_profiles_ON.png")

        self.fram_indicator_img = PhotoImage(file="images\\Fram_indicator.png")
        self.vertical_line_img = PhotoImage(file="images\\Vertical_line.png")        

        # Create buttons with images
        self.dashboard_button = ctk.CTkButton(master=self.left_frame, 
                                                image=self.image_dashboard_ON_img, 
                                                text="DASHBOARD", 
                                                compound="top", 
                                                command=self.show_dashboard, 
                                                fg_color='white', 
                                                text_color='#0000ff', 
                                                font=("Arial Rounded MT Bold",14), 
                                                hover_color="white")
        self.dashboard_button.pack(pady=10, padx=0, anchor='w')

        self.bk_profiles_button = ctk.CTkButton(master=self.left_frame,
                                                 image=self.image_bk_profiles_OFF_img, 
                                                 text="BK PROFILES", 
                                                 compound="top", 
                                                 command=self.show_bk_profiles, 
                                                 fg_color='white', 
                                                 text_color='#b7b7fc', 
                                                 font=("Arial Rounded MT Bold",14), 
                                                 hover_color="white")
        self.bk_profiles_button.pack(pady=10, padx=0, anchor='w')

        # Create a frame for the main content
        self.main_frame = ctk.CTkFrame(master=self.window, fg_color='#D7E1E7')
        self.main_frame.pack(side="right", expand=True, fill="both", padx=0, pady=0)

        self.dashboard_frame = ctk.CTkFrame(master=self.main_frame, fg_color='#D7E1E7')
        self.bk_profiles_frame = ctk.CTkFrame(master=self.main_frame, fg_color='#D7E1E7')

        # Initial display
        self.dashboard_frame.pack(fill="both", expand=True)

        # Setup content of dashboard
        self.setup_dashboard_content()
        self.setup_bk_profiles_content()

        # Start update loop for maximum power and time

        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)  # Handle window closing event
        self.window.mainloop()


    # Funtion to Get Max power
    def update_max_power(self) :

        if self.Power_var.get() > self.max_power_var.get():
            self.max_power_var.set(self.Power_var.get())
            self.Impp_var.set(self.Current_var.get())
            self.Vmpp_var.set(self.Voltage_var.get())

            self.max_power_formated.set(f"{self.max_power_var.get():.2f}")
            self.Impp_formated.set(f"{self.Impp_var.get():.2f}")
            self.Vmpp_formated.set(f"{self.Vmpp_var.get():.2f}")



    def update_data(self) :
        if self.running :
            self.update_max_power()
            self.calculate_isc_voc()
            self.calculate_FF()
            self.calculate_grade()


    # Calculate Isc and Voc
    def calculate_isc_voc(self):


        # Get Value of Isc if it found if not make it 0
        Isc_found = False
        for v, i in zip(self.data_list_voltage, self.data_list_current):
            if v == 0:
                self.Isc_var.set(i)
                Isc_found = True
                break

        if not Isc_found :
            self.Isc_var.set(0.00)

        self.Isc_formated.set(f"{self.Isc_var.get():.2f}")



        # Get Value of Voc if it found if not make it 0
        Voc_found = False

        for v, i in zip(self.data_list_voltage, self.data_list_current):
            if i == 0:
                self.Voc_var.set(v)
                Voc_found = True
                break

        if not Voc_found :
            self.Voc_var.set(0.00)

        self.Voc_formated.set(f"{self.Voc_var.get():.2f}")


    def calculate_FF(self) :
        isc = self.Isc_var.get()
        voc = self.Voc_var.get()
        maxPower = self.max_power_var.get()

        if isc and voc and maxPower :
            FF = (maxPower / isc * voc) * 100
        else :
            FF = 0.00

        self.FF_var.set(FF)
        self.FF_formated.set(f"{self.FF_var.get():.2f}")

    def calculate_grade(self) :
        pass

    def configure_plot(self, ax):
        ax.set_facecolor('#e4e4e4')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        ax.set_xlabel("Voltage (V)", fontsize=12, fontweight='bold')
        ax.set_ylabel("Current (A) / Power (W)", fontsize=12, fontweight='bold')

        ax.tick_params(axis='x', colors='#000000')
        ax.tick_params(axis='y', colors='#000000')

        ax.xaxis.label.set_color('#000000')
        ax.yaxis.label.set_color('#000000')

        ax.grid(True, color='#3A3A3A', linestyle='--', linewidth=0.5)

    def setup_chart(self) :

        fig_combined, ax_combined = plt.subplots(figsize=(8, 4.5))
        fig_combined.patch.set_facecolor("white")
        fig_combined.patch.set_linewidth(2)

        self.configure_plot(ax_combined)

        canvas_combined = FigureCanvasTkAgg(fig_combined, master=self.dashboard_frame)
        canvas_combined.draw()
        canvas_combined.get_tk_widget().place(x=65, y=95)

    
        self.ani_combined = animation.FuncAnimation(
            fig_combined,
            self.animate_combined,
            fargs=(ax_combined,),
            interval=100
        )

    # Funtion to update Time
    def update_time(self):
        current_time = datetime.now().strftime('%H:%M:%S')   
        current_date = datetime.now().strftime('%d/%m/%Y')        
        self.time_label.configure(text=current_time)  
        self.date_label.configure(text = current_date)                        
        self.dashboard_frame.after(1000, self.update_time)    

    # Funtion to Get data from Bk-precision
    def get_data(self):
        try:
            measured_current = self.bk_device.get_current()
            voltage = self.bk_device.get_voltage()
            if measured_current and voltage :
                return measured_current, voltage
            else :
                return 0.00,0.00
        except Exception as e:
            print(f"Error retrieving data: {e}")
            return 0.00, 0.00

    # Animation function for updating Combined (current & voltage) plot
    def animate_combined(self, i, ax):
        if self.running :
            # Simulated data
            current = 10 * math.sin(i * 0.1) + random.uniform(-1, 1)
            voltage = 20 * math.cos(i * 0.1) + random.uniform(-1, 1)
            # current, voltage = self.get_data()

            power = current * voltage

            self.Current_var.set(current)
            self.Voltage_var.set(voltage)
            self.Power_var.set(power)
            self.Current_formated.set(f"{self.Current_var.get():.2f}")
            self.Voltage_formated.set(f"{self.Voltage_var.get():.2f}")
            self.Power_formated.set(f"{self.Power_var.get():.2f}")


            self.update_data()

            self.data_list_current.append(current)
            self.data_list_voltage.append(voltage)
            self.data_list_power.append(power)
            self.data_list_current = self.data_list_current[-50:]  # Limit to the last 50 data points
            self.data_list_voltage = self.data_list_voltage[-50:]   # Limit to the last 50 data points
            self.data_list_power = self.data_list_power[-50:]
            
            ax.clear()
            self.configure_plot(ax)

                    
            # Plot the current data
            ax.plot(self.data_list_voltage,self.data_list_current,color='Blue', linewidth=2, label='I-V curve')        
            ax.plot(self.data_list_voltage,self.data_list_power,color='Red', linewidth=2, label='I-P curve')

            ax.legend(loc='upper right')

    # Method to get Serial Number
    def get_serialNum(self):
        serial_number = self.entry_serialNum.get()
        return serial_number

    # Function to Get Data and send it to CollectData function That Add it to excel file
    def SaveData(self) :

        current_date = datetime.now()
        formatted_date = current_date.strftime("%m/%d/%Y")
        formatted_time = current_date.strftime("%H:%M:%S")

        serial_number = self.get_serialNum()
        Voc =self.Voc_var.get()
        Isc = self.Isc_var.get()
        max_power = self.max_power
        Impp = self.Impp_var.get()
        Vmpp = self.Vmpp_var.get()
        
        CollectData(formatted_date,formatted_time,serial_number,max_power,Vmpp,Impp,Voc,Isc)
        
    # Keep Updating progress bar, and if it arrives 100%, it Will save data
    def update_progress(self):
        if self.progress < 1:
            self.progress += 0.01
            self.progress_bar.set(self.progress)
            self.progress_label.configure(text=f"{int(self.progress * 100)}%")
            self.window.after(100, self.update_progress)  # Update self.progress every 100 milliseconds
        else:
            self.running = False
            self.progress_label.configure(text="100%")
            self.status_label.configure(text="  Saved", text_color="#06F30B")
            self.run_button.configure(state = 'normal')
            self.run_button.configure(image = self.run_test_img)

    # Function to Start Test if RUN TEST button is pressed
    def run_test(self):
        self.running = True
        self.progress = 0
        self.progress_bar.set(self.progress)
        self.status_label.configure(text="Running...", text_color="orange")
        self.run_button.configure(state = 'disabled')
        self.run_button.configure(image = self.test_running_img)
        self.update_progress()
    
    # Function to turn ON the light
    def ON_Lamps(self):
        self.ON_button.configure(image = self.lamps_on_img)
        self.OFF_button.configure(image = self.lamps_button_img)

    # Function to turn OFF the light
    def OFF_Lamps(self):
        self.ON_button.configure(image = self.lamps_button_img)
        self.OFF_button.configure(image = self.lamps_off_img)
    

    def show_table(self) :
        pass







    # Function to show content of dashboard frame if DASHBOARD button is pressed
    def show_dashboard(self):
        self.dashboard_frame.pack(fill="both", expand=True)
        self.dashboard_button.configure(image = self.image_dashboard_ON_img, text_color = "#0000ff")
        self.bk_profiles_button.configure(image = self.image_bk_profiles_OFF_img, text_color = "#b7b7fc")
        self.bk_profiles_frame.pack_forget()

    # Function to show content of Bk Profiles frame if BK PROFILES button is pressed
    def show_bk_profiles(self):
        self.bk_profiles_frame.pack(fill="both", expand=True)
        self.dashboard_button.configure(image = self.image_dashboard_OFF_img, text_color = "#b7b7fc")
        self.bk_profiles_button.configure(image = self.image_bk_profiles_ON_img, text_color = "#0000ff")

        self.dashboard_frame.pack_forget()

    # Funtion to Setup content of Dashboard Frame
    def setup_dashboard_content(self):

        # initialis dashborad canvas
        canvas = Canvas(
            self.dashboard_frame,
            bg = "#D7E1E7",
            height = 990,
            width = 1620,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )

        # Add images and Text
        canvas.place(x = 0, y = 0)
        self.run_test_img = PhotoImage(file="images\\run_test.png")
        self.test_running_img = PhotoImage(file = "images\\test_running.png")
        self.Add_to_tab_img = PhotoImage(file = "images\\Add_to_test_tab.png")
        self.insert_SN_img = PhotoImage(file="images\\insert_SN.png")
        self.plot_img = PhotoImage(file = "images\\plot_background.png")
        self.data_img = PhotoImage(file = "images\\data_background.png")
        self.vol_curr_pow_img = PhotoImage(file = "images\\vol_curr_pow_background.png")
        self.lamps_img = PhotoImage(file = "images\\lamps_background.png")
        self.table_img = PhotoImage(file = "images\\table_background.png")
        self.progress_bar_img = PhotoImage(file = "images\\progress_bar_background.png")
        self.status_img = PhotoImage(file = "images\\status.png")
        self.vol_curr_pow_values_img = PhotoImage(file = "images\\vol_curr_pow_values_background.png")
        self.data_values_img = PhotoImage(file = "images\\data_values_background.png")
        self.line_data_img = PhotoImage(file = "images\\line.png")
        self.lamps_on_img = PhotoImage(file = "images\\lamps_on.png")
        self.lamps_off_img = PhotoImage(file = "images\\lamps_off.png")
        self.lamps_button_img = PhotoImage(file = "images\\lamps_button.png")



        insert_SN_background = canvas.create_image(           260,36,image=self.insert_SN_img)
        plot_background = canvas.create_image(     480,340,image = self.plot_img)
        data_background = canvas.create_image(     1250,340,image = self.data_img)
        plot_background = canvas.create_image(     480,640,image = self.vol_curr_pow_img)
        lamp_ON_background = canvas.create_image(  1125,640,image = self.lamps_img)
        lamp_OFF_background = canvas.create_image( 1380,640,image = self.lamps_img)
        table_background = canvas.create_image(    770,875,image = self.table_img)
        fram_indicator = canvas.create_image(      4,62.0,image=self.fram_indicator_img)
        vertical_line = canvas.create_image(       0,400,image=self.vertical_line_img)
        progress_bar_background = canvas.create_image(        1000, 56, image=self.progress_bar_img )
        test_status_background = canvas.create_image(1380,55, image = self.status_img)
        line_data = canvas.create_image(1210, 120, image = self.line_data_img)
        line_data = canvas.create_image(1340, 120, image = self.line_data_img)
        isc_value_background = canvas.create_image(1275, 160, image = self.data_values_img)
        Voc_value_background = canvas.create_image(1275, 210, image = self.data_values_img)
        Mpp_value_background = canvas.create_image(1275, 260, image = self.data_values_img)
        Ipm_value_background = canvas.create_image(1275, 310, image = self.data_values_img)
        Vpm_value_background = canvas.create_image(1275, 360, image = self.data_values_img)
        FF_value_background = canvas.create_image(1275, 410, image = self.data_values_img)
        Grade_value_background = canvas.create_image(1275, 460, image = self.data_values_img)
        temperature_value_background = canvas.create_image(1275, 510, image = self.data_values_img)
        test_recurrence_background = canvas.create_image(1275, 560, image = self.data_values_img)

        Vol_background = canvas.create_image(230,640, image = self.vol_curr_pow_values_img)
        Curr_background = canvas.create_image(500,640, image = self.vol_curr_pow_values_img)
        Pow_background = canvas.create_image(800,640, image = self.vol_curr_pow_values_img)


        


        canvas.create_text(40,30, anchor="nw", text="Insert SN", fill="black", font=("",15))
        canvas.create_text(1180,25, anchor="nw", text="Test status", fill="black", font=("",15))
        canvas.create_text(1080,110, anchor="nw", text="Description", fill="#9ea5d2", font=("Helvetica",12, "bold"))
        canvas.create_text(1250,110, anchor="nw", text="Value", fill="#9ea5d2", font=("Helvetica",12, "bold"))
        canvas.create_text(1375,110, anchor="nw", text="Unit", fill="#9ea5d2", font=("Helvetica",12, "bold"))


        canvas.create_text(1180,150, anchor="nw", text="Isc", fill="#0000FF", font=("",10))
        canvas.create_text(1180,200, anchor="nw", text="Voc", fill="#0000FF", font=("",10))
        canvas.create_text(1180,250, anchor="nw", text="Mpp", fill="#0000FF", font=("",10))
        canvas.create_text(1180,300, anchor="nw", text="Ipm", fill="#0000FF", font=("",10))
        canvas.create_text(1180,350, anchor="nw", text="Vpm", fill="#0000FF", font=("",10))
        canvas.create_text(1185,400, anchor="nw", text="FF", fill="#0000FF", font=("",10))
        canvas.create_text(1185,450, anchor="nw", text="G", fill="#0000FF", font=("",10))
        canvas.create_text(1185,500, anchor="nw", text="T", fill="#0000FF", font=("",10))

        canvas.create_text(1055,150, anchor="nw", text="Short Circuit Current", fill="Black", font=("",10))
        canvas.create_text(1055,200, anchor="nw", text="Open Circuit Voltage", fill="Black", font=("",10))
        canvas.create_text(1042,251, anchor="nw", text="Maximum Power Point", fill="Black", font=("",10))
        canvas.create_text(1082,301, anchor="nw", text="Current at MPP", fill="Black", font=("",10))
        canvas.create_text(1082,350, anchor="nw", text="Voltage at MPP", fill="Black", font=("",10))
        canvas.create_text(1120,400, anchor="nw", text="Fil Factor", fill="Black", font=("",10))
        canvas.create_text(1145,450, anchor="nw", text="Grade", fill="Black", font=("",10))
        canvas.create_text(1108,500, anchor="nw", text="Temperature", fill="Black", font=("",10))
        canvas.create_text(1090,550, anchor="nw", text="Test Recurrence", fill="Black", font=("",10))

        canvas.create_text(1390,150, anchor="nw", text="A", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,200, anchor="nw", text="V", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,250, anchor="nw", text="W", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,300, anchor="nw", text="A", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,350, anchor="nw", text="V", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,400, anchor="nw", text="%", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,450, anchor="nw", text="⭐", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,500, anchor="nw", text="°C", fill="#0000FF", font=("Helvetica",12, "bold"))
        canvas.create_text(1390,550, anchor="nw", text="#", fill="#0000FF", font=("Helvetica",12, "bold"))


        canvas.create_text(90,630, anchor="nw", text="Voltage", fill="Black", font=("",13))
        canvas.create_text(360,630, anchor="nw", text="Current", fill="Black", font=("",13))
        canvas.create_text(660,630, anchor="nw", text="Power", fill="Black", font=("",13))

        canvas.create_text(270,627, anchor="nw", text="V", fill="#0000FF", font=("Helvetica",14, "bold"))
        canvas.create_text(540,627, anchor="nw", text="A", fill="#0000FF", font=("Helvetica",14, "bold"))
        canvas.create_text(840,627, anchor="nw", text="W", fill="#0000FF", font=("Helvetica",14, "bold"))

        canvas.create_text(1193,630, anchor="nw", text="On", fill="#0000FF", font=("Helvetica",14, "bold"))
        canvas.create_text(1449,630, anchor="nw", text="Off", fill="#0000FF", font=("Helvetica",14, "bold"))


        self.setup_chart()


        # Entry Text for serial number of solar module
        self.entry_serialNum = ctk.CTkEntry(master=self.dashboard_frame, placeholder_text="Serial Number",border_width = 0, fg_color = "white", bg_color="white", width=150)
        self.entry_serialNum.place(x=120, y=15)

        # RUN TEST button to start test
        self.run_button = ctk.CTkButton(master=self.dashboard_frame, text = "RUN TEST" ,image=self.run_test_img, command=self.run_test, fg_color='transparent', text_color='#0000FF', font=("Arial Rounded MT Bold",14),hover="transparent")
        self.run_button.place(x=540, y=9)

        self.Add_to_tab_button = ctk.CTkButton(master=self.dashboard_frame, text = "Add to Test Tab" ,image=self.Add_to_tab_img, command=self.SaveData, fg_color='#D7E1E7', text_color='#000000', font=("Arial Rounded MT Bold",14),hover="#b7b7fc")
        self.Add_to_tab_button.place(x=340, y=10)

        # Progress bar to show the progress of the test
        self.progress_bar = ctk.CTkProgressBar(master=self.dashboard_frame, width = 150,height=8.5, bg_color= "white",progress_color = ("#00d9ff", "black"))
        self.progress_bar.set(self.progress)
        self.progress_bar.place(x=708, y=26)

        # Progress label to display the percentage
        self.progress_label = ctk.CTkLabel(master=self.dashboard_frame, text="0%", font=("Arial", 13, "bold"), text_color='#2021fd', bg_color="white")
        self.progress_label.place(x=868,y=14)

        # Status Label to display the state of the test
        self.status_label = ctk.CTkLabel(master=self.dashboard_frame, text="Waiting", text_color="orange", font=("Arial", 14, "bold"), bg_color="white")
        self.status_label.place(x=1073, y=15)

        # Temperature Label
        self.temp_label = ctk.CTkLabel(master = self.dashboard_frame, text = f"{round(self.temp_var.get(),2)}", height=5 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.temp_label.place(x=1003, y=400)

        self.Isc_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Isc_formated, height=5 , width=20 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.Isc_label.place(x=1003, y=120)

        self.Voc_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Voc_formated, height=5 , width=20 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.Voc_label.place(x=1003, y=160)

        self.Mpp_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.max_power_formated, height=5 , width=20 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.Mpp_label.place(x=1003, y=200)

        self.Impp_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Impp_formated, height=5 , width=20 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.Impp_label.place(x=1003, y=240)

        self.Vmpp_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Vmpp_formated, height=5 , width=20 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.Vmpp_label.place(x=1003, y=280)
        
        self.FF_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.FF_formated, height=5 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.FF_label.place(x=1003, y=320)

        self.grade_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.grade_var, height=5 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 13.5,"bold"))
        self.grade_label.place(x=1003, y=360)

        self.Voltage_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Voltage_formated, height=5, width=60 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 15,"bold"))
        self.Voltage_label.place(x=143, y=502)

        self.Current_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Current_formated, height=5 ,width=60 ,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 15,"bold"))
        self.Current_label.place(x=368, y=502)

        self.Power_label = ctk.CTkLabel(master = self.dashboard_frame, textvariable = self.Power_formated, height=5, width=60,text_color="Black",bg_color = "#EBECF0" ,font=("Arial", 15,"bold"))
        self.Power_label.place(x=600, y=502)


    


        # Lamps Button
        self.ON_button = ctk.CTkButton(master = self.dashboard_frame, 
                                        text="Turn Lamps", 
                                        image = self.lamps_button_img, 
                                        text_color="Black", 
                                        fg_color='white', 
                                        command= self.ON_Lamps, 
                                        font=("Arial",14), 
                                        compound = "right",
                                        hover_color="white",
                                        bg_color = "white")
        self.ON_button.place(x=814,y=490)

        self.OFF_button = ctk.CTkButton(master = self.dashboard_frame, 
                                        text="Turn Lamps", 
                                        image = self.lamps_off_img, 
                                        text_color="Black", 
                                        fg_color='white', 
                                        compound = "right", 
                                        command= self.OFF_Lamps,  
                                        font=("Arial",14),
                                        hover_color="white",
                                        bg_color="white")
        self.OFF_button.place(x=1018,y=490)

    # Function to Setup content of bk_profiles Frame
    def setup_bk_profiles_content(self):

        canvas = Canvas(
            self.bk_profiles_frame,
            bg = "#D7E1E7",
            height = 930,
            width = 1721,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )
        canvas.place(x = 0, y = 0)

        fram_indicator = canvas.create_image(
            4,
            182.0,
            image=self.fram_indicator_img
        )
        vertical_line = canvas.create_image(
            0.0,
            400.0,
            image=self.vertical_line_img
        )

    # Funtion to close the application correctly
    def on_closing(self):
        self.bk_device.reset_to_manual()
        self.window.destroy()
        self.window.quit()

# Instantiate the GUI class to start the application
if __name__ == "__main__":
    GUI()