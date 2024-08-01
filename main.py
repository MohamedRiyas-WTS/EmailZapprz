import customtkinter
from customtkinter import CTkFrame, CTkLabel, CTkEntry, CTkButton, CTkSegmentedButton, CTkScrollableFrame, CTkImage, CTkToplevel,CTkFont
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from redmail import gmail
import customtkinter
import re
import webview
from tkinter import ttk
from tkinter import scrolledtext
from pathlib import Path
import os
from CTkMessagebox import CTkMessagebox
import openpyxl



class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.excel_file_df_from_mail = None
        self.excel_file_df_to_mail = None
        self.html_full_content = None
        self.body_params = None
        self.excel_file_path = None
        self.excel_file_to_mail_header_list = []
        self.scrollable_frame_switches = []
        self.title("EmailZapprz")
        self.attachment_file_path_list = []
        self.static_attachment_file_count = 0
        self.html_state = ["Dynamic","Static"]
        self.current_html_state = None
        self.individual_attachments_header = []
        self.email_subject = ""
        # self.iconbitmap("WTS.ico")
        self.resizable(False, False)  
        # Disable window resizing
        window_width = 1000 
        # Window Width - Main frame
        window_height = 580  
        #Window Height - Main frame
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_coordinate = (screen_width - window_width) / 2  
        # Main frame setting in X axis
        y_coordinate = (screen_height - window_height) / 2  
        # Main frame setting in Y axis
        self.geometry("%dx%d+%d+%d" % (window_width, window_height, x_coordinate, y_coordinate))
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(8, weight=1) 

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="EmailZapprz", 
                                                             compound="right", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=2, column=0, padx=20, pady=50)
        
        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Excel Uploader",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="w", command=self.home_button_event)
        self.home_button.grid(row=3, column=0, sticky="ew")

        self.frame_2_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Html Uploader",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                       anchor="w", command=self.frame_2_button_event)
        self.frame_2_button.grid(row=4, column=0, sticky="ew")

        self.start_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=50, height=40, border_spacing=10, text="Start Email",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                       anchor="w", command=self.start_email)
        self.start_button.grid(row=5, column=0, padx=20, pady=(150,0), sticky="s")

        self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"],
                                                                command=self.change_appearance_mode_event)
        self.appearance_mode_menu.grid(row=8, column=0, padx=20, pady=(10,50), sticky="s")

        self.home_frame = customtkinter.CTkFrame(self, corner_radius=50, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=20)
        self.home_frame_label = customtkinter.CTkLabel(self.home_frame, text="Email Credentials Files",font=CTkFont(family="times",size=34,weight="bold"),height=40)
        self.home_frame_label.grid(row=0, column=0, padx=20, pady=80)
        self.home_frame_qoute1_label = customtkinter.CTkLabel(self.home_frame, text="Select the following options",font=CTkFont(family="times",size=20,weight="bold",slant="italic"))
        self.home_frame_qoute1_label.grid(row=1, column=0, padx=20, pady=20)

        self.path_frame_button = CTkButton(self.home_frame, text="Excel Path",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=40,border_color="dark",text_color="#1c1c1c",corner_radius=10,command=self.upload_file)
        self.path_frame_button.grid(row=2, column=0, padx=20, pady=10)
        self.url_frame_button = CTkButton(self.home_frame, text="File URL",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=40,border_color="dark",text_color="#1c1c1c",corner_radius=10)
        self.url_frame_button.grid(row=3, column=0, padx=20, pady=20)
        self.url_frame_button = customtkinter.CTkLabel(self.home_frame, text="Expecting\nSheet1 --> From Mail Credential \nSheet2 --> To Mail Credentials",font=CTkFont(family="times",size=15,weight="bold"),fg_color='#3b8ed0',height=40,text_color="#1c1c1c",corner_radius=10)
        self.url_frame_button.grid(row=4, column=0, padx=20, pady=40)

        # create second frame
        self.second_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="grey")
        self.second_frame.grid_columnconfigure(0, weight=1)
        # Set row and column weights
        self.second_frame.grid_rowconfigure(1, weight=1)

        self.seg_button_1 = customtkinter.CTkSegmentedButton(self.second_frame,font=CTkFont(family="times",size=18,weight="bold"),width=500,command=self.change_segment_event)
        self.seg_button_1.grid(row=0, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.seg_button_1.configure(values=["Dynamic", "Static"])
        self.seg_button_1.set("Dynamic")

        self.dynamic_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.dynamic_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.dynamic_frame.grid_rowconfigure(1, weight=1)  # Allows the textbox row to expand

        self.static_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.static_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.static_frame.grid_forget()
        
        self.textbox_dynamic = customtkinter.CTkTextbox(self.dynamic_frame, width=700,fg_color="grey",text_color="white", height=340)
        self.textbox_dynamic.grid(row=1, column=0, columnspan=3, padx=(40, 40), pady=(10, 10), sticky="nsew")
        self.textbox_dynamic.insert("2.0","Html Code goes here.../ Upload the html file")
        self.textbox_dynamic.bind("<FocusIn>",self.clear_placeholder)
        self.textbox_dynamic.bind("<KeyRelease>",lambda event: self.clear_entry_text())
        
        self.textbox_static = customtkinter.CTkTextbox(self.static_frame, width=700,fg_color="grey",text_color="white", height=340)
        self.textbox_static.grid(row=1, column=0, columnspan=3, padx=(40, 40), pady=(10, 10), sticky="nsew")
        self.textbox_static.insert("2.0","Html Code goes here.../ Upload the html file")
        self.textbox_static.bind("<FocusIn>",self.clear_placeholder)
        self.textbox_static.bind("<KeyRelease>",lambda event: self.clear_entry_text_static())
        # self.tabview = customtkinter.CTkTabview(self.second_frame, width=700,height=450,corner_radius=50)
        # self.tabview.grid(row=0, column=1, padx=(65, 0), pady=(50,100), sticky="nsew")
        # self.tabview.add("Static")
        # self.tabview.add("Dynamic")
        # self.tabview.tab("Static").grid_columnconfigure(2, weight=1)  # configure grid of individual tabs
        # self.tabview.tab("Dynamic").grid_columnconfigure(2, weight=1)
        self.dynamic_upload_button=CTkButton(self.dynamic_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10,command=self.upload_html_file)
        self.dynamic_upload_button.grid(row=0, column=2, columnspan=3, padx=(10, 40), pady=(10, 10), sticky="ew")
        self.dynamic_sub_button=CTkButton(self.dynamic_frame,text="Submit",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.upload_custom_file)
        self.dynamic_sub_button.grid(row=2, column=1, columnspan=2, padx=(40, 300), pady=(10, 10), sticky="ew")
        self.entry_dynamic = customtkinter.CTkEntry(self.dynamic_frame, placeholder_text="Upload....")
        self.entry_dynamic.grid(row=0, column=0, columnspan=2, padx=(40, 1), pady=(10, 10), sticky="ew")
        self.entry_dynamic.bind("<KeyRelease>",lambda event: self.clear_textbox_dynomic())


        self.static_upload_button=CTkButton(self.static_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.upload_static_html_file)
        self.static_upload_button.grid(row=0, column=2, columnspan=3, padx=(10, 40), pady=(10, 10), sticky="ew")
        self.static_sub_button=CTkButton(self.static_frame,text="Submit",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.upload_static_custom_file)
        self.static_sub_button.grid(row=2, column=1, columnspan=2, padx=(40, 300), pady=(10, 10), sticky="ew")
        self.entry_static = customtkinter.CTkEntry(self.static_frame, placeholder_text="Upload....")
        self.entry_static.grid(row=0, column=0, columnspan=2, padx=(40, 1), pady=(10, 10), sticky="ew")
        self.entry_static.bind("<KeyRelease>",lambda event: self.clear_textbox_static())

        self.subject_attach_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.subject_attach_frame.grid(row=0, column=0, padx=(20, 20), pady=(10, 10), sticky="ns")
        self.subject_attach_frame.grid_rowconfigure(1, weight=1)  
        self.subject_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Subject",font=CTkFont(family="times", size=20, weight="bold"), height=40)
        self.subject_label.grid(row=0, column=0, padx=(40, 1), pady=(20, 10), sticky="ew")

        self.subject_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="Enter the subject")
        self.subject_entry.grid(row=0, column=1, padx=(10, 10), pady=(20, 10), sticky="ew")

        self.attach_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Static Attachments",font=CTkFont(family="times", size=20, weight="bold"), height=40)
        self.attach_label.grid(row=1, column=0, padx=(40, 10), pady=(10, 230), sticky="ew")

        self.attachment_files = CTkButton(self.subject_attach_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.static_attach_files_function)
        self.attachment_files.grid(row=1, column=1, padx=(10, 460), pady=(10, 230))

        self.dynamic_scroll_checkbox_frame = customtkinter.CTkScrollableFrame(self.subject_attach_frame, label_text="Individual Attachment",width=10,height=50)
        self.dynamic_scroll_checkbox_frame.grid(row=1, column=1,padx=(1,10), pady=(100,10), sticky="ew")
        self.dynamic_scrollable_frame_checkbox = []

        self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2, text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
        self.attachment_sub_button.grid(row=2, column=1,pady=(10, 50))
        self.attachment_back_button = CTkButton(self.subject_attach_frame,corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
        self.attachment_back_button.grid(row=2, column=0,pady=(10, 50))

        self.subject_attach_frame.grid_columnconfigure(0, weight=0)
        self.subject_attach_frame.grid_columnconfigure(1, weight=1)

        self.subject_attach_frame.grid_forget()


        self.list_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.list_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.list_frame.grid_forget()

    def attachment_back_function(self):
        self.subject_attach_frame.grid_forget()
        self.dynamic_frame.configure(corner_radius=16)
        self.dynamic_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.seg_button_1.configure(state=customtkinter.NORMAL)

    def list_frame_show_call(self):
        params_variable = re.findall(r"\{\{(\w+)\}\}",self.html_full_content)
        if params_variable and self.current_html_state == self.html_state[0]:
            if self.excel_file_to_mail_header_list:
                self.list_frame_show(params_variable, self.excel_file_to_mail_header_list)
            else:
                print("Empty header list")
        else:
            self.static_preview_frame_function()

    def attachment_sub_function(self):
        self.email_subject = self.subject_entry.get()
        if len(self.subject_entry.get()) >= 1 and self.subject_entry.get().strip() != "":                
            self.list_frame_show_call()
        else:
            msg_closing = CTkMessagebox(title="Email Subject is empty!", message="Do you want to continue?",
                        icon="question", option_1="No", option_2="Yes")
            response = msg_closing.get()
            if response=="Yes":
                self.list_frame_show_call()      

    def clear_placeholder(self, event): # function for clearing the place holder
        if self.textbox_dynamic.get("1.0", "end-1c") == "Html Code goes here.../ Upload the html file":
            self.textbox_dynamic.delete("1.0", "end-1c")

        if self.textbox_static.get("1.0", "end-1c") == "Html Code goes here.../ Upload the html file":
            self.textbox_static.delete("1.0", "end-1c")

    def change_segment_event(self,name):
        if name == "Dynamic":
            self.dynamic_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
            self.dynamic_frame.grid(row=1, column=0,  padx=(20, 20), pady=(10, 20), sticky="ew")
        else:
            self.dynamic_frame.grid_forget()
        if name == "Static":
            self.static_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
            self.static_frame.grid(row=1, column=0,  padx=(20, 20), pady=(10, 20), sticky="ew")
        else:
            self.static_frame.grid_forget()

    def get_entry_data(self):
        data_dict = {}
        for label, entry_widget, combo_box in self.scrollable_frame_switches:
            key = label.cget("text")
            entry_value = entry_widget.get()
            combo_value = combo_box.get()
            if entry_value:
                data_dict[key] = entry_value
            else:
                data_dict[key] = f"row['{combo_value}']"
        
        return data_dict
    
    def get_checkbox_values(self):
        data_dict = []
        for checkbox in self.dynamic_scrollable_frame_checkbox:
            if checkbox.get() == 1:
                data_dict.append(checkbox.cget("text"))
        self.individual_attachments_header = data_dict 
    
    def sub_attach_function(self):
        print("Attach entry")
        self.static_frame.grid_forget()
        self.dynamic_frame.grid_forget()
        # self.list_frame.grid_forget()
        self.subject_attach_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
        self.subject_attach_frame.grid(row=1, column=0,  padx=(20, 20), pady=(20, 20), sticky="ns")

        for index, key in enumerate(self.excel_file_to_mail_header_list):
            # switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
            # switch.grid(row=i, column=0, padx=10, pady=(0, 20))
            # self.scrollable_frame_switches.append(switch)
            checkbox = customtkinter.CTkCheckBox(master=self.dynamic_scroll_checkbox_frame,text=key)
            checkbox.grid(row=index, column=0, pady=20, padx=20, sticky="w")
            self.dynamic_scrollable_frame_checkbox.append(checkbox)
        self.dynamic_scroll_checkbox_submit = customtkinter.CTkButton(self.subject_attach_frame, text="Submit",width=10, command=self.get_checkbox_values)
        self.dynamic_scroll_checkbox_submit.grid(row=3, column=0, padx=(200, 0), pady=(20, 20), sticky="nsew")

        # params_variable= re.findall(r"\{\{(\w+)\}\}",str(self.html_full_content))
        # if params_variable:
        #     if self.excel_file_to_mail_header_list:
        #         self.list_frame_show(params_variable, self.excel_file_to_mail_header_list)
        #     else:
        #         print("Empty header list")
        # else:
        #     print("Empty params variable list")

    def list_frame_show(self,params_variable, params_name):
        self.static_frame.grid_forget()
        self.dynamic_frame.grid_forget()
        self.subject_attach_frame.grid_forget()
        self.list_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
        self.list_frame.grid(row=1, column=0,  padx=(20, 20), pady=(20, 20), sticky="ew")
        self.scrollable_frame = customtkinter.CTkScrollableFrame(self.list_frame, label_text="Email Mapping",height=350,width=700)
        self.scrollable_frame.grid(row=1, column=2, padx=(25, 10), pady=(20,150), sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame_switches = []
        replacer_list = params_variable
        header_list = params_name
        # for i in range(5):
        #     switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
        #     switch.grid(row=i, column=0, padx=10, pady=(0, 20))
            # self.scrollable_frame_switches.append(switch)
        for index,context in enumerate(replacer_list):
            
            context_label = customtkinter.CTkLabel(master=self.scrollable_frame,text=context)
            context_label.grid(row=index,column=0, padx=10, pady=(0, 20))
            context_list = customtkinter.CTkComboBox(self.scrollable_frame,values=header_list)
            context_list.grid(row=index,column=2, padx=10, pady=(0, 20))
            entry_dynamic = customtkinter.CTkEntry(self.scrollable_frame)
            entry_dynamic.grid(row=index, column=1, padx=10, pady=(0, 20))
            self.scrollable_frame_switches.append((context_label, entry_dynamic, context_list))

            # Bind the entry widget to the callback function
            entry_dynamic.bind("<KeyRelease>", lambda event, e=entry_dynamic, c=context_list: self.on_entry_change(e, c))
        self.submit_list_button = customtkinter.CTkButton(self.list_frame, corner_radius=30, text="Submit",
                                                         fg_color="white", border_color="green", border_width=2, text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.dynamic_submit_button)
        self.submit_list_button.grid(row=1, column=1, columnspan=2, rowspan=2,padx=(350,10), pady=(350, 10))
        self.dynamic_back_button = CTkButton(self.list_frame,corner_radius=30, text="Back",
                                                         fg_color="white", border_color="gray", border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.dynamic_back_button_function)
        self.dynamic_back_button.grid(row=1, column=2,columnspan=2, rowspan=2,padx=(10,350), pady=(350,10))

        # create preview frame==>dummy frame created for future use
        self.preview_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="grey")
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)

    def clear_entry_text_static(self):
        self.entry_static.delete(0,"end")  
    def clear_textbox_dynomic(self):
        self.textbox_dynamic.delete(1.0,"end")

    def clear_textbox_static(self):
        self.textbox_static.delete(1.0,"end")
    def dynamic_back_button_function(self):
        self.list_frame.grid_forget()
        self.subject_attach_frame.configure(corner_radius=16)
        self.subject_attach_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nwes")
        # self.change_segment_event("Dynamic")

    def static_preview_frame_function(self):
        self.webview_window  = None
        html_code = self.html_full_content
        if not html_code.startswith("<!DOCTYPE html>"):
            html_code = f"<!DOCTYPE html><html><body>{html_code}</body></html>"

        if self.webview_window:
            self.webview_window.destroy()
        
        # Create a webview window to display HTML content
        self.webview_window = webview.create_window("HTML Preview", html=html_code)
        webview.start()

    def dynamic_preview_frame_function(self):
        html_code = self.html_full_content
        self.webview_window  = None
        if not html_code.startswith("<!DOCTYPE html>"):
            html_code = f"<!DOCTYPE html><html><body>{html_code}</body></html>"

        values = None
        for index, row in self.excel_file_df_to_mail.iterrows():
            values = self.evaluate_body_params(row)
            break

        # Replace placeholders with actual values
        for key, value in values.items():
            # Replace {{key}} with actual value
            html_code = html_code.replace(f"{{{{{key}}}}}", str(value))

        # If the webview window is already open, close it before reopening
        if self.webview_window:
            self.webview_window.destroy()
        
        # Create a webview window to display HTML content
        self.webview_window = webview.create_window("HTML Preview", html=html_code)
        webview.start()

    def dynamic_submit_button(self):
        self.body_params = self.get_entry_data()
        self.preview_frame.configure(corner_radius=0)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)
        self.dynamic_preview_frame_function()
        
        
    def on_entry_change(self, entry, combo_box):
        if entry.get():  # Check if the entry has text
            combo_box.configure(state="disabled")
        else:
            combo_box.configure(state="normal")

            
    def select_frame_by_name(self, name):

        # set button color for selected button
        self.home_button.configure(fg_color=("gray75", "gray25") if name == "home" else "transparent")
        self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "frame_2" else "transparent")

        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.second_frame.grid_forget()

    # def segment_dynamic(self):
    #     self.change_segment_event("Dynamic")
    
    # def segment_static(self):
    #     self.change_segment_event("Static")
  
    def home_button_event(self):
        # Frame Home
        self.select_frame_by_name("home")
 
        

    def frame_2_button_event(self):
        # Frame 2
        self.select_frame_by_name("frame_2")


    def change_appearance_mode_event(self, new_appearance_mode):
        # Apperance Mode
        customtkinter.set_appearance_mode(new_appearance_mode)
    
    def clear_entry_text(self,event=None):
        self.entry_dynamic.delete(0,"end")

    def static_attach_files_function(self):
        file_path = filedialog.askopenfilenames(
            title="Select files upto 25MB"
        )
        total_file_size = 0.0
        for file in file_path:
            file_size = os.path.getsize(file)
            total_file_size += file_size
        if total_file_size >= 26000000:
            messagebox.showwarning("Error",f"Please select the file size upto 25MB\nYour selected file size:{total_file_size/1048576:.2f}MB")
        else:
            messagebox.showinfo("File Selected", "Successfully uploaded")
            for static_path in file_path:
                if os.path.exists(static_path):
                    self.attachment_file_path_list.append(static_path)
            self.static_attachment_file_count = len(self.attachment_file_path_list)            

    def upload_file(self):
        file_path = filedialog.askopenfilename(
          filetypes=[("Excel files", "*.xlsx *.xls")],
          title="Select an Excel file"
        )
        if file_path:
        # You can add the logic to process the Excel file here
            try:
                
                self.excel_file_df_from_mail = pd.read_excel(file_path, sheet_name="Sheet1")
                self.excel_file_df_to_mail = pd.read_excel(file_path,sheet_name="Sheet2")
                self.excel_file_path = file_path
                excel_header = pd.read_excel(file_path,sheet_name="Sheet2")
                self.excel_file_to_mail_header_list = excel_header.columns.tolist()
                messagebox.showinfo("File Selected", "Excel file read successfully")
                self.frame_2_button_event()
            except PermissionError:
                messagebox.showwarning("Error",f"The file '{file_path}' is already open. Please close it and try again.")
            except Exception as e:
                messagebox.showwarning("Error","Excel file error")
        else:
           messagebox.showwarning("No File", "Please select an Excel file.")

    def upload_html_file(self):
        self.entry_dynamic.delete(0,"end")
        file_path = filedialog.askopenfilename(
        filetypes=[("Html files", "*.html *.htm *.txt")],
        title="Select an Excel file"
        )
        if file_path:
                self.entry_dynamic.insert(0,file_path)
                self.textbox_dynamic.delete(1.0,"end")
                if re.search(".txt$",file_path):
                    with open(file_path,"r") as html_content:
                        file_content_str= html_content.read()   
                else:
                    with open(file_path,"rb") as html_content:
                        file_content = html_content.read()
                    file_content_str = file_content.decode('utf-8')
                self.html_full_content = file_content_str
                self.seg_button_1.configure(state=customtkinter.DISABLED)
                #self.dynamic_frame.grid_forget()
                self.current_html_state = self.html_state[0]
                #self.sub_attach_function()
            # except Exception as e:
            #     messagebox.showwarning("Error", "Html/Text file Error")
        else:
           messagebox.showwarning("No File", "Please select an Html/Text file.")
        
    def upload_custom_file(self):
        
        input_value = self.textbox_dynamic.get("1.0", "end")  # Get text from textbox
        if len(input_value) > 1 and input_value.strip() != "":
            self.seg_button_1.configure(state=customtkinter.DISABLED)
            self.html_full_content = input_value
            self.dynamic_frame.grid_forget()
            self.current_html_state = self.html_state[0]
            self.sub_attach_function()
        else:
            print("Empty")

    def upload_static_html_file(self):
        
        file_path = filedialog.askopenfilename(
        filetypes=[("Html files", "*.html *.htm *.txt")],
        title="Select an Excel file"
        )
        if file_path:
                self.entry_dynamic.delete(0,"end")
                self.entry_static.insert(0,file_path)
                self.textbox_static.delete(1.0,"end")
                if re.search(".txt$",file_path):
                    with open(file_path,"r") as html_content:
                        file_content_str= html_content.read()   
                else:
                    with open(file_path,"rb") as html_content:
                        file_content = html_content.read()
                    file_content_str = file_content.decode('utf-8')
                self.html_full_content = file_content_str
                self.seg_button_1.configure(state=customtkinter.DISABLED)
                # self.dynamic_frame.grid_forget()
                self.current_html_state = self.html_state[1]
                #self.sub_attach_function()
            # except Exception as e:
            #     messagebox.showwarning("Error", "Html/Text file Error")
        else:
           messagebox.showwarning("No File", "Please select an Html/Text file.")

    def upload_static_custom_file(self):
        input_value = self.textbox_static.get("1.0", "end")  # Get text from textbox
        
        if len(input_value) > 1 and input_value.strip() != "":
            self.html_full_content = input_value
            self.seg_button_1.configure(state=customtkinter.DISABLED)
            self.dynamic_frame.grid_forget()
            self.current_html_state = self.html_state[1]
            self.sub_attach_function()
        else:
            print("Empty")
        
    # def clear_text(self):
    #     self.entry_dynamic.delete(1)

    def evaluate_body_params(self, row):
        body_params = {}
        if self.body_params:
            for key, value in self.body_params.items():
                if value.startswith("row['") and value.endswith("']"):
                    col_name = value[5:-2]
                    body_params[key] = row[col_name]
                else:
                    body_params[key] = value
        return body_params

    def mail_processer(self):
            error_mail_id = []
            excel_header = pd.read_excel(self.excel_file_path,sheet_name="Sheet2")
            excelfile_to_mail_header_list = excel_header.columns.tolist()
            for index, row in self.excel_file_df_to_mail.iterrows():
                    # print(f"To mail {index+1},{row[0]}, {row[1]}")
                    try:
                        for i, email_data in self.excel_file_df_from_mail.iterrows():
                            if email_data.iloc[0] not in error_mail_id:
                                if "MailStatus" in excelfile_to_mail_header_list:
                                    if row["MailStatus"] == "Completed":
                                        continue
                                try:
                                    print(index, row.iloc[0])
                                    for path in self.individual_attachments_header:
                                        if os.path.exists(rf"{row[path]}"):
                                            self.attachment_file_path_list.append(rf"{row[path]}")
                                    recipient_email = row.iloc[0] # Assuming your Excel file has a column named 'Email'
                                    gmail.username =   email_data.iloc[0]# Notification mail sent to registered mail of customer
                                    gmail.password = email_data.iloc[1]
                                    gmail.send(subject = self.email_subject,
                                                receivers = [recipient_email],
                                                html =self.html_full_content,
                                    body_params=self.evaluate_body_params(row),
                                    attachments=self.attachment_file_path_list)
                                    self.attachment_file_path_list = self.attachment_file_path_list[:self.static_attachment_file_count]
                                    self.excel_file_df_to_mail['FromMail'] =  email_data.iloc[0] 
                                    self.excel_file_df_to_mail['MailStatus'] =  "Completed"
                                    break  
                                except Exception as e:
                                    error_mail_id.append(email_data.iloc[0])
                                    print(f"{email_data.iloc[0]} added to error list")
                    except Exception as e:
                        print(e)

            # Step 2: Create an ExcelWriter object
            with pd.ExcelWriter(self.excel_file_path, engine='openpyxl') as writer:
                # Write each DataFrame to a different sheet
                self.excel_file_df_from_mail.to_excel(writer, sheet_name='Sheet1', index=False)
                self.excel_file_df_to_mail.to_excel(writer, sheet_name='Sheet2', index=False)         
           
    def start_email(self):
        # Need to set the try catch methods
        self.mail_processer()

if __name__ == "__main__":
    app = App()
    app.mainloop()


