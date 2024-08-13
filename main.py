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
import threading
from PIL import Image, ImageTk
import copy



class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # Created global variables for storing the user given data
        self.excel_file_df_from_mail = None
        self.excel_file_df_to_mail = None
        self.html_full_content = None
        self.full_body_text_content = None
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
        self.break_flag = 0
        self.total_email_data_count = 0
        self.completed_count = 0
        
        # self.iconbitmap("WTS.ico")
        # self.resizable(False, False)  
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
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0,fg_color="#0B0C10")
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(8, weight=1) 

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="Email Zapprz", compound="right", font=customtkinter.CTkFont(size=30, weight="bold",family="times"),text_color="#66FCF1")
        self.navigation_frame_label.grid(row=2, column=0, padx=20, pady=50)
        
        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=5, height=40, border_spacing=0, width=200,font=customtkinter.CTkFont(size=15,family="times"),text="Excel Uploader", fg_color="#0B0C10", text_color=("#C5C6C7", "#C5C6C7"), anchor="w", command=self.home_button_event)
        self.home_button.grid(row=3, column=0)

        self.frame_2_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=5, height=40, width=200,font=customtkinter.CTkFont(size=15,family="times"), border_spacing=0, text="Html Uploader", fg_color="#1F2833", text_color=("#C5C6C7", "#C5C6C7"), anchor="w", command=self.frame_2_button_event)
        self.frame_2_button.grid(row=4, column=0)

        # self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        # self.appearance_mode_menu.grid(row=8, column=0, padx=20, pady=(10,50), sticky="s")


        # Home Frame 
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#B3B4BD")
        self.home_frame.grid_columnconfigure(0, weight=20)



        self.home_frame_label = customtkinter.CTkLabel(self.home_frame, text="Email Credentials Files",font=CTkFont(family="times",size=40,weight="bold"),height=40,text_color="#66FCF1")
        self.home_frame_label.grid(row=0, column=0, padx=20, pady=80)

        self.home_frame_qoute1_label = customtkinter.CTkLabel(self.home_frame, text="Select the following options",font=CTkFont(family="times",size=20,weight="bold",slant="italic"),text_color="#C5C6C7")
        self.home_frame_qoute1_label.grid(row=1, column=0, padx=20, pady=20)

        self.path_frame_button = CTkButton(self.home_frame, text="Excel Path",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#45A29E',hover=True,fg_color='#1F2833',height=40,border_color="#66FCF1",border_width=1,text_color="#C5C6C7",corner_radius=0,command=self.upload_file)
        self.path_frame_button.grid(row=2, column=0, padx=20, pady=10)

        # self.url_frame_button = CTkButton(self.home_frame, text="File URL",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=40,border_color="dark",text_color="#1c1c1c",corner_radius=10)
        # self.url_frame_button.grid(row=3, column=0, padx=20, pady=20)
        
        # self.url_frame_button = customtkinter.CTkLabel(self.home_frame, text="Expecting\nSheet1 --> From Mail Credential \nSheet2 --> To Mail Credentials",font=CTkFont(family="times",size=15,weight="bold"),fg_color='#3b8ed0',height=40,text_color="#1c1c1c",corner_radius=10)
        # self.url_frame_button.grid(row=4, column=0, padx=20, pady=40)


        # Second Frame
        self.second_frame = customtkinter.CTkScrollableFrame(self, corner_radius=0, fg_color="#1F2833")
        self.second_frame.grid_columnconfigure(0, weight=1)
        self.second_frame.grid_rowconfigure(1, weight=1)

        self.seg_button_1 = customtkinter.CTkSegmentedButton(self.second_frame,font=CTkFont(family="times",size=18,weight="bold"),width=500,command=self.change_segment_event,fg_color="#0B0C10",bg_color="#0B0C10",corner_radius=20,selected_color="#45A29E",text_color="#0B0C10",unselected_color="#1F2833")
        self.seg_button_1.grid(row=0, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.seg_button_1.configure(values=["Dynamic", "Static"])
        self.seg_button_1.set("Dynamic")


        # Second Frame(Dynamic Frame)
        self.dynamic_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="#1F2833",width=500,height=500)
        self.dynamic_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.dynamic_frame.grid_rowconfigure(1, weight=1)  # Allows the textbox row to expand

        self.entry_dynamic = customtkinter.CTkEntry(self.dynamic_frame, placeholder_text="Upload....")
        self.entry_dynamic.grid(row=0, column=0, columnspan=2, padx=(40, 1), pady=(10, 10), sticky="ew")
        self.entry_dynamic.bind("<KeyRelease>",lambda event: self.clear_textbox_dynamic())

        self.dynamic_upload_button=CTkButton(self.dynamic_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10,command=self.upload_html_file)
        self.dynamic_upload_button.grid(row=0, column=2, columnspan=3, padx=(10, 40), pady=(10, 10), sticky="ew")

        self.textbox_dynamic = customtkinter.CTkTextbox(self.dynamic_frame, width=700,fg_color="grey",text_color="white", height=340)
        self.textbox_dynamic.grid(row=1, column=0, columnspan=3, padx=(40, 40), pady=(10, 10), sticky="nsew")
        self.textbox_dynamic.insert("2.0","Html Code goes here.../ Upload the html file")
        self.textbox_dynamic.bind("<FocusIn>",self.clear_placeholder)
        self.textbox_dynamic.bind("<KeyRelease>",lambda event: self.clear_entry_text())

        self.dynamic_main_back_button=CTkButton(self.dynamic_frame,text="Back",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10,command=self.main_back_button_func)
        self.dynamic_main_back_button.grid(row=2, column=0, padx=(10, 40), pady=(10, 10))

        self.dynamic_sub_button=CTkButton(self.dynamic_frame,text="Submit",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.dynamic_sub_button_func)
        self.dynamic_sub_button.grid(row=2, column=1, columnspan=2, padx=(150, 80), pady=(10, 10), sticky="ew")


        # Second Frame(Static Frame)
        self.static_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.static_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.static_frame.grid_forget()

        self.entry_static = customtkinter.CTkEntry(self.static_frame, placeholder_text="Upload....")
        self.entry_static.grid(row=0, column=0, columnspan=2, padx=(40, 1), pady=(10, 10), sticky="ew")
        self.entry_static.bind("<KeyRelease>",lambda event: self.clear_textbox_static())
        
        self.static_upload_button=CTkButton(self.static_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.upload_static_html_file)
        self.static_upload_button.grid(row=0, column=2, columnspan=3, padx=(10, 40), pady=(10, 10), sticky="ew")

        self.textbox_static = customtkinter.CTkTextbox(self.static_frame, width=700,fg_color="grey",text_color="white", height=340)
        self.textbox_static.grid(row=1, column=0, columnspan=3, padx=(40, 40), pady=(10, 10), sticky="nsew")
        self.textbox_static.insert("2.0","Html Code goes here.../ Upload the html file")
        self.textbox_static.bind("<FocusIn>",self.clear_placeholder)
        self.textbox_static.bind("<KeyRelease>",lambda event: self.clear_entry_text_static())

        self.static_main_back_button=CTkButton(self.static_frame,text="Back",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10,command=self.main_back_button_func)
        self.static_main_back_button.grid(row=2, column=0, padx=(10, 40), pady=(10, 10))

        self.static_sub_button=CTkButton(self.static_frame,text="Submit",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.static_sub_button_func)
        self.static_sub_button.grid(row=2, column=1, columnspan=2, padx=(150, 80), pady=(10, 10), sticky="ew")


        # Second Frame(Subject and Attachments Frame)
        self.subject_attach_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white", width=700,height=500)
        self.subject_attach_frame.grid(row=0, column=0, padx=(20, 20), pady=(10, 10), sticky="nsew")
        self.subject_attach_frame.grid_rowconfigure(1, weight=0)
        self.subject_attach_frame.grid_columnconfigure(1, weight=0)
        

        self.subject_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Subject",font=customtkinter.CTkFont(family="times", size=14, weight="bold"), height=40)
        self.subject_label.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="w")

        self.subject_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="Enter the subject")
        self.subject_entry.grid(row=0, column=1, columnspan=2, padx=(10, 20), pady=(10, 10), sticky="ew")
#####################################
        self.cc_label = customtkinter.CTkLabel(self.subject_attach_frame, text="CC",font=customtkinter.CTkFont(family="times", size=14, weight="bold"), height=40)
        self.cc_label.grid(row=1, column=0, padx=(20, 10), pady=(5, 10), sticky="w")

        self.cc_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="")
        self.cc_entry.grid(row=1, column=1, columnspan=2, padx=(10, 20), pady=(5, 10), sticky="ew")
        
        self.bcc_label = customtkinter.CTkLabel(self.subject_attach_frame, text="BCC",font=customtkinter.CTkFont(family="times", size=14, weight="bold"), height=40)
        self.bcc_label.grid(row=2, column=0, padx=(20, 10), pady=(5, 10), sticky="w")

        self.bcc_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="")
        self.bcc_entry.grid(row=2, column=1, columnspan=2, padx=(10, 20), pady=(5, 10), sticky="ew")
################################

        self.attach_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Static Attachments",font=customtkinter.CTkFont(family="times", size=14, weight="bold"), height=40)
        self.attach_label.grid(row=3, column=0, padx=(20, 10), pady=(5, 10), sticky="w")

        self.attachment_files = customtkinter.CTkButton(self.subject_attach_frame, text="Upload",font=customtkinter.CTkFont(family="times", size=14, weight="bold"),hover_color='#808080', fg_color='#3b8ed0', height=10,text_color="#1c1c1c", corner_radius=10,command=self.static_attach_files_function)
        self.attachment_files.grid(row=3, column=1, padx=(10, 20), pady=(5, 10))

        self.attachment_preview_button = customtkinter.CTkButton(self.subject_attach_frame, text="Preview",font=customtkinter.CTkFont(family="times", size=14, weight="bold"),hover_color='#808080', fg_color='#3b8ed0', height=10,text_color="#1c1c1c", corner_radius=10,command=self.attachment_preview_button_function)
        self.attachment_preview_button.grid(row=3, column=2, padx=(10, 20), pady=(5, 10) )

        self.individual_attach_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Individual Attachments",font=customtkinter.CTkFont(family="times", size=14, weight="bold"), height=40)
        self.individual_attach_label.grid(row=4, column=0, padx=(20, 10), pady=(5, 10))

       
        self.dynamic_scroll_checkbox_frame = customtkinter.CTkScrollableFrame(self.subject_attach_frame,label_text="Select Below columns", height=100,width=500)
        self.dynamic_scroll_checkbox_frame.grid(row=4, column=1,columnspan=2, padx=(10, 20), pady=(5, 10))
        self.dynamic_scrollable_frame_checkbox = []
        self.subject_attach_frame.grid_forget()


        # Second Frame(Email Mapping Frame( For Dynamic Contents))
        self.list_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white", width=700, height=500)
        self.list_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.list_frame.grid_forget()

        # Final frame and Email Loading Bar
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=50, fg_color="black")
        self.third_frame.grid_columnconfigure(0, weight=20)
        self.third_frame.grid_forget()

        self.select_frame_by_name(name="home")

        self.frame_2_button.configure(state="disabled",fg_color= "transparent",text_color="black")

    
    def attachment_preview_button_function(self):
        print("dbrgefdf")
        pass

    # home_button
    def home_button_event(self):
        self.select_frame_by_name("home")

    
    # frame_2_button
    def frame_2_button_event(self):
        # Frame 2
        self.frame_2_button.configure(state="normal")
        self.home_button.configure(state="disabled")
        self.select_frame_by_name("frame_2")

    
    #appearance_mode_menu
    # def change_appearance_mode_event(self, new_appearance_mode):
    #     # Apperance Mode
    #     customtkinter.set_appearance_mode(new_appearance_mode)


    # Home Frame (path_frame_button)  
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
                self.seg_button_1.configure(state="normal")
                if "MailStatus" in self.excel_file_to_mail_header_list:
                    for index, row in self.excel_file_df_to_mail.iterrows():
                        if row.get("MailStatus"):
                            if row["MailStatus"] != "Completed":
                                self.total_email_data_count += 1
                else:
                    self.total_email_data_count = len(self.excel_file_df_to_mail)
                if self.total_email_data_count != 0:
                    self.frame_2_button_event()
                else:
                    messagebox.showwarning("Warning","Emails Sent Already")

            except PermissionError:
                messagebox.showwarning("Error",f"The file '{file_path}' is already open. Please close it and try again.")
            except Exception as e:
                messagebox.showwarning("Error","Excel file error")
        else:
           messagebox.showwarning("No File", "Please select an Excel file.")


    # Second Frame(seg_button_1)
    def change_segment_event(self,name):
        if name == "Dynamic":
            self.textbox_static.delete(1.0,"end")
            self.entry_static.delete(0, "end")
            self.dynamic_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
            self.dynamic_frame.grid(row=1, column=0,  padx=(20, 20), pady=(10, 20), sticky="ew")
        else:
            self.dynamic_frame.grid_forget()
        if name == "Static":
            self.textbox_dynamic.delete(1.0,"end")
            self.entry_dynamic.delete(0, "end")
            self.static_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
            self.static_frame.grid(row=1, column=0,  padx=(20, 20), pady=(10, 20), sticky="ew")
        else:
            self.static_frame.grid_forget()


    # Second Frame(Dynamic Frame(entry_dynamic.bind))
    def clear_textbox_dynamic(self):
        self.textbox_dynamic.delete(1.0,"end")


    # Second Frame(Dynamic Frame(dynamic_upload_button))
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
                    with open(file_path,"r", encoding='utf-8') as html_content:
                        file_content_str= html_content.read()
                        self.full_body_text_content = file_content_str
                        self.html_full_content = None
                else:
                    with open(file_path,"rb") as html_content:
                        file_content = html_content.read()
                    file_content_str = file_content.decode('utf-8')
                    self.html_full_content = file_content_str
                    self.full_body_text_content = None
        else:
           messagebox.showwarning("No File", "Please select an Html/Text file.")
           
           


    # Second Frame(Dynamic Frame(textbox_dynamic.bind))
    def clear_placeholder(self, event): # function for clearing the place holder
        if self.textbox_dynamic.get("1.0", "end-1c") == "Html Code goes here.../ Upload the html file":
            self.textbox_dynamic.delete("1.0", "end-1c")

        if self.textbox_static.get("1.0", "end-1c") == "Html Code goes here.../ Upload the html file":
            self.textbox_static.delete("1.0", "end-1c")

    
    # Second Frame(Dynamic Frame(textbox_dynamic.bind))
    def clear_entry_text_static(self):
        self.entry_static.delete(0,"end") 


    # Second Frame(Dynamic Frame(textbox_dynamic.bind))
    def clear_entry_text(self,event=None):
        self.entry_dynamic.delete(0,"end")


    # Second Frame(Dynamic Frame(dynamic_main_back_button))
    def main_back_button_func(self):
        self.second_frame.grid_forget()
        self.home_button.configure(state="normal")
        self.frame_2_button.configure(state="disabled")
        self.select_frame_by_name(name="home")

    # Second Frame(Dynamic Frame(dynamic_sub_button))
    def dynamic_sub_button_func(self):
        dynamic_text_value = self.textbox_dynamic.get("1.0", "end")  # Get text from textbox
        dynamic_upload_value = self.entry_dynamic.get()
        self.current_html_state = self.html_state[0]
        self.seg_button_1.configure(state=customtkinter.DISABLED)
        try:
            self.attachment_back_button.destroy()
            self.attachment_sub_button.destroy()
            self.static_preview_logo_button.destroy()
        except Exception as e:
            pass

        if (dynamic_text_value.strip() == "" or dynamic_text_value.strip() == "Html Code goes here.../ Upload the html file") and dynamic_upload_value.strip() == "":
            self.seg_button_1.configure(state="normal")
            messagebox.showwarning("Error","Please enter a value/select the file")
        else:
            if dynamic_upload_value.strip() != "":
                if os.path.exists(rf"{dynamic_upload_value.strip("\"")}"):

                    # Second Frame(Subject and Attachments Frame (Dynamic - Back and Submit Button))
                    self.attachment_back_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
                    self.attachment_back_button.grid(row=5, column=0, padx=(20, 10), pady=(20, 10), sticky="w")

                    self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                    self.attachment_sub_button.grid(row=5, column=1, columnspan=2, padx=(10, 20), pady=(20, 10), sticky="e")
                    
                    self.sub_attach_function()
                else:
                    messagebox.showwarning("Error","File not found")
            else:
                if (dynamic_text_value.startswith("<!DOCTYPE html>")) or (dynamic_text_value.startswith("<html>")) :
                    self.html_full_content = dynamic_text_value
                    self.full_body_text_content = None
                else:
                    self.full_body_text_content = dynamic_text_value
                    self.html_full_content = None
                
                # Second Frame(Subject and Attachments Frame (Dynamic - Back and Submit Button))
                self.attachment_back_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
                self.attachment_back_button.grid(row=5, column=0, padx=(10, 10), pady=(10, 10), sticky="w")

                self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                self.attachment_sub_button.grid(row=5, column=1, columnspan=2, padx=(10, 10), pady=(10, 10), sticky="e")
                self.sub_attach_function()


    # Second Frame(Static Frame(entry_static.bind))
    def clear_textbox_static(self):
        self.textbox_static.delete(1.0,"end")

    
    # Second Frame(Static Frame(static_upload_button))
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
                    with open(file_path,"r", encoding='utf-8') as html_content:
                        file_content_str= html_content.read()   
                        self.full_body_text_content = file_content_str
                        self.html_full_content = None
                else:
                    with open(file_path,"rb") as html_content:
                        file_content = html_content.read()
                    file_content_str = file_content.decode('utf-8')
                    self.html_full_content = file_content_str
                    self.full_body_text_content = None
        else:
           messagebox.showwarning("No File", "Please select an Html/Text file.")

    
    # Second Frame(Static Frame(static_sub_button))
    def static_sub_button_func(self):
        static_text_value = self.textbox_static.get("1.0", "end")
        static_upload_value = self.entry_static.get()
        self.current_html_state = self.html_state[1]
        self.seg_button_1.configure(state=customtkinter.DISABLED)
        try:
            self.attachment_back_button.destroy()
            self.attachment_sub_button.destroy()
        except Exception as e:
            pass
        if (static_text_value.strip() == "" or static_text_value.strip() == "Html Code goes here.../ Upload the html file") and static_upload_value.strip() == "":
            self.seg_button_1.configure(state="normal")
            messagebox.showwarning("Error","Please enter a value/select the file")
        else:
            if static_upload_value.strip() != "":
                if os.path.exists(rf"{static_upload_value.strip("\"")}"):

                    # Subject and Attachments Frame (Static - Back and Submit Button)
                    self.attachment_back_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
                    self.attachment_back_button.grid(row=3, column=0, padx=(10, 10), pady=(10, 10), sticky="w")

                    self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                    self.attachment_sub_button.grid(row=3, column=1, padx=(100, 10), pady=(10, 10))

                    self.static_preview_logo = Image.open(r"logo\eye.png")
                    self.static_preview_logo_button = CTkButton(self.subject_attach_frame,corner_radius=30, text="",image=CTkImage(self.static_preview_logo), fg_color="transparent", border_color="green",border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),width=10, height=10, command=self.static_preview_function)
                    self.static_preview_logo_button.grid(row=3, column=1,padx=(400, 10), pady=(10, 10))

                    self.sub_attach_function()
                else:
                    messagebox.showwarning("Error","File not found")
            else:
                if (static_text_value.startswith("<!DOCTYPE html>")) or (static_text_value.startswith("<html>")) :
                    self.html_full_content = static_text_value
                    self.full_body_text_content = None
                else:
                    self.full_body_text_content = static_text_value
                    self.html_full_content = None

                # Subject and Attachments Frame (Static - Back and Submit Button)
                self.attachment_back_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
                self.attachment_back_button.grid(row=3, column=0, padx=(10, 10), pady=(10, 10), sticky="w")

                self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                self.attachment_sub_button.grid(row=3, column=1, padx=(100, 10), pady=(10, 10))

                self.static_preview_logo = Image.open(r"logo\eye.png")
                self.static_preview_logo_button = CTkButton(self.subject_attach_frame,corner_radius=30, text="",image=CTkImage(self.static_preview_logo),fg_color="transparent", border_color="green",border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),width=10, height=10, command=self.static_preview_function)
                self.static_preview_logo_button.grid(row=3, column=1,padx=(400, 10), pady=(10, 10))
                
                self.sub_attach_function()


    # Second Frame(Subject and Attachments Frame(attachment_files))
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
            for static_path in file_path:
                if os.path.exists(static_path):
                    self.attachment_file_path_list.append(static_path)
            self.static_attachment_file_count = len(self.attachment_file_path_list)  
            list(file_path).clear()          
            if self.static_attachment_file_count > 0:
                messagebox.showinfo("File Selected", "Successfully uploaded")


    # Second Frame(Subject and Attachments Frame(attachment_back_button))
    def attachment_back_function(self):
        self.subject_attach_frame.grid_forget()
        self.dynamic_frame.configure(corner_radius=16)
        self.dynamic_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.seg_button_1.configure(state=customtkinter.NORMAL)


    # Second Frame(Subject and Attachments Frame(attachment_sub_button))
    def attachment_sub_function(self):
        data_dict = []
        for checkbox in self.dynamic_scrollable_frame_checkbox:
            if checkbox.get() == 1:
                data_dict.append(checkbox.cget("text"))
        self.individual_attachments_header = data_dict
        data_dict = []
        self.email_subject = self.subject_entry.get()
        if len(self.subject_entry.get()) >= 1 and self.subject_entry.get().strip() != "":                
            self.list_frame_show_call()
        else:
            messagebox.showwarning("Error", "Please fill the email subject")

    
    # Second Frame(Subject and Attachments Frame)
    def sub_attach_function(self):
        self.static_frame.grid_forget()
        self.dynamic_frame.grid_forget()
        # self.list_frame.grid_forget()
        self.subject_attach_frame.configure(corner_radius=16, fg_color="white",width=500,height=500)
        self.subject_attach_frame.grid(row=1, column=0,  padx=(20, 20), pady=(20, 20), sticky="ns")

        for index, key in enumerate(self.excel_file_to_mail_header_list):
            checkbox = customtkinter.CTkCheckBox(master=self.dynamic_scroll_checkbox_frame,text=key)
            checkbox.grid(row=index, column=0, pady=5, padx=20, sticky="w")
            self.dynamic_scrollable_frame_checkbox.append(checkbox)

    
    # Second Frame(Subject and Attachments Frame(static_preview_logo_button))
    def static_preview_function(self): 
        self.preview_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="grey")
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)

        self.static_preview_frame_function()


    # Static Preview Frame
    def static_preview_frame_function(self):
        self.webview_window  = None
        if self.html_full_content != None:
            html_code = self.html_full_content
        else:
            html_code = self.full_body_text_content

        if not html_code.startswith("<!DOCTYPE html>"):
            html_code = f"<!DOCTYPE html><html><body>{html_code}</body></html>"

        if self.webview_window:
            self.webview_window.destroy()
        
        # Create a webview window to display HTML content
        self.webview_window = webview.create_window("HTML Preview", html=html_code)
        webview.start()


    # Second Frame(Email Mapping Frame( For Dynamic Contents))
    def list_frame_show_call(self):
        if self.html_full_content != None:
            body_content = self.html_full_content
        else:
            body_content = self.full_body_text_content
        params_variable = re.findall(r"\{\{(\w+)\}\}",body_content)
        if params_variable and self.current_html_state == self.html_state[0]:
            if self.excel_file_to_mail_header_list:
                self.list_frame_show(params_variable, self.excel_file_to_mail_header_list)
            else:
                print("Empty header list")
        else:
            self.dynamic_submit_button()
    

    # Second Frame(Email Mapping Frame( For Dynamic Contents))
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

        # Second Frame(Email Mapping Frame(For Dynamic Contents))
        self.dynamic_back_button = CTkButton(self.list_frame,corner_radius=30, text="Back", fg_color="white", border_color="gray", border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.dynamic_back_button_function)
        self.dynamic_back_button.grid(row=1, column=2,columnspan=2, rowspan=2,padx=(10,350), pady=(350,10))

        self.submit_list_button = customtkinter.CTkButton(self.list_frame, corner_radius=30, text="Submit", fg_color="white", border_color="green", border_width=2, text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.dynamic_submit_button)
        self.submit_list_button.grid(row=1, column=1, columnspan=2, rowspan=2,padx=(350,10), pady=(350, 10))

        self.preview_logo = Image.open(r"logo\eye.png")
        self.preview_logo_button = CTkButton(self.list_frame,corner_radius=30, text="",image=CTkImage(self.preview_logo), fg_color="transparent", border_color="green",border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),width=10, height=10, command=self.preview_button)
        self.preview_logo_button.grid(row=1, column=2,padx=(630,10), pady=(350, 10))

        # create preview frame==>dummy frame created for future use
        self.preview_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="grey")
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)

    
    # Second Frame(Email Mapping Frame( For Dynamic Contents(dynamic_back_button)))
    def dynamic_back_button_function(self):
        self.list_frame.grid_forget()
        self.subject_attach_frame.configure(corner_radius=16)
        self.subject_attach_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nwes")


    # Second Frame(Email Mapping Frame(For Dynamic Contents(submit_list_button)))
    def dynamic_submit_button(self):
        self.body_params = self.get_entry_data()

        self.second_frame.grid_forget()

        # Final frame and Email Loading Bar
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="gray",width=700,height=500)
        self.third_frame.grid(row=0, column=1, sticky="nsew")

        self.progressbar_empty = customtkinter.CTkLabel(master=self.third_frame,text="Click Start Button",height=20, width=500)
        self.progressbar_empty.grid(row=0, column=0, columnspan=2,padx=(150,0), pady=(200,0))
        
        self.progressbar_text = customtkinter.CTkLabel(master=self.third_frame,text=" ",height=20, width=500)
        self.progressbar_text.grid(row=0, column=0, columnspan=2,padx=(150,0), pady=(250,0))

        self.start_button = customtkinter.CTkButton(self.third_frame, corner_radius=50, height=40, border_spacing=10, text="Start", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=self.start_email)
        self.start_button.grid(row=1, column=1, padx=0, pady=(100,0))
        
        self.email_back_button = customtkinter.CTkButton(self.third_frame, corner_radius=50, height=40, border_spacing=10, text="Back", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command = self.email_back_button_func)
        self.email_back_button.grid(row=1, column=0, padx=(150,0), pady=(100,0))


    # Second Frame(Email Mapping Frame(For Dynamic Contents(preview_logo_button)))
    def preview_button(self):
        self.body_params = self.get_entry_data()
        self.preview_frame.configure(corner_radius=0)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)
        self.dynamic_preview_frame_function()


    # Dynamic Preview Frame
    def dynamic_preview_frame_function(self):
        if self.html_full_content != None:
            html_code = self.html_full_content
        else:
            html_code = self.full_body_text_content
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


    # Final frame and Email Loading Bar(email_back_button)
    def email_back_button_func(self):
        self.third_frame.grid_forget()
        self.second_frame.grid(row=0, column=1, sticky="nsew")
        if self.current_html_state == self.html_state[0]:
            self.attachment_sub_function()
        else:
            self.static_sub_button_func()

    
    # Body params function
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
    

    # Main Mail sender function
    def mail_processor(self):
            email_progressbar = customtkinter.CTkProgressBar(master=self.third_frame,height=20, width=500)
            email_progressbar.grid(row=0, column=0, columnspan=2, padx=(150,0), pady=(200,0))
            self.progressbar_text.configure(text=f"0/{self.total_email_data_count}")
            email_progressbar.set(0)
            error_mail_id = []
            excel_header = pd.read_excel(self.excel_file_path,sheet_name="Sheet2")
            excelfile_to_mail_header_list = excel_header.columns.tolist()
            if "FromMail" not in excelfile_to_mail_header_list:
                self.excel_file_df_to_mail['FromMail'] = ""
                if "MailStatus" not in excelfile_to_mail_header_list:
                    
                    self.excel_file_df_to_mail['MailStatus'] = ""
            else:
                if "MailStatus" not in excelfile_to_mail_header_list:
                    self.excel_file_df_to_mail['MailStatus'] = ""

            for index, row in self.excel_file_df_to_mail.iterrows():
                    if self.break_flag == 0:
                        try:
                            for i, email_data in self.excel_file_df_from_mail.iterrows():
                                if email_data.iloc[0] not in error_mail_id:
                                    if "MailStatus" in excelfile_to_mail_header_list:
                                        if row["MailStatus"] == "Completed":
                                            continue
                                    try:
                                        for path in self.individual_attachments_header:
                                            if os.path.exists(rf"{row[path]}"):
                                                self.attachment_file_path_list.append(rf"{row[path]}")
                                        recipient_email = row.iloc[0] # Assuming your Excel file has a column named 'Email'
                                        gmail.username =   email_data.iloc[0]# Notification mail sent to registered mail of customer
                                        gmail.password = email_data.iloc[1]
                                        gmail.send(subject = self.email_subject,
                                                    receivers = [recipient_email],
                                                    bcc = "",
                                                    cc = "dhineshwisetechsource@gmail.com",
                                                    text = self.full_body_text_content,
                                                    html =self.html_full_content,
                                                    body_params=self.evaluate_body_params(row),
                                                    attachments=self.attachment_file_path_list)
                                        self.attachment_file_path_list = self.attachment_file_path_list[:self.static_attachment_file_count]
                                        self.excel_file_df_to_mail.loc[index,'FromMail'] = email_data.iloc[0] 
                                        self.excel_file_df_to_mail.loc[index,'MailStatus'] =  "Completed"
                                        self.completed_count += 1
                                        progressbar_value = (self.completed_count)/self.total_email_data_count
                                        email_progressbar = customtkinter.CTkProgressBar(master=self.third_frame,height=20, width=500)
                                        email_progressbar.grid(row=0, column=0, columnspan=2, padx=(150,0), pady=(200,0))
                                        email_progressbar.set(progressbar_value)
                                        self.progressbar_text.configure(text=f"{self.completed_count}/{self.total_email_data_count}")
                                        break
                                    except Exception as e:
                                        error_mail_id.append(email_data.iloc[0])
                        except Exception as e:
                            pass
                    else:
                        messagebox.showwarning("Stop", "Process Stopped")
                        self.break_flag = 0
                        self.home_frame.grid_forget()
                        self.second_frame.grid_forget()
                        self.third_frame.grid_forget()
                        self.subject_attach_frame.grid_forget()
                        self.list_frame.grid_forget()

                        self.select_frame_by_name(name="home")
                        self.home_button.configure(state="normal")
                        self.frame_2_button.configure(state="disabled")
                        break
            else:
                messagebox.showinfo("Succcess", "Email Sent Successfully")
                self.break_flag = 0
                self.home_frame.grid_forget()
                self.second_frame.grid_forget()
                self.third_frame.grid_forget()
                self.subject_attach_frame.grid_forget()
                self.list_frame.grid_forget()

                self.select_frame_by_name(name="home")
                self.home_button.configure(state="normal")
                self.frame_2_button.configure(state="disabled")

            # Step 2: Create an ExcelWriter object
            with pd.ExcelWriter(self.excel_file_path, engine='openpyxl') as writer:
                # Write each DataFrame to a different sheet
                self.excel_file_df_from_mail.to_excel(writer, sheet_name='Sheet1', index=False)
                self.excel_file_df_to_mail.to_excel(writer, sheet_name='Sheet2', index=False)

            self.excel_file_df_from_mail = None
            self.excel_file_df_to_mail = None
            self.html_full_content = None
            self.full_body_text_content = None
            self.body_params = None
            self.excel_file_path = None
            self.excel_file_to_mail_header_list = []
            self.scrollable_frame_switches = []
            self.attachment_file_path_list = []
            self.static_attachment_file_count = 0
            self.html_state = ["Dynamic","Static"]
            self.current_html_state = None
            self.individual_attachments_header=[]
            self.dynamic_scrollable_frame_checkbox = []
            self.email_subject = ""
            self.break_flag = 0
            self.total_email_data_count = 0
            self.completed_count = 0
    
    # Final frame and Email Loading Bar(start_button)
    def start_email(self):
        self.email_back_button.grid_forget()
        self.start_button.grid_forget()
        self.email_stop_button = customtkinter.CTkButton(self.third_frame, corner_radius=50, height=40, border_spacing=10, text="Stop", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=self.stop_back_button_func)
        self.email_stop_button.grid(row=1, column=1, padx=(0,75), pady=(100,0))
        threading.Thread(target=self.mail_processor).start()


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
    

    def on_entry_change(self, entry, combo_box):
        if entry.get():  # Check if the entry has text
            combo_box.configure(state="disabled")
        else:
            combo_box.configure(state="normal")

       
    def select_frame_by_name(self, name):

        # Deleting old values while second run(if present)
        self.entry_dynamic.delete(0,"end")
        self.textbox_dynamic.delete(1.0,"end")
        self.entry_static.delete(0,"end")
        self.textbox_static.delete(1.0,"end")
        self.subject_entry.delete(0,"end")


        # set button color for selected button
        self.home_button.configure(fg_color=("#1F2833", "#1F2833") if name == "home" else "#0B0C10")
        self.frame_2_button.configure(fg_color=("#1F2833", "#1F2833") if name == "frame_2" else "#0B0C10")

        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
            self.dynamic_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        else:
            self.second_frame.grid_forget()
  
    # Final frame and Email Loading Bar(email_stop_button)    
    def stop_back_button_func(self):
        self.break_flag = 1

if __name__ == "__main__":
    app = App()
    app.mainloop()


