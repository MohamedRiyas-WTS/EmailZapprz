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
import requests
import datetime as dt

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.project_name = "Email Zapperz"
        # Created global variables for storing the user given data
        self.excel_file_df_from_mail = None
        self.excel_file_df_to_mail = None
        self.html_full_content = None
        self.full_body_text_content = None
        self.body_params = None
        self.excel_file_path = None
        self.excel_file_to_mail_header_list = []
        self.scrollable_frame_switches = []
        self.title(self.project_name)
        self.attachment_file_path_list = []
        self.static_attachment_file_count = 0
        # self.html_state = ["Dynamic","Static"]
        self.html_state = ["Advanced","Normal"]
        self.current_html_state = None
        self.individual_attachments_header = []
        self.replacer_keyword_list = []
        self.email_subject = ""
        self.break_flag = 0
        self.total_email_data_count = 0
        self.completed_count = 0
        self.url = "http://www.google.com"
        self.timeout = 5
        self.email_cc = None
        self.email_bcc = None
        self.excel_sheet_name = ["From_Mail", "To_Mail", "Guide"]
        self.excel_to_mail_header_changing_data = ["FromMail", "MailStatus"]
        self.iconbitmap(r"logo\zapperz_logo.ico")
        
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
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0,fg_color="#00246B")
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(8, weight=1) 

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text=self.project_name, compound="right", font=customtkinter.CTkFont(size=30, weight="bold",family="Comic Sans MS"),text_color="white")
        self.navigation_frame_label.grid(row=2, column=0, padx=20, pady=50)
        
        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=5, height=40, border_spacing=0, width=200,font=customtkinter.CTkFont(size=17,family="Times New Roman"),text="Excel Uploader", text_color="#00246B", anchor="w",hover=True,hover_color="white", command=self.home_button_event)
        self.home_button.grid(row=3, column=0)

        self.frame_2_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=5, height=40, width=200,font=customtkinter.CTkFont(size=17,family="Times New Roman"), border_spacing=0, text="Html Uploader", text_color="#00246B", anchor="w",hover=True,hover_color="white", command=self.frame_2_button_event)
        self.frame_2_button.grid(row=4, column=0)

        # self.template_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=5, height=30, width=150,font=customtkinter.CTkFont(size=17,family="Times New Roman"), border_spacing=0, text="Template", text_color="#00246B",hover=True,hover_color="white", command=self.template_button_func, fg_color="#CADCFC")
        # self.template_button.grid(row=7, column=0, pady=(300,50), sticky="s")

        # self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        # self.appearance_mode_menu.grid(row=8, column=0, padx=20, pady=(10,50), sticky="s")


        # Home Frame 
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#CADCFC")
        self.home_frame.grid_columnconfigure(0, weight=20)

        self.logo_image = customtkinter.CTkImage(Image.open(r"logo\zapperz_logo.png"), size=(150, 150))
        self.logo = customtkinter.CTkLabel(self.home_frame, text="", image=self.logo_image,
                                                             compound="right", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.logo.grid(row=0, column=0, pady=10)

        self.home_frame_label = customtkinter.CTkLabel(self.home_frame, text="Email Credentials Files",font=CTkFont(family="Times New Roman",size=40,weight="bold"),height=40,text_color="#00246B")
        self.home_frame_label.grid(row=1, column=0, padx=20, pady=80)

        self.home_frame_qoute1_label = customtkinter.CTkLabel(self.home_frame, text="Select the following options",font=CTkFont(family="Times New Roman",size=20,weight="bold",slant="italic"),text_color="black")
        self.home_frame_qoute1_label.grid(row=2, column=0, padx=20, pady=20)

        self.path_frame_button = CTkButton(self.home_frame, text="Excel Path",font=CTkFont(family="Times New Roman",size=20,weight="bold"),hover_color='#A7BEAE',hover=True,fg_color='#00246B',height=40,border_color="#A7BEAE",border_width=1,text_color="white",corner_radius=0,command=self.upload_file)
        self.path_frame_button.grid(row=3, column=0, padx=20, pady=10)

        # self.url_frame_button = CTkButton(self.home_frame, text="File URL",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=40,border_color="dark",text_color="#1c1c1c",corner_radius=10)
        # self.url_frame_button.grid(row=3, column=0, padx=20, pady=20)
        
        # self.url_frame_button = customtkinter.CTkLabel(self.home_frame, text="Expecting\nSheet1 --> From Mail Credential \nSheet2 --> To Mail Credentials",font=CTkFont(family="times",size=15,weight="bold"),fg_color='#3b8ed0',height=40,text_color="#1c1c1c",corner_radius=10)
        # self.url_frame_button.grid(row=4, column=0, padx=20, pady=40)

        # Second Frame
        self.second_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#CADCFC")
        self.second_frame.grid_columnconfigure(0, weight=1)
        self.second_frame.grid_rowconfigure(1, weight=1)

        self.seg_button_1 = customtkinter.CTkSegmentedButton(self.second_frame,font=CTkFont(family="Times New Roman",size=18,weight="bold"),width=500,command=self.change_segment_event,fg_color="#CADCFC",bg_color="#CADCFC",corner_radius=20,selected_color="#00246B",text_color="white",unselected_color="gray")
        self.seg_button_1.grid(row=0, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        # self.seg_button_1.configure(values=["Dynamic", "Static"])
        # self.seg_button_1.set("Dynamic")
        self.seg_button_1.configure(values=["Advanced", "Normal"])
        self.seg_button_1.set("Advanced")

        # Adjust the dynamic frame to full screen
        self.dynamic_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="#CADCFC")
        self.dynamic_frame.grid(row=1, column=0, padx=(0, 0), pady=(10, 10), sticky="nsew")
        self.dynamic_frame.grid_rowconfigure(1, weight=1)  # Allows the textbox row to expand
        self.dynamic_frame.grid_columnconfigure(0, weight=1)  # Allows the entry and buttons to expand


        # Adjust the entry field to take full width
        self.entry_dynamic = customtkinter.CTkEntry(self.dynamic_frame, placeholder_text="Upload....")
        self.entry_dynamic.grid(row=0, column=0, padx=(40, 10), pady=(10, 10), sticky="ew") 
        self.entry_dynamic.bind("<KeyRelease>", lambda event: self.clear_textbox_dynamic())

        # Adjust the upload button
        self.dynamic_upload_button = CTkButton(self.dynamic_frame, text="Upload", font=CTkFont(family="Times New Roman", size=18, weight="bold"), hover_color='#808080', hover=True, fg_color='#3b8ed0', height=10, border_color="dark", text_color="#1c1c1c", corner_radius=10, command=self.upload_html_file)
        self.dynamic_upload_button.grid(row=0, column=1, padx=(0, 40), pady=(10, 10), sticky="e")

        # Adjust the textbox to take full space
        self.textbox_dynamic = customtkinter.CTkTextbox(self.dynamic_frame, fg_color="white", text_color="black", width=700,height=340)
        self.textbox_dynamic.grid(row=1, column=0, columnspan=2, padx=(40, 40), pady=(10, 10), sticky="nsew")
        self.textbox_dynamic.insert("2.0", "Html Code goes here.../ Upload the html file")
        self.textbox_dynamic.bind("<FocusIn>", self.clear_placeholder)
        self.textbox_dynamic.bind("<KeyRelease>", lambda event: self.clear_entry_text())

        # Adjust the back button position and size
        self.dynamic_main_back_button = customtkinter.CTkButton(self.dynamic_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=self.main_back_button_func)
        self.dynamic_main_back_button.grid(row=2, column=0, padx=(40, 10), pady=(10, 10), sticky="w")

        # Adjust the submit button to take full width of its section
        self.dynamic_sub_button = CTkButton(self.dynamic_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"), command=self.dynamic_sub_button_func)
        self.dynamic_sub_button.grid(row=2, column=1, padx=(0, 40), pady=(10, 10), sticky="e")


        # Second Frame(Static Frame)
        self.static_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="#CADCFC")
        self.static_frame.grid(row=1, column=0, padx=(0, 0), pady=(10, 10), sticky="nsew")
        self.static_frame.grid_rowconfigure(1, weight=1)  # Allows the textbox row to expand
        self.static_frame.grid_columnconfigure(0, weight=1)  # Allows the entry and buttons to expand
        self.static_frame.grid_forget()

        # Adjust the entry field to take full width
        self.entry_static = customtkinter.CTkEntry(self.static_frame, placeholder_text="Upload....")
        self.entry_static.grid(row=0, column=0, padx=(40, 10), pady=(10, 10), sticky="ew")
        self.entry_static.bind("<KeyRelease>", lambda event: self.clear_textbox_static())

        self.static_upload_button = CTkButton(self.static_frame, text="Upload", font=CTkFont(family="Times New Roman", size=18, weight="bold"), hover_color='#808080', hover=True, fg_color='#3b8ed0', height=10, border_color="dark", text_color="#1c1c1c", corner_radius=10, command=self.upload_static_html_file)
        self.static_upload_button.grid(row=0, column=1, padx=(0, 40), pady=(10, 10), sticky="e")

        self.textbox_static = customtkinter.CTkTextbox(self.static_frame, fg_color="white", text_color="black", width=700,height=340)
        self.textbox_static.grid(row=1, column=0, columnspan=2, padx=(40, 40), pady=(10, 10), sticky="nsew")
        self.textbox_static.insert("2.0", "Html Code goes here.../ Upload the html file")
        self.textbox_static.bind("<FocusIn>", self.clear_placeholder)
        self.textbox_static.bind("<KeyRelease>", lambda event: self.clear_entry_text_static())

        self.static_main_back_button = CTkButton(self.static_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command=self.main_back_button_func)
        self.static_main_back_button.grid(row=2, column=0, padx=(40, 10), pady=(10, 10), sticky="w")

        self.static_sub_button = CTkButton(self.static_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"), command=self.static_sub_button_func)
        self.static_sub_button.grid(row=2, column=1, padx=(0, 40), pady=(10, 10), sticky="e")


        # Second Frame(Subject and Attachments Frame)
        self.subject_attach_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="#CADCFC")
        self.subject_attach_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nsew")
        self.subject_attach_frame.grid_rowconfigure(6, weight=1)  # Adjusted to accommodate the buttons
        self.subject_attach_frame.grid_columnconfigure(2, weight=1)

        # Subject label and entry
        self.subject_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Subject", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), height=40)
        self.subject_label.grid(row=0, column=0, padx=(10, 5), pady=(5, 5), sticky="w")

        self.subject_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="Enter the subject")
        self.subject_entry.grid(row=0, column=1, columnspan=2, padx=(5, 10), pady=(5, 5), sticky="ew")

        # CC label and entry
        self.cc_label = customtkinter.CTkLabel(self.subject_attach_frame, text="CC", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), height=40)
        self.cc_label.grid(row=1, column=0, padx=(10, 5), pady=(5, 5), sticky="w")

        self.cc_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="Enter email addresses, separated by commas for multiple mail...")
        self.cc_entry.grid(row=1, column=1, columnspan=2, padx=(5, 10), pady=(5, 5), sticky="ew")

        # BCC label and entry
        self.bcc_label = customtkinter.CTkLabel(self.subject_attach_frame, text="BCC", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), height=40)
        self.bcc_label.grid(row=2, column=0, padx=(10, 5), pady=(5, 5), sticky="w")

        self.bcc_entry = customtkinter.CTkEntry(self.subject_attach_frame, placeholder_text="Enter email addresses, separated by commas for multiple mail...")
        self.bcc_entry.grid(row=2, column=1, columnspan=2, padx=(5, 10), pady=(5, 5), sticky="ew")

        # Static Attachments label and buttons
        self.attach_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Common Attachments", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), height=40)
        self.attach_label.grid(row=3, column=0, padx=(5, 5), pady=(5, 5), sticky="w")

        self.attachment_files = customtkinter.CTkButton(self.subject_attach_frame, text="Upload", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), hover_color='#808080', fg_color='#3b8ed0', text_color="#1c1c1c", corner_radius=10, command=self.static_attach_files_function)
        self.attachment_files.grid(row=3, column=1, padx=(5, 5), pady=(5, 5), sticky="w")

        self.attachment_preview_button = customtkinter.CTkButton(self.subject_attach_frame, text="Preview", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), hover_color='#808080', fg_color='#3b8ed0', text_color="#1c1c1c", corner_radius=10, command=self.attachment_preview_button_function)
        self.attachment_preview_button.grid(row=3, column=2, padx=(5, 10), pady=(5, 5), sticky="e")

        # Individual Attachments label and scrollable frame
        self.individual_attach_label = customtkinter.CTkLabel(self.subject_attach_frame, text="Individual Attachments", font=customtkinter.CTkFont(family="Times New Roman", size=14, weight="bold"), height=40)
        self.individual_attach_label.grid(row=4, column=0, padx=(10, 5), pady=(5, 5), sticky="w")

        self.dynamic_scroll_checkbox_frame = customtkinter.CTkScrollableFrame(self.subject_attach_frame, label_text="Select Below columns", fg_color="white")
        self.dynamic_scroll_checkbox_frame.grid(row=4, column=1, columnspan=2, padx=(5, 10), pady=(5, 5), sticky="ew")
        self.dynamic_scrollable_frame_checkbox = []
        self.subject_attach_frame.grid_forget()

        # Second Frame(Email Mapping Frame( For Dynamic Contents))
        self.list_frame = customtkinter.CTkFrame(self.second_frame, fg_color="#CADCFC")
        self.list_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nsew")
        self.list_frame.grid_forget()

        # Final frame and Email Loading Bar
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=50, fg_color="#CADCFC")
        self.third_frame.grid_columnconfigure(0, weight=20)
        self.third_frame.grid_forget()

        self.select_frame_by_name(name="home")

        self.frame_2_button.configure(state="disabled",fg_color= "transparent",text_color="black")

    # Navigation Frame (To create a config file)
    def template_button_func(self):
        try:
            # Create data for all sheets
            data1 = {"From_Mail_Id":[],"App_Password":[]}
            data2 = {"To_Mail_Id": []}
            data3 = {"Configuration Instructions:":["","**Excel Instruction:**", 
                                                    "First Sheet - (From_Mail):",
                                                    "    Column 1 (From_Mail_Id): Fill in the sender's email address.",
                                                    "    Column 2 (App_password) : Fill in the app password.",
                                                    "    Note: We can add multiple sender email address along with their correct app password.",
                                                    "",
                                                    "Second Sheet - (To_Mail):",
                                                    "    Column 1 (To_Mail_Id) : Fill in the recipient's email address (This field is mandatory).",
                                                    "    Note: Then add multiple columns according to your needs.",
                                                    "    Note: We can also give attachment path for individual receiver on separate column.",
                                                    "",
                                                    "",
                                                    "**Mail Body Instruction:**",
                                                    "    Note: The advanced email (text file/excel file) body contains dynamic content placeholders marked by {{word}} without spaces. Ensure all dynamic fields are correctly populated before sending. Eg.{{candidate_name}}"]}


            # Create DataFrames for all sheets
            df1 = pd.DataFrame(data1)
            df2 = pd.DataFrame(data2)
            df3 = pd.DataFrame(data3)


            date_time = dt.datetime.now().strftime("%m_%d_%Y_%H%M%S")

            # Define the path to save the Excel file in the Downloads directory
            downloads_path = str(Path.home() / "Downloads" / f"Email_Zapper_{date_time}.xlsx")

            # Write the DataFrames to two separate sheets in the Excel file
            with pd.ExcelWriter(downloads_path) as writer:
                df1.to_excel(writer, sheet_name=self.excel_sheet_name[0], index=False)
                df2.to_excel(writer, sheet_name=self.excel_sheet_name[1], index=False)
                df3.to_excel(writer, sheet_name=self.excel_sheet_name[2], index=False)
            messagebox.showinfo("Success", f"Created successfully\n File path:\n{downloads_path}")
        except Exception as e:
            messagebox.showwarning("Error", "Error while creating template excel")


    # Second Frame(Subject and Attachments Frame(attachment_preview_button))
    def attachment_preview_button_function(self):
        # Minimize the main window (self)
        self.iconify()  # Use self.withdraw() if you want to completely hide the window

        # Create the new window
        self.new_window = customtkinter.CTkToplevel(self)
        self.new_window.title("Attachment Files")
        self.new_window.geometry("850x450")
        self.new_window.resizable(False, False)  
        # Bring the new window to the front
        self.new_window.lift()
        self.new_window.focus_force()

        # Add a scrollable frame to the new window
        self.dynamic_scroll_frame = customtkinter.CTkScrollableFrame(
            self.new_window,
            width=800,
            height=400, fg_color = "white"
        )

        self.dynamic_scroll_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Configure grid weights to allocate more space to the label column
        self.dynamic_scroll_frame.grid_columnconfigure(0, weight=1)  # Label column
        self.dynamic_scroll_frame.grid_columnconfigure(1, weight=0) 

        self.populate_scroll_frame()

        # Restore the main window when the new window is closed
        self.new_window.protocol("WM_DELETE_WINDOW", self.on_new_window_close)


    # Second Frame(Subject and Attachments Frame(attachment_preview_button))
    def populate_scroll_frame(self):
        for widget in self.dynamic_scroll_frame.winfo_children():
            widget.destroy()  # Clear the frame before populating

        for idx, item in enumerate(self.attachment_file_path_list):
            # Create a label for each item
            label = customtkinter.CTkLabel(self.dynamic_scroll_frame, text=item,font=("Arial", 16),
                wraplength=750)
            label.grid(row=idx, column=0, padx=10, pady=5,sticky="w")

            # Create an edit button for each item
            edit_button = customtkinter.CTkButton(self.dynamic_scroll_frame, text="X",fg_color="red",width=60, 
                                        command=lambda idx=idx: self.delete_item(idx))
            edit_button.grid(row=idx, column=1, padx=50, pady=5,sticky="e")


    # Second Frame(Subject and Attachments Frame(attachment_preview_button))
    def delete_item(self, idx):
        # Delete the item from the list
        del self.attachment_file_path_list[idx]

        # Repopulate the scrollable frame with updated items
        self.populate_scroll_frame()


    def on_new_window_close(self):
        # Destroy the new window
        self.new_window.destroy()

        # Restore the main window (self)
        self.deiconify()


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
                self.excel_file_df_from_mail = pd.read_excel(file_path, sheet_name=self.excel_sheet_name[0])
                self.excel_file_df_to_mail = pd.read_excel(file_path,sheet_name=self.excel_sheet_name[1])
                self.excel_file_path = file_path
                excel_header = pd.read_excel(file_path,sheet_name=self.excel_sheet_name[1])
                self.excel_file_to_mail_header_list = excel_header.columns.tolist()
                self.seg_button_1.configure(state="normal")
                if self.excel_to_mail_header_changing_data[1] in self.excel_file_to_mail_header_list:
                    for index, row in self.excel_file_df_to_mail.iterrows():
                        if row.get(self.excel_to_mail_header_changing_data[1]):
                            if row[self.excel_to_mail_header_changing_data[1]] != "Completed":
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
        if name == self.html_state[0]:
            self.textbox_static.delete(1.0,"end")
            self.entry_static.delete(0, "end")
            self.dynamic_frame.configure(corner_radius=16, fg_color="#CADCFC",width=500,height=500)
            self.dynamic_frame.grid(row=1, column=0,  padx=(20, 20), pady=(10, 20), sticky="ew")
        else:
            self.dynamic_frame.grid_forget()
        if name == self.html_state[1]:
            self.textbox_dynamic.delete(1.0,"end")
            self.entry_dynamic.delete(0, "end")
            self.static_frame.configure(corner_radius=16, fg_color="#CADCFC",width=500,height=500)
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
                if os.path.exists(r"{}".format(dynamic_upload_value.strip('\"'))):

                    # Second Frame(Subject and Attachments Frame (Dynamic - Back and Submit Button))
                    self.attachment_back_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
                    self.attachment_back_button.grid(row=5, column=0, padx=(20, 10), pady=(5, 5), sticky="w")

                    self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                    self.attachment_sub_button.grid(row=5, column=2, padx=(10, 20), pady=(5, 5), sticky="e")
                    
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
                self.attachment_back_button.grid(row=5, column=0, padx=(20, 10), pady=(5, 5), sticky="w")

                self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                self.attachment_sub_button.grid(row=5, column=2, padx=(10, 20), pady=(5, 5), sticky="e")
                
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
                if os.path.exists(r"{}".format(static_upload_value.strip('\"'))):

                    # Subject and Attachments Frame (Static - Back and Submit Button)
                    self.attachment_back_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.attachment_back_function)
                    self.attachment_back_button.grid(row=5, column=0, padx=(10, 20), pady=(5, 5), sticky="w")

                    self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                    self.attachment_sub_button.grid(row=5, column=2, padx=(10, 80), pady=(5, 5), sticky="e")

                    self.static_preview_logo = Image.open(r"logo\eye.png")
                    self.static_preview_logo_button = CTkButton(self.subject_attach_frame,corner_radius=30, text="",image=CTkImage(self.static_preview_logo), fg_color="transparent", border_color="green",border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),width=10, height=10, command=self.static_preview_function)
                    self.static_preview_logo_button.grid(row=5, column=2, padx=(10, 20), pady=(5, 5), sticky="e")

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
                self.attachment_back_button.grid(row=5, column=0, padx=(10, 20), pady=(5, 5), sticky="w")

                self.attachment_sub_button = customtkinter.CTkButton(self.subject_attach_frame, corner_radius=30, text="Submit",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.attachment_sub_function)
                self.attachment_sub_button.grid(row=5, column=2, padx=(10, 80), pady=(5, 5), sticky="e")

                self.static_preview_logo = Image.open(r"logo\eye.png")
                self.static_preview_logo_button = CTkButton(self.subject_attach_frame,corner_radius=30, text="",image=CTkImage(self.static_preview_logo), fg_color="transparent", border_color="green",border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),width=10, height=10, command=self.static_preview_function)
                self.static_preview_logo_button.grid(row=5, column=2, padx=(10, 20), pady=(5, 5), sticky="e")
                
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
        self.email_cc = self.cc_entry.get().split(",")
        self.email_bcc = self.bcc_entry.get().split(",")
        if len(self.subject_entry.get()) >= 1 and self.subject_entry.get().strip() != "":                
            self.list_frame_show_call()
        else:
            messagebox.showwarning("Error", "Please fill the email subject")

    
    # Second Frame(Subject and Attachments Frame)
    def sub_attach_function(self):
        self.static_frame.grid_forget()
        self.dynamic_frame.grid_forget()
        # self.list_frame.grid_forget()
        self.subject_attach_frame.configure()
        self.subject_attach_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nsew")

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
        self.replacer_keyword_list = params_variable

        if params_variable and self.current_html_state == self.html_state[0]:
            if self.excel_file_to_mail_header_list:
                self.list_frame_show(params_variable, self.excel_file_to_mail_header_list)
            else:
                # print("Empty header list")
                pass
        else:
            self.dynamic_submit_button()
    

    # Second Frame(Email Mapping Frame( For Dynamic Contents))
    def list_frame_show(self,params_variable, params_name):
        self.static_frame.grid_forget()
        self.dynamic_frame.grid_forget()
        self.subject_attach_frame.grid_forget()
        self.list_frame.configure()
        self.list_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nsew")  # Set sticky to "nsew" for expansion
        self.list_frame.grid_rowconfigure(1, weight=1)  # Allow row 1 to expand
        self.list_frame.grid_columnconfigure(0, weight=1)  # Allow column 0 to expand
 
        # Scrollable Frame inside List Frame
        self.scrollable_frame = customtkinter.CTkScrollableFrame(self.list_frame, label_text="Email Mapping", fg_color="#CADCFC", height=350, width=700)
        self.scrollable_frame.grid(row=1, column=0, padx=(25, 10), pady=(20, 20), sticky="nsew")  # Set sticky to "nsew" for expansion

        self.scrollable_frame_switches = []
        replacer_list = params_variable
        header_list = params_name

        self.scrollable_frame.grid_columnconfigure(0, weight=1)  # First column (label) expands
        self.scrollable_frame.grid_columnconfigure(1, weight=2)  # Second column (entry) expands more
        self.scrollable_frame.grid_columnconfigure(2, weight=1)

        for index,context in enumerate(replacer_list):
            
            context_label = customtkinter.CTkLabel(master=self.scrollable_frame,text=context)
            context_label.grid(row=index,column=0, padx=10, pady=(0, 20))

            entry_dynamic = customtkinter.CTkEntry(self.scrollable_frame)
            entry_dynamic.grid(row=index, column=1, padx=(10,100), pady=(0, 20),sticky="e")

            context_list = customtkinter.CTkComboBox(self.scrollable_frame,values=header_list)
            context_list.grid(row=index,column=2, padx=10, pady=(0, 20), sticky="e")
            
            self.scrollable_frame_switches.append((context_label, entry_dynamic, context_list))

            # Bind the entry widget to the callback function
            entry_dynamic.bind("<KeyRelease>", lambda event, e=entry_dynamic, c=context_list: self.on_entry_change(e, c))

        self.list_frame.grid_rowconfigure(2, weight=0)
                                          
        # Second Frame(Email Mapping Frame(For Dynamic Contents))
        self.dynamic_back_button = CTkButton(self.list_frame,corner_radius=30, text="Back", fg_color="white", border_color="gray", border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.dynamic_back_button_function)
        self.dynamic_back_button.grid(row=2, column=0, padx=(10, 10), pady=(10, 10), sticky="w")

        self.submit_list_button = customtkinter.CTkButton(self.list_frame, corner_radius=30, text="Submit", fg_color="white", border_color="green", border_width=2, text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.dynamic_submit_button)
        self.submit_list_button.grid(row=2, column=0, padx=(10, 100), pady=(10, 10), sticky="e")

        self.preview_logo = Image.open(r"logo\eye.png")
        self.preview_logo_button = CTkButton(self.list_frame,corner_radius=30, text="",image=CTkImage(self.preview_logo), fg_color="transparent", border_color="green",border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),width=10, height=10, command=self.preview_button)
        self.preview_logo_button.grid(row=2, column=0, padx=(10, 10), pady=(10, 10), sticky="e")

        # create preview frame==>dummy frame created for future use
        self.preview_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="grey")
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(1, weight=1)

    
    # Second Frame(Email Mapping Frame( For Dynamic Contents(dynamic_back_button)))
    def dynamic_back_button_function(self):
        self.list_frame.grid_forget()
        self.subject_attach_frame.configure()
        self.subject_attach_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="nsew")


    # Second Frame(Email Mapping Frame(For Dynamic Contents(submit_list_button)))
    def dynamic_submit_button(self):
        self.body_params = self.get_entry_data()

        self.second_frame.grid_forget()

        # Final frame and Email Loading Bar
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#CADCFC",width=700,height=500)
        self.third_frame.grid(row=0, column=1, sticky="nsew")

        self.progressbar_empty = customtkinter.CTkLabel(master=self.third_frame,text="Click Start Button",height=20, width=500)
        self.progressbar_empty.grid(row=0, column=0, columnspan=2,padx=(150,0), pady=(200,0))
        
        self.progressbar_text = customtkinter.CTkLabel(master=self.third_frame,text=" ",height=20, width=500)
        self.progressbar_text.grid(row=0, column=0, columnspan=2,padx=(150,0), pady=(250,0))
        
        self.email_back_button = customtkinter.CTkButton(self.third_frame, corner_radius=30, text="Back",fg_color="white", border_color="gray", border_width=2,text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), command = self.email_back_button_func)
        self.email_back_button.grid(row=1, column=0, padx=(150,0), pady=(100,0))

        self.start_button = customtkinter.CTkButton(self.third_frame, corner_radius=30, text="Start",fg_color="white", border_color="green", border_width=2,text_color=("gray10", "gray90"), hover_color=("green", "green"),command=self.start_email)
        self.start_button.grid(row=1, column=1, padx=0, pady=(100,0))


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
            if self.replacer_keyword_list:
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
            excel_header = pd.read_excel(self.excel_file_path,sheet_name=self.excel_sheet_name[1])
            excelfile_to_mail_header_list = excel_header.columns.tolist()
            if self.excel_to_mail_header_changing_data[0] not in excelfile_to_mail_header_list:
                self.excel_file_df_to_mail[self.excel_to_mail_header_changing_data[0]] = ""
                if self.excel_to_mail_header_changing_data[1] not in excelfile_to_mail_header_list:
                    
                    self.excel_file_df_to_mail[self.excel_to_mail_header_changing_data[1]] = ""
            else:
                if self.excel_to_mail_header_changing_data[1] not in excelfile_to_mail_header_list:
                    self.excel_file_df_to_mail[self.excel_to_mail_header_changing_data[1]] = ""

            for index, row in self.excel_file_df_to_mail.iterrows():
                    if self.check_internet_connection(self.url, self.timeout):
                        if self.break_flag == 0:
                            try:
                                for i, email_data in self.excel_file_df_from_mail.iterrows():
                                    if email_data.iloc[0] not in error_mail_id:
                                        if self.excel_to_mail_header_changing_data[1] in excelfile_to_mail_header_list:
                                            if row[self.excel_to_mail_header_changing_data[1]] == "Completed":
                                                continue
                                        try:
                                            for path in self.individual_attachments_header:
                                                if os.path.exists(r"{}".format(row[path])):
                                                    self.attachment_file_path_list.append(r"{}".format(row[path]))
                                            recipient_email = row.iloc[0] # Assuming your Excel file has a column named 'Email'
                                            gmail.username =   email_data.iloc[0]# Notification mail sent to registered mail of customer
                                            gmail.password = email_data.iloc[1]
                                            gmail.send(subject = self.email_subject,
                                                        bcc = self.email_bcc,
                                                        cc = self.email_cc,
                                                        receivers = [recipient_email],
                                                        text = self.full_body_text_content,
                                                        html =self.html_full_content,
                                                        body_params=self.evaluate_body_params(row),
                                                        attachments=self.attachment_file_path_list)
                                            self.attachment_file_path_list = self.attachment_file_path_list[:self.static_attachment_file_count]
                                            self.excel_file_df_to_mail.loc[index,self.excel_to_mail_header_changing_data[0]] = email_data.iloc[0] 
                                            self.excel_file_df_to_mail.loc[index,self.excel_to_mail_header_changing_data[1]] = "Completed"
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
                            self.back_to_normal()
                            break
                    else:
                        messagebox.showwarning("Error", "No Internet Connection")
                        self.back_to_normal()
                        break
            else:
                messagebox.showinfo("Succcess", "Email Sent Successfully")
                self.back_to_normal()

            # Step 2: Create an ExcelWriter object
            with pd.ExcelWriter(self.excel_file_path, engine='openpyxl') as writer:
                # Write each DataFrame to a different sheet
                self.excel_file_df_from_mail.to_excel(writer, sheet_name=self.excel_sheet_name[0], index=False)
                self.excel_file_df_to_mail.to_excel(writer, sheet_name=self.excel_sheet_name[1], index=False)

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
            # self.html_state = ["Dynamic","Static"]
            self.html_state = ["Advanced","Normal"]
            self.current_html_state = None
            self.individual_attachments_header=[]
            self.dynamic_scrollable_frame_checkbox = []
            self.replacer_keyword_list = []
            self.email_subject = ""
            self.break_flag = 0
            self.total_email_data_count = 0
            self.completed_count = 0
            self.email_cc = None
            self.email_bcc = None

            # Deleting old values while second run(if present)
            self.entry_dynamic.delete(0,"end")
            self.textbox_dynamic.delete(1.0,"end")
            self.entry_static.delete(0,"end")
            self.textbox_static.delete(1.0,"end")
            self.subject_entry.delete(0,"end")
            self.cc_entry.delete(0,"end")
            self.bcc_entry.delete(0,"end")
        
    # Final frame and Email Loading Bar(start_button)
    def start_email(self):
        self.email_back_button.grid_forget()
        self.start_button.grid_forget()
        self.email_stop_button = customtkinter.CTkButton(self.third_frame, corner_radius=30, text="Stop", fg_color="white", border_color="red", border_width=2,text_color="black", hover_color="red",command=self.stop_back_button_func)
        self.email_stop_button.grid(row=1, column=1, padx=(0,75), pady=(100,0))
        threading.Thread(target=self.mail_processor).start()

    
    # To check internet connection
    def check_internet_connection(self, url, timeout):
        try:
            response = requests.get(url, timeout=timeout)
            # Check if the response status code is 200 (OK)
            if response.status_code == 200:
                return True
            else:
                return False
        except requests.ConnectionError:
            # If there's a ConnectionError, the connection failed
            return False
        except requests.Timeout:
            # If there's a Timeout, the request timed out
            return False
    

    # Second Frame(Subject and Attachments Frame(Individual Attachments))
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
        # self.entry_dynamic.delete(0,"end")
        # self.textbox_dynamic.delete(1.0,"end")
        # self.entry_static.delete(0,"end")
        # self.textbox_static.delete(1.0,"end")
        # self.subject_entry.delete(0,"end")


        # set button color for selected button
        self.home_button.configure(fg_color=("#CADCFC", "white") if name == "home" else "#00246B")
        self.frame_2_button.configure(fg_color=("#CADCFC", "white") if name == "frame_2" else "#00246B")

        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
            self.template_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=5, height=30, width=150,font=customtkinter.CTkFont(size=16,family="Times New Roman"), border_spacing=0, text="Download Template Here", text_color="#00246B",hover=True,hover_color="white", command=self.template_button_func, fg_color="#CADCFC")
            self.template_button.grid(row=7, column=0, pady=(300,50), sticky="s")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
            self.dynamic_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
            self.template_button.grid_forget()
        else:
            self.second_frame.grid_forget()
  
    # Final frame and Email Loading Bar(email_stop_button)    
    def stop_back_button_func(self):
        self.break_flag = 1

    def back_to_normal(self):
        self.break_flag = 0
        self.home_frame.grid_forget()
        self.second_frame.grid_forget()
        self.third_frame.grid_forget()
        self.subject_attach_frame.grid_forget()
        self.list_frame.grid_forget()

        self.select_frame_by_name(name="home")
        self.home_button.configure(state="normal")
        self.frame_2_button.configure(state="disabled")

if __name__ == "__main__":
    app = App()
    app.mainloop()


