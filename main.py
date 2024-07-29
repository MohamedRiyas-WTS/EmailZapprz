import customtkinter
from customtkinter import CTkFrame, CTkLabel, CTkEntry, CTkButton, CTkSegmentedButton, CTkScrollableFrame, CTkImage, CTkToplevel,CTkFont
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from redmail import gmail
import customtkinter
import re



class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.excel_file_df_from_mail = None
        self.excel_file_df_to_mail = None
        self.html_full_content = None
        self.body_params = None
        self.excel_file_to_mail_header_list = []
        self.scrollable_frame_switches = []
        self.title("EmailZapprz")
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

        self.static_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.static_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")

        self.static_frame.grid_forget()
        
        self.textbox_dynamic = customtkinter.CTkTextbox(self.dynamic_frame, width=700,fg_color="grey",text_color="white", height=340)
        self.textbox_dynamic.grid(row=0, column=1, padx=(40, 0), pady=(10,140))
        self.textbox_dynamic.insert("2.0","Html Code goes here.../ Upload the html file")
        self.textbox_dynamic.bind("<FocusIn>",self.clear_placeholder)
        
        self.textbox_static = customtkinter.CTkTextbox(self.static_frame, width=700,fg_color="grey",text_color="white", height=340)
        self.textbox_static.grid(row=0, column=1, padx=(40, 0), pady=(10,140))
        self.textbox_static.insert("2.0","Html Code goes here.../ Upload the html file")
        self.textbox_static.bind("<FocusIn>",self.clear_placeholder)
        # self.tabview = customtkinter.CTkTabview(self.second_frame, width=700,height=450,corner_radius=50)
        # self.tabview.grid(row=0, column=1, padx=(65, 0), pady=(50,100), sticky="nsew")
        # self.tabview.add("Static")
        # self.tabview.add("Dynamic")
        # self.tabview.tab("Static").grid_columnconfigure(2, weight=1)  # configure grid of individual tabs
        # self.tabview.tab("Dynamic").grid_columnconfigure(2, weight=1)
        self.dynamic_upload_button=CTkButton(self.dynamic_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10,command=self.upload_html_file)
        self.dynamic_upload_button.grid(row=0, column=1,columnspan=2, padx=(600,0), pady=(279, 10)) 
        self.dynamic_sub_button=CTkButton(self.dynamic_frame,text="Submit",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10, command=self.upload_custom_file)
        self.dynamic_sub_button.grid(row=0, column=1,columnspan=2, padx=(80, 10), pady=(400,10))
        self.entry_dynamic = customtkinter.CTkEntry(self.dynamic_frame, placeholder_text="Upload....")
        self.entry_dynamic.grid(row=0, column=0, columnspan=2, padx=(40, 150), pady=(365, 95), sticky="nsew")
        self.static_upload_button=CTkButton(self.static_frame,text="Upload",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10)
        self.static_upload_button.grid(row=0, column=1,columnspan=2, padx=(600,0), pady=(279, 10)) 
        self.static_sub_button=CTkButton(self.static_frame,text="Submit",font=CTkFont(family="times",size=20,weight="bold"),hover_color='#808080',hover=True,fg_color='#3b8ed0',height=10,border_color="dark",text_color="#1c1c1c",corner_radius=10)
        self.static_sub_button.grid(row=0, column=1,columnspan=2, padx=(80, 10), pady=(400,10))
        self.entry_static = customtkinter.CTkEntry(self.static_frame, placeholder_text="Upload....")
        self.entry_static.grid(row=0, column=0, columnspan=2, padx=(40, 150), pady=(365, 95), sticky="nsew")
        self.list_frame = customtkinter.CTkFrame(self.second_frame, corner_radius=16, fg_color="white",width=500,height=500)
        self.list_frame.grid(row=1, column=0, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.list_frame.grid_forget()

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

    def list_frame_show(self,params_variable, params_name):
        self.static_frame.grid_forget()
        self.dynamic_frame.grid_forget()
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
                                                         fg_color="white", border_color="green", border_width=2, text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),command=self.dynamic_submit_button)
        self.submit_list_button.grid(row=1, column=1, columnspan=2, rowspan=2, pady=(350, 10))

    def dynamic_submit_button(self):
        self.body_params = self.get_entry_data()
        
        
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
        file_path = filedialog.askopenfilename(
        filetypes=[("Html files", "*.html *.htm *.txt")],
        title="Select an Excel file"
        )
        if file_path:
                messagebox.showinfo("File Selected", f"Successfully uploaded: {file_path}")
                if re.search(".txt$",file_path):
                    with open(file_path,"r") as html_content:
                        file_content_str= html_content.read()   
                else:
                    with open(file_path,"rb") as html_content:
                        file_content = html_content.read()
                    file_content_str = file_content.decode('utf-8')
                self.html_full_content = file_content_str
                params_variable= re.findall("\{\{(\w+)\}\}",str(file_content_str))
                if params_variable:
                    if self.excel_file_to_mail_header_list:
                        self.list_frame_show(params_variable, self.excel_file_to_mail_header_list)
                    else:
                        print("Empty header list")
                else:
                    print("Empty params variable list")
            # except Exception as e:
            #     messagebox.showwarning("Error", "Html/Text file Error")
        else:
           messagebox.showwarning("No File", "Please select an Html/Text file.")
        
    def upload_custom_file(self):
        input_value = self.textbox_dynamic.get("1.0", "end")  # Get text from textbox
        self.html_full_content = input_value
        if len(input_value) > 1 and input_value.strip() != "":
            params_variable = re.findall("\{\{(\w+)\}\}",input_value)
            if params_variable:
                if self.excel_file_to_mail_header_list:
                    self.list_frame_show(params_variable, self.excel_file_to_mail_header_list)
                else:
                    print("Empty header list")
            else:
                print("Empty params variable list")
        else:
            print("Empty")

    def evaluate_body_params(self, row):
        body_params = {}
        for key, value in self.body_params.items():
            if value.startswith("row['") and value.endswith("']"):
                col_name = value[5:-2]
                body_params[key] = row[col_name]
            else:
                body_params[key] = value
        return body_params

    def mail_preprocesser(self):
            error_mail_id = []
            for index, row in self.excel_file_df_to_mail.iterrows():
                    # print(f"To mail {index+1},{row[0]}, {row[1]}")
                    try:
                        for i, email_data in self.excel_file_df_from_mail.iterrows():
                            if email_data[0] not in error_mail_id:
                                try:
                                    print(index, row[0])
                                    recipient_email = row[0] # Assuming your Excel file has a column named 'Email'
                                    gmail.username =   email_data[0]# Notification mail sent to registered mail of customer
                                    gmail.password = email_data[1]
                                    gmail.send(subject = "Checking subject4",
                                                receivers = [recipient_email],
                                                html =self.html_full_content,
                                    body_params=self.evaluate_body_params(row),)
                                    break
                                except Exception as e:
                                    error_mail_id.append(email_data[0])
                                    print(f"{email_data[0]} added to error list")
                    except Exception as e:
                        print(e)
           
    def start_email(self):
        # Need to set the try catch methods
        self.mail_preprocesser()

if __name__ == "__main__":
    app = App()
    app.mainloop()


