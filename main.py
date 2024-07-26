import customtkinter
from customtkinter import CTkFrame, CTkLabel, CTkEntry, CTkButton, CTkSegmentedButton, CTkScrollableFrame, CTkImage, CTkToplevel,CTkFont
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from redmail import gmail


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
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
                                                       anchor="w")
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
        self.url_frame_button = customtkinter.CTkLabel(self.home_frame, text="Expecting\nSheet1 -> From Mail Credential \nSheet2 -> To Mail Credentials",font=CTkFont(family="times",size=15,weight="bold"),fg_color='#3b8ed0',height=40,text_color="#1c1c1c",corner_radius=10)
        self.url_frame_button.grid(row=4, column=0, padx=20, pady=40)

        # create second frame
        self.second_frame = customtkinter.CTkFrame(self, corner_radius=50, fg_color="blue")
        self.tabview = customtkinter.CTkTabview(self.second_frame, width=450,height=400)
        self.tabview.grid(row=0, column=1, padx=(170, 50), pady=(80,50), sticky="nsew")
        self.tabview.add("Email")
        self.tabview.add("Skype")
        self.tabview.tab("Email").grid_columnconfigure(2, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Skype").grid_columnconfigure(2, weight=1)


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
            messagebox.showinfo("File Selected", f"Successfully uploaded: {file_path}")
            try:
                excel_file_df_to_mail = pd.read_excel(file_path,sheet_name="Sheet2")
                excel_file_df_from_mail = pd.read_excel(file_path, sheet_name="Sheet1")
            except Exception as e:
                messagebox.showwarning("Error","Excel file error")
        else:
           messagebox.showwarning("No File", "Please select an Excel file.")

def mail_preprocesser(excel_file_path):
        error_mail_id = []
        # try:
            # excel_file_df_to_mail = pd.read_excel(excel_file_path,sheet_name="Sheet2")
            # excel_file_df_from_mail = pd.read_excel(excel_file_path, sheet_name="Sheet1")
            # print(excel_file_df_to_mail)
        for index, row in excel_file_df_to_mail.iterrows():
                # print(f"To mail {index+1},{row[0]}, {row[1]}")
                try:
                    for i, email_data in excel_file_df_from_mail.iterrows():
                        if email_data[0] not in error_mail_id:
                            try:
                                print(f"From mail {i+1},{email_data[0]}, {email_data[1]}")
                                recipient_email = row[0] # Assuming your Excel file has a column named 'Email'
                                gmail.username =   email_data[0]# Notification mail sent to registered mail of customer
                                gmail.password = email_data[1]
                                gmail.send(subject = "Checking subject4",
                                                    receivers = [recipient_email],
                                                        ##################### Content for email ########################### 
                                                        html ="""
<html>
<head>
</head>
<body>
Checking
</body>
</html>
                            """,
                                body_params={
                                },)
                                break
                            except Exception as e:
                                error_mail_id.append(email_data[0])
                                print(f"{email_data[0]} added to error list")
                except Exception as e:
                    print(e)
        # except Exception as e:
        #     messagebox.showwarning("Error","Excel file error")


if __name__ == "__main__":
    app = App()
    app.mainloop()