import customtkinter as ctk
from tkinter import filedialog
import time
import asyncio
import threading
import dictionaryFrame as df
import functions as fns  # Abstracted Functionality

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window Setup ======================================================================================================

        self.title("Pricing Pack Quantification Tool")
        self.geometry("800x600")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Main frame
        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # TabView
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.mainTab = self.tabview.add("Main")
        self.settingsTab = self.tabview.add("Settings")
        self.creditsTab = self.tabview.add("Credits")
        
        # Settings variables
        self.num_rows_settings = 10
        
        
         # Create asyncio event loop
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        

            

        # END Window Setup ==================================================================================================


        #               _                                     _____      _     
        #   /\/\   __ _(_)_ __     /\/\   ___ _ __  _   _    /__   \__ _| |__  
        #  /    \ / _` | | '_ \   /    \ / _ \ '_ \| | | |     / /\/ _` | '_ \ 
        # / /\/\ \ (_| | | | | | / /\/\ \  __/ | | | |_| |    / / | (_| | |_) |
        # \/    \/\__,_|_|_| |_| \/    \/\___|_| |_|\__,_|    \/   \__,_|_.__/ 
        # =====================================================================



        # Main Menu Tab  ====================================================================================================
        
        self.mainTab.grid_columnconfigure(0, weight=1)



        # row 8
        
        self.frame1 = ctk.CTkFrame(self.mainTab)
        self.frame1.grid(row=8, column=0, padx=10, pady=0, sticky="ew")
        self.modeLabel = ctk.CTkLabel(self.frame1, text="Mode")
        self.modeLabel.pack(padx=10, pady=0)
        
        # row 9
        
        self.modeComboBox = ctk.CTkComboBox(self.mainTab, values=["Generate Quantifications Excel File from CSV"])
        self.modeComboBox.grid(row=9, column=0, padx=10, pady=0, sticky="ew")

        # row 10

        # Run Tool Button
        self.button = ctk.CTkButton(self.mainTab, text="Run Tool", command=self.run_tool_callback)
        self.button.grid(row=10, column=0, padx=10, pady=10, sticky="ew")

        # row 11
        
        # CTkProgressBar
        self.progressBar = ctk.CTkProgressBar(self.mainTab)
        self.progressBar.grid(row=11, column=0, padx=10, pady=10, sticky="ew")
        self.progressBar.set(0.0)

        
        
        # END Main Menu Tab =================================================================================================
        

        #  __      _   _   _                     _____      _     
        # / _\ ___| |_| |_(_)_ __   __ _ ___    /__   \__ _| |__  
        # \ \ / _ \ __| __| | '_ \ / _` / __|     / /\/ _` | '_ \ 
        # _\ \  __/ |_| |_| | | | | (_| \__ \    / / | (_| | |_) |
        # \__/\___|\__|\__|_|_| |_|\__, |___/    \/   \__,_|_.__/ 
        #                         |___/                          
        # ========================================================

        
        # Settings Tab ======================================================================================================
        
        self.settingsTab.grid_columnconfigure(0, weight=1)


        self.addRowButton = ctk.CTkButton(self.settingsTab, text="+", command=self.add_settings_row_callback)
        self.addRowButton.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.removeRowButton = ctk.CTkButton(self.settingsTab, text="-", command=self.remove_settings_row_callback, fg_color="#EE0000", hover_color="#660000")
        self.removeRowButton.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # CTkLabel in a CTkFrame
        self.frame = ctk.CTkFrame(self.settingsTab)
        self.frame.grid(row=8, column=0, padx=10, pady=0, sticky="ew")
        self.label = ctk.CTkLabel(self.frame, text="Settings")
        self.label.pack(padx=10, pady=0)
        
        
        # NOTE: We need to make a table (like Naviswork Selection Inspector) to represent name conversions, then carry them out using external functions (fns) 
        
        self.nameUpdate = df.DictionaryFrame(self.settingsTab, num_rows=self.num_rows_settings)
        self.nameUpdate.grid(row=9, column=0, padx=10, pady=0, sticky="ew")
        
        
        
        
        
        # Output Column Names
        
        # (E) ID
        # (I) GUID
        # (E) Category
        # (C) Location (or KGE_Location)
        # (E) Name
        # (E) Size
        # (E) Rod Length
        # (E) Length
        #     - Any other length values too (Length, length, Length 2, etc.)
        # (E) Unistrut Length
        # (+) Units (mm/No.)
        # (E) Angle
        #     - Any other angles too (Angle, angle, Angle 2, etc.)
        # (E) Service type
        # (C) Raceway Name
        # (C) Raceway Ref. Number 
        
        
        # END Settings Tab ==================================================================================================
        
        
        #    ___             _ _ _           _____      _     
        #   / __\ __ ___  __| (_) |_ ___    /__   \__ _| |__  
        #  / / | '__/ _ \/ _` | | __/ __|     / /\/ _` | '_ \ 
        # / /__| | |  __/ (_| | | |_\__ \    / / | (_| | |_) |
        # \____/_|  \___|\__,_|_|\__|___/    \/   \__,_|_.__/ 
        # ====================================================                                                 
        
        
        # Credits Tab =======================================================================================================
        
        
        self.creditsTab.grid_columnconfigure(0, weight=1)
        
        self.frame = ctk.CTkFrame(self.creditsTab)
        self.frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.creditsLabel = ctk.CTkLabel(self.frame, text="Made by Seamus Mullan @ Kirby Group Engineering, August 2024")
        self.creditsLabel.pack(padx=10, pady=10)
        
        # END Credits Tab ===================================================================================================
        

    #                                 ___      _ _ _                _        
    #   __ _ ___ _   _ _ __   ___    / __\__ _| | | |__   __ _  ___| | _____ 
    #  / _` / __| | | | '_ \ / __|  / /  / _` | | | '_ \ / _` |/ __| |/ / __|
    # | (_| \__ \ |_| | | | | (__  / /__| (_| | | | |_) | (_| | (__|   <\__ \
    #  \__,_|___/\__, |_| |_|\___| \____/\__,_|_|_|_.__/ \__,_|\___|_|\_\___/
    #            |___/                                                       
    # ========================================================================
        
    # Async Callbacks =======================================================================================================
        
    def run_tool_callback(self):
        asyncio.run_coroutine_threadsafe(self.run_tool_async(), self.loop)

    async def run_tool_async(self):
        fns.log("Running Tool")
        await self.update_progress_async(0.5)
        fns.log("Set value to 0.5", 'message')
        
        # Simulate some work
        await asyncio.sleep(2)
        
        await self.update_progress_async(0.0)
        fns.log("Set value to 0.0", 'message')


    async def update_progress_async(self, amount):
        self.after(0, self.update_progress_callback, amount)

    def update_progress_callback(self, amount):
        """Sets progress bar amount

        Args:
            amount (float): number between 0 - 1
        """
        self.progressBar.set(amount)


    def add_settings_row_callback(self):
        asyncio.run_coroutine_threadsafe(self.add_settings_row(), self.loop)
        
    async def add_settings_row(self):
        self.num_rows_settings += 1
        await self.nameUpdate._draw()
        fns.log(self.num_rows_settings, 'log')
        
    def remove_settings_row_callback(self):
        asyncio.run_coroutine_threadsafe(self.remove_settings_row(), self.loop)
        
    async def remove_settings_row(self):
        self.num_rows_settings -= 1
        await self.nameUpdate._draw()
        fns.log(self.num_rows_settings, 'log')



    def start_async_loop(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def on_closing(self):
        self.loop.call_soon_threadsafe(self.loop.stop)
        self.destroy()
    

    # END Async Callbacks ===================================================================================================


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Run the asyncio event loop in a separate thread
    threading.Thread(target=app.start_async_loop, daemon=True).start()
    
    # Run the Tkinter main loop
    app.mainloop()