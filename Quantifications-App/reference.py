import customtkinter as ctk
from tkinter import filedialog

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CustomTkinter Widget Showcase")
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
        self.tab1 = self.tabview.add("Widgets")
        self.tab2 = self.tabview.add("Text & Progress")

        # Tab 1: Widgets
        self.tab1.grid_columnconfigure(0, weight=1)

        # CTkButton
        self.button = ctk.CTkButton(self.tab1, text="CTkButton", command=self.button_callback)
        self.button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # CTkCheckBox
        self.checkbox = ctk.CTkCheckBox(self.tab1, text="CTkCheckBox")
        self.checkbox.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        # CTkComboBox
        self.combobox = ctk.CTkComboBox(self.tab1, values=["Option 1", "Option 2", "Option 3"])
        self.combobox.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        # CTkEntry
        self.entry = ctk.CTkEntry(self.tab1, placeholder_text="CTkEntry")
        self.entry.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        # CTkOptionMenu
        self.optionmenu = ctk.CTkOptionMenu(self.tab1, values=["OptionMenu 1", "OptionMenu 2", "OptionMenu 3"])
        self.optionmenu.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        # CTkRadioButton
        self.radio_var = ctk.IntVar(value=0)
        self.radio1 = ctk.CTkRadioButton(self.tab1, text="RadioButton 1", variable=self.radio_var, value=1)
        self.radio2 = ctk.CTkRadioButton(self.tab1, text="RadioButton 2", variable=self.radio_var, value=2)
        self.radio1.grid(row=5, column=0, padx=10, pady=(10,0), sticky="w")
        self.radio2.grid(row=6, column=0, padx=10, pady=(0,10), sticky="w")

        # CTkSegmentedButton
        self.segemented_button = ctk.CTkSegmentedButton(self.tab1, values=["Segment 1", "Segment 2"])
        self.segemented_button.grid(row=7, column=0, padx=10, pady=10, sticky="ew")

        # CTkSlider
        self.slider = ctk.CTkSlider(self.tab1, from_=0, to=100, number_of_steps=10)
        self.slider.grid(row=8, column=0, padx=10, pady=10, sticky="ew")

        # CTkSwitch
        self.switch = ctk.CTkSwitch(self.tab1, text="CTkSwitch")
        self.switch.grid(row=9, column=0, padx=10, pady=10, sticky="w")

        # Tab 2: Text & Progress
        self.tab2.grid_columnconfigure(0, weight=1)

        # CTkTextbox
        self.textbox = ctk.CTkTextbox(self.tab2, height=100)
        self.textbox.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.textbox.insert("0.0", "This is a CTkTextbox\n\nIt can contain multiple lines of text.")

        # CTkProgressBar
        self.progressbar = ctk.CTkProgressBar(self.tab2)
        self.progressbar.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.progressbar.set(0.5)

        # CTkLabel in a CTkFrame
        self.frame = ctk.CTkFrame(self.tab2)
        self.frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        self.label = ctk.CTkLabel(self.frame, text="This is a CTkLabel inside a CTkFrame")
        self.label.pack(padx=10, pady=10)

        # CTkScrollbar (Note: CustomTkinter doesn't have a separate scrollbar widget, 
        # it's integrated into scrollable widgets like CTkScrollableFrame)

    def button_callback(self):
        print("Button clicked!")

if __name__ == "__main__":
    app = App()
    app.mainloop()