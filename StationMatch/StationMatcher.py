from tkinterdnd2 import TkinterDnD, DND_FILES
from pathHelper import ensure_user_csv
from PIL import Image, ImageTk
from cProfile import label
from match import matcher
from pathlib import Path
import customtkinter
import subprocess
import tkinter
import sys
import os

class GUI(customtkinter.CTk, TkinterDnD.DnDWrapper):
    def start(self) -> None:
        self.mainloop()

    def onClose(self):
        self.quit()     # close the GUI
        self.destroy()  # destroy the GUI
    
    def clean_path(self, raw_path: str) -> str:
        if raw_path.startswith("{") and raw_path.endswith("}"):
            return raw_path[1:-1]
        return raw_path

    def drop(self, event):
        try:
            filename = self.clean_path(event.data)
            self.label.configure(text=filename)
            self.info.configure(state='normal')
            self.info.delete("0.0", "end")
            self.info.insert("0.0", f"Processing file:\n{filename}")
            self.process_file(filename)
            self.info.configure(state='disabled', text_color="#22C55E")
        except Exception as e:
            self.info.configure(state='normal')
            self.info.delete("0.0", "end")
            self.info.insert("0.0", f"Error processing file:\n{str(e)}")
            self.info.configure(state='disabled', text_color="#EF4444")

    def process_file(self, filepath):
        matcher(filepath)
    
    def open_downloads(self):
        downloads_path = str(Path.home() / "Downloads" / "StationMatcher")
        if os.path.exists(downloads_path):
            subprocess.Popen(["explorer", downloads_path])
        else:
            print(f"Path does not exist: {downloads_path}")

    def edit_station_map(self):
        path = ensure_user_csv()
        try:
            # If it still doesn't exist, create a file now.
            if not os.path.exists(path):
                Path(path).parent.mkdir(parents=True, exist_ok=True)
                Path(path).write_text("Code,Name\n", encoding="utf-8")

            self.info.configure(state='normal')
            self.info.delete("0.0", "end")
            self.info.insert("0.0", f"Station map is at:\n{path}\n\n"
                                    "Tip: Edit and Save. Then run your file.")
            self.info.configure(state='disabled', text_color="#22C55E")

            os.startfile(path)  # opens in Excel/Notepad with the default app
        except Exception as e:
            self.info.configure(state='normal')
            self.info.delete("0.0", "end")
            self.info.insert("0.0", f"Error opening station map:\n{e}")
            self.info.configure(state='disabled', text_color="#EF4444")
    
    def resource_path(self, relative_path):
        base_path = getattr(sys, '_MEIPASS', Path(__file__).parent)
        return str(Path(base_path) / relative_path)

    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        customtkinter.set_appearance_mode('system')
        customtkinter.set_default_color_theme('dark-blue')

        self.protocol("WM_DELETE_WINDOW", self.onClose)
        self.title("Station Matcher")
        self.iconbitmap(default=self.resource_path("Resources/TCLogo.ico"))
        self.geometry("800x400")

        # Grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        # Header
        self.header = customtkinter.CTkLabel(self, text="Station Matcher", font=("Arial", 24, "bold"))
        self.header.grid(row=0, column=0, columnspan=2, pady=(10, 0))

        # Left panel (buttons)
        self.ButtonFrame = customtkinter.CTkFrame(self)
        self.ButtonFrame.grid(row=1, column=0, sticky="ns", padx=10, pady=10)

        self.openFolderButton = customtkinter.CTkButton(
            self.ButtonFrame, 
            text="Open Downloads Folder", 
            fg_color="#3B82F6",  # Primary accent
            hover_color="#1D4ED8",  # Darker blue on hover
            text_color="#FFFFFF",
            command=self.open_downloads
        )
        self.openFolderButton.pack(pady=20, fill="x")

        self.editMapButton = customtkinter.CTkButton(
            self.ButtonFrame, 
            text="Edit Station Map", 
            fg_color="#3B82F6",  # Primary accent
            hover_color="#1D4ED8",  # Darker blue on hover
            text_color="#FFFFFF",
            command=self.edit_station_map
        )
        self.editMapButton.pack(pady=20, fill="x")

        # Right panel (drag & drop + info)
        self.RightPanel = customtkinter.CTkFrame(self)
        self.RightPanel.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        self.RightPanel.grid_rowconfigure(1, weight=1)

        self.label = customtkinter.CTkLabel(
            self.RightPanel, 
            text="➕ Drag & Drop Excel File Here", 
            corner_radius=10, 
            fg_color="#2563EB",  # Drop zone background
            text_color="#FFFFFF",   # Drop zone text color
            font=("Arial", 16), 
            height=100)
        self.label.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind("<<Drop>>", self.drop)

        self.info = customtkinter.CTkTextbox(self.RightPanel)
        self.info.configure(state='disabled', wrap='word', font=('Arial', 14))
        self.info.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        # Footer
        self.footer = customtkinter.CTkLabel(self, text="© Travel Counsellors", font=("Arial", 10))
        self.footer.grid(row=2, column=0, columnspan=2, pady=(0, 10))

        self.start()

    
if __name__ == "__main__":
    app = GUI()
