import tkinter as tk
from tkinter import messagebox, filedialog
import pywinstyles
import os
import re  # For sanitizing file names
import json
import shutil  # Import shutil to handle file operations
from PIL import Image, ImageTk, UnidentifiedImageError  # For image display
import sys
from tkinter import ttk
import traceback  # To help with error logging

def resource_path(relative_path):
    """ Get the absolute path to resource, works for dev and for PyInstaller bundle """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller sets _MEIPASS to the temp folder where it unpacks files (read-only)
        return os.path.join(sys._MEIPASS, relative_path)
    else:
        return os.path.join(os.path.abspath("."), relative_path)

debug_messages = False  # Toggle to enable/disable debug messages

# Determine a user-writable base directory for writable files.
if hasattr(sys, '_MEIPASS'):
    # Running from PyInstaller bundle
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running in a normal Python environment
    BASE_DIR = os.path.abspath(".")

BANNER_FILE = resource_path("banner.png")
INVENTORY_FILE = os.path.join(BASE_DIR, "inventory.json")
EXPORT_FILE = os.path.join(BASE_DIR, "vinyl_collection.txt")
IMAGES_DIR = os.path.join(BASE_DIR, "images")

def handle_error(message, exception=None):
    """Display a messagebox with the error and optionally log details to a file."""
    messagebox.showerror("Error", message)
    if debug_messages:
        print(f"[ERROR] {message}")
        if exception:
            print(traceback.format_exc())
            # Additionally, log errors to a file for debugging
            with open(os.path.join(BASE_DIR, "error_log.txt"), "a", encoding="utf-8") as log_file:
                log_file.write(f"ERROR: {message}\n")
                if exception:
                    log_file.write(traceback.format_exc() + "\n")

class VinylTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vinyl Tracker")
        self.root.geometry("1024x700")
        self.root.resizable(False, False)
        
        # Center the main window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 1024) // 2
        y = ((screen_height - 768) // 2) - 50
        self.root.geometry(f"1024x700+{x}+{y}")
        
        # Change title bar color
        pywinstyles.change_header_color(self.root, color="#7F6E5A")
        
        self.data = {"bands": {}}
        
        self.selected_band = None
        self.selected_album = None
        
        self.load_data()
        self.setup_gui()
        self.refresh_bands()
        
        # Call self.save_on_exit() when application closes
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def center_window(self, window, width, height):
        """Center a window on the screen or relative to the root window."""
        self.root.update_idletasks()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        x = parent_x + (parent_width // 2) - (width // 2)
        y = parent_y + (parent_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def load_data(self):
        """Load data from INVENTORY_FILE or create a new structure if not found."""
        try:
            with open(INVENTORY_FILE, "r", encoding="utf-8") as f:
                try:
                    self.data = json.load(f)
                except json.JSONDecodeError as e:
                    handle_error(f"Failed to parse '{INVENTORY_FILE}'. The file may be corrupted. Starting with empty data.", e)
                    self.data = {"bands": {}}
        except FileNotFoundError:
            if debug_messages:
                print(f"[DEBUG] '{INVENTORY_FILE}' not found. Initialized empty data structure.")
            self.data = {"bands": {}}
        except Exception as e:
            handle_error(f"Failed to load data from '{INVENTORY_FILE}'.", e)
            self.data = {"bands": {}}

        # Normalize image fields
        for band, band_data in self.data.get("bands", {}).items():
            for album, album_data in band_data.get("albums", {}).items():
                if album_data.get("image") == "":
                    self.data["bands"][band]["albums"][album]["image"] = None

        self.save_data()

    def save_data(self):
        """Write the entire data dictionary back to the inventory file."""
        try:
            with open(INVENTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4)
            if debug_messages:
                print(f"[DEBUG] Data saved successfully to '{INVENTORY_FILE}'.")
        except IOError as e:
            handle_error("Failed to save data. Check file permissions and disk space.", e)
        except Exception as e:
            handle_error("An unexpected error occurred while saving data.", e)

    def on_closing(self):
        # Save data before exit.
        self.save_data()
        self.root.destroy()

    def setup_gui(self):
        # Colors and Styles
        bg_color = "#EEE9E5"
        self.root.config(bg=bg_color)

        # Banner frame
        banner_frame = tk.Frame(self.root, bg=bg_color, height=120)
        banner_frame.pack(side=tk.TOP, fill=tk.X)
        banner_frame.pack_propagate(False)

        if os.path.exists(BANNER_FILE):
            try:
                banner_image = Image.open(BANNER_FILE)
                resample_filter = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
                banner_image = banner_image.resize((1024, 120), resample_filter)
                self.banner_photo = ImageTk.PhotoImage(banner_image)
                banner_label = tk.Label(banner_frame, image=self.banner_photo)
                banner_label.pack(fill=tk.BOTH, expand=True)
            except Exception as e:
                handle_error("Failed to load banner image.", e)
                banner_frame.config(bg=bg_color)
        else:
            banner_frame.config(bg=bg_color)

        # Header frame
        header_frame = tk.Frame(self.root, bg=bg_color, height=50)
        header_frame.pack(side=tk.TOP, fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(header_frame, text="Vinyl Tracker", font=("Helvetica", 16, "bold"), bg=bg_color, fg="#4B3B2A")
        title_label.pack(side=tk.LEFT, padx=20, pady=10)

        import_button = tk.Button(header_frame, text="Import Collection", command=self.import_collection)
        import_button.pack(side=tk.RIGHT, padx=25, pady=10)

        export_button = tk.Button(header_frame, text="Export Collection", command=self.export_collection)
        export_button.pack(side=tk.RIGHT, padx=0, pady=10)

        separator_frame = tk.Frame(self.root, bg=bg_color)
        separator_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        separator = ttk.Separator(separator_frame, orient='horizontal')
        separator.pack(fill=tk.X, padx=20)

        main_frame = tk.Frame(self.root, bg=bg_color)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        main_frame.grid_columnconfigure(0, weight=1, uniform="equal")
        main_frame.grid_columnconfigure(1, weight=1, uniform="equal")
        main_frame.grid_columnconfigure(2, weight=1, uniform="equal")

        self.left_frame = tk.Frame(main_frame, bg=bg_color)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=(50, 10), pady=20)

        self.middle_frame = tk.Frame(main_frame, bg=bg_color)
        self.middle_frame.grid(row=0, column=1, sticky="nsew", padx=40, pady=20)

        self.right_frame = tk.Frame(main_frame, bg=bg_color)
        self.right_frame.grid(row=0, column=2, sticky="nsew", padx=(10, 50), pady=20)

        # Left column: Band List
        band_label = tk.Label(self.left_frame, text="Band Name", bg=bg_color, font=("Helvetica", 12, "bold"))
        band_label.pack(anchor="w")

        self.band_listbox = tk.Listbox(self.left_frame, width=30, height=25, exportselection=False)
        self.band_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.band_listbox.bind("<<ListboxSelect>>", self.on_band_select)

        left_button_frame = tk.Frame(self.left_frame, bg=bg_color)
        left_button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

        left_buttons_subframe = tk.Frame(left_button_frame, bg=bg_color)
        left_buttons_subframe.pack(anchor="center")

        self.add_band_button = tk.Button(left_buttons_subframe, text="Add Band", command=self.add_band)
        self.add_band_button.pack(side=tk.LEFT, padx=5)

        self.delete_band_button = tk.Button(left_buttons_subframe, text="Delete Band", command=self.delete_band, state=tk.DISABLED)
        self.delete_band_button.pack(side=tk.LEFT, padx=5)

        # Middle column: Album List
        album_label = tk.Label(self.middle_frame, text="Album Name", bg=bg_color, font=("Helvetica", 12, "bold"))
        album_label.pack(anchor="w")

        self.album_listbox = tk.Listbox(self.middle_frame, width=30, height=25, exportselection=False)
        self.album_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.album_listbox.bind("<<ListboxSelect>>", self.on_album_select)

        middle_button_frame = tk.Frame(self.middle_frame, bg=bg_color)
        middle_button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

        middle_buttons_subframe = tk.Frame(middle_button_frame, bg=bg_color)
        middle_buttons_subframe.pack(anchor="center")

        self.add_album_button = tk.Button(middle_buttons_subframe, text="Add Album", command=self.add_album, state=tk.DISABLED)
        self.add_album_button.pack(side=tk.LEFT, padx=5)

        self.delete_album_button = tk.Button(middle_buttons_subframe, text="Delete Album", command=self.delete_album, state=tk.DISABLED)
        self.delete_album_button.pack(side=tk.LEFT, padx=5)

        # Right column: Album Image
        self.image_label = tk.Label(self.right_frame, text="No album selected", bg=bg_color, font=("Helvetica", 10))
        self.image_label.pack(anchor="center", pady=30)

        self.image_canvas = tk.Canvas(self.right_frame, width=300, height=300, bg="#FFFFFF")
        self.image_canvas.pack(pady=0)

        self.right_frame.pack_propagate(False)

        spacer = tk.Frame(self.right_frame, height=30, bg=bg_color)
        spacer.pack(side=tk.TOP, fill=tk.X)

        right_button_frame = tk.Frame(self.right_frame, bg=bg_color)
        right_button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 5))

        right_buttons_subframe = tk.Frame(right_button_frame, bg=bg_color)
        right_buttons_subframe.pack(anchor="center")

        self.associate_image_button = tk.Button(right_buttons_subframe, text="Associate Image", command=self.associate_image, state=tk.DISABLED)
        self.associate_image_button.pack(side=tk.LEFT, padx=5)

        self.remove_image_button = tk.Button(right_buttons_subframe, text="Remove Image", command=self.remove_image, state=tk.DISABLED)
        self.remove_image_button.pack(side=tk.LEFT, padx=5)

        if not os.path.exists(IMAGES_DIR):
            os.makedirs(IMAGES_DIR)

    def refresh_bands(self):
        self.band_listbox.delete(0, tk.END)
        band_names = sorted(self.data["bands"].keys())
        for name in band_names:
            self.band_listbox.insert(tk.END, name)
        self.selected_band = None
        self.selected_album = None
        self.delete_band_button.config(state=tk.DISABLED)
        self.add_album_button.config(state=tk.DISABLED)
        self.delete_album_button.config(state=tk.DISABLED)
        self.associate_image_button.config(state=tk.DISABLED)
        self.remove_image_button.config(state=tk.DISABLED)
        self.album_listbox.delete(0, tk.END)
        self.image_label.config(text="No album selected")
        self.image_canvas.delete("all")

    def refresh_albums(self, band_name):
        self.album_listbox.delete(0, tk.END)
        if band_name and band_name in self.data["bands"]:
            albums = sorted(self.data["bands"][band_name]["albums"].keys())
            for album_name in albums:
                self.album_listbox.insert(tk.END, album_name)
        self.selected_album = None
        self.delete_album_button.config(state=tk.DISABLED)
        self.associate_image_button.config(state=tk.DISABLED)
        self.remove_image_button.config(state=tk.DISABLED)
        self.image_label.config(text="No album selected")
        self.image_canvas.delete("all")

    def on_band_select(self, event):
        if not self.band_listbox.curselection():
            return
        selection = self.band_listbox.curselection()
        if selection:
            band_name = self.band_listbox.get(selection[0])
            self.selected_band = band_name
            self.delete_band_button.config(state=tk.NORMAL)
            self.add_album_button.config(state=tk.NORMAL)
            self.refresh_albums(band_name)
        else:
            self.selected_band = None
            self.delete_band_button.config(state=tk.DISABLED)
            self.add_album_button.config(state=tk.DISABLED)
            self.refresh_albums(None)

    def on_album_select(self, event):
        if not self.selected_band:
            return
        try:
            selection = self.album_listbox.curselection()
            if selection:
                album_name = self.album_listbox.get(selection[0])
                if self.selected_band not in self.data["bands"]:
                    raise KeyError(f"Band '{self.selected_band}' not found in data.")
                
                self.selected_album = album_name
                self.delete_album_button.config(state=tk.NORMAL)
                self.associate_image_button.config(state=tk.NORMAL)

                album_data = self.data["bands"][self.selected_band]["albums"].get(album_name, {})
                image_rel_path = album_data.get("image")

                if image_rel_path:
                    image_path = os.path.join(BASE_DIR, image_rel_path)
                    if os.path.exists(image_path):
                        self.show_image(image_path)
                        self.remove_image_button.config(state=tk.NORMAL)
                    else:
                        self.image_label.config(text=f"No image for '{self.selected_album}'")
                        self.image_canvas.delete("all")
                        self.remove_image_button.config(state=tk.DISABLED)
                else:
                    self.image_label.config(text=f"No image for '{self.selected_album}'")
                    self.image_canvas.delete("all")
                    self.remove_image_button.config(state=tk.DISABLED)
            else:
                self.selected_album = None
                self.delete_album_button.config(state=tk.DISABLED)
                self.associate_image_button.config(state=tk.DISABLED)
                self.remove_image_button.config(state=tk.DISABLED)
                self.image_label.config(text="No album selected")
                self.image_canvas.delete("all")
        except KeyError as e:
            handle_error(f"Error: {e}", e)
            self.selected_album = None
            self.delete_album_button.config(state=tk.DISABLED)
            self.associate_image_button.config(state=tk.DISABLED)
            self.remove_image_button.config(state=tk.DISABLED)
            self.image_label.config(text="No album selected")
            self.image_canvas.delete("all")
        except Exception as e:
            handle_error("An error occurred while selecting the album.", e)
            self.selected_album = None
            self.delete_album_button.config(state=tk.DISABLED)
            self.associate_image_button.config(state=tk.DISABLED)
            self.remove_image_button.config(state=tk.DISABLED)
            self.image_label.config(text="No album selected")
            self.image_canvas.delete("all")

    def show_image(self, path):
        self.image_canvas.delete("all")
        try:
            img = Image.open(path)
            resample_filter = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
            img = img.resize((280, 310), resample_filter)
            self.album_img = ImageTk.PhotoImage(img)
            self.image_canvas.create_image(0, 0, image=self.album_img, anchor='nw')
            self.image_label.config(text="")
        except FileNotFoundError:
            handle_error(f"Image file not found: {path}")
            self.image_label.config(text="Image not found")
        except UnidentifiedImageError:
            handle_error(f"Unsupported image format: {path}")
            self.image_label.config(text="Unsupported image format")
        except Exception as e:
            handle_error(f"Error loading image: {e}", e)
            self.image_label.config(text="Error loading image")

    def validate_input(self, name):
        """Validate the input name for bands/albums."""
        if not name or len(name.strip()) == 0:
            return False, "Name cannot be empty."
        if len(name) > 100:
            return False, "Name is too long (maximum 100 characters)."
        if re.search(r'[<>:"/\\|?*]', name):
            return False, "Name contains forbidden characters: <>:\"/\\|?*"
        return True, ""

    def add_band(self):
        band_name = self.simple_input_dialog("Add Band", "Enter new band name:")
        if band_name:
            band_name = band_name.strip()
            if band_name in self.data["bands"]:
                handle_error(f"Band '{band_name}' already exists.")
                return
            self.data["bands"][band_name] = {"albums": {}}
            self.save_data()
            self.refresh_bands()
            self.select_band_by_name(band_name)

    def delete_band(self):
        if self.selected_band:
            res = messagebox.askyesno("Confirm", f"Delete band '{self.selected_band}'?", parent=self.root)
            if res:
                del self.data["bands"][self.selected_band]
                self.save_data()
                self.refresh_bands()

    def add_album(self):
        if not self.selected_band:
            handle_error("No band selected.")
            return
        album_name = self.simple_input_dialog("Add Album", f"Enter new album name for '{self.selected_band}':")
        if album_name:
            album_name = album_name.strip()
            valid, error_msg = self.validate_input(album_name)
            if not valid:
                handle_error(error_msg)
                return
            if album_name in self.data["bands"][self.selected_band]["albums"]:
                handle_error("Album already exists for this band.")
                return
            self.data["bands"][self.selected_band]["albums"][album_name] = {"image": None}
            self.save_data()
            self.refresh_albums(self.selected_band)
            self.select_album_by_name(album_name)

    def delete_album(self):
        if self.selected_band and self.selected_album:
            if self.selected_band not in self.data["bands"]:
                handle_error(f"Band '{self.selected_band}' not found.")
                return
            res = messagebox.askyesno("Confirm", f"Delete album '{self.selected_album}' from '{self.selected_band}'?", parent=self.root)
            if res:
                album_data = self.data["bands"][self.selected_band]["albums"].get(self.selected_album, {})
                image_rel_path = album_data.get("image")
                if image_rel_path:
                    image_path = os.path.join(BASE_DIR, image_rel_path)
                    if os.path.exists(image_path):
                        try:
                            os.remove(image_path)
                        except Exception as e:
                            handle_error("Failed to delete image.", e)
                del self.data["bands"][self.selected_band]["albums"][self.selected_album]
                self.save_data()
                self.refresh_albums(self.selected_band)

    def associate_image(self):
        if not (self.selected_band and self.selected_album):
            handle_error("No band or album selected.")
            return

        filetypes = [("Image files", "*.jpg *.jpeg *.png"), ("All files", "*.*")]
        img_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select Album Image", filetypes=filetypes, parent=self.root)
        if img_path:
            try:
                # Validate that the selected file is a valid image before proceeding
                try:
                    Image.open(img_path).close()  # Check if it's a valid image file
                except Exception as e:
                    handle_error("Selected file is not a valid image.", e)
                    return

                sanitized_album_name = re.sub(r'[<>:"/\\|?*]', '', self.selected_album)
                sanitized_album_name = sanitized_album_name.replace(' ', '_')
                _, ext = os.path.splitext(img_path)
                ext = ext.lower()
                if ext not in ['.jpg', '.jpeg', '.png']:
                    handle_error("Unsupported file type selected. Please choose a JPG or PNG file.")
                    return

                dest_path = os.path.join(IMAGES_DIR, f"{sanitized_album_name}{ext}")
                counter = 1
                while os.path.exists(dest_path):
                    dest_path = os.path.join(IMAGES_DIR, f"{sanitized_album_name}_{counter}{ext}")
                    counter += 1
                
                try:
                    shutil.copy(img_path, dest_path)
                except Exception as e:
                    handle_error(f"Failed to copy image to {dest_path}", e)
                    return

                final_rel_path = os.path.relpath(dest_path, start=BASE_DIR)
                self.data["bands"][self.selected_band]["albums"][self.selected_album]["image"] = final_rel_path
                self.save_data()

                self.show_image(dest_path)
                self.remove_image_button.config(state=tk.NORMAL)

            except Exception as e:
                handle_error("Failed to associate image.", e)

    def remove_image(self):
        if self.selected_band and self.selected_album:
            album_data = self.data["bands"][self.selected_band]["albums"].get(self.selected_album, {})
            image_rel_path = album_data.get("image")
            if image_rel_path:
                image_path = os.path.join(BASE_DIR, image_rel_path)
                if os.path.exists(image_path):
                    try:
                        os.remove(image_path)
                    except Exception as e:
                        handle_error("Failed to delete image.", e)
            self.data["bands"][self.selected_band]["albums"][self.selected_album]["image"] = None
            self.save_data()
            self.image_label.config(text=f"No image for '{self.selected_album}'")
            self.image_canvas.delete("all")
            self.remove_image_button.config(state=tk.DISABLED)

    def simple_input_dialog(self, title, prompt):
        """Dialog to get a single line of input from the user with validation."""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        dialog.grab_set()

        self.center_window(dialog, 300, 150)

        tk.Label(dialog, text=prompt, font=("Helvetica", 10)).pack(pady=10)
        entry_var = tk.StringVar()
        entry = tk.Entry(dialog, textvariable=entry_var, width=30)
        entry.pack(pady=5)

        def on_ok():
            # Validate input before closing
            valid, error_msg = self.validate_input(entry_var.get())
            if valid:
                dialog.result = entry_var.get()
                dialog.destroy()
            else:
                handle_error(error_msg)

        def on_cancel():
            dialog.result = None
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=5)
        ok_btn = tk.Button(btn_frame, text="OK", command=on_ok)
        ok_btn.pack(side=tk.LEFT, padx=5)
        cancel_btn = tk.Button(btn_frame, text="Cancel", command=on_cancel)
        cancel_btn.pack(side=tk.LEFT, padx=5)

        dialog.bind('<Return>', lambda event: on_ok())

        entry.focus_set()
        self.root.wait_window(dialog)
        return dialog.result

    def select_band_by_name(self, band_name):
        bands = self.band_listbox.get(0, tk.END)
        if band_name in bands:
            idx = bands.index(band_name)
            self.band_listbox.selection_clear(0, tk.END)
            self.band_listbox.selection_set(idx)
            self.band_listbox.activate(idx)
            self.band_listbox.event_generate("<<ListboxSelect>>")

    def select_album_by_name(self, album_name):
        albums = self.album_listbox.get(0, tk.END)
        if album_name in albums:
            idx = albums.index(album_name)
            self.album_listbox.selection_clear(0, tk.END)
            self.album_listbox.selection_set(idx)
            self.album_listbox.activate(idx)
            self.album_listbox.event_generate("<<ListboxSelect>>")

    def export_collection(self):
        try:
            bands = sorted(self.data["bands"].keys())
            lines = []
            for band in bands:
                albums = sorted(self.data["bands"][band]["albums"].keys())
                for album in albums:
                    lines.append(f"{band} - {album}")
            
            with open(EXPORT_FILE, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            
            messagebox.showinfo("Export Complete", f"The collection has been exported to {EXPORT_FILE}", parent=self.root)
        except Exception as e:
            handle_error("Failed to export collection.", e)

    def import_collection(self):
        file_path = filedialog.askopenfilename(
            title="Select a TXT File",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            parent=self.root
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                lines = f.readlines()

            added_count = 0
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if " - " in line:
                    parts = line.split(" - ")
                    if len(parts) == 2:
                        band_name, album_name = parts[0].strip(), parts[1].strip()
                        if band_name and album_name:
                            if band_name not in self.data["bands"]:
                                self.data["bands"][band_name] = {"albums": {}}
                            if album_name not in self.data["bands"][band_name]["albums"]:
                                self.data["bands"][band_name]["albums"][album_name] = {"image": None}
                                added_count += 1
                            # If album already exists, skip
                    # If invalid line, skip
                # If invalid line, skip

            self.save_data()
            self.refresh_bands()
            messagebox.showinfo("Import Complete", f"Imported {added_count} new entries from the file.", parent=self.root)

        except Exception as e:
            handle_error("Failed to import collection.", e)

def main():
    root = tk.Tk()
    app = VinylTrackerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()