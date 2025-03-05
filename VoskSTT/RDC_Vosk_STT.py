import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import queue
import sounddevice as sd
import vosk
import sys
import json
import threading
import os
import keyboard  # Add this import
import logging
import win32gui
import win32con
import win32api
import win32process
import win32com.client
import win32api as wapi
import win32con as wcon
import time
import numpy as np
import traceback  # Add this import
import shutil  # For moving dropped model folders
from tkinter.font import Font  # Add for customizing fonts
import webbrowser  # Add this import

# Get the path to the model folder relative to the script location
BASE_PATH = os.path.dirname(os.path.abspath(__file__))

# Define available models
MODELS = {
    "English": "model-en",  # Default English model folder name
    "Small": "vosk-model-small-en-us-0.15",
    "Large": "vosk-model-en-us-0.22"
}

# Add preset applications
APP_PRESETS = {
    "Notepad": "Notepad",
    "RDC Vision": "RDC Vision",
    "Word": "Microsoft Word",
    "Chrome": "Google Chrome",
    "Default": None
}

# Configure logging
logging.basicConfig(level=logging.DEBUG,
                   format='%(asctime)s - %(levelname)s - %(message)s',
                   handlers=[logging.StreamHandler(),
                           logging.FileHandler('speech_debug.log')])

class DebugWindow:
    def __init__(self, parent):  # Add parent parameter
        self.window = tk.Toplevel()
        self.window.title("Debug Info")
        self.window.geometry("400x300")
        self.parent = parent  # Store parent reference
        
        self.log_area = scrolledtext.ScrolledText(self.window, wrap=tk.WORD)
        self.log_area.pack(expand=True, fill='both')
        
        # Hide window by default
        self.window.withdraw()
        
        # Add window close handler
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def show(self):
        self.window.deiconify()
        
    def on_closing(self):
        self.window.withdraw()  # Just hide the window
    
    def log(self, message):
        self.log_area.insert(tk.END, f"{message}\n")
        self.log_area.see(tk.END)

class WindowSelector(tk.Toplevel):
    def __init__(self, parent, presets):
        super().__init__(parent)
        self.title("Select Target Window")
        self.geometry("400x400")
        
        # Add presets dropdown
        preset_frame = ttk.Frame(self)
        preset_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(preset_frame, text="Quick Select:").pack(side=tk.LEFT)
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(
            preset_frame,
            textvariable=self.preset_var,
            values=list(presets.keys()),
            state='readonly'
        )
        self.preset_combo.pack(side=tk.LEFT, padx=5)
        self.preset_combo.bind('<<ComboboxSelected>>', self.on_preset_selected)
        
        # Existing window list
        self.listbox = tk.Listbox(self)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Window info display
        self.info_text = scrolledtext.ScrolledText(self, height=4)
        self.info_text.pack(fill=tk.X, padx=5)
        
        ttk.Button(self, text="Refresh", command=self.refresh_windows).pack(pady=5)
        
        self.selected_hwnd = None
        self.windows = []
        self.presets = presets
        self.refresh_windows()
        
        self.listbox.bind('<Double-Button-1>', self.on_select)
        self.listbox.bind('<<ListboxSelect>>', self.show_window_info)

    def show_window_info(self, event=None):
        selection = self.listbox.curselection()
        if selection:
            hwnd = self.windows[selection[0]][0]
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                self.info_text.delete(1.0, tk.END)
                self.info_text.insert(tk.END, 
                    f"Window Handle: {hwnd}\n"
                    f"Process ID: {pid}\n"
                    f"Class Name: {class_name}\n"
                )
            except Exception as e:
                self.info_text.insert(tk.END, f"Error getting window info: {e}")

    def on_preset_selected(self, event):
        preset_name = self.preset_var.get()
        preset_title = self.presets[preset_name]
        if preset_title:
            for idx, (hwnd, title) in enumerate(self.windows):
                if preset_title.lower() in title.lower():
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(idx)
                    self.listbox.see(idx)
                    self.show_window_info()
                    break

    def refresh_windows(self):
        self.listbox.delete(0, tk.END)
        self.windows = []
        win32gui.EnumWindows(self._window_callback, None)
        
    def _window_callback(self, hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                self.windows.append((hwnd, title))
                self.listbox.insert(tk.END, title)
                
    def on_select(self, _):
        selection = self.listbox.curselection()
        if selection:
            self.selected_hwnd = self.windows[selection[0]][0]
            self.destroy()

class LoadingWindow:
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Loading Model")
        self.window.geometry("300x100")
        self.window.transient(parent)
        self.window.grab_set()
        
        # Center the window
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'+{x}+{y}')
        
        self.label = ttk.Label(self.window, text="Loading model...\nPlease wait...")
        self.label.pack(pady=20)
        
        self.progress = ttk.Progressbar(self.window, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=20)
        self.progress.start()

    def destroy(self):
        self.window.destroy()

class ModelSelector(tk.Toplevel):
    def __init__(self, parent, models):
        super().__init__(parent)
        self.title("Select Model")
        self.geometry("500x300")  # Keep some width if needed
        self.selected_model = None
        self.models = models
        
        tk.Label(self, text="Available Models:").pack(pady=5)
        self.listbox = tk.Listbox(self)
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        for m in self.models:
            self.listbox.insert(tk.END, m)

        # Create a frame to hold the buttons
        button_frame = tk.Frame(self)  # Use tk.Frame instead of ttk.Frame
        button_frame.pack(pady=5)

        # OK button
        ok_button = tk.Button(button_frame, text="OK", command=self.on_select)
        ok_button.pack(side=tk.LEFT, padx=5)

        # Download More Models button
        download_button = tk.Button(button_frame, text="Download More Models", command=self.open_model_link)
        download_button.pack(side=tk.LEFT, padx=5)

    def open_model_link(self):
        webbrowser.open_new("https://alphacephei.com/vosk/models")

    def on_select(self):
        selection = self.listbox.curselection()
        if selection:
            self.selected_model = self.models[selection[0]]
        self.destroy()

def find_available_models():
    import os
    one_level_up = os.path.abspath(os.path.join(BASE_PATH, '..'))
    models_path = os.path.join(one_level_up, 'models')
    if not os.path.isdir(models_path):
        return []
    all_dirs = [d for d in os.listdir(models_path)
                if os.path.isdir(os.path.join(models_path, d))]
    return [os.path.join(models_path, d) for d in all_dirs]

def calculate_folder_size(folder_path):
    total_size = 0
    for root, dirs, files in os.walk(folder_path):
        for f in files:
            fp = os.path.join(root, f)
            if os.path.isfile(fp):
                total_size += os.path.getsize(fp)
    return total_size

def find_smallest_model():
    models = find_available_models()
    if not models:
        return None
    sizes = [(m, calculate_folder_size(m)) for m in models]
    smallest = min(sizes, key=lambda x: x[1])[0]
    return smallest

class SpeechToTextApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vosk Speech to Text - made by RDC")  # Changed from RDC_VoskGiga_STT
        self.root.geometry("630x400")  # Increase width by 5% (600 * 0.05 = 30)
        # Change to lighter gray (20% lighter than #1A1A1A)
        self.root.configure(bg='#3D3D3D')  # 15% lighter than before

        self.is_recording = False
        self.is_muted = False
        self.q = queue.Queue()
        self.current_model = None
        self.rec = None
        self.target_window = None
        self.output_lock = threading.Lock()  # Add this line
        self.debug_window = DebugWindow(self)  # Pass self as parent
        self.debug_log("Application started")
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.last_focused_window = None
        self.window_mode_active = False
        self.live_mode_active = False
        self.last_audio_time = time.time()
        self.silence_threshold = 500  # Adjust based on your mic sensitivity
        self.silence_duration = 3.0  # Seconds of silence before Enter
        self.last_enter_time = 0
        self.min_enter_interval = 1.0  # Minimum seconds between Enter presses
        self.debug_mode = True  # Enable detailed logging
        self.live_mode_test = True  # Test mode without Enter simulation
        self.is_delivering_text = False  # Prevent premature Enter
        self.delay_var = tk.IntVar(value=0)  # Changed default to 0 for Manual mode
        self.phrase_var = tk.StringVar(value="")  # Custom key phrase
        self.activate_on_phrase = tk.BooleanVar(value=False)
        self.show_key_phrase_var = tk.BooleanVar(value=True)  # New toggle for showing phrase

        # Key phrase settings
        self.default_phrases = ["Send it", "I'm done talking", "That's it"]
        self.key_phrases = set(self.default_phrases)
        self.waiting_for_silence = False
        
        # Improved state tracking
        self.key_phrase_detected = False
        self.last_speech_time = time.time()
        self.silence_threshold = 300  # Lowered threshold for better detection
        self.last_phrase_time = 0
        self.min_phrase_interval = 0.5  # Minimum seconds between phrase triggers
        
        self.cursor_mode_active = False  # Add this line
        self.silence_mode_var = tk.StringVar(value="Manual")  # Add this line

        # Create GUI elements
        self.create_widgets()
        self.create_context_menu()

        # Add clipboard support
        self.root.clipboard_clear()

        # Disable all widgets initially
        self.root.withdraw()
        
        # Show loading window and initialize model
        self.loading_window = LoadingWindow(self.root)
        self.root.after(100, self.initialize_model)
        
        # Audio stream setup
        self.stream = None
        self.last_final_text = ""
        self.last_partial_text = ""
        self.disable_partials = True  # new flag to ignore partial results

    def create_context_menu(self):
        """Create right-click context menu"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Cut", command=self.cut_text)
        self.context_menu.add_command(label="Copy", command=self.copy_selected_text)
        self.context_menu.add_command(label="Paste", command=self.paste_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Select All", command=self.select_all_text)

    def create_widgets(self):
        # Configure frame backgrounds to match lighter theme
        style = ttk.Style()
        style.configure('Darker.TFrame', background='#333333')  # Lighter gray
        style.configure('TLabel', background='#333333', foreground='white')  # White text
        style.configure('TLabelframe', background='#333333')
        style.configure('TLabelframe.Label', background='#333333', foreground='white')  # White text
        
        # Configure specific elements
        style.configure('TButton', background='#333333')
        style.configure('TCheckbutton', background='#333333', foreground='white')  # White text
        style.configure('TCombobox', background='#333333', foreground='white')  # White text
        
        # Define a new style for ttk.Buttons
        style.configure('Custom.TButton',
                        background='#6A6A6A',
                        foreground='black')

        top_frame = ttk.Frame(self.root, style='Darker.TFrame')
        top_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Left side - Debug button
        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side=tk.LEFT, padx=2)
        self.debug_button = ttk.Button(
            left_frame, 
            text="Debug",
            command=self.show_debug,
            width=10,
            style='Custom.TButton'  # Apply custom style
        )
        self.debug_button.pack()
        
        # Middle frame for title with absolute positioning
        middle_frame = ttk.Frame(top_frame, style='Darker.TFrame')  # Add style here
        middle_frame.pack(side=tk.LEFT, expand=True)
        
        # Create large title label with adjusted positioning and matching background
        title_container = ttk.Frame(middle_frame, style='Darker.TFrame')  # Add style here
        title_container.pack(fill=tk.X)
        title_label = ttk.Label(
            title_container,
            text="VOSK STT",
            font=('Arial', 24, 'bold'),
            foreground='#A0A0A0',  # Keep light grey text
            background='#333333',   # Match dark background
            padding=(135, 5)        # Adjusted from 150 to 135 pixels for better centering
        )
        title_label.pack(pady=5)
        
        # Right side - Live mode controls
        right_frame = ttk.Frame(top_frame)
        right_frame.pack(side=tk.RIGHT, padx=2)
        
        # Create live mode subframe
        live_mode_frame = ttk.LabelFrame(right_frame, text="Live Mode (Experimental)")
        live_mode_frame.pack(pady=2)
        
        # Live mode button and status
        self.live_button = tk.Button(
            live_mode_frame,
            text="Live Mode: Off",
            command=self.toggle_live_mode,
            width=15,
            bg='#6A6A6A',
            fg='black'
        )
        self.live_button.pack(pady=2)
        
        # Add silence mode selector
        silence_frame = ttk.Frame(live_mode_frame)
        silence_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(silence_frame, text="Auto-Enter:").pack(side=tk.LEFT, padx=2)
        self.silence_mode_var = tk.StringVar(value="Manual")
        modes = ["Manual", "Quick (3s)", "Medium (5s)", "Slow (10s)"]
        self.silence_combo = ttk.Combobox(
            silence_frame,
            values=modes,
            textvariable=self.silence_mode_var,
            state='readonly',
            width=10
        )
        self.silence_combo.pack(side=tk.LEFT, padx=2)
        self.silence_combo.bind('<<ComboboxSelected>>', self.on_silence_mode_change)
        
        # Remove old delay_combo and its frame since we're replacing it
        if hasattr(self, 'delay_combo'):
            self.delay_combo.destroy()
        
        # Middle frame for silence delay slider
        middle_frame = ttk.Frame(top_frame)
        middle_frame.pack(side=tk.LEFT, expand=True)
        
        # Enhanced key phrase frame
        phrase_frame = ttk.LabelFrame(self.root, text="Key Phrase Settings")
        phrase_frame.pack(fill=tk.X, padx=5, pady=5)

        row_frame = ttk.Frame(phrase_frame)
        row_frame.pack(fill=tk.X, pady=2)

        ttk.Checkbutton(
            row_frame,
            text="Use Key Phrases",
            variable=self.activate_on_phrase,
            command=self.toggle_phrase_mode
        ).pack(side=tk.LEFT, padx=5)

        ttk.Label(
            row_frame,
            text="Default: " + ", ".join(self.default_phrases)
        ).pack(side=tk.LEFT, padx=5)

        ttk.Label(row_frame, text="Custom:").pack(side=tk.LEFT)
        self.phrase_entry = ttk.Entry(row_frame, textvariable=self.phrase_var, width=15)
        self.phrase_entry.pack(side=tk.LEFT, padx=2)
        
        # Move the "Add" button to the right side
        ttk.Button(row_frame, text="Add", command=self.add_custom_phrase).pack(side=tk.RIGHT, padx=5)

        # Add a toggle for showing key phrases in transcript
        show_phrase_frame = ttk.Frame(self.root, style='Darker.TFrame')
        show_phrase_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Checkbutton(
            show_phrase_frame,
            text="Show key phrase in output",
            variable=self.show_key_phrase_var
        ).pack(side=tk.LEFT, padx=5)

        # Control buttons frame (existing code)
        control_frame = ttk.Frame(self.root)
        control_frame.pack(pady=10)

        self.start_button = ttk.Button(control_frame, text="Start", command=self.toggle_recording)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.mute_button = ttk.Button(control_frame, text="Mute", command=self.toggle_mute)
        self.mute_button.pack(side=tk.LEFT, padx=5)

        self.window_button = ttk.Button(
            control_frame, 
            text="Window Linked: Off", 
            command=self.toggle_window_mode
        )
        self.window_button.pack(side=tk.LEFT, padx=5)

        self.cursor_button = ttk.Button(
            control_frame, 
            text="Cursor Mode: Off", 
            command=self.toggle_cursor_mode
        )
        self.cursor_button.pack(side=tk.LEFT, padx=5)

        # Status frame
        status_frame = ttk.Frame(self.root, style='Darker.TFrame')
        status_frame.pack(fill=tk.X, padx=10)
        
        # Left side - Status label
        self.status_label = ttk.Label(status_frame, text="Mode: Default")
        self.status_label.pack(side=tk.LEFT)
        
        # Right side - Copy and Clear buttons (moved to switch_frame)
        
        self.model_label = ttk.Label(status_frame, text="Current Model: None")
        self.model_label.pack(side=tk.LEFT, padx=10)

        # Create a new frame just above the status_frame for the Switch Model button
        switch_frame = ttk.Frame(self.root, style='Darker.TFrame')
        switch_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        self.switch_model_button = ttk.Button(
            switch_frame,
            text="Models",  # Renamed
            command=self.switch_model,
            width=10
        )
        self.switch_model_button.pack(side=tk.LEFT, padx=5)

        # Rename the placeholder "Aa" button to "Text Settings"
        self.placeholder_size_button = ttk.Button(
            switch_frame,
            text="Text Settings",
            command=self.open_text_settings,
            width=10
        )
        self.placeholder_size_button.pack(side=tk.LEFT, padx=5)
        
        # Add Copy All and Clear All buttons to switch_frame
        self.copy_button = ttk.Button(
            switch_frame,
            text="Copy All",
            command=self.copy_all_text,
            width=10
        )
        self.copy_button.pack(side=tk.RIGHT, padx=5)
        
        self.clear_button = ttk.Button(
            switch_frame,
            text="Clear All",
            command=self.clear_all_text,
            width=10
        )
        self.clear_button.pack(side=tk.RIGHT, padx=5)
        
        # Modify text area with adjusted lighter theme and make it editable
        self.text_area = scrolledtext.ScrolledText(
            self.root, 
            wrap=tk.WORD, 
            width=50, 
            height=20,
            bg='#707070',  # 5% lighter than before
            fg='black',
            font=('TkDefaultFont', 12),
            state='normal'  # Make sure it's editable
        )
        self.text_area.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)
        
        # Enable mouse interaction
        self.text_area.config(cursor="ibeam")  # Show text cursor
        
        # Bind right-click event
        self.text_area.bind("<Button-3>", self.show_context_menu)
        
        # Also bind common keyboard shortcuts
        self.text_area.bind("<Control-c>", lambda e: self.copy_selected_text())
        self.text_area.bind("<Control-x>", lambda e: self.cut_text())
        self.text_area.bind("<Control-v>", lambda e: self.paste_text())
        self.text_area.bind("<Control-a>", lambda e: self.select_all_text())

    def show_context_menu(self, event):
        """Show the context menu at mouse position"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def copy_selected_text(self):
        """Copy selected text to clipboard"""
        try:
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.debug_log("Selected text copied to clipboard")
        except tk.TclError:  # No selection
            self.debug_log("No text selected for copying")

    def cut_text(self):
        """Cut selected text"""
        try:
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.debug_log("Text cut to clipboard")
        except tk.TclError:  # No selection
            self.debug_log("No text selected for cutting")

    def paste_text(self):
        """Paste text from clipboard"""
        try:
            text = self.root.clipboard_get()
            try:
                self.text_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass  # No selection to replace
            self.text_area.insert(tk.INSERT, text)
            self.debug_log("Text pasted from clipboard")
        except tk.TclError:
            self.debug_log("No text available in clipboard")

    def select_all_text(self, event=None):
        """Select all text in the text area"""
        self.text_area.tag_add(tk.SEL, "1.0", tk.END)
        self.text_area.mark_set(tk.INSERT, "1.0")
        self.text_area.see(tk.INSERT)
        return 'break'  # Prevent default behavior

    def show_debug(self):
        """Show debug window"""
        self.debug_window.show()

    def toggle_recording(self):
        if not self.is_recording:
            self.start_recording()
            self.start_button.configure(text="Stop")
        else:
            self.stop_recording()
            self.start_button.configure(text="Start")

    def toggle_mute(self):
        self.is_muted = not self.is_muted
        self.mute_button.configure(text="Unmute" if self.is_muted else "Mute")

    def audio_callback(self, indata, frames, time, status):
        """Enhanced audio callback with error handling"""
        try:
            if not self.is_muted:
                audio_data = bytes(indata)
                self.q.put(audio_data)
                
                try:
                    if self.activate_on_phrase.get():
                        self.check_key_phrase(audio_data)
                    elif self.live_mode_active:
                        self.check_silence_and_enter(audio_data)
                except Exception as e:
                    self.debug_log(f"Error in live or key phrase mode processing: {str(e)}\n{traceback.format_exc()}")
                    self.disable_live_mode()
                        
        except Exception as e:
            self.debug_log(f"Error in audio callback: {str(e)}\n{traceback.format_exc()}")

    def process_audio(self):
        """Process audio stream"""
        while self.is_recording:
            try:
                data = self.q.get(timeout=1.0)
                if self.rec and self.rec.AcceptWaveform(data):
                    result = json.loads(self.rec.Result())
                    text = result.get("text", "").strip().lower()
                    
                    if text:
                        # Check for key phrase if enabled
                        if self.activate_on_phrase.get():
                            for phrase in self.key_phrases:
                                if phrase.lower() == text:
                                    self.debug_log(f"Key phrase detected: {text}")
                                    # Only show phrase if the option is enabled
                                    if self.show_key_phrase_var.get():
                                        self.root.after(0, self.output_text, text)
                                    self.simulate_enter_key()
                                    break
                            else:  # No key phrase found, show text normally
                                if text != self.last_final_text:
                                    self.last_final_text = text
                                    self.root.after(0, self.output_text, text)
                        else:  # Normal mode, always show text
                            if text != self.last_final_text:
                                self.last_final_text = text
                                self.root.after(0, self.output_text, text)
                            
            except queue.Empty:
                continue
            except Exception as e:
                self.debug_log(f"Error processing audio: {str(e)}")
                if not self.is_recording:
                    break

    def simulate_enter_key(self):
        """Just press Enter key"""
        try:
            self.debug_log("Pressing Enter key")
            wapi.keybd_event(wcon.VK_RETURN, 0, 0, 0)
            wapi.keybd_event(wcon.VK_RETURN, 0, wcon.KEYEVENTF_KEYUP, 0)
        except Exception as e:
            self.debug_log(f"Error pressing Enter: {str(e)}")

    def update_streaming_text(self, partial_text):
        """Append new partial text only if different from the last update"""
        self.text_area.insert(tk.END, " " + partial_text)
        self.text_area.see(tk.END)

    def ensure_window_focus(self, hwnd):
        """Ensure window has focus and return success status"""
        try:
            # Get current focus
            current_focus = win32gui.GetForegroundWindow()
            self.debug_log(f"Current focused window: {win32gui.GetWindowText(current_focus)}")
            
            if current_focus != hwnd:
                # Bring window to front
                win32gui.ShowWindow(hwnd, wcon.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                
                # Wait briefly for focus
                for _ in range(10):
                    if win32gui.GetForegroundWindow() == hwnd:
                        break
                    self.root.after(50)
                
            # Verify focus
            new_focus = win32gui.GetForegroundWindow()
            success = (new_focus == hwnd)
            self.debug_log(f"Focus attempt {'successful' if success else 'failed'} for window: {win32gui.GetWindowText(hwnd)}")
            return success
            
        except Exception as e:
            self.debug_log(f"Error setting focus: {str(e)}")
            return False

    def simulate_typing(self, text):
        """Simulate keyboard input using keyboard.write for Unicode support"""
        try:
            keyboard.write(text, delay=0.01)
            keyboard.send('space')
            return True
        except Exception as e:
            self.debug_log(f"Error simulating typing: {str(e)}")
            return False

    def output_text(self, text):
        # Prevent pressing Enter while text is being delivered
        self.is_delivering_text = True
        self.debug_log(f"Attempting to output text: {text}")
        words = text.split()
        delay = 100  # delay in milliseconds between words

        def deliver_word(i):
            if i < len(words):
                word = words[i]
                # Try to send word via typing simulation if target window and cursor mode are active
                if self.cursor_mode_active:
                    self.debug_log("Cursor mode active - delivering word via simulate_typing")
                    self.simulate_typing(word + " ")
                elif self.target_window and win32gui.IsWindow(self.target_window):
                    window_title = win32gui.GetWindowText(self.target_window)
                    self.debug_log(f"Target window: {window_title} (hwnd: {self.target_window})")
                    if self.ensure_window_focus(self.target_window):
                        self.simulate_typing(word + " ")
                    else:
                        self.debug_log("Focusing failed - falling back to text area")
                        self.text_area.insert(tk.END, word + " ")
                        self.text_area.see(tk.END)
                else:
                    # Fallback to text area delivery
                    self.text_area.insert(tk.END, word + " ")
                    self.text_area.see(tk.END)
                # Schedule next word delivery
                self.root.after(delay, lambda: deliver_word(i + 1))
            else:
                self.is_delivering_text = False

        deliver_word(0)

    def initialize_model(self):
        smallest = find_smallest_model()
        if not smallest:
            messagebox.showerror("Error", "No models found in the folder. Please add a model to continue.")
            self.root.destroy()
            return
        try:
            self.current_model = vosk.Model(smallest)
            self.rec = vosk.KaldiRecognizer(self.current_model, 16000)
            self.loading_window.destroy()
            self.root.deiconify()
            self.model_label.config(text=f"Current Model: {os.path.basename(smallest)}")
            self.status_label.config(text="Model loaded successfully")
            self.debug_log(f"Model loaded from: {smallest}")
        except Exception as e:
            self.loading_window.destroy()
            messagebox.showerror("Error", "Error loading model: Please ensure the model files are valid and try again.")
            self.debug_log(f"Failed to load model: {str(e)}")
            self.root.destroy()

    def load_model(self, model_name=None):
        """Modified to show loading indicator"""
        try:
            loading = LoadingWindow(self.root)
            self.root.update()
            
            was_recording = self.is_recording
            if was_recording:
                self.stop_recording()

            # Load model
            model_path = os.path.join(BASE_PATH, MODELS["English"])
            if not os.path.exists(model_path):
                raise FileNotFoundError(f"Model directory not found: {model_path}")
            
            self.current_model = vosk.Model(model_path)
            self.rec = vosk.KaldiRecognizer(self.current_model, 16000)
            
            if was_recording:
                self.start_recording()
            
            loading.destroy()
            self.status_label.config(text="Model loaded successfully")
            self.debug_log("Model reloaded successfully")
            return True
            
        except Exception as e:
            if 'loading' in locals():
                loading.destroy()
            error_msg = f"Failed to load model: {str(e)}"
            messagebox.showerror("Error", error_msg)
            self.debug_log(error_msg)
            return False

    def switch_model(self):
        """Bare minimum model switching"""
        models = find_available_models()
        if not models:
            messagebox.showerror("Error", "No models found")
            return

        # Stop everything
        was_recording = self.is_recording
        self.stop_recording()
        
        # Show selector
        selector = ModelSelector(self.root, models)
        selector.wait_window()
        
        try:
            if selector.selected_model:
                # Kill old model completely
                self.rec = None
                if self.current_model:
                    del self.current_model
                
                # Load new one
                self.current_model = vosk.Model(selector.selected_model)
                self.rec = vosk.KaldiRecognizer(self.current_model, 16000)
                
                # Update UI
                name = os.path.basename(selector.selected_model)
                self.model_label.config(text=f"Model: {name}")
                
                # Restart if needed
                if was_recording:
                    self.start_recording()
                    
        except Exception as e:
            self.debug_log(f"Model switch failed: {str(e)}")
            messagebox.showerror("Error", "Model switch failed")
            self.initialize_model()  # Try to recover

    def debug_log(self, message):
        logging.debug(message)
        # Still log messages even when window is hidden
        self.debug_window.log(message)

    def toggle_window_mode(self):
        try:
            if not self.window_mode_active:
                # Activate window mode
                selector = WindowSelector(self.root, APP_PRESETS)
                selector.wait_window()
                
                if selector.selected_hwnd:
                    self.target_window = selector.selected_hwnd
                    title = win32gui.GetWindowText(self.target_window)
                    self.window_mode_active = True
                    self.window_button.config(text="Window Linked: On")
                    self.debug_log(f"Window mode activated: {title}")
                    self.status_label.config(text=f"Target: {title[:20]}...")
                    # Reset silence detection when window mode is activated
                    self.last_audio_time = time.time()
                    self.last_enter_time = time.time()
            else:
                # Deactivate window mode
                self.target_window = None
                self.window_mode_active = False
                self.window_button.config(text="Window Linked: Off")
                self.debug_log("Window mode deactivated")
                self.status_label.config(text="Mode: Default")
                
        except Exception as e:
            self.debug_log(f"Error toggling window mode: {str(e)}")
            self.target_window = None
            self.window_mode_active = False
            self.window_button.config(text="Window Linked: Off")
            self.status_label.config(text="Mode: Default")

    def toggle_cursor_mode(self):
        """Toggle cursor-based text sending mode"""
        self.cursor_mode_active = not self.cursor_mode_active
        self.cursor_button.config(text="Cursor Mode: On" if self.cursor_mode_active else "Cursor Mode: Off")
        self.debug_log(f"Cursor mode {'activated' if self.cursor_mode_active else 'deactivated'}")

    def toggle_live_mode(self):
        """Updated live mode toggle with new UI"""
        try:
            self.live_mode_active = not self.live_mode_active
            if self.live_mode_active:
                self.live_button.config(
                    text="Live Mode: On",
                    bg='lightblue'
                )
                self.silence_combo.config(state='readonly')
            else:
                self.live_button.config(
                    text="Live Mode: Off",
                    bg='#6A6A6A'
                )
                self.silence_combo.config(state='disabled')
                self.silence_mode_var.set("Manual")
            
            self.debug_log(f"Live mode {'activated' if self.live_mode_active else 'deactivated'}")
            self.reset_silence_detection()
            
        except Exception as e:
            self.debug_log(f"Error toggling live mode: {str(e)}")
            self.disable_live_mode()

    def check_silence_and_enter(self, audio_data):
        """Simple silence detection and Enter press"""
        if not self.live_mode_active or self.delay_var.get() == 0:
            return

        try:
            audio_array = np.frombuffer(audio_data, dtype=np.int16)
            rms = np.sqrt(np.mean(np.square(audio_array)))
            current_time = time.time()

            if rms > self.silence_threshold:
                self.last_speech_time = current_time
            else:
                silence_duration = current_time - self.last_speech_time
                delay = self.delay_var.get()

                if silence_duration >= delay:
                    self.debug_log(f"Live mode - pressing Enter after {delay}s silence")
                    self.simulate_enter_key()
                    self.last_speech_time = current_time

        except Exception as e:
            self.debug_log(f"Live mode error: {str(e)}")

    def disable_live_mode(self):
        """Safely disable live mode"""
        try:
            self.live_mode_active = False
            self.live_button.config(bg='SystemButtonFace')  # Reset to default color
            self.live_indicator.config(text="Mode: Default")
            self.debug_log("Live mode safely disabled due to error")
        except Exception as e:
            self.debug_log(f"Error disabling live mode: {str(e)}\n{traceback.format_exc()}")

    def toggle_phrase_mode(self):
        """Improved mode switching"""
        if self.activate_on_phrase.get():
            self.debug_log("Key phrase mode activated - will press Enter immediately on phrase detection")
            if hasattr(self, 'delay_combo'):
                self.delay_combo.config(state='disabled')
            # Reset any ongoing silence detection
            self.waiting_for_silence = False
            self.live_mode_active = False  # Ensure live mode is off
            if hasattr(self, 'live_button'):
                self.live_button.config(bg='#6A6A6A', text="Live Mode: Off")
        else:
            self.debug_log("Silence delay mode activated")
            if hasattr(self, 'delay_combo'):
                self.delay_combo.config(state='readonly')  
            self.reset_silence_detection()

    def add_custom_phrase(self):
        """Add custom phrase to recognized phrases"""
        phrase = self.phrase_var.get().strip()
        if phrase:
            self.key_phrases.add(phrase)
            self.debug_log(f"Added custom key phrase: '{phrase}'")
            self.phrase_var.set("")  # Clear entry
        else:
            messagebox.showwarning("Invalid Phrase", "Please enter a valid phrase")

    def reset_silence_detection(self):
        """Reset silence detection state"""
        self.waiting_for_silence = False
        self.last_audio_time = time.time()
        self.last_enter_time = time.time()
        self.silence_start_time = None
        self.debug_log("Silence detection reset")

    def copy_all_text(self):
        """Copy all text from the text area to clipboard"""
        try:
            text = self.text_area.get(1.0, tk.END).strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.debug_log("Text copied to clipboard successfully")
        except Exception as e:
            self.debug_log(f"Error copying text: {str(e)}")
            messagebox.showerror("Copy Error", "Failed to copy text to clipboard")

    def clear_all_text(self):
        """Clear all text from the text area"""
        try:
            self.text_area.delete(1.0, tk.END)
            self.debug_log("Text area cleared successfully")
        except Exception as e:
            self.debug_log(f"Error clearing text: {str(e)}")
            messagebox.showerror("Clear Error", "Failed to clear text area")

    def open_text_settings(self):
        """Open a new window to adjust text appearance."""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Text Settings")
        settings_window.geometry("300x300")

        # Font size
        tk.Label(settings_window, text="Font Size:").pack()
        self.font_size_var = tk.IntVar(value=12)  # Default
        tk.Spinbox(settings_window, from_=8, to=48, textvariable=self.font_size_var).pack()

        # Font style
        tk.Label(settings_window, text="Font Style:").pack()
        self.font_style_var = tk.StringVar(value="TkDefaultFont")
        tk.OptionMenu(settings_window, self.font_style_var, "TkDefaultFont", "TkFixedFont", "Arial", "Helvetica", "Times").pack()

        # Basic formatting options (bold, italic, underline)
        self.bold_var = tk.BooleanVar(value=False)
        tk.Checkbutton(settings_window, text="Bold", variable=self.bold_var).pack()
        self.italic_var = tk.BooleanVar(value=False)
        tk.Checkbutton(settings_window, text="Italic", variable=self.italic_var).pack()
        self.underline_var = tk.BooleanVar(value=False)
        tk.Checkbutton(settings_window, text="Underline", variable=self.underline_var).pack()

        # (Optional) Apply/OK button to confirm changes
        def apply_changes():
            new_size = self.font_size_var.get()
            new_style = self.font_style_var.get()
            
            # Determine weight and underline
            weight = "bold" if self.bold_var.get() else "normal"
            underline = self.underline_var.get()
            
            # Create and apply new font
            current_font = Font(
                family=new_style,
                size=new_size,
                weight=weight,
                underline=underline
            )
            self.text_area.configure(font=current_font)

        ttk.Button(settings_window, text="Apply", command=apply_changes).pack(pady=5)

    def on_silence_mode_change(self, event=None):
        """Handle silence mode changes"""
        mode = self.silence_mode_var.get()
        if (mode == "Quick (3s)"):
            self.delay_var.set(3)
        elif (mode == "Medium (5s)"):
            self.delay_var.set(5)
        elif (mode == "Slow (10s)"):
            self.delay_var.set(10)
        else:  # Manual mode
            self.delay_var.set(0)
        self.debug_log(f"Silence mode changed to: {mode}")

    def start_recording(self):
        """Start audio recording and processing"""
        try:
            self.stream = sd.RawInputStream(
                samplerate=16000,
                blocksize=8000,
                dtype='int16',
                channels=1,
                callback=self.audio_callback
            )
            self.stream.start()
            self.is_recording = True
            self.debug_log("Started recording")
            threading.Thread(target=self.process_audio, daemon=True).start()
        except Exception as e:
            self.debug_log(f"Failed to start recording: {str(e)}")
            self.is_recording = False
            raise

    def stop_recording(self):
        """Completely stop recording and cleanup"""
        if self.stream:
            self.stream.stop()
            self.stream.close()
            self.stream = None
        self.is_recording = False
        self.q.queue.clear()
        self.debug_log("Recording stopped and cleaned up")

def main():
    root = tk.Tk()
    app = SpeechToTextApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()