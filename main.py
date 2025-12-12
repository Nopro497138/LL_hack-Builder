import tkinter as tk
from tkinter import ttk, messagebox
import keyboard
import pyautogui
import threading
import time
import json
import os
import sys
import ctypes
from collections import defaultdict
from PIL import ImageGrab, Image, ImageEnhance
import cv2
import numpy as np
import pyttsx3
try:
    import pytesseract
    # Try multiple common installation paths
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        r'C:\Tesseract-OCR\tesseract.exe',
        r'C:\Users\Public\Tesseract-OCR\tesseract.exe'
    ]
    
    tesseract_found = False
    for path in possible_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            try:
                pytesseract.get_tesseract_version()
                tesseract_found = True
                print(f"Tesseract found at: {path}")
                break
            except:
                continue
    
    TESSERACT_AVAILABLE = tesseract_found
except:
    TESSERACT_AVAILABLE = False

class WordAutofiller:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word Autofiller Pro")
        self.root.geometry("700x920")
        self.root.configure(bg='#0d1117')
        
        # Set icon if available
        self.set_icon()
        
        # Settings
        self.settings = {
            'backspace_delay': 0.25,
            'typing_delay': 0.15,
            'start_delay': 0.3,
            'after_delete_delay': 0.3,
            'prefer_longer_words': 1.0,
            'min_word_length': 4,
            'max_suggestions_per_prefix': 50
        }
        
        # State variables
        self.is_listening = False
        self.current_buffer = ""
        self.used_words = defaultdict(set)
        self.last_completion = ""
        self.completion_lock = threading.Lock()
        self.current_tab = "main"
        
        # OCR variables
        self.ocr_active = False
        self.ocr_thread = None
        self.is_admin = self.check_admin()
        self.tts_engine = None
        if self.is_admin:
            try:
                self.tts_engine = pyttsx3.init()
                self.tts_engine.setProperty('rate', 150)
            except:
                pass
        
        # Load words
        self.words = self.load_words_from_json()
        if not self.words:
            return
            
        self.word_index = self.build_word_index()
        
        self.setup_ui()
        self.start_keyboard_monitoring()
    
    def set_icon(self):
        """Set application icon"""
        try:
            # Check if running as exe
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
            else:
                icon_path = 'icon.ico'
            
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
    
    def check_admin(self):
        """Check if running as administrator"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
    
    def speak(self, text):
        """Text to speech"""
        if self.tts_engine:
            try:
                self.tts_engine.say(text)
                self.tts_engine.runAndWait()
            except:
                pass
    
    def detect_letter_tiles(self, screenshot):
        """Detect letter tiles from screenshot"""
        if not TESSERACT_AVAILABLE:
            return []
        
        try:
            # Convert to numpy array
            img = np.array(screenshot)
            
            # Convert to grayscale
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
            
            # Apply threshold to get white tiles
            _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)
            
            # Find contours (white squares)
            contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            tiles = []
            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                
                # Filter for square-ish shapes (letter tiles)
                aspect_ratio = w / float(h) if h > 0 else 0
                if 0.8 < aspect_ratio < 1.2 and w > 50 and w < 200:
                    # Extract tile region
                    tile_img = screenshot.crop((x, y, x + w, y + h))
                    
                    # OCR on tile
                    text = pytesseract.image_to_string(
                        tile_img,
                        config='--psm 10 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ'
                    ).strip()
                    
                    if len(text) == 1 and text.isalpha():
                        tiles.append({
                            'letter': text.upper(),
                            'x': x,
                            'y': y,
                            'w': w,
                            'h': h
                        })
            
            return tiles
        except Exception as e:
            self.log_message(f"OCR Error: {str(e)}", "#f85149")
            return []
    
    def analyze_tile_layout(self, tiles):
        """Check if tiles are in 2x2 grid or horizontal line"""
        if len(tiles) == 0:
            return "none", []
        elif len(tiles) == 1:
            return "single", [tiles[0]['letter']]
        elif len(tiles) >= 2:
            # Sort tiles by x coordinate (left to right)
            sorted_by_x = sorted(tiles, key=lambda t: t['x'])
            
            # Check if tiles form a 2x2 grid
            if len(tiles) == 4:
                sorted_by_y = sorted(tiles, key=lambda t: t['y'])
                top_two = sorted_by_y[:2]
                bottom_two = sorted_by_y[2:]
                
                # Check if arranged in 2x2 grid pattern
                if abs(top_two[0]['y'] - top_two[1]['y']) < 50 and \
                   abs(bottom_two[0]['y'] - bottom_two[1]['y']) < 50:
                    return "grid", []
            
            # Check if tiles are in a horizontal line (same Y coordinate approximately)
            y_coords = [t['y'] for t in tiles]
            avg_y = sum(y_coords) / len(y_coords)
            max_y_diff = max([abs(y - avg_y) for y in y_coords])
            
            # If all tiles have similar Y coordinate (horizontal line)
            if max_y_diff < 50:
                # Extract letters in left-to-right order
                letters = [t['letter'] for t in sorted_by_x]
                return "horizontal", letters
        
        return "other", []
    
    def ocr_scanner_loop(self):
        """Continuous OCR scanning loop"""
        self.log_message("🔍 OCR Scanner started", "#58a6ff")
        
        while self.ocr_active:
            try:
                # Take screenshot
                screenshot = ImageGrab.grab()
                
                # Detect letter tiles
                tiles = self.detect_letter_tiles(screenshot)
                
                if tiles:
                    layout, letters = self.analyze_tile_layout(tiles)
                    
                    if layout in ["single", "horizontal"] and letters:
                        # Single letter or horizontal line detected
                        word = ''.join(letters).lower()
                        self.log_message(f"📸 Detected: {word.upper()}", "#f0883e")
                        
                        # Speak the letters
                        threading.Thread(target=lambda: self.speak(word.upper()), daemon=True).start()
                        
                        # Set buffer and trigger completion
                        self.current_buffer = word
                        self.root.after(0, self.update_buffer_display)
                        
                        time.sleep(0.5)  # Small delay before completion
                        threading.Thread(target=self.trigger_completion, daemon=True).start()
                        
                        time.sleep(2)  # Wait before next scan
                        
                    elif layout == "grid":
                        # 2x2 grid detected - do nothing
                        self.log_message(f"⏭️ Skipping 2x2 grid", "#6e7681")
                        time.sleep(1)
                    
                time.sleep(0.3)  # Scan interval
                
            except Exception as e:
                self.log_message(f"Scanner error: {str(e)}", "#f85149")
                time.sleep(1)
        
        self.log_message("🔍 OCR Scanner stopped", "#f85149")
    
    def toggle_ocr_scanner(self):
        """Toggle OCR scanner on/off"""
        if not self.is_admin:
            messagebox.showwarning(
                "Admin Required",
                "OCR Scanner requires administrator privileges.\nPlease restart the application as administrator."
            )
            return
        
        if not TESSERACT_AVAILABLE:
            messagebox.showerror(
                "Tesseract Not Found",
                "Tesseract OCR not found!\n\n"
                "Installation steps:\n"
                "1. Download: https://github.com/UB-Mannheim/tesseract/wiki\n"
                "2. Install to: C:\\Program Files\\Tesseract-OCR\\\n"
                "3. Make sure 'tesseract.exe' is in that folder\n"
                "4. Restart this application\n\n"
                "Alternative locations checked:\n"
                "• C:\\Program Files\\Tesseract-OCR\\tesseract.exe\n"
                "• C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe"
            )
            return
        
        if not self.is_listening:
            messagebox.showwarning(
                "Not Listening",
                "Please activate 'START LISTENING' first!"
            )
            return
        
        self.ocr_active = not self.ocr_active
        
        if self.ocr_active:
            self.ocr_btn.config(
                text="STOP OCR SCANNER",
                bg='#da3633',
                activebackground='#f85149'
            )
            self.ocr_thread = threading.Thread(target=self.ocr_scanner_loop, daemon=True)
            self.ocr_thread.start()
        else:
            self.ocr_btn.config(
                text="START OCR SCANNER",
                bg='#9e6a03',
                activebackground='#c69026'
            )
            if self.ocr_thread:
                self.ocr_thread = None
        
    def load_words_from_json(self):
        """Load words from words.json file"""
        try:
            if os.path.exists('words.json'):
                with open('words.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, dict) and 'words' in data:
                        words = data['words']
                    elif isinstance(data, list):
                        words = data
                    else:
                        raise ValueError("Invalid JSON format")
                    print(f"Loaded {len(words)} words")
                    return words
            else:
                messagebox.showerror("Error", "words.json not found!")
                self.root.destroy()
                return []
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load: {str(e)}")
            self.root.destroy()
            return []
        
    def build_word_index(self):
        """Build word index"""
        index = defaultdict(list)
        for word in self.words:
            if len(word) >= self.settings['min_word_length']:
                for i in range(1, min(len(word) + 1, 5)):
                    index[word[:i].lower()].append(word.lower())
        return index
        
    def setup_ui(self):
        """Setup modern UI"""
        # Main container
        main_container = tk.Frame(self.root, bg='#0d1117')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = tk.Frame(main_container, bg='#161b22', height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="WORD AUTOFILLER",
            font=("Segoe UI", 24, "bold"),
            bg='#161b22',
            fg='#58a6ff'
        ).pack(pady=20)
        
        # Navigation
        nav = tk.Frame(main_container, bg='#0d1117', height=50)
        nav.pack(fill=tk.X, pady=10)
        
        nav_inner = tk.Frame(nav, bg='#0d1117')
        nav_inner.pack()
        
        self.main_tab_btn = self.create_nav_button(nav_inner, "MAIN", lambda: self.switch_tab("main"), 0)
        self.usage_tab_btn = self.create_nav_button(nav_inner, "USAGE", lambda: self.switch_tab("usage"), 1)
        self.settings_tab_btn = self.create_nav_button(nav_inner, "SETTINGS", lambda: self.switch_tab("settings"), 2)
        self.stats_tab_btn = self.create_nav_button(nav_inner, "STATS", lambda: self.switch_tab("stats"), 3)
        
        # Content area
        content_area = tk.Frame(main_container, bg='#0d1117')
        content_area.pack(fill=tk.BOTH, expand=True)
        
        # Frames
        self.main_frame = tk.Frame(content_area, bg='#0d1117')
        self.usage_frame = tk.Frame(content_area, bg='#0d1117')
        self.settings_frame = tk.Frame(content_area, bg='#0d1117')
        self.stats_frame = tk.Frame(content_area, bg='#0d1117')
        
        self.setup_main_tab()
        self.setup_usage_tab()
        self.setup_settings_tab()
        self.setup_stats_tab()
        
        self.switch_tab("main")
        
        # Footer
        footer = tk.Frame(main_container, bg='#161b22', height=40)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)
        
        tk.Label(
            footer,
            text='Made with ♥️ for the game "Last Letter"',
            font=("Segoe UI", 10, "italic"),
            bg='#161b22',
            fg='#8b949e'
        ).pack(pady=10)
        
    def create_nav_button(self, parent, text, command, col):
        """Create navigation button"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI", 10, "bold"),
            bg='#21262d',
            fg='#c9d1d9',
            activebackground='#30363d',
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            padx=20,
            pady=10
        )
        btn.grid(row=0, column=col, padx=3)
        return btn
        
    def switch_tab(self, tab):
        """Switch tabs"""
        self.current_tab = tab
        self.main_frame.pack_forget()
        self.usage_frame.pack_forget()
        self.settings_frame.pack_forget()
        self.stats_frame.pack_forget()
        
        # Reset button colors
        for btn in [self.main_tab_btn, self.usage_tab_btn, self.settings_tab_btn, self.stats_tab_btn]:
            btn.config(bg='#21262d', fg='#c9d1d9')
        
        if tab == "main":
            self.main_tab_btn.config(bg='#58a6ff', fg='#0d1117')
            self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        elif tab == "usage":
            self.usage_tab_btn.config(bg='#58a6ff', fg='#0d1117')
            self.usage_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        elif tab == "settings":
            self.settings_tab_btn.config(bg='#58a6ff', fg='#0d1117')
            self.settings_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        else:
            self.stats_tab_btn.config(bg='#58a6ff', fg='#0d1117')
            self.stats_frame.pack(fill=tk.BOTH, expand=True, padx=20)
    
    def setup_main_tab(self):
        """Main tab"""
        # Status card
        status_card = self.create_card(self.main_frame, "STATUS")
        
        status_inner = tk.Frame(status_card, bg='#161b22')
        status_inner.pack(pady=10)
        
        self.status_dot = tk.Label(
            status_inner,
            text="●",
            font=("Arial", 20),
            bg='#161b22',
            fg='#f85149'
        )
        self.status_dot.pack(side=tk.LEFT, padx=10)
        
        self.status_text = tk.Label(
            status_inner,
            text="INACTIVE",
            font=("Segoe UI", 14, "bold"),
            bg='#161b22',
            fg='#c9d1d9'
        )
        self.status_text.pack(side=tk.LEFT)
        
        # Buffer card
        buffer_card = self.create_card(self.main_frame, "CURRENT BUFFER")
        
        self.buffer_display = tk.Label(
            buffer_card,
            text="[EMPTY]",
            font=("Consolas", 16, "bold"),
            bg='#161b22',
            fg='#7ee787'
        )
        self.buffer_display.pack(pady=10)
        
        # Controls
        controls = tk.Frame(self.main_frame, bg='#0d1117')
        controls.pack(pady=20)
        
        self.toggle_btn = self.create_button(
            controls,
            "START LISTENING",
            self.toggle_listening,
            '#238636',
            '#2ea043'
        )
        self.toggle_btn.grid(row=0, column=0, padx=10)
        
        self.manual_btn = self.create_button(
            controls,
            "COMPLETE NOW",
            self.manual_complete,
            '#1f6feb',
            '#388bfd'
        )
        self.manual_btn.grid(row=0, column=1, padx=10)
        
        # OCR Scanner button (only show if admin)
        if self.is_admin:
            self.ocr_btn = self.create_button(
                controls,
                "START OCR SCANNER",
                self.toggle_ocr_scanner,
                '#9e6a03',
                '#c69026'
            )
            self.ocr_btn.grid(row=1, column=0, columnspan=2, pady=(10, 0))
            
            # Admin badge
            admin_label = tk.Label(
                controls,
                text="🛡️ ADMIN MODE",
                font=("Segoe UI", 8, "bold"),
                bg='#0d1117',
                fg='#7ee787'
            )
            admin_label.grid(row=2, column=0, columnspan=2, pady=5)
        else:
            # Not admin warning
            warning_label = tk.Label(
                controls,
                text="⚠️ Run as Admin for OCR Scanner",
                font=("Segoe UI", 8, "italic"),
                bg='#0d1117',
                fg='#f0883e'
            )
            warning_label.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Log card
        log_card = self.create_card(self.main_frame, "ACTIVITY LOG")
        
        log_header = tk.Frame(log_card, bg='#161b22')
        log_header.pack(fill=tk.X, padx=10, pady=5)
        
        clear_btn = tk.Button(
            log_header,
            text="CLEAR",
            command=self.clear_log,
            font=("Segoe UI", 9, "bold"),
            bg='#da3633',
            fg='#ffffff',
            activebackground='#f85149',
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            padx=15,
            pady=3
        )
        clear_btn.pack(side=tk.RIGHT)
        
        log_scroll = tk.Scrollbar(log_card)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10))
        
        self.log_text = tk.Text(
            log_card,
            height=10,
            bg='#0d1117',
            fg='#7ee787',
            font=("Consolas", 9),
            yscrollcommand=log_scroll.set,
            state=tk.DISABLED,
            wrap=tk.WORD,
            relief=tk.FLAT,
            borderwidth=0
        )
        self.log_text.pack(padx=10, pady=(0, 10), fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)
    
    def setup_usage_tab(self):
        """Usage guide tab"""
        usage_card = self.create_card(self.usage_frame, "HOW TO USE")
        
        # Create scrollable frame
        canvas = tk.Canvas(usage_card, bg='#161b22', highlightthickness=0)
        scrollbar = tk.Scrollbar(usage_card, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#161b22')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Usage content
        usage_sections = [
            ("🚀 GETTING STARTED", [
                "1. Click the 'START LISTENING' button to activate the autofiller",
                "2. The status indicator will turn green when active",
                "3. Navigate to any text field (Roblox, Notepad, etc.)",
                "4. Start typing your word prefix"
            ]),
            ("⌨️ BASIC OPERATION", [
                "• Type the beginning of a word (e.g., 'po')",
                "• Press the INSERT key on your keyboard",
                "• The app will delete your input and complete the word",
                "• Alternative: Use the 'COMPLETE NOW' button instead of INSERT"
            ]),
            ("🔍 OCR SCANNER (ADMIN ONLY)", [
                "• Requires running the app as Administrator",
                "• Automatically detects letter tiles on screen",
                "• Can detect single letters (e.g., 'Y') or multiple (e.g., 'OL')",
                "• Recognizes letters in horizontal lines",
                "• Ignores 2x2 letter grids (selection screens)",
                "• Speaks detected letters out loud",
                "• Perfect for fast-paced gameplay!",
                "• Works best with clear, white letter tiles"
            ]),
            ("🎯 FOR 'LAST LETTER' GAME", [
                "• Perfect for Roblox's Last Letter typing game",
                "• Type starting letters based on the previous word's last letter",
                "• Press INSERT to quickly complete challenging words",
                "• The app remembers used words and suggests new ones",
                "• OCR mode detects game letters automatically"
            ]),
            ("⚙️ CUSTOMIZATION", [
                "• Visit the SETTINGS tab to adjust timing",
                "• Increase delays if typing doesn't work in some apps",
                "• For games like Roblox: Use 0.3-0.5s delays",
                "• Prefer longer words for higher scores"
            ]),
            ("📊 KEY FEATURES", [
                "✓ Never suggests the same word twice for a prefix",
                "✓ Prefers longer words for better gameplay",
                "✓ Works in most applications and games",
                "✓ Real-time activity logging",
                "✓ Customizable timing for compatibility",
                "✓ OCR screen scanning (Admin mode)",
                "✓ Text-to-speech for detected letters"
            ]),
            ("💡 TIPS & TRICKS", [
                "• Use 2-3 letter prefixes for best results",
                "• The buffer shows what you've typed",
                "• Check the log to see completion history",
                "• Adjust 'Prefer Longer Words' to LONGEST for competitive play",
                "• Set minimum word length to filter short words",
                "• OCR Scanner works best with clear letter tiles"
            ]),
            ("⚠️ TROUBLESHOOTING", [
                "Problem: Words not typing in game",
                "Solution: Increase all timing delays to 0.4-0.5 seconds",
                "",
                "Problem: Backspaces not working",
                "Solution: Try running the app as administrator",
                "",
                "Problem: INSERT key not detected",
                "Solution: Use the 'COMPLETE NOW' button instead",
                "",
                "Problem: OCR not available",
                "Solution: Install Tesseract OCR from official website",
                "",
                "Problem: OCR not detecting letters",
                "Solution: Ensure running as admin and letters are clearly visible"
            ])
        ]
        
        for title, items in usage_sections:
            # Section header
            header_frame = tk.Frame(scrollable_frame, bg='#21262d')
            header_frame.pack(fill=tk.X, pady=(15, 5), padx=15)
            
            tk.Label(
                header_frame,
                text=title,
                font=("Segoe UI", 12, "bold"),
                bg='#21262d',
                fg='#58a6ff',
                anchor="w"
            ).pack(fill=tk.X, padx=10, pady=8)
            
            # Section content
            content_frame = tk.Frame(scrollable_frame, bg='#161b22')
            content_frame.pack(fill=tk.X, padx=15)
            
            for item in items:
                tk.Label(
                    content_frame,
                    text=item,
                    font=("Segoe UI", 10),
                    bg='#161b22',
                    fg='#c9d1d9',
                    anchor="w",
                    justify=tk.LEFT,
                    wraplength=600
                ).pack(fill=tk.X, padx=20, pady=3)
        
        # Pack canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10, padx=(0, 10))
        
    def setup_settings_tab(self):
        """Settings tab"""
        # Timing settings
        timing_card = self.create_card(self.settings_frame, "TIMING SETTINGS")
        
        self.create_slider(timing_card, "Backspace Delay", 'backspace_delay', 0.05, 1.0, 0.05)
        self.create_slider(timing_card, "Typing Delay", 'typing_delay', 0.05, 1.0, 0.05)
        self.create_slider(timing_card, "Start Delay", 'start_delay', 0.1, 2.0, 0.1)
        self.create_slider(timing_card, "After Delete Delay", 'after_delete_delay', 0.1, 2.0, 0.1)
        
        # Word settings
        word_card = self.create_card(self.settings_frame, "WORD PREFERENCES")
        
        self.create_slider(word_card, "Prefer Longer Words", 'prefer_longer_words', 0.0, 1.0, 0.1, 
                          labels=["SHORTEST", "LONGEST"])
        self.create_slider(word_card, "Minimum Word Length", 'min_word_length', 3, 10, 1)
        self.create_slider(word_card, "Max Suggestions", 'max_suggestions_per_prefix', 10, 100, 10)
        
        # Reset button
        reset_frame = tk.Frame(self.settings_frame, bg='#0d1117')
        reset_frame.pack(pady=20)
        
        self.create_button(
            reset_frame,
            "RESET TO DEFAULTS",
            self.reset_settings,
            '#da3633',
            '#f85149'
        ).pack()
        
    def setup_stats_tab(self):
        """Statistics tab"""
        stats_card = self.create_card(self.stats_frame, "STATISTICS")
        
        stats_inner = tk.Frame(stats_card, bg='#161b22')
        stats_inner.pack(pady=20, fill=tk.BOTH, expand=True)
        
        self.stats_labels = {}
        
        ocr_status = "Available ✓" if TESSERACT_AVAILABLE else "Not Installed ✗"
        admin_status = "Yes 🛡️" if self.is_admin else "No"
        
        stats_data = [
            ("Total Words Loaded", len(self.words)),
            ("Unique Prefixes", len(self.word_index)),
            ("Total Completions", 0),
            ("Admin Mode", admin_status),
            ("OCR Available", ocr_status),
            ("Current Session", "Active")
        ]
        
        for i, (label, value) in enumerate(stats_data):
            frame = tk.Frame(stats_inner, bg='#161b22')
            frame.pack(pady=10, fill=tk.X, padx=20)
            
            tk.Label(
                frame,
                text=label,
                font=("Segoe UI", 12),
                bg='#161b22',
                fg='#8b949e'
            ).pack(side=tk.LEFT)
            
            val_label = tk.Label(
                frame,
                text=str(value),
                font=("Segoe UI", 12, "bold"),
                bg='#161b22',
                fg='#58a6ff'
            )
            val_label.pack(side=tk.RIGHT)
            self.stats_labels[label] = val_label
        
    def create_card(self, parent, title):
        """Create a card container"""
        card = tk.Frame(parent, bg='#161b22', relief=tk.FLAT)
        card.pack(pady=10, fill=tk.BOTH, expand=True)
        
        title_frame = tk.Frame(card, bg='#21262d', height=35)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text=title,
            font=("Segoe UI", 10, "bold"),
            bg='#21262d',
            fg='#58a6ff'
        ).pack(side=tk.LEFT, padx=15, pady=8)
        
        return card
        
    def create_button(self, parent, text, command, bg, hover_bg):
        """Create modern button"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI", 11, "bold"),
            bg=bg,
            fg='#ffffff',
            activebackground=hover_bg,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            padx=25,
            pady=12
        )
        return btn
        
    def create_slider(self, parent, label, key, min_val, max_val, res, labels=None):
        """Create setting slider"""
        frame = tk.Frame(parent, bg='#161b22')
        frame.pack(pady=12, padx=20, fill=tk.X)
        
        header = tk.Frame(frame, bg='#161b22')
        header.pack(fill=tk.X)
        
        tk.Label(
            header,
            text=label,
            font=("Segoe UI", 10),
            bg='#161b22',
            fg='#c9d1d9'
        ).pack(side=tk.LEFT)
        
        if labels:
            value_text = labels[0] if self.settings[key] <= 0.5 else labels[1]
        else:
            value_text = f"{self.settings[key]:.2f}"
        
        value_label = tk.Label(
            header,
            text=value_text,
            font=("Consolas", 10, "bold"),
            bg='#161b22',
            fg='#58a6ff'
        )
        value_label.pack(side=tk.RIGHT)
        
        slider = tk.Scale(
            frame,
            from_=min_val,
            to=max_val,
            resolution=res,
            orient=tk.HORIZONTAL,
            bg='#161b22',
            fg='#c9d1d9',
            highlightthickness=0,
            troughcolor='#0d1117',
            activebackground='#58a6ff',
            showvalue=False,
            command=lambda v: self.update_setting(key, v, value_label, labels)
        )
        slider.set(self.settings[key])
        slider.pack(fill=tk.X, pady=5)
        
    def update_setting(self, key, value, label, labels=None):
        """Update setting"""
        self.settings[key] = float(value) if '.' in str(value) else int(value)
        if labels:
            text = labels[0] if float(value) <= 0.5 else labels[1]
        else:
            text = f"{float(value):.2f}" if isinstance(self.settings[key], float) else str(int(value))
        label.config(text=text)
        
        if key == 'min_word_length':
            self.word_index = self.build_word_index()
            self.log_message(f"⚙️ Rebuilt index with min length {int(value)}", "#58a6ff")
    
    def reset_settings(self):
        """Reset settings"""
        self.settings = {
            'backspace_delay': 0.25,
            'typing_delay': 0.15,
            'start_delay': 0.3,
            'after_delete_delay': 0.3,
            'prefer_longer_words': 1.0,
            'min_word_length': 4,
            'max_suggestions_per_prefix': 50
        }
        self.log_message("⚙️ Settings reset", "#f0883e")
        self.switch_tab("settings")
        
    def update_buffer_display(self):
        """Update buffer"""
        if self.current_buffer:
            self.buffer_display.config(text=f'"{self.current_buffer}"', fg='#7ee787')
        else:
            self.buffer_display.config(text="[EMPTY]", fg='#6e7681')
    
    def clear_log(self):
        """Clear log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def log_message(self, msg, color="#7ee787"):
        """Log message"""
        self.log_text.config(state=tk.NORMAL)
        start = self.log_text.index(tk.END)
        self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        end = self.log_text.index(tk.END)
        tag = f"tag_{time.time()}"
        self.log_text.tag_add(tag, start, end)
        self.log_text.tag_config(tag, foreground=color)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def toggle_listening(self):
        """Toggle listening"""
        self.is_listening = not self.is_listening
        
        if self.is_listening:
            self.status_dot.config(fg='#3fb950')
            self.status_text.config(text="ACTIVE")
            self.toggle_btn.config(text="STOP LISTENING", bg='#da3633', activebackground='#f85149')
            self.log_message("✓ Listening activated", "#3fb950")
            self.current_buffer = ""
            self.update_buffer_display()
        else:
            self.status_dot.config(fg='#f85149')
            self.status_text.config(text="INACTIVE")
            self.toggle_btn.config(text="START LISTENING", bg='#238636', activebackground='#2ea043')
            self.log_message("✗ Listening deactivated", "#f85149")
            self.current_buffer = ""
            self.update_buffer_display()
    
    def manual_complete(self):
        """Manual completion"""
        if not self.is_listening:
            self.log_message("⚠ Activate listening first", "#f0883e")
            return
        self.log_message("🖱️ Manual trigger", "#58a6ff")
        threading.Thread(target=self.trigger_completion, daemon=True).start()
    
    def start_keyboard_monitoring(self):
        """Monitor keyboard"""
        def monitor():
            self.log_message("⌨️ Keyboard monitoring started", "#58a6ff")
            
            def on_insert(e):
                if self.is_listening and e.event_type == keyboard.KEY_DOWN:
                    self.log_message("⌨️ INSERT pressed", "#f0883e")
                    threading.Thread(target=self.trigger_completion, daemon=True).start()
            
            keyboard.on_press_key('insert', on_insert)
            try:
                keyboard.add_hotkey('insert', lambda: self.on_insert_hotkey())
            except:
                pass
        
        threading.Thread(target=monitor, daemon=True).start()
    
    def on_insert_hotkey(self):
        """INSERT hotkey"""
        if self.is_listening:
            threading.Thread(target=self.trigger_completion, daemon=True).start()
    
    def start_keyboard_listener(self):
        """Listen to typing"""
        def on_key(e):
            if not self.is_listening:
                return
            
            try:
                if e.event_type == keyboard.KEY_DOWN and len(e.name) == 1 and e.name.isalpha():
                    self.current_buffer += e.name.lower()
                    self.root.after(0, self.update_buffer_display)
                    self.log_message(f"+ '{e.name}' → '{self.current_buffer}'", "#58a6ff")
                elif e.name == 'backspace' and e.event_type == keyboard.KEY_DOWN:
                    if self.current_buffer:
                        self.current_buffer = self.current_buffer[:-1]
                        self.root.after(0, self.update_buffer_display)
                elif e.name in ['space', 'enter'] and e.event_type == keyboard.KEY_DOWN:
                    if self.current_buffer:
                        self.log_message(f"⟲ Buffer reset", "#6e7681")
                    self.current_buffer = ""
                    self.root.after(0, self.update_buffer_display)
            except:
                pass
        
        keyboard.hook(on_key)
        
    def find_completion(self, prefix):
        """Find completion - prefers longer words"""
        if not prefix:
            return None
            
        prefix = prefix.lower()
        possible = set()
        
        for key in self.word_index:
            if key.startswith(prefix):
                possible.update(self.word_index[key])
        
        candidates = [w for w in possible 
                     if w.startswith(prefix) and w != prefix and w not in self.used_words[prefix]]
        
        if not candidates:
            all_possible = [w for w in possible if w.startswith(prefix) and w != prefix]
            if all_possible:
                self.used_words[prefix].clear()
                self.log_message(f"↻ Cycled through all words for '{prefix}'", "#f0883e")
                candidates = all_possible
            else:
                return None
        
        # Sort by length based on preference
        preference = self.settings['prefer_longer_words']
        candidates.sort(key=len, reverse=(preference > 0.5))
        
        # Limit suggestions
        candidates = candidates[:self.settings['max_suggestions_per_prefix']]
        
        if candidates:
            word = candidates[0]
            self.used_words[prefix].add(word)
            return word
        return None
        
    def trigger_completion(self):
        """Trigger completion"""
        with self.completion_lock:
            if not self.current_buffer:
                self.log_message("⚠ No prefix", "#f85149")
                return
                
            prefix = self.current_buffer
            completion = self.find_completion(prefix)
            
            if completion:
                remaining = completion[len(prefix):]
                time.sleep(self.settings['start_delay'])
                
                try:
                    self.log_message(f"← Deleting...", "#f0883e")
                    for i in range(5):
                        keyboard.press_and_release('backspace')
                        time.sleep(self.settings['backspace_delay'])
                    
                    time.sleep(self.settings['after_delete_delay'])
                    
                    self.log_message(f"→ Typing '{remaining}'...", "#58a6ff")
                    for char in remaining:
                        try:
                            keyboard.send(char)
                        except:
                            try:
                                pyautogui.press(char)
                            except:
                                keyboard.write(char)
                        time.sleep(self.settings['typing_delay'])
                    
                    self.log_message(f"✓ '{prefix}' → '{completion}'", "#3fb950")
                    self.current_buffer = ""
                    self.root.after(0, self.update_buffer_display)
                except Exception as e:
                    self.log_message(f"✗ Error: {str(e)}", "#f85149")
            else:
                self.log_message(f"✗ No completion for '{prefix}'", "#f85149")
            
    def run(self):
        """Run app"""
        if self.words:
            self.log_message("═" * 40, "#c9d1d9")
            self.log_message("✓ Application started", "#3fb950")
            self.log_message(f"📚 {len(self.words)} words loaded", "#58a6ff")
            self.log_message("🎯 Press START to begin", "#f0883e")
            self.log_message("═" * 40, "#c9d1d9")
            self.start_keyboard_listener()
            self.root.mainloop()

if __name__ == "__main__":
    app = WordAutofiller()
    app.run()