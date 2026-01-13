"""
SlideSnap - Simple Smart Slide Extractor
Single-instance app with system tray
"""

import cv2
import os
import numpy as np
from skimage.metrics import structural_similarity as ssim
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from PIL import Image, ImageDraw
import pystray
import sys
import atexit
import tempfile
import platform
import subprocess

# Platform-specific imports for file locking
if platform.system() == 'Windows':
    import msvcrt
else:
    import fcntl

# ----------------------------- Config (Hidden from user) -----------------------------
SSIM_THRESHOLD = 0.85      # Lower = more sensitive to changes
BLUR_THRESHOLD = 50        # Lower = accepts slightly blurry frames
MIN_STABLE_SECONDS = 0.8   # Stability duration (balanced for lectures and trailers)
FRAME_SKIP = 2             # Check more frames for accuracy
HIST_THRESHOLD = 0.25      # Histogram distance threshold for layout changes
COOLDOWN_SECONDS = 0.3     # Cooldown after saving a slide
# ------------------------------------------------------------------------------------

# Single instance lock and signal
_lock_file = None
_lock_path = os.path.join(tempfile.gettempdir(), "slidesnap.lock")
_signal_path = os.path.join(tempfile.gettempdir(), "slidesnap.show")

def acquire_lock():
    """Try to acquire single-instance lock (cross-platform)."""
    global _lock_file
    try:
        _lock_file = open(_lock_path, 'w')
        if platform.system() == 'Windows':
            msvcrt.locking(_lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            fcntl.flock(_lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except:
        return False

def release_lock():
    """Release single-instance lock (cross-platform)."""
    global _lock_file
    if _lock_file:
        try:
            if platform.system() == 'Windows':
                msvcrt.locking(_lock_file.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(_lock_file.fileno(), fcntl.LOCK_UN)
            _lock_file.close()
            os.remove(_lock_path)
        except:
            pass
    # Clean up signal file
    try:
        os.remove(_signal_path)
    except:
        pass


def open_folder(path):
    """Open a folder in the system file manager (cross-platform)."""
    if platform.system() == 'Windows':
        os.startfile(path)
    elif platform.system() == 'Darwin':  # macOS
        subprocess.run(['open', path])
    else:  # Linux
        subprocess.run(['xdg-open', path])

def signal_show_window():
    """Signal the running instance to show its window."""
    try:
        with open(_signal_path, 'w') as f:
            f.write('show')
    except:
        pass

# Global tray icon reference
_tray_icon = None
_tray_thread = None


def get_slide_region(frame_gray):
    """Crop frame to slide region, ignoring dynamic areas (speaker, subtitles, edges)."""
    h, w = frame_gray.shape
    top = int(h * 0.05)      # Ignore top 5%
    bottom = int(h * 0.85)   # Ignore bottom 15% (speaker/subtitles)
    left = int(w * 0.05)     # Ignore left 5%
    right = int(w * 0.95)    # Ignore right 5%
    return frame_gray[top:bottom, left:right]


def histogram_diff(a, b):
    """Calculate histogram distance between two grayscale images."""
    hist_a = cv2.calcHist([a], [0], None, [64], [0, 256])
    hist_b = cv2.calcHist([b], [0], None, [64], [0, 256])
    cv2.normalize(hist_a, hist_a)
    cv2.normalize(hist_b, hist_b)
    return cv2.compareHist(hist_a, hist_b, cv2.HISTCMP_BHATTACHARYYA)


def detect_scene_change(ref_frame, curr_frame, ssim_threshold=SSIM_THRESHOLD, hist_threshold=HIST_THRESHOLD):
    """Return True if significant change detected using SSIM + histogram."""
    if ref_frame is None:
        return True

    # Crop to slide region (ignore dynamic areas)
    ref = get_slide_region(ref_frame)
    curr = get_slide_region(curr_frame)

    # Resize for faster comparison
    ref_small = cv2.resize(ref, (320, 240))
    curr_small = cv2.resize(curr, (320, 240))

    # SSIM comparison
    ssim_score, _ = ssim(ref_small, curr_small, full=True)

    # Histogram comparison (catches layout changes SSIM might miss)
    hist_score = histogram_diff(ref_small, curr_small)

    # Scene changed if SSIM is low OR histogram distance is high
    return (ssim_score < ssim_threshold) or (hist_score > hist_threshold)


def is_blurry(frame_gray):
    """Return True if frame is too blurry."""
    return cv2.Laplacian(frame_gray, cv2.CV_64F).var() < BLUR_THRESHOLD


def frames_similar(a, b, ssim_threshold):
    """Check if two frames are similar using SSIM only (no histogram).
    Used for stability confirmation to avoid histogram noise."""
    a_crop = get_slide_region(a)
    b_crop = get_slide_region(b)
    a_small = cv2.resize(a_crop, (320, 240))
    b_small = cv2.resize(b_crop, (320, 240))
    score, _ = ssim(a_small, b_small, full=True)
    return score >= ssim_threshold


def extract_slides(video_path, output_dir, progress_callback=None, ssim_threshold=None):
    """Extract slides using smart detection with reference slide and cooldown."""
    # Use provided threshold or default
    threshold = ssim_threshold if ssim_threshold is not None else SSIM_THRESHOLD

    # Dynamic histogram threshold - scales with SSIM threshold
    # Lower SSIM threshold = more aggressive, so reduce histogram sensitivity
    hist_thresh = 0.15 + (1.0 - threshold) * 0.2

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return {"error": "Cannot open video", "slides": 0}

    fps = cap.get(cv2.CAP_PROP_FPS)
    if fps <= 0:
        fps = 25  # Safe fallback for videos with missing FPS metadata
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    min_stable_frames = int(MIN_STABLE_SECONDS * fps / FRAME_SKIP)
    cooldown_frames = int(COOLDOWN_SECONDS * fps / FRAME_SKIP)  # Post-capture cooldown

    os.makedirs(output_dir, exist_ok=True)

    # Reference slide stores full grayscale frame
    # All comparisons use get_slide_region() to crop before comparing
    reference_slide = None
    candidate_frame = None
    candidate_frame_color = None
    stable_count = 0
    slide_count = 0
    frame_idx = 0
    cooldown = 0               # Cooldown counter after saving

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        frame_idx += 1

        if frame_idx % FRAME_SKIP != 0:
            continue

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        # First frame - save if not blurry
        if reference_slide is None:
            if progress_callback:
                progress = (frame_idx / total_frames) * 100
                progress_callback(progress, f"Analyzing... {slide_count} slides found")
            # Skip blurry first frames (e.g., mid-fade transitions)
            if is_blurry(gray):
                continue
            slide_count += 1
            cv2.imwrite(os.path.join(output_dir, f"slide_{slide_count:03d}.png"), frame)
            reference_slide = gray.copy()
            continue

        # Skip detection during cooldown period
        if cooldown > 0:
            cooldown -= 1
            if progress_callback:
                progress = (frame_idx / total_frames) * 100
                progress_callback(progress, f"Stabilizing... {slide_count} slides found")
            continue

        if progress_callback:
            progress = (frame_idx / total_frames) * 100
            progress_callback(progress, f"Analyzing... {slide_count} slides found")

        # Reject blurry frames early - before any detection logic
        if is_blurry(gray):
            stable_count = 0
            candidate_frame = None
            continue

        # Check for scene change against reference slide (not last saved)
        if detect_scene_change(reference_slide, gray, threshold, hist_thresh):
            if candidate_frame is None:
                candidate_frame = gray.copy()
                candidate_frame_color = frame.copy()
                stable_count = 1
            else:
                # Check if candidate is stable using SSIM only (no histogram noise)
                if frames_similar(candidate_frame, gray, threshold):
                    stable_count += 1
                else:
                    # New candidate detected
                    candidate_frame = gray.copy()
                    candidate_frame_color = frame.copy()
                    stable_count = 1

            if stable_count >= min_stable_frames:
                slide_count += 1
                cv2.imwrite(os.path.join(output_dir, f"slide_{slide_count:03d}.png"), candidate_frame_color)
                reference_slide = candidate_frame.copy()  # Update reference
                candidate_frame = None
                stable_count = 0
                cooldown = cooldown_frames  # Start cooldown
        else:
            stable_count = 0
            candidate_frame = None

    cap.release()

    if progress_callback:
        progress_callback(100, f"Done! {slide_count} slides extracted")

    return {"slides": slide_count, "output_dir": output_dir}


def get_resource_path(filename):
    """Get path to bundled resource file."""
    if getattr(sys, 'frozen', False):
        # Running as bundled exe
        base_path = sys._MEIPASS
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)


def create_tray_icon_image():
    """Load custom tray icon image (cross-platform)."""
    # Try PNG first for macOS/Linux, then ICO for Windows
    icon_files = ["camera.png", "slidesnap.ico"] if platform.system() != 'Windows' else ["slidesnap.ico", "camera.png"]

    for icon_name in icon_files:
        try:
            icon_path = get_resource_path(icon_name)
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                # macOS menu bar uses smaller icons (22x22 is standard)
                # Windows uses 64x64
                size = 22 if platform.system() == 'Darwin' else 64
                img = img.resize((size, size), Image.Resampling.LANCZOS)
                return img
        except:
            continue

    # Fallback to simple generated icon
    size = 22 if platform.system() == 'Darwin' else 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = size // 8
    draw.rectangle([margin, margin, size - margin, size - margin], fill='#AF29F5', outline='#8B1FBF', width=1)
    # Play triangle
    draw.polygon([(size//3, size//4), (size//3, size*3//4), (size*3//4, size//2)], fill='white')
    return img


def cleanup_tray():
    """Clean up global tray icon."""
    global _tray_icon
    if _tray_icon:
        try:
            _tray_icon.stop()
        except:
            pass
        _tray_icon = None


# Register cleanup at exit
atexit.register(cleanup_tray)


class SlideSnapApp:
    # Modern color scheme (pink/purple theme matching icon)
    COLORS = {
        'bg': '#1a1025',            # Dark purple background
        'card': '#2d1f3d',          # Card background
        'primary': '#AF29F5',       # Purple button
        'primary_hover': '#c55bff', # Lighter purple on hover
        'secondary': '#AF29F5',     # Purple
        'text': '#ffffff',          # White text
        'text_muted': '#a89bb5',    # Muted purple-gray text
        'success': '#10b981',       # Green for success
        'border': '#3d2a54',        # Border color
        'button_disabled': '#2d1f3d', # Disabled button
        'accent': '#ff6ac1'         # Bright pink for title
    }

    def __init__(self):
        global _tray_icon, _tray_thread

        self.root = tk.Tk()
        self.root.title("SlideSnap")
        self.root.geometry("500x620")
        self.root.resizable(True, True)  # Allow resize and maximize
        self.root.minsize(450, 550)  # Minimum size
        self.root.configure(bg=self.COLORS['bg'])

        # Set window icon (platform-specific)
        try:
            if platform.system() == 'Windows':
                icon_path = get_resource_path("slidesnap.ico")
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
            else:
                # macOS/Linux - use PNG for window icon
                icon_path = get_resource_path("camera.png")
                if os.path.exists(icon_path):
                    icon_img = tk.PhotoImage(file=icon_path)
                    self.root.iconphoto(True, icon_img)
                    self._icon_img = icon_img  # Keep reference
        except:
            pass

        # Center window on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 500) // 2
        y = (self.root.winfo_screenheight() - 620) // 2
        self.root.geometry(f"500x620+{x}+{y}")

        self.video_path = None
        self.output_path = None
        self.is_hidden = False
        self.similarity_threshold = tk.StringVar(value="85")  # Default 85%

        self.setup_styles()
        self.create_ui()

        # Create tray icon once at startup
        if _tray_icon is None:
            self.setup_tray()

        # Handle window close - quit the app (tray acts as launcher only)
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app)

    def setup_styles(self):
        """Configure ttk styles for modern look."""
        style = ttk.Style()
        style.theme_use('clam')

        # Style the dropdown listbox popup
        self.root.option_add('*TCombobox*Listbox.background', self.COLORS['card'])
        self.root.option_add('*TCombobox*Listbox.foreground', self.COLORS['text'])
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.COLORS['primary'])
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.COLORS['text'])

        # Progress bar style
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=self.COLORS['card'],
            background=self.COLORS['primary'],
            bordercolor=self.COLORS['border'],
            lightcolor=self.COLORS['primary'],
            darkcolor=self.COLORS['secondary']
        )

        # Combobox style
        style.configure(
            "Custom.TCombobox",
            fieldbackground=self.COLORS['border'],
            background=self.COLORS['border'],
            foreground=self.COLORS['text'],
            arrowcolor=self.COLORS['text'],
            bordercolor=self.COLORS['primary'],
            lightcolor=self.COLORS['border'],
            darkcolor=self.COLORS['border']
        )
        style.map("Custom.TCombobox",
            fieldbackground=[('readonly', self.COLORS['border']), ('active', self.COLORS['border'])],
            selectbackground=[('readonly', self.COLORS['border'])],
            selectforeground=[('readonly', self.COLORS['text'])],
            background=[('active', self.COLORS['border']), ('pressed', self.COLORS['border'])],
            arrowcolor=[('active', self.COLORS['text']), ('pressed', self.COLORS['text'])]
        )

    def setup_tray(self):
        """Create system tray icon once (with fallback for Linux without tray support)."""
        global _tray_icon, _tray_thread

        try:
            icon_image = create_tray_icon_image()

            menu = pystray.Menu(
                pystray.MenuItem("Show SlideSnap", self.show_window, default=True),
                pystray.MenuItem("Exit", self.quit_app)
            )

            _tray_icon = pystray.Icon("SlideSnap", icon_image, "SlideSnap - Click to show", menu)

            # Run tray icon in separate thread
            _tray_thread = threading.Thread(target=_tray_icon.run, daemon=True)
            _tray_thread.start()
        except Exception:
            # Tray not supported (some Linux desktops) - app still works
            _tray_icon = None

    def create_styled_button(self, parent, text, command, primary=False):
        """Create a modern styled button with rounded look."""
        if primary:
            bg_color = self.COLORS['secondary']  # Purple
            hover_color = self.COLORS['primary']  # Pink on hover
        else:
            bg_color = self.COLORS['border']  # Dark purple
            hover_color = '#4d3a64'  # Lighter purple on hover

        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI", 10, "bold" if primary else "normal"),
            fg=self.COLORS['text'],
            bg=bg_color,
            activeforeground=self.COLORS['text'],
            activebackground=hover_color,
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=20,
            pady=10,
            highlightthickness=0
        )

        # Hover effects
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_color))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg_color))

        # Store colors for later
        btn.bg_color = bg_color
        btn.hover_color = hover_color

        return btn

    def create_ui(self):
        # Outer container that fills the window
        outer_frame = tk.Frame(self.root, bg=self.COLORS['bg'])
        outer_frame.pack(fill="both", expand=True)

        # Centered container with max width
        container = tk.Frame(outer_frame, bg=self.COLORS['bg'])
        container.place(relx=0.5, rely=0.5, anchor="center")

        # Main content frame with max width
        main_frame = tk.Frame(container, bg=self.COLORS['bg'], width=480)
        main_frame.pack(padx=30, pady=30)
        main_frame.pack_propagate(False)
        main_frame.config(width=480, height=580)

        # Header section
        header_frame = tk.Frame(main_frame, bg=self.COLORS['bg'])
        header_frame.pack(fill="x", pady=(0, 25))

        # Title with cyan accent
        title = tk.Label(
            header_frame,
            text="SlideSnap",
            font=("Segoe UI", 32, "bold"),
            fg=self.COLORS['accent'],
            bg=self.COLORS['bg']
        )
        title.pack()

        subtitle = tk.Label(
            header_frame,
            text="Extract slides from videos automatically",
            font=("Segoe UI", 10),
            fg=self.COLORS['text_muted'],
            bg=self.COLORS['bg']
        )
        subtitle.pack(pady=(5, 0))

        # Card container with dashed border effect
        card = tk.Frame(main_frame, bg=self.COLORS['card'], padx=25, pady=20,
                       highlightbackground=self.COLORS['primary'], highlightthickness=1)
        card.pack(fill="x", pady=10)

        # Video Selection Field
        video_frame = tk.Frame(card, bg=self.COLORS['card'])
        video_frame.pack(fill="x", pady=(0, 12))

        # Title row with clear button
        video_title_frame = tk.Frame(video_frame, bg=self.COLORS['card'])
        video_title_frame.pack(fill="x", pady=(0, 5))

        video_label = tk.Label(
            video_title_frame,
            text="Video File",
            font=("Segoe UI", 10),
            fg=self.COLORS['text_muted'],
            bg=self.COLORS['card']
        )
        video_label.pack(side="left")

        self.video_clear_btn = tk.Button(
            video_title_frame,
            text="✕",
            command=self.clear_video,
            font=("Segoe UI", 8, "bold"),
            fg='#ffffff',
            bg='#e53935',
            activeforeground='#ffffff',
            activebackground='#c62828',
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=6,
            pady=1
        )
        # Hidden initially
        self.video_clear_btn.bind("<Enter>", lambda e: self.video_clear_btn.config(bg='#c62828'))
        self.video_clear_btn.bind("<Leave>", lambda e: self.video_clear_btn.config(bg='#e53935'))

        self.video_btn = self.create_styled_button(
            video_frame,
            "Click to select video...",
            self.select_video
        )
        self.video_btn.pack(fill="x", ipady=3)

        # Output Folder Field
        output_frame = tk.Frame(card, bg=self.COLORS['card'])
        output_frame.pack(fill="x", pady=(0, 12))

        # Title row with clear button
        output_title_frame = tk.Frame(output_frame, bg=self.COLORS['card'])
        output_title_frame.pack(fill="x", pady=(0, 5))

        output_label = tk.Label(
            output_title_frame,
            text="Output Folder",
            font=("Segoe UI", 10),
            fg=self.COLORS['text_muted'],
            bg=self.COLORS['card']
        )
        output_label.pack(side="left")

        self.output_clear_btn = tk.Button(
            output_title_frame,
            text="✕",
            command=self.clear_output,
            font=("Segoe UI", 8, "bold"),
            fg='#ffffff',
            bg='#e53935',
            activeforeground='#ffffff',
            activebackground='#c62828',
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=6,
            pady=1
        )
        # Hidden initially
        self.output_clear_btn.bind("<Enter>", lambda e: self.output_clear_btn.config(bg='#c62828'))
        self.output_clear_btn.bind("<Leave>", lambda e: self.output_clear_btn.config(bg='#e53935'))

        self.output_btn = self.create_styled_button(
            output_frame,
            "Click to select folder...",
            self.select_output
        )
        self.output_btn.pack(fill="x", ipady=3)

        # Sensitivity Section
        sensitivity_frame = tk.Frame(card, bg=self.COLORS['card'])
        sensitivity_frame.pack(fill="x", pady=(5, 0))

        sens_label = tk.Label(
            sensitivity_frame,
            text="Similarity Threshold (higher value = more slides captured)",
            font=("Segoe UI", 10),
            fg=self.COLORS['text_muted'],
            bg=self.COLORS['card']
        )
        sens_label.pack(anchor="w", pady=(0, 5))

        # Dropdown full width with % sign
        dropdown_frame = tk.Frame(sensitivity_frame, bg=self.COLORS['border'])
        dropdown_frame.pack(fill="x")

        self.threshold_combo = ttk.Combobox(
            dropdown_frame,
            textvariable=self.similarity_threshold,
            values=["25", "50", "75", "85", "90", "95", "100"],
            state="readonly",
            font=("Segoe UI", 10),
            style="Custom.TCombobox"
        )
        self.threshold_combo.pack(side="left", fill="x", expand=True)
        self.threshold_combo.set("85")

        percent_label = tk.Label(
            dropdown_frame,
            text=" % ",
            font=("Segoe UI", 10, "bold"),
            fg=self.COLORS['text'],
            bg=self.COLORS['border']
        )
        percent_label.pack(side="right")

        # Extract Button
        self.extract_btn = self.create_styled_button(
            main_frame,
            "EXTRACT SLIDES",
            self.start_extraction,
            primary=True
        )
        self.extract_btn.pack(fill="x", pady=20, ipady=8)
        # Start disabled
        self.extract_btn.config(state="disabled", bg=self.COLORS['border'], cursor="arrow")
        self.extract_btn.unbind("<Enter>")
        self.extract_btn.unbind("<Leave>")

        # Progress section (hidden initially)
        self.progress_frame = tk.Frame(main_frame, bg=self.COLORS['bg'])
        # Don't pack yet - will show when extraction starts

        self.progress = ttk.Progressbar(
            self.progress_frame,
            length=440,
            mode='determinate',
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x")

        self.status = tk.Label(
            self.progress_frame,
            text="",
            font=("Segoe UI", 10),
            fg=self.COLORS['text_muted'],
            bg=self.COLORS['bg']
        )
        self.status.pack(pady=(10, 0))

    def select_video(self):
        filename = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=[("Video Files", "*.mp4 *.avi *.mov *.mkv *.webm")]
        )
        if filename:
            self.video_path = filename
            name = os.path.basename(filename)
            short = name[:35] + "..." if len(name) > 35 else name
            self.video_btn.config(
                text=short,
                bg=self.COLORS['success'],
                fg=self.COLORS['bg']
            )
            # Update hover colors for selected state
            self.video_btn.bind("<Enter>", lambda e: self.video_btn.config(bg='#34d399'))
            self.video_btn.bind("<Leave>", lambda e: self.video_btn.config(bg=self.COLORS['success']))
            # Show clear button on right of title
            self.video_clear_btn.pack(side="right")
            self.check_ready()

    def clear_video(self):
        """Clear video selection."""
        self.video_path = None
        self.video_btn.config(
            text="Click to select video...",
            bg=self.COLORS['border'],
            fg=self.COLORS['text']
        )
        self.video_btn.bind("<Enter>", lambda e: self.video_btn.config(bg='#4d3a64'))
        self.video_btn.bind("<Leave>", lambda e: self.video_btn.config(bg=self.COLORS['border']))
        # Hide clear button
        self.video_clear_btn.pack_forget()
        self.check_ready()

    def select_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_path = folder
            short = "..." + folder[-30:] if len(folder) > 30 else folder
            self.output_btn.config(
                text=short,
                bg=self.COLORS['success'],
                fg=self.COLORS['bg']
            )
            # Update hover colors for selected state
            self.output_btn.bind("<Enter>", lambda e: self.output_btn.config(bg='#34d399'))
            self.output_btn.bind("<Leave>", lambda e: self.output_btn.config(bg=self.COLORS['success']))
            # Show clear button on right of title
            self.output_clear_btn.pack(side="right")
            self.check_ready()

    def clear_output(self):
        """Clear output folder selection."""
        self.output_path = None
        self.output_btn.config(
            text="Click to select folder...",
            bg=self.COLORS['border'],
            fg=self.COLORS['text']
        )
        self.output_btn.bind("<Enter>", lambda e: self.output_btn.config(bg='#4d3a64'))
        self.output_btn.bind("<Leave>", lambda e: self.output_btn.config(bg=self.COLORS['border']))
        # Hide clear button
        self.output_clear_btn.pack_forget()
        self.check_ready()

    def check_ready(self):
        if self.video_path and self.output_path:
            self.extract_btn.config(
                state="normal",
                bg=self.COLORS['primary'],
                cursor="hand2"
            )
            # Restore hover effects for extract button
            self.extract_btn.bind("<Enter>", lambda e: self.extract_btn.config(bg=self.COLORS['primary_hover']))
            self.extract_btn.bind("<Leave>", lambda e: self.extract_btn.config(bg=self.COLORS['primary']))
            self.status.config(text="Ready! Click EXTRACT SLIDES", fg=self.COLORS['success'])
        else:
            self.extract_btn.config(
                state="disabled",
                bg=self.COLORS['border'],
                cursor="arrow"
            )
            self.extract_btn.unbind("<Enter>")
            self.extract_btn.unbind("<Leave>")
            self.status.config(text="")

    def update_progress(self, progress, message):
        self.progress['value'] = progress
        self.status.config(text=message, fg=self.COLORS['text_muted'])
        self.root.update_idletasks()

    def start_extraction(self):
        self.extract_btn.config(
            state="disabled",
            text="EXTRACTING...",
            bg=self.COLORS['secondary'],
            fg='#ffffff',
            cursor="wait"
        )
        self.extract_btn.unbind("<Enter>")
        self.extract_btn.unbind("<Leave>")

        # Show progress bar
        self.progress_frame.pack(fill="x", pady=(10, 0))
        self.progress['value'] = 0
        self.status.config(text="Starting...", fg=self.COLORS['text_muted'])

        thread = threading.Thread(target=self.run_extraction)
        thread.daemon = True
        thread.start()

    def run_extraction(self):
        video_name = os.path.splitext(os.path.basename(self.video_path))[0]
        output_dir = os.path.join(self.output_path, f"{video_name}-slides")

        # Get threshold from dropdown (convert percentage to decimal)
        threshold = int(self.similarity_threshold.get()) / 100.0

        try:
            result = extract_slides(
                self.video_path,
                output_dir,
                progress_callback=lambda p, m: self.root.after(0, lambda: self.update_progress(p, m)),
                ssim_threshold=threshold
            )

            def show_result():
                self.extract_btn.config(
                    state="normal",
                    text="EXTRACT SLIDES",
                    bg=self.COLORS['primary'],
                    cursor="hand2"
                )
                # Restore hover effects
                self.extract_btn.bind("<Enter>", lambda e: self.extract_btn.config(bg=self.COLORS['primary_hover']))
                self.extract_btn.bind("<Leave>", lambda e: self.extract_btn.config(bg=self.COLORS['primary']))

                if result.get("slides", 0) > 0:
                    self.progress['value'] = 100
                    self.status.config(
                        text=f"Done! {result['slides']} slides extracted",
                        fg=self.COLORS['success']
                    )
                    messagebox.showinfo(
                        "Success!",
                        f"Extracted {result['slides']} slides!\n\nSaved to:\n{output_dir}"
                    )
                    open_folder(output_dir)
                    # Hide progress bar after a delay
                    self.root.after(3000, lambda: self.progress_frame.pack_forget())
                else:
                    self.progress_frame.pack_forget()  # Hide immediately
                    messagebox.showinfo("Done", "No distinct slides found in this video.")

            self.root.after(0, show_result)

        except Exception as e:
            def show_error():
                self.extract_btn.config(
                    state="normal",
                    text="EXTRACT SLIDES",
                    bg=self.COLORS['primary'],
                    cursor="hand2"
                )
                self.extract_btn.bind("<Enter>", lambda e: self.extract_btn.config(bg=self.COLORS['primary_hover']))
                self.extract_btn.bind("<Leave>", lambda e: self.extract_btn.config(bg=self.COLORS['primary']))
                self.progress_frame.pack_forget()  # Hide progress bar
                messagebox.showerror("Error", str(e))

            self.root.after(0, show_error)

    def on_minimize(self, event):
        """Handle window minimize - hide to tray."""
        if self.root.state() == 'iconic':
            self.hide_to_tray()

    def hide_to_tray(self):
        """Hide window to system tray."""
        self.root.withdraw()
        self.is_hidden = True

    def show_window(self, icon=None, item=None):
        """Show the main window from tray."""
        self.root.after(0, self._restore_window)

    def _restore_window(self):
        """Restore window on main thread."""
        self.root.deiconify()
        self.root.state('normal')
        self.root.lift()
        self.root.focus_force()
        self.is_hidden = False

    def quit_app(self, icon=None, item=None):
        """Quit the application completely."""
        cleanup_tray()
        self.root.after(0, self._do_quit)

    def _do_quit(self):
        """Actually quit on main thread."""
        self.root.destroy()
        os._exit(0)

    def check_show_signal(self):
        """Check if another instance wants us to show the window."""
        try:
            if os.path.exists(_signal_path):
                os.remove(_signal_path)
                self._restore_window()
        except:
            pass
        # Check again in 500ms
        self.root.after(500, self.check_show_signal)

    def run(self):
        # Start checking for show signals
        self.root.after(500, self.check_show_signal)
        self.root.mainloop()


if __name__ == "__main__":
    # Check for single instance
    if not acquire_lock():
        # Another instance is already running - signal it to show window
        signal_show_window()
        sys.exit(0)

    # Register cleanup
    atexit.register(release_lock)

    app = SlideSnapApp()
    app.run()
