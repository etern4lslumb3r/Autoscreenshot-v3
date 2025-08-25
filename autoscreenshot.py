import cv2
import numpy as np
import pyautogui
import time
import io
import os.path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import queue # New: Import the queue module
import re # Import re for regular expressions
import collections # New: Import collections for deque

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

from PIL import Image, ImageTk
try:
    import win32clipboard
except ImportError:
    win32clipboard = None

# --- User Definable Bounding Box ---
# Set these to None for full-screen capture, or specify coordinates for a region.
BBOX_LEFT = None    # Example: 100
BBOX_TOP = None     # Example: 100
BBOX_WIDTH = None   # Example: 800
BBOX_HEIGHT = None  # Example: 600

# If you uncommented the BBOX variables above, uncomment the line below:
# user_defined_bbox = (BBOX_LEFT, BBOX_TOP, BBOX_WIDTH, BBOX_HEIGHT) if all([BBOX_LEFT, BBOX_TOP, BBOX_WIDTH, BBOX_HEIGHT]) else None
user_defined_bbox = None # Set to None for full screen capture by default

# Google Docs API imports would be required for real use, but omitted here

RECENT_CHANGES_WINDOW_SIZE = 20 # Number of recent local change fractions to keep for mean calculation

def screenshot_to_clipboard(pil_img):
    """
    Copy a PIL image to the clipboard (Windows only).
    """
    if win32clipboard is None:
        print("win32clipboard not available. Clipboard copy not supported on this platform.")
        return
    output = io.BytesIO()
    pil_img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]  # BMP header fix
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def take_screenshot(bbox=None):
    """
    Take a screenshot of the screen or a region (bbox: (left, top, width, height)).
    Returns a PIL Image.
    """
    if bbox and all(bbox): # Check if all bbox components are not None
        x, y, w, h = bbox
        screenshot = pyautogui.screenshot(region=(x, y, w, h))
    else:
        screenshot = pyautogui.screenshot()
    return screenshot

def images_significantly_different(img1, img2, threshold=0.05, pixel_diff_threshold=25):
    """
    Compare two PIL images. Return True if they are significantly different.
    threshold: fraction of pixels that must change to trigger.
    pixel_diff_threshold: minimum absolute pixel value difference to consider a pixel changed.
    """
    arr1 = np.array(img1.resize((min(img1.width, img2.width), min(img1.height, img2.height))))
    arr2 = np.array(img2.resize((min(img1.width, img2.width), min(img1.height, img2.height))))
    if arr1.shape != arr2.shape:
        return True
    diff = np.abs(arr1.astype(np.int16) - arr2.astype(np.int16))
    changed = np.any(diff > pixel_diff_threshold, axis=2)  # Use dynamic pixel_diff_threshold
    frac_changed = np.sum(changed) / changed.size
    return (frac_changed > threshold, frac_changed) # Return both significance and raw fraction

def _take_and_clipboard_screenshot(bbox=None, stop_event=None):
    """
    Takes a screenshot and, if not in adaptive mode, copies it to the clipboard.
    Returns a PIL Image or None.
    """
    # adaptive_mode is no longer a parameter, so it's effectively always False
    # The previous adaptive logic is also removed, so this just takes a screenshot
    curr_img = take_screenshot(bbox)
    return curr_img # Directly return the screenshot, clipboard copy is handled by _monitor_loop if needed

def add_screenshot_to_google_doc(docs_service, drive_service, doc_id, pil_img):
    """
    Add the screenshot (PIL Image) to a Google Doc.
    'docs_service' is an authenticated Google Docs API service.
    'drive_service' is an authenticated Google Drive API service.
    'doc_id' is the ID of the Google Doc.
    """
    print("Uploads are currently disabled. Skipping Google Drive and Google Docs operations.")
    # If you want to re-enable uploads, uncomment the following code block:

    # # Convert PIL image to bytes (PNG)
    # img_byte_arr = io.BytesIO()
    # pil_img.save(img_byte_arr, format='PNG')
    # img_bytes = img_byte_arr.getvalue()
    # img_byte_arr.seek(0) # Reset stream position for reading

    # # Upload image to Google Drive
    # file_metadata = {
    #     'name': 'screenshot.png', 
    #     'mimeType': 'image/png',
    #     'parents': [folder_id] # This 'folder_id' would need to be passed to this function
    # }
    # media = MediaIoBaseUpload(img_byte_arr, mimetype='image/png', resumable=True)
    # file = drive_service.files().create(body=file_metadata, media_body=media, fields='id,webContentLink').execute()
    # image_url = file.get('webContentLink') # Use webContentLink for direct access
    # image_id = file.get('id') # Get the ID of the uploaded file

    # # Make the image publicly accessible
    # drive_service.permissions().create(
    #     fileId=image_id,
    #     body={'type': 'anyone', 'role': 'reader'},
    #     fields='id'
    # ).execute()

    # print(f"Image uploaded to Google Drive: {image_url}")

    # # Insert image into Google Doc at the end of the document
    # requests = [
    #     {
    #         'insertInlineImage': {
    #             'endOfSegmentLocation': {}, # Use endOfSegmentLocation to append to the end
    #             'uri': image_url,
    #             'objectSize': {'height': {'magnitude': 300, 'unit': 'PT'}}
    #         }
    #     }
    # ]
    # docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    # print("Screenshot added to Google Doc successfully!")

    # # Delete the uploaded image from Google Drive
    # drive_service.files().delete(fileId=image_id).execute()
    # print("Screenshot deleted from Google Drive successfully!")


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/documents"]

def authenticate_google_api():
    """
    Authenticates with Google and returns the Docs and Drive service objects.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    credentials_path = os.path.join(current_script_dir, "credentials.json")
    token_path = os.path.join(current_script_dir, "token.json")

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_path, "w") as token:
            token.write(creds.to_json())

    try:
        docs_service = build("docs", "v1", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)
        return docs_service, drive_service
    except HttpError as error:
        print(f"An error occurred during authentication: {error}")
        return None, None

def _get_or_create_autoscreenshot_folder(drive_service):
    """
    Checks if 'autoscreenshot' folder exists in Google Drive. If not, creates it.
    Returns the folder ID.
    """
    folder_name = "autoscreenshot"
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    
    try:
        # Search for the folder
        results = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])

        if items:
            print(f"Found existing '{folder_name}' folder with ID: {items[0]['id']}")
            return items[0]['id']
        else:
            # Create the folder if it doesn't exist
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            folder = drive_service.files().create(body=file_metadata, fields='id').execute()
            print(f"Created new '{folder_name}' folder with ID: {folder.get('id')}")
            return folder.get('id')
    except HttpError as error:
        print(f"An error occurred while getting/creating 'autoscreenshot' folder: {error}")
        return None

def extract_doc_id_from_url(url):
    """
    Extracts the Google Doc ID from a given URL.
    """
    match = re.search(r'/document/d/([a-zA-Z0-9_-]+)', url)
    if match:
        return match.group(1)
    return url # Return original string if no match (it might be just the ID)


class ScreenshotApp:
    def __init__(self, master):
        self.master = master
        master.title("Auto Screenshot to Google Docs")

        self.monitoring = False
        self.monitor_thread = None
        self.processor_thread = None # New: Thread for processing screenshots
        self.screenshot_queue = queue.Queue() # New: Queue for screenshots
        self.screenshot_counter = 0 # New: To generate unique IDs for screenshots
        self.upload_statuses = {} # New: To track status and progress of each upload
        self.treeview_items = {} # New: To map screenshot_id to treeview iid
        self.autoscreenshot_folder_id = None # New: To store the folder ID
        self.stop_event = threading.Event() # New: Event to signal threads to stop
        self.exit_app_on_stop = False # New: Flag to indicate if the app should exit after stopping
        self.previous_local_img = None # New: Stores the last screenshot for local comparison
        self.previous_uploaded_img = None # New: Stores the last screenshot uploaded to Drive/Docs
        self.recent_local_frac_changes = collections.deque(maxlen=RECENT_CHANGES_WINDOW_SIZE) # New: Deque for recent local change fractions

        # --- Google Docs ID ---
        ttk.Label(master, text="Google Doc ID:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.doc_id_entry = ttk.Entry(master, width=50)
        self.doc_id_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.doc_id_entry.insert(0, extract_doc_id_from_url("11CdinqkOFzIbSx18oxN4WHvadFsKnR0IJ0G_x5f7pfc")) # Default value
        self.doc_id_entry.bind("<Control-v>", self._on_paste_doc_id)

        # --- Bounding Box Inputs ---
        bbox_frame = ttk.LabelFrame(master, text="Screenshot Region (Leave empty for full screen)")
        bbox_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        ttk.Label(bbox_frame, text="Left (x):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.bbox_left = ttk.Entry(bbox_frame, width=10)
        self.bbox_left.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(bbox_frame, text="Top (y):").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.bbox_top = ttk.Entry(bbox_frame, width=10)
        self.bbox_top.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(bbox_frame, text="Width:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.bbox_width = ttk.Entry(bbox_frame, width=10)
        self.bbox_width.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(bbox_frame, text="Height:").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.bbox_height = ttk.Entry(bbox_frame, width=10)
        self.bbox_height.grid(row=1, column=3, padx=5, pady=2)
        
        self.select_region_button = ttk.Button(bbox_frame, text="Select Region on Screen", command=self.select_region_on_screen)
        self.select_region_button.grid(row=2, column=0, columnspan=4, sticky="ew", padx=5, pady=5)

        # --- Interval Input ---
        ttk.Label(master, text="Check Interval (seconds):").grid(row=2, column=0, sticky="w", padx=5, pady=5) # Changed row
        self.interval_entry = ttk.Entry(master, width=10)
        self.interval_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5) # Changed row
        self.interval_entry.insert(0, "1") # Default value changed to 1 second

        # --- Advanced Section (New) ---
        advanced_frame = ttk.LabelFrame(master, text="Advanced Settings")
        advanced_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5, pady=5) # Changed row

        ttk.Label(advanced_frame, text="Last Local Change (Fraction):").grid(row=0, column=0, sticky="w", padx=5, pady=2) # Changed text
        self.last_local_change_var = tk.StringVar(value="N/A")
        self.last_local_change_label = ttk.Label(advanced_frame, textvariable=self.last_local_change_var)
        self.last_local_change_label.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(advanced_frame, text="Mean Deviation Threshold * :").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.mean_deviation_threshold_var = tk.DoubleVar(value=0.95) # Default changed to 0.95
        self.mean_deviation_threshold_slider = ttk.Scale(advanced_frame, from_=0.0, to=2.0, orient="horizontal", variable=self.mean_deviation_threshold_var, command=self._update_mean_deviation_label)
        self.mean_deviation_threshold_slider.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        self.mean_deviation_threshold_label = ttk.Label(advanced_frame, text=f"{self.mean_deviation_threshold_var.get():.2f}")
        self.mean_deviation_threshold_label.grid(row=1, column=2, sticky="w", padx=0, pady=2)

        self.mean_deviation_threshold_text_label = ttk.Label(advanced_frame, text="(e.g., 0.5 for 50% above mean)", font=("TkDefaultFont", 8))
        self.mean_deviation_threshold_text_label.grid(row=1, column=3, columnspan=2, sticky="w", padx=5, pady=0)

        self.mean_deviation_threshold_text_label.bind("<Enter>", self._show_mean_deviation_tooltip)
        self.mean_deviation_threshold_text_label.bind("<Leave>", self._hide_mean_deviation_tooltip)

        ttk.Label(advanced_frame, text="Current Mean Fraction:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.current_mean_change_var = tk.StringVar(value="N/A")
        self.current_mean_change_label = ttk.Label(advanced_frame, textvariable=self.current_mean_change_var)
        self.current_mean_change_label.grid(row=2, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(advanced_frame, text="Min Mean Threshold * :").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.min_mean_threshold_var = tk.DoubleVar(value=0.0050) # Default changed to 0.0050
        self.min_mean_threshold_slider = ttk.Scale(advanced_frame, from_=0.0, to=0.1, orient="horizontal", variable=self.min_mean_threshold_var, command=self._update_min_mean_label)
        self.min_mean_threshold_slider.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        self.min_mean_threshold_label = ttk.Label(advanced_frame, text=f"{self.min_mean_threshold_var.get():.4f}")
        self.min_mean_threshold_label.grid(row=3, column=2, sticky="w", padx=0, pady=2)

        self.min_mean_threshold_text_label = ttk.Label(advanced_frame, text="(e.g., 0.005 to prevent very low mean)", font=("TkDefaultFont", 8))
        self.min_mean_threshold_text_label.grid(row=3, column=3, columnspan=2, sticky="w", padx=5, pady=0)

        self.min_mean_threshold_text_label.bind("<Enter>", self._show_min_mean_tooltip)
        self.min_mean_threshold_text_label.bind("<Leave>", self._hide_min_mean_tooltip)

        # --- Buttons ---
        self.start_button = ttk.Button(master, text="Start Monitoring", command=self.start_monitoring)
        self.start_button.grid(row=4, column=0, sticky="ew", padx=5, pady=5) # Changed row

        self.stop_button = ttk.Button(master, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.grid(row=4, column=1, sticky="ew", padx=5, pady=5) # Changed row

        # --- Status Label ---
        self.status_label = ttk.Label(master, text="Status: Ready")
        self.status_label.grid(row=4, column=2, columnspan=3, sticky="w", padx=5, pady=5) # Changed row

        # --- Pending Uploads Display (New) ---
        uploads_frame = ttk.LabelFrame(master, text="Pending Uploads")
        uploads_frame.grid(row=5, column=0, columnspan=3, sticky="ew", padx=5, pady=5) # Changed row

        self.uploads_tree = ttk.Treeview(uploads_frame, columns=("ID", "Status", "Progress"), show="headings")
        self.uploads_tree.heading("ID", text="ID")
        self.uploads_tree.heading("Status", text="Status")
        self.uploads_tree.heading("Progress", text="Progress")
        self.uploads_tree.column("ID", width=50, stretch=tk.NO)
        self.uploads_tree.column("Status", width=150, stretch=tk.NO)
        self.uploads_tree.column("Progress", width=70, stretch=tk.NO)
        self.uploads_tree.pack(side="left", fill="both", expand=True)

        uploads_scrollbar = ttk.Scrollbar(uploads_frame, orient="vertical", command=self.uploads_tree.yview)
        self.uploads_tree.configure(yscrollcommand=uploads_scrollbar.set)
        uploads_scrollbar.pack(side="right", fill="y")


        # Override the window closing protocol
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def update_status(self, message):
        self.status_label.config(text=f"Status: {message}")
        self.master.update_idletasks() # Refresh GUI

    def _update_treeview_item(self, screenshot_id):
        """Safely updates a Treeview item from the upload_statuses dictionary."""
        status_info = self.upload_statuses.get(screenshot_id)
        if status_info:
            status = status_info['status']
            progress = status_info['progress']
            iid = self.treeview_items.get(screenshot_id)
            if iid:
                self.uploads_tree.item(iid, values=(screenshot_id, status, f"{progress}%"))
            if status == "Completed" or status == "Failed":
                # Schedule removal after a short delay for "Completed" items
                self.master.after(5000, lambda s_id=screenshot_id: self._remove_treeview_item(s_id))

    def _remove_treeview_item(self, screenshot_id):
        """Removes a completed item from the Treeview."""
        iid = self.treeview_items.pop(screenshot_id, None)
        if iid and self.uploads_tree.exists(iid):
            self.uploads_tree.delete(iid)
        self.upload_statuses.pop(screenshot_id, None) # Also remove from tracking dict

    def start_monitoring(self):
        self.stop_event.clear() # Clear the stop event when starting monitoring
        doc_id = self.doc_id_entry.get().strip()
        if not doc_id:
            messagebox.showerror("Error", "Please enter a Google Doc ID.")
            return

        try:
            interval = float(self.interval_entry.get().strip())
            if interval <= 0:
                messagebox.showerror("Error", "Interval must be a positive number.")
                return
        except ValueError:
            messagebox.showerror("Error", "Invalid interval. Please enter a number.")
            return

        bbox_values = []
        for entry in [self.bbox_left, self.bbox_top, self.bbox_width, self.bbox_height]:
            val = entry.get().strip()
            if val:
                try:
                    bbox_values.append(int(val))
                except ValueError:
                    messagebox.showerror("Error", f"Invalid BBox value: '{val}'. Please enter integers or leave empty.")
                    return
            else:
                bbox_values.append(None)
        
        # If all bbox values are None, then bbox is None, otherwise create tuple
        if all(val is None for val in bbox_values):
            bbox = None
        elif all(val is not None for val in bbox_values):
            bbox = tuple(bbox_values)
        else:
            messagebox.showerror("Error", "Please fill all BBox fields or leave all empty for full screen capture.")
            return

        self.update_status("Authenticating with Google...")
        docs_service, drive_service = authenticate_google_api()

        if not docs_service or not drive_service:
            self.update_status("Authentication failed. Check console for details.")
            messagebox.showerror("Authentication Error", "Could not authenticate with Google. See console for error messages.")
            return

        self.update_status("Getting/Creating 'autoscreenshot' folder in Google Drive...")
        self.autoscreenshot_folder_id = _get_or_create_autoscreenshot_folder(drive_service)
        if not self.autoscreenshot_folder_id:
            self.update_status("Failed to get/create 'autoscreenshot' folder. Monitoring stopped.")
            messagebox.showerror("Drive Folder Error", "Could not get or create 'autoscreenshot' folder in Google Drive. Monitoring stopped.")
            return


        self.monitoring = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.doc_id_entry.config(state=tk.DISABLED)
        self.interval_entry.config(state=tk.DISABLED)
        for entry in [self.bbox_left, self.bbox_top, self.bbox_width, self.bbox_height]:
            entry.config(state=tk.DISABLED)
        self.select_region_button.config(state=tk.DISABLED) # Disable select region button too
        self.current_mean_change_var.set("N/A") # Reset mean display on start
        self.recent_local_frac_changes.clear() # Clear recent changes history

        self.update_status("Monitoring started...")
        
        # Start the processor thread first
        self.processor_thread = threading.Thread(target=self._processor_loop, args=(docs_service, drive_service, doc_id))
        self.processor_thread.daemon = True
        self.processor_thread.start()

        # Start the monitor thread
        self.monitor_thread = threading.Thread(target=self._monitor_loop, args=(bbox, interval))
        self.monitor_thread.daemon = True
        self.monitor_thread.start()

    def stop_monitoring(self) -> bool:
        self.monitoring = False
        self.stop_event.set() # Signal all threads to stop

        self.update_status("Stopping monitoring. Waiting for background tasks...")

        # Reset UI elements that can be re-enabled immediately
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.doc_id_entry.config(state=tk.NORMAL)
        self.interval_entry.config(state=tk.NORMAL)
        for entry in [self.bbox_left, self.bbox_top, self.bbox_width, self.bbox_height]:
            entry.config(state=tk.NORMAL)
        self.current_mean_change_var.set("N/A") # Reset mean display on stop

        # Schedule a check for thread termination
        self.master.after(100, self._check_threads_for_shutdown)
        return True # Indicate that shutdown process has started

    def _check_threads_for_shutdown(self):
        monitor_alive = self.monitor_thread and self.monitor_thread.is_alive()
        processor_alive = self.processor_thread and self.processor_thread.is_alive()

        if monitor_alive or processor_alive:
            status_message = "Waiting for:"
            if monitor_alive: status_message += " Monitor "
            if processor_alive: status_message += " Processor "
            self.update_status(status_message.strip() + "...")
            self.master.after(500, self._check_threads_for_shutdown) # Check again after 500ms
        else:
            # All threads have stopped, proceed with final cleanup
            # --- Clear the queue and associated GUI/tracking on stop/exit ---
            while not self.screenshot_queue.empty():
                try:
                    self.screenshot_queue.get_nowait()
                except queue.Empty:
                    pass
            
            for item in self.uploads_tree.get_children():
                self.uploads_tree.delete(item)
                
            self.upload_statuses.clear()
            self.treeview_items.clear()
            # --- End clearing logic ---

            self.update_status("Monitoring stopped.")
            if self.exit_app_on_stop: # Only destroy if exit was intended
                self.master.destroy()

    def _on_paste_doc_id(self, event):
        """Handles pasting of a Google Doc ID into the entry field, extracting the ID if it's a URL."""
        try:
            pasted_text = self.master.clipboard_get()
            extracted_id = extract_doc_id_from_url(pasted_text)
            self.doc_id_entry.delete(0, tk.END)
            self.doc_id_entry.insert(0, extracted_id)
            if extracted_id != pasted_text: # Only show info if extraction actually happened
                messagebox.showinfo("Info", "Google Doc ID extracted and updated from clipboard.")
        except tk.TclError: # Clipboard might be empty or not contain text
            messagebox.showwarning("Warning", "Could not paste from clipboard or clipboard is empty.")
        # Prevent default paste behavior
        return "break"

    def _update_mean_deviation_label(self, value):
        """Updates the mean deviation label when the slider is moved."""
        self.mean_deviation_threshold_label.config(text=f"{float(value):.2f}")

    def _show_mean_deviation_tooltip(self, event):
        """Shows a tooltip when the info button is hovered."""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

        tooltip_text = "This slider controls the threshold for how much the current mean change fraction can deviate from the average.\nLower values (e.g., 0.5) allow more deviation, making the mean more stable.\nHigher values (e.g., 2.0) make the mean more sensitive to recent changes."
        self.tooltip_window = tk.Toplevel(self.master)
        self.tooltip_window.attributes("-topmost", True)
        self.tooltip_window.overrideredirect(True)
        self.tooltip_window.withdraw() # Hide it initially

        label = ttk.Label(self.tooltip_window, text=tooltip_text, justify=tk.LEFT, wraplength=200)
        label.pack(padx=10, pady=10)

        # Position the tooltip window relative to the info button
        x = self.mean_deviation_threshold_text_label.winfo_x() + self.mean_deviation_threshold_text_label.winfo_width() + 5
        y = self.mean_deviation_threshold_text_label.winfo_y() + self.mean_deviation_threshold_text_label.winfo_height() // 2
        self.tooltip_window.geometry(f"+{x}+{y}")

        self.tooltip_window.deiconify()

    def _hide_mean_deviation_tooltip(self, event):
        """Hides the tooltip when the info button is left."""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def _update_min_mean_label(self, value):
        """Updates the minimum mean label when the slider is moved."""
        self.min_mean_threshold_label.config(text=f"{float(value):.4f}")

    def _show_min_mean_tooltip(self, event):
        """Shows a tooltip when the info button is hovered."""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

        tooltip_text = "This slider controls the minimum mean change fraction that must be exceeded for an upload to occur.\nLower values (e.g., 0.005) prevent very low mean changes from triggering uploads.\nHigher values (e.g., 0.1) allow more frequent uploads for very small but frequent changes."
        self.tooltip_window = tk.Toplevel(self.master)
        self.tooltip_window.attributes("-topmost", True)
        self.tooltip_window.overrideredirect(True)
        self.tooltip_window.withdraw() # Hide it initially

        label = ttk.Label(self.tooltip_window, text=tooltip_text, justify=tk.LEFT, wraplength=200)
        label.pack(padx=10, pady=10)

        # Position the tooltip window relative to the info button
        x = self.min_mean_threshold_text_label.winfo_x() + self.min_mean_threshold_text_label.winfo_width() + 5
        y = self.min_mean_threshold_text_label.winfo_y() + self.min_mean_threshold_text_label.winfo_height() // 2
        self.tooltip_window.geometry(f"+{x}+{y}")

        self.tooltip_window.deiconify()

    def _hide_min_mean_tooltip(self, event):
        """Hides the tooltip when the info button is left."""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def _update_current_mean_change_label(self, value):
        """
        Updates the current mean change label when the slider is moved.
        """
        self.current_mean_change_label.config(text=f"{float(value):.4f}")

    def _calculate_mean_local_frac_change(self):
        """
        Calculates the mean of the recent local change fractions.
        """
        if not self.recent_local_frac_changes:
            return 0.0
        return sum(self.recent_local_frac_changes) / len(self.recent_local_frac_changes)

    def _queue_screenshot_for_upload(self, pil_img, status_reason):
        """
        Queues a screenshot for upload to Google Docs.
        'pil_img' is the PIL Image.
        'status_reason' is a string indicating why the screenshot was taken (e.g., "Initial", "Significant Change").
        """
        self.screenshot_counter += 1
        screenshot_id = f"SS-{self.screenshot_counter}"
        item_to_queue = {'id': screenshot_id, 'img': pil_img, 'folder_id': self.autoscreenshot_folder_id}
        self.screenshot_queue.put(item_to_queue)
        
        # Initialize status for GUI update
        self.upload_statuses[screenshot_id] = {'status': 'Queued', 'progress': 0}
        iid = self.uploads_tree.insert("", "end", values=(screenshot_id, "Queued", "0%"))
        self.treeview_items[screenshot_id] = iid
        self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))

        self.update_status(f"Screenshot queued for upload ({status_reason}).")

    def _monitor_loop(self, bbox, interval):
        while self.monitoring and not self.stop_event.is_set():
            try:
                self.update_status("Checking for changes...")
                current_img = _take_and_clipboard_screenshot(bbox, self.stop_event) # adaptive_mode is now always False

                if current_img is None: # This can happen if stop_event is set within _take_and_clipboard_screenshot
                    continue

                if self.previous_local_img is None: # First run or after a restart
                    self.previous_local_img = current_img
                    self.previous_uploaded_img = current_img
                    self._queue_screenshot_for_upload(current_img, "Initial")
                    self.update_status("Monitoring started: Initial screenshot taken and queued for upload.")
                    self.recent_local_frac_changes.append(0.0) # Add 0 to deque for initial state

                else:
                    # Store the image that was 'previous_local_img' for potential upload
                    image_for_potential_upload = self.previous_local_img

                    # Fixed pixel_diff_threshold (no extreme sensitivity option)
                    pixel_diff_threshold_for_scan = 10 # Adjust as needed for typical subtle changes

                    is_changed_enough, raw_local_frac_changed = images_significantly_different(
                        self.previous_local_img, current_img, 
                        threshold=0.0, # Always consider a change for raw_local_frac_changed calc
                        pixel_diff_threshold=pixel_diff_threshold_for_scan
                    )
                    
                    # Always update deque and display, regardless of significance
                    self.recent_local_frac_changes.append(raw_local_frac_changed)
                    self.master.after(0, lambda: self.last_local_change_var.set(f"{raw_local_frac_changed:.4f}"))

                    current_mean_frac_change = self._calculate_mean_local_frac_change()
                    # Apply minimum mean threshold
                    current_mean_frac_change = max(current_mean_frac_change, self.min_mean_threshold_var.get())
                    self.master.after(0, lambda: self.current_mean_change_var.set(f"{current_mean_frac_change:.4f}"))

                    # Decision to upload based on local change significance and deviation from mean
                    if is_changed_enough: # Only consider upload if raw_local_frac_changed is non-zero
                        # Update previous_local_img for next iteration's local scan
                        self.previous_local_img = current_img 

                        # Upload Logic: if local change > mean * (1 + deviation_threshold)
                        deviation_threshold = self.mean_deviation_threshold_var.get()

                        if (current_mean_frac_change == 0.0 and raw_local_frac_changed > 0.0001) or \
                           (raw_local_frac_changed > current_mean_frac_change * (1 + deviation_threshold)):
                            
                            # Before uploading, check if the `image_for_potential_upload`
                            # is significantly different from the last uploaded image.
                            # This prevents re-uploading the same image if deviation threshold is met by small noise
                            if images_significantly_different(self.previous_uploaded_img, image_for_potential_upload, threshold=0.0001, pixel_diff_threshold=pixel_diff_threshold_for_scan)[0]:
                                print(f"Upload triggered: Change {raw_local_frac_changed:.4f} > Mean {current_mean_frac_change:.4f} * (1 + {deviation_threshold:.2f}). Uploading 'last' image.")
                                self._queue_screenshot_for_upload(image_for_potential_upload, "Significant Deviation")
                                self.previous_uploaded_img = image_for_potential_upload # Update uploaded reference
                            else:
                                print("Upload suppressed: Not enough pixel change from last uploaded image (despite deviation).")
                        else:
                            print(f"Upload suppressed: Change {raw_local_frac_changed:.4f} not > Mean {current_mean_frac_change:.4f} * (1 + {deviation_threshold:.2f}).")
                    else:
                        print("No local change detected (no upload attempt).")

                # Use wait for interval, so it can be interrupted
                self.stop_event.wait(interval) # This replaces time.sleep(interval)

            except Exception as e:
                self.update_status(f"Error in monitor loop: {e}")
                messagebox.showerror("Monitor Error", f"An error occurred in monitor: {e}. Monitoring stopped.")
                self.monitoring = False
                self.stop_event.set() # Ensure stop_event is set if loop breaks due to error
                self.master.after(0, self.stop_monitoring) # Schedule stop_monitoring in main thread

    def _processor_loop(self, docs_service, drive_service, doc_id):
        while self.monitoring and not self.stop_event.is_set():
            try:
                # Use a timeout for get() and check stop_event to allow graceful exit
                item = self.screenshot_queue.get(timeout=0.1) 
                screenshot_id = item['id']
                img = item['img']
                folder_id = item['folder_id']

                self.update_status(f"Processing queued screenshot {screenshot_id} ({self.screenshot_queue.qsize()} remaining).")
                self.upload_statuses[screenshot_id] = {'status': 'Uploading to Drive', 'progress': 25}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))

                # Upload image to Google Drive
                img_byte_arr = io.BytesIO()
                print(f"DEBUG: Image dimensions before saving to Drive for {screenshot_id}: {img.size}")
                img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)

                file_metadata = {
                    'name': f'screenshot_{screenshot_id}.png', 
                    'mimeType': 'image/png',
                    'parents': [folder_id] # New: Specify the parent folder ID
                }
                
                media = MediaIoBaseUpload(img_byte_arr, mimetype='image/png', resumable=True)
                request = drive_service.files().create(body=file_metadata, media_body=media, fields='id,webContentLink')

                response = None
                while response is None:
                    status, response = request.next_chunk()
                    if status:
                        progress_percentage = int(status.progress() * 100)
                        self.upload_statuses[screenshot_id] = {'status': f'Uploading {progress_percentage}%', 'progress': progress_percentage}
                        self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))

                file = response
                image_url = file.get('webContentLink')
                image_id = file.get('id')

                self.upload_statuses[screenshot_id] = {'status': 'Setting Permissions', 'progress': 50}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))

                # Make the image publicly accessible
                drive_service.permissions().create(
                    fileId=image_id,
                    body={'type': 'anyone', 'role': 'reader'},
                    fields='id'
                ).execute()

                self.upload_statuses[screenshot_id] = {'status': 'Inserting into Doc', 'progress': 75}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))

                # Insert image into Google Doc
                requests = [{'insertInlineImage': {'endOfSegmentLocation': {}, 'uri': image_url, 'objectSize': {'height': {'magnitude': 300, 'unit': 'PT'}}}}]
                docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

                self.upload_statuses[screenshot_id] = {'status': 'Deleting from Drive', 'progress': 90}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))

                # Delete the uploaded image from Google Drive
                drive_service.files().delete(fileId=image_id).execute()
                
                self.upload_statuses[screenshot_id] = {'status': 'Completed', 'progress': 100}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id)) # Final update before removal

                self.update_status(f"Screenshot {screenshot_id} added to Doc and deleted from Drive. Waiting for next...")
                self.screenshot_queue.task_done()
            except queue.Empty:
                if self.stop_event.is_set(): # Exit if stop_event is set and queue is empty
                    break
                continue # Continue waiting if queue is empty but not stopping
            except HttpError as error:
                self.update_status(f"Google API Error in processor for {screenshot_id}: {error}")
                self.upload_statuses[screenshot_id] = {'status': 'Failed', 'progress': 0}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))
                messagebox.showerror("Google API Error", f"An API error occurred during upload for {screenshot_id}: {error}. Monitoring stopped.")
                self.monitoring = False
                self.stop_event.set()
                if not self.screenshot_queue.empty(): # Mark task done even if failed to avoid deadlock if other tasks are pending
                     self.screenshot_queue.task_done()
            except Exception as e:
                self.update_status(f"Error in processor loop for {screenshot_id}: {e}")
                self.upload_statuses[screenshot_id] = {'status': 'Failed', 'progress': 0}
                self.master.after(0, lambda s_id=screenshot_id: self._update_treeview_item(s_id))
                messagebox.showerror("Processor Error", f"An unexpected error occurred in processor for {screenshot_id}: {e}. Monitoring stopped.")
                self.monitoring = False
                self.stop_event.set()
                if not self.screenshot_queue.empty():
                    self.screenshot_queue.task_done()
        
        if not self.monitoring and self.stop_event.is_set():
            self.master.after(0, self.stop_monitoring)

    def on_closing(self):
        if self.monitoring:
            self.update_status("Attempting graceful shutdown. Please wait...")
            self.exit_app_on_stop = True # Indicate that the app should exit after stopping
            self.stop_monitoring() # Initiate the asynchronous shutdown
        else:
            self.master.destroy()
            
    def select_region_on_screen(self):
        # Temporarily hide the main window
        self.master.withdraw()

        # Create a new Toplevel window for region selection
        self.region_selector = tk.Toplevel(self.master)
        self.region_selector.attributes("-fullscreen", True)
        self.region_selector.attributes("-alpha", 0.3) # Make it semi-transparent
        self.region_selector.attributes("-topmost", True) # Keep on top
        self.region_selector.attributes("-toolwindow", True) # Hide from taskbar

        # Disable window decorations (title bar, borders)
        self.region_selector.overrideredirect(True)

        self.canvas = tk.Canvas(self.region_selector, cursor="cross", bg="white", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.start_x = None
        self.start_y = None
        self.rect_id = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        
        # Add a way to cancel selection (e.g., Escape key)
        self.region_selector.bind("<Escape>", self.cancel_selection)

    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)

        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def on_mouse_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        # Calculate bounding box
        left = min(self.start_x, end_x)
        top = min(self.start_y, end_y)
        width = abs(end_x - self.start_x)
        height = abs(end_y - self.start_y)

        # Update the main GUI's entry fields
        self.bbox_left.delete(0, tk.END)
        self.bbox_left.insert(0, str(int(left)))
        self.bbox_top.delete(0, tk.END)
        self.bbox_top.insert(0, str(int(top)))
        self.bbox_width.delete(0, tk.END)
        self.bbox_width.insert(0, str(int(width)))
        self.bbox_height.delete(0, tk.END)
        self.bbox_height.insert(0, str(int(height)))

        self.close_region_selector()

    def cancel_selection(self, event=None):
        # Clear any partial selection drawing
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.close_region_selector()

    def close_region_selector(self):
        if self.region_selector:
            self.region_selector.destroy()
        self.master.deiconify() # Restore the main window

if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()
