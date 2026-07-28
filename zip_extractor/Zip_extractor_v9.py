import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import gc

def extract_files():
    def task():
        # Select individual ZIP files
        zip_files = filedialog.askopenfilenames(
            title="Select ZIP Files",
            filetypes=[("ZIP Files", "*.zip")]
        )
        if not zip_files:
            return

        # Select folder to save extracted files
        save_folder = filedialog.askdirectory(title="Select Folder to Save Extracted Files")
        if not save_folder:
            return

        extracted_count = 0
        total_files = 0

        # Calculate total number of files
        for zip_path in zip_files:
            if zip_path.endswith(".zip"):
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    total_files += len(zip_ref.namelist())

        # Process each ZIP file
        for zip_path in zip_files:
            if zip_path.endswith(".zip"):
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for file in zip_ref.namelist():
                        file_name = os.path.basename(file)
                        if file_name:  # Skip empty entries
                            # Extract the file with its original name
                            save_path = os.path.join(save_folder, file_name)

                            # Check if file already exists, and rename if needed
                            base_name, ext = os.path.splitext(file_name)
                            counter = 1
                            while os.path.exists(save_path):
                                save_path = os.path.join(save_folder, f"{base_name}_{counter}{ext}")
                                counter += 1

                            # Write file to save folder
                            with zip_ref.open(file) as source, open(save_path, "wb") as target:
                                target.write(source.read())

                            extracted_count += 1
                            update_progress((extracted_count / total_files) * 100)
                            gc.collect()  # 手动触发垃圾回收

        # Show success message
        messagebox.showinfo("Extraction Complete", f"Extracted {extracted_count} files to {save_folder}.")
        root.quit()

    # Run the extraction task in a separate thread
    threading.Thread(target=task).start()


# GUI Interface
root = tk.Tk()
root.title("ZIP File Extractor")
root.geometry("400x200")

label = tk.Label(root, text="Extract All Files from ZIP Archives", font=("Arial", 12))
label.pack(pady=20)

progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=10)

btn_extract = tk.Button(root, text="Extract Files", command=extract_files, font=("Arial", 12))
btn_extract.pack(pady=10)


def update_progress(value):
    progress['value'] = value
    root.update_idletasks()


root.mainloop()
