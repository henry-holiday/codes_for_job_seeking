import os
import pandas as pd
import zipfile
from tkinter import Tk, Button, filedialog, messagebox, Toplevel, Radiobutton, StringVar, Label, Text, Scrollbar, END
from datetime import datetime
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
import threading
import gc

# Function to select multiple ZIP files
def select_zip_files():
    zip_files = filedialog.askopenfilenames(
        title="Select ZIP Files",
        filetypes=[("ZIP Files", "*.zip")]
    )
    return zip_files

# Function to select output folder
def select_output_folder():
    folder_selected = filedialog.askdirectory(title="Select Output Folder")
    return folder_selected

# Function to apply black borders to Excel cells
def apply_borders(worksheet):
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border

# Function to adjust column width in Excel
def adjust_column_width(worksheet):
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)  # Add some padding
        worksheet.column_dimensions[column].width = adjusted_width

# Function to process VCSEL files (flood or dot)
def process_vcsel(input_folder, output_folder, file_identifier, report_name, wavelength_threshold, serial_number, log_terminal):
    try:
        # Find files that contain the specified identifier and have the correct extension
        vcsel_files = [f for f in os.listdir(input_folder) if file_identifier in f and (f.endswith('.csv') or f.endswith('.xlsx'))]
        dfs = []
        for file in vcsel_files:
            file_path = os.path.join(input_folder, file)
            log_terminal.insert(END, f"Processing file: {file_path}\n")
            log_terminal.see(END)

            # Read the file
            if file.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            # Check if required columns exist
            required_columns = ["PS_AVI_PF", "FAB_WF_ID", "RW_WF_ID", "CAW_INT_50C_(nm)"]
            if not all(col in df.columns for col in required_columns):
                log_terminal.insert(END, f"Skipping file {file}: Missing required columns\n")
                log_terminal.see(END)
                continue

            # Filter rows where PS_AVI_PF == 1
            df = df[df["PS_AVI_PF"] == 1]

            # Skip if DataFrame is empty after filtering
            if df.empty:
                log_terminal.insert(END, f"Skipping file {file}: No rows with PS_AVI_PF == 1\n")
                log_terminal.see(END)
                continue

            dfs.append(df)

        if dfs:
            # Concatenate all DataFrames
            concatenated_df = pd.concat(dfs, axis=0)

            # Classify rows based on CAW_INT_50C_(nm) and the given threshold
            concatenated_df['type'] = concatenated_df['CAW_INT_50C_(nm)'].apply(lambda x: 'ng' if x < wavelength_threshold else 'ok')

            # Group by FAB_WF_ID and RW_WF_ID, and count 'ng' and 'ok'
            grouped_df = concatenated_df.groupby(['FAB_WF_ID', 'RW_WF_ID', 'type']).size().unstack(fill_value=0)

            # Ensure 'ng' and 'ok' columns always exist
            grouped_df = grouped_df.reindex(columns=['ng', 'ok'], fill_value=0)

            # Add total column (ng + ok)
            grouped_df['total'] = grouped_df['ng'] + grouped_df['ok']

            # Calculate percentage (ng / total)
            grouped_df['percentage'] = (grouped_df['ng'] / grouped_df['total']) * 100

            # Format percentage column to include '%' symbol
            grouped_df['percentage'] = grouped_df['percentage'].apply(lambda x: f"{x:.2f}%")

            # Filter out rows where Percentage < 50%
            grouped_df['percentage_numeric'] = grouped_df['percentage'].apply(lambda x: float(x.replace('%', '')))
            grouped_df = grouped_df[grouped_df['percentage_numeric'] >= 5]
            grouped_df.drop(columns=['percentage_numeric'], inplace=True)

            # Reorder and rename columns
            grouped_df = grouped_df[['ng', 'ok', 'total', 'percentage']]
            grouped_df.columns = ['NG', 'OK', 'Total', 'Percentage']

            # Reset index to include FAB_WF_ID and RW_WF_ID as columns
            grouped_df = grouped_df.reset_index()

            # Save to output file with serial number
            current_time = datetime.now().strftime("%Y%m%d")
            output_file = os.path.join(output_folder, f"{report_name} wavelength report-{current_time}-v{serial_number}.xlsx")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                grouped_df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                apply_borders(worksheet)
                adjust_column_width(worksheet)  # Adjust column width

            log_terminal.insert(END, f"Report generated and saved to {output_file}\n")
            log_terminal.see(END)
        else:
            log_terminal.insert(END, f"No valid {report_name} VCSEL files found.\n")
            log_terminal.see(END)

    except Exception as e:
        log_terminal.insert(END, f"Error during processing: {str(e)}\n")
        log_terminal.see(END)

# Function to extract files from ZIP archives
def extract_files(zip_files, output_folder, log_terminal):
    extracted_count = 0
    total_files = 0

    # Calculate total number of files in all ZIP archives
    for zip_file in zip_files:
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            total_files += len(zip_ref.namelist())

    log_terminal.insert(END, f"Extracting {total_files} files from {len(zip_files)} ZIP archives...\n")
    log_terminal.see(END)

    # Extract files from each ZIP archive
    for zip_file in zip_files:
        log_terminal.insert(END, f"Extracting files from {zip_file}...\n")
        log_terminal.see(END)
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            for file in zip_ref.namelist():
                file_name = os.path.basename(file)
                if file_name:  # Skip empty entries
                    save_path = os.path.join(output_folder, file_name)
                    base_name, ext = os.path.splitext(file_name)
                    counter = 1

                    # Rename file if it already exists
                    while os.path.exists(save_path):
                        save_path = os.path.join(output_folder, f"{base_name}_{counter}{ext}")
                        counter += 1

                    # Write the file to the output folder
                    with zip_ref.open(file) as source, open(save_path, "wb") as target:
                        target.write(source.read())
                    extracted_count += 1

                    log_terminal.insert(END, f"Extracted file: {save_path}\n")
                    log_terminal.see(END)

                    gc.collect()  # Free memory after writing each file

    log_terminal.insert(END, f"Extracted {extracted_count} files successfully.\n")
    log_terminal.see(END)
    return extracted_count

# GUI Interface
def main():
    root = Tk()
    root.title("VCSEL Report Generator")

    # Center the GUI window
    window_width = 600
    window_height = 500
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Variables to store user selections
    zip_files = []
    output_folder = ""
    process_type = StringVar(value="flood")

    # Log terminal
    log_terminal = Text(root, wrap="word", height=15, width=70, bg="black", fg="white")
    log_terminal.pack(pady=10, fill="both", expand=True)

    scrollbar = Scrollbar(log_terminal)
    scrollbar.pack(side="right", fill="y")
    log_terminal.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=log_terminal.yview)

    def select_files():
        nonlocal zip_files
        zip_files = select_zip_files()
        if zip_files:
            log_terminal.insert(END, f"Selected {len(zip_files)} ZIP files.\n")
            log_terminal.see(END)
        else:
            log_terminal.insert(END, "No ZIP files selected.\n")
            log_terminal.see(END)

    def select_output():
        nonlocal output_folder
        output_folder = select_output_folder()
        if output_folder:
            log_terminal.insert(END, f"Output folder set to {output_folder}.\n")
            log_terminal.see(END)
        else:
            log_terminal.insert(END, "No output folder selected.\n")
            log_terminal.see(END)

    def choose_process_type():
        if not zip_files:
            log_terminal.insert(END, "Please select ZIP files first.\n")
            log_terminal.see(END)
            return
        if not output_folder:
            log_terminal.insert(END, "Please select an output folder first.\n")
            log_terminal.see(END)
            return

        popup = Toplevel(root)
        popup.title("Choose Process Type")
    # Set popup window size
        popup_width = 300
        popup_height = 150

    # Calculate the position to center the popup on the screen
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width // 2) - (popup_width // 2)
        y = (screen_height // 2) - (popup_height // 2)

    # Set the geometry of the popup to center it
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        label = Label(popup, text="Select the type of process:")
        label.pack(pady=10)

        flood_radio = Radiobutton(popup, text="Flood", variable=process_type, value="flood")
        flood_radio.pack()

        dot_radio = Radiobutton(popup, text="Dot", variable=process_type, value="dot")
        dot_radio.pack()

        popup.destroy_button = Button(popup, text="Confirm", command=popup.destroy)
        popup.destroy_button.pack(pady=10)

    def start_processing():
        if not zip_files:
            log_terminal.insert(END, "Please select ZIP files first.\n")
            log_terminal.see(END)
            return
        if not output_folder:
            log_terminal.insert(END, "Please select an output folder first.\n")
            log_terminal.see(END)
            return

        # Start extraction
        def task():
            extract_files(zip_files, output_folder, log_terminal)
            # After extraction is complete, start processing
            if process_type.get() == "flood":
                process_vcsel(output_folder, output_folder, 'PB21', "Lumentum flood vcsel", 938.5, 1, log_terminal)
            elif process_type.get() == "dot":
                process_vcsel(output_folder, output_folder, 'PA23', "Lumentum dot vcsel", 940, 1, log_terminal)

        # Run the extraction and processing task in a separate thread
        threading.Thread(target=task).start()

    def clear_logs():
        log_terminal.delete(1.0, END)

    # Buttons for the main GUI
    select_files_button = Button(root, text="Select ZIP Files", command=select_files)
    select_files_button.pack(pady=5)

    select_output_button = Button(root, text="Select Output Folder", command=select_output)
    select_output_button.pack(pady=5)

    choose_process_button = Button(root, text="Choose Process Type", command=choose_process_type)
    choose_process_button.pack(pady=5)

    start_processing_button = Button(root, text="Start Processing", command=start_processing)
    start_processing_button.pack(pady=5)

    clear_logs_button = Button(root, text="Clear Logs", command=clear_logs)
    clear_logs_button.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
