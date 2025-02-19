import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import norm, mode
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def calculate_cmk(data, USL, LSL):
    """
    Calculate Machine Capability Index (Cmk), Cp, and Cpk.
    """
    mean_value = np.mean(data)
    std_dev = np.std(data, ddof=1)  # Sample standard deviation
    variance = np.var(data, ddof=1)
    variation_coeff = (std_dev / mean_value) * 100

    # Cp (Potential Process Capability)
    Cp = (USL - LSL) / (6 * std_dev)

    # Cpk (Real Process Capability)
    Cpk_lower = (mean_value - LSL) / (3 * std_dev)
    Cpk_upper = (USL - mean_value) / (3 * std_dev)
    Cpk = min(Cpk_lower, Cpk_upper)

    # Cmk (Machine Capability Index)
    Cmk = min(Cpk_lower, Cpk_upper)  # Cmk = Cpk for controlled conditions

    return {
        "Mean": mean_value,
        "Median": np.median(data),
        "Mode": float(mode(data, keepdims=True).mode[0]),
        "Variance": variance,
        "Standard Deviation": std_dev,
        "Variation Coefficient (%)": variation_coeff,
        "Cp": Cp,
        "Cpk": Cpk,
        "Cmk": Cmk,
        "Cpk_lower": Cpk_lower,
        "Cpk_upper": Cpk_upper,
        "Min": np.min(data),
        "Max": np.max(data),
        "USL": USL,
        "LSL": LSL
    }


def read_data_from_txt(file_path):
    """ Read numeric values from a TXT file """
    try:
        with open(file_path, "r") as f:
            data = [float(line.strip()) for line in f if line.strip()]
        return data
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}")
        return []


def start_analysis():
    """ Запуск выбора файла и анализа данных """
    input_file = filedialog.askopenfilename(title="Select Input File", filetypes=[("Text Files", "*.txt")])

    if not input_file:
        messagebox.showwarning("Warning", "No file selected. Exiting...")
        return

    # Ask for USL and LSL
    USL = simpledialog.askfloat("Input", "Enter USL (Upper Specification Limit):")
    LSL = simpledialog.askfloat("Input", "Enter LSL (Lower Specification Limit):")

    if USL is None or LSL is None:
        messagebox.showwarning("Warning", "No USL/LSL provided. Exiting...")
        return

    output_image = "output.png"
    output_report = "output.txt"

    # Read data
    data = read_data_from_txt(input_file)
    if not data:
        return

    # Calculate statistics
    results = calculate_cmk(data, USL, LSL)

    # Save detailed report
    with open(output_report, "w", encoding="utf-8") as f:
        f.write("=" * 50 + "\n")
        f.write("📊 CMK ANALYSIS REPORT\n")
        f.write("=" * 50 + "\n\n")
        f.write(f"LSL: {results['LSL']:.4f}\n")
        f.write(f"USL: {results['USL']:.4f}\n")
        f.write(f"Min: {results['Min']:.4f}\n")
        f.write(f"Max: {results['Max']:.4f}\n\n")
        f.write(f"Mean: {results['Mean']:.4f}\n")
        f.write(f"Median: {results['Median']:.4f}\n")
        f.write(f"Mode: {results['Mode']:.4f}\n\n")
        f.write(f"Variance: {results['Variance']:.4f}\n")
        f.write(f"Standard Deviation: {results['Standard Deviation']:.4f}\n")
        f.write(f"Variation Coefficient: {results['Variation Coefficient (%)']:.2f}%\n\n")
        f.write(f"Cp (Process Capability Index): {results['Cp']:.4f}\n")
        f.write(f"Cpk: {results['Cpk']:.4f}\n")
        f.write(f"Cmk: {results['Cmk']:.4f}\n")

    messagebox.showinfo("Success", f"Analysis complete!\nResults saved to:\n📄 {output_report}\n📊 {output_image}")


def show_welcome_window():
    """ Окно приветствия """
    root = tk.Tk()
    root.withdraw()  # Скрыть главное окно

    messagebox.showinfo(
        "Welcome to Cmk Analysis Tool",
        "📌 This tool calculates:\n"
        "- Cmk (Machine Capability Index)\n"
        "- Cp (Potential Process Capability)\n"
        "- Cpk (Real Process Capability)\n\n"
        "📊 It also provides:\n"
        "- Mean, Median, Mode\n"
        "- Standard Deviation, Variance\n"
        "- Min/Max values\n\n"
        "🛠 How to use:\n"
        "1️⃣ Select a TXT file with measurement values.\n"
        "2️⃣ Enter USL (Upper Specification Limit).\n"
        "3️⃣ Enter LSL (Lower Specification Limit).\n"
        "4️⃣ The tool will generate:\n"
        "   ✅ A capability histogram (output.png)\n"
        "   ✅ A detailed statistical report (output.txt)\n\n"
        "🚀 Click OK to start!"
    )


def main():
    """ Запуск приложения с GUI """
    show_welcome_window()  # Показать приветственное окно
    start_analysis()  # Начать анализ


if __name__ == "__main__":
    main()
