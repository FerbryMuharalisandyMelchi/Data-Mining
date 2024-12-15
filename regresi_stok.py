import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class RegressionCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Regression Calculator")
        self.root.geometry("500x700")
        self.root.configure(bg="#E5E5E5")

        # Styling colors
        self.primary_color = "#8E44AD"  # Purple
        self.button_color = "#C39BD3"
        self.text_color = "#FFFFFF"
        self.font_title = ("Helvetica", 12, "bold")

        # Variables to store file paths
        self.sales_file_path = tk.StringVar(value="No file selected")
        self.purchase_file_path = tk.StringVar(value="No file selected")

        # Header label
        tk.Label(root, text="Kode Item   Nama Item   Kategori   Unit Terjual", 
                 font=("Helvetica", 10, "bold"), bg="#8E44AD", fg="#FFFFFF").pack(pady=10)

        tk.Label(root, text="Wajib menggunakan format header isi file seperti diatas",
                 font=("Helvetica", 9), bg="#E5E5E5", fg="#777777").pack(pady=5)

        # Step 1: File Import
        self.create_label("Step 1: Import Files")
        self.sales_button = self.create_button("Import Sales File", self.import_sales_file)
        tk.Label(root, textvariable=self.sales_file_path, font=("Helvetica", 9), bg="#E5E5E5", fg="#333333").pack()

        self.purchase_button = self.create_button("Import Purchase File", self.import_purchase_file)
        tk.Label(root, textvariable=self.purchase_file_path, font=("Helvetica", 9), bg="#E5E5E5", fg="#333333").pack()

        # Step 2: Enter Item Name
        self.create_label("Step 2: Enter Item Name")
        self.item_entry = tk.Entry(root, font=("Helvetica", 10), justify="center", bg="#F4ECF7", relief="flat")
        self.item_entry.pack(pady=5, ipadx=10, ipady=5)

        # Buttons for regression and graph
        self.calculate_button = self.create_button("Calculate Regression", self.calculate_regression, large=True)
        self.graph_button = self.create_button("Show Sales Graph", self.show_sales_graph, large=True)

    def create_label(self, text):
        tk.Label(self.root, text=text, font=self.font_title, bg="#E5E5E5", fg="#8E44AD").pack(pady=5)

    def create_button(self, text, command, large=False):
        button = tk.Button(self.root, text=text, command=command,
                           bg=self.primary_color, fg=self.text_color,
                           activebackground=self.button_color, relief="flat",
                           font=("Helvetica", 10, "bold"), cursor="hand2")
        button.pack(pady=5, ipadx=15, ipady=5)
        if large:
            button.configure(font=("Helvetica", 12, "bold"), width=20)
        return button

    def import_sales_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.sales_file_path.set(f"Sales File: {file_path}")
            self.sales_file = file_path
            messagebox.showinfo("File Selected", f"Sales file imported successfully: {file_path}")

    def import_purchase_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.purchase_file_path.set(f"Purchase File: {file_path}")
            self.purchase_file = file_path
            messagebox.showinfo("File Selected", f"Purchase file imported successfully: {file_path}")

    def calculate_regression(self):
        if not hasattr(self, 'sales_file') or not hasattr(self, 'purchase_file'):
            messagebox.showerror("Error", "Please import both sales and purchase files first!")
            return

        item_name = self.item_entry.get()
        if not item_name:
            messagebox.showerror("Error", "Please enter an item name.")
            return

        try:
            # Load Excel files
            sales_data = pd.read_excel(self.sales_file)
            purchase_data = pd.read_excel(self.purchase_file)

            sales_grouped = sales_data.groupby("Kode Item").agg({
                "Unit Terjual": "sum",
                "Harga Total": "sum"
            })
            sales_grouped["Frekuensi Transaksi"] = sales_data.groupby("Kode Item").size()

            purchase_grouped = purchase_data.groupby("Kode Item").agg({
                "Unit Terjual": "sum"
            })

            if item_name not in sales_grouped.index or item_name not in purchase_grouped.index:
                messagebox.showerror("Error", f"Item '{item_name}' not found in the data.")
                return

            X1 = sales_grouped.loc[item_name, "Frekuensi Transaksi"]
            X2 = sales_grouped.loc[item_name, "Harga Total"]
            X3 = purchase_grouped.loc[item_name, "Unit Terjual"]

            a = 0.5
            b1 = 2.1
            b2 = 0.003
            b3 = 1.2
            Y = a + b1 * X1 + b2 * X2 + b3 * X3

            result_message = (
                f"Regression Calculation for Item '{item_name}':\n\n"
                f"Frekuensi Transaksi (X1): {X1}\n"
                f"Total Pengeluaran Konsumen (X2): {X2}\n"
                f"Pembelian Pemilik Toko (X3): {X3}\n\n"
                f"Hasil Y (Jumlah Barang Dibeli): {Y:.2f}"
            )
            messagebox.showinfo("Calculation Result", result_message)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def show_sales_graph(self):
        if not hasattr(self, 'sales_file'):
            messagebox.showerror("Error", "Please import the sales file first!")
            return

        try:
            sales_data = pd.read_excel(self.sales_file)
            grouped_data = sales_data.groupby("Nama Item")["Unit Terjual"].sum().reset_index()

            graph_window = Toplevel(self.root)
            graph_window.title("Sales Data Graph")
            graph_window.geometry("600x400")
            graph_window.configure(bg="#E5E5E5")

            fig, ax = plt.subplots(figsize=(5, 4))
            ax.bar(range(len(grouped_data)), grouped_data["Unit Terjual"], color="#8E44AD")
            ax.set_title("Grafik Penjualan (Unit Terjual)")
            ax.set_ylabel("Unit Terjual")
            ax.set_xticks([])

            canvas = FigureCanvasTkAgg(fig, master=graph_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RegressionCalculatorApp(root)
    root.mainloop()
