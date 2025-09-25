import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc

def get_connection():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=localhost\\SQLEXPRESS;'
        'DATABASE=InstaOnlineShopFinal;'
        'Trusted_Connection=yes;'
    )
    return conn

def run_query(sql):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
        conn.close()
        return rows
    except Exception as e:
        return f"خطا: {str(e)}"

QUERIES = {
    "1. مشتریان با ایمیل دامنه example.com": "SELECT name, email FROM Customer WHERE email LIKE '%@example.com'",
    "2. تعداد کالاهای هر دسته‌بندی": "SELECT category, COUNT(*) AS count FROM Product GROUP BY category",
    "3. لیست سفارشات و مشتری (JOIN)": """
        SELECT c.name AS customer_name, o.order_id, o.order_date
        FROM [Order] o
        JOIN Customer c ON o.customer_id = c.customer_id
    """,
    "4. مجموع درآمد فروشگاه": "SELECT SUM(amount) AS total_income FROM Payment",
    "5. کالاهای الکترونیک مرتب بر اساس قیمت (نزولی)": "SELECT name, price FROM Product WHERE category = 'الکترونیک' ORDER BY price DESC",
    "6. مشتریانی که گوشی خریده‌اند": """
        SELECT DISTINCT c.name
        FROM Customer c
        JOIN [Order] o ON c.customer_id = o.customer_id
        JOIN OrderItem oi ON o.order_id = oi.order_id
        JOIN Product p ON oi.product_id = p.product_id
        WHERE p.name LIKE N'%گوشی%'
    """,
    "7. بیشترین قیمت کالا": "SELECT MAX(price) AS max_price FROM Product",
    "8. کالاهای با قیمت بین 1 تا 5 میلیون": "SELECT name, price FROM Product WHERE price BETWEEN 1000000 AND 5000000",
    "9. تعداد کل سفارشات": "SELECT COUNT(*) AS total_orders FROM [Order]",
    "10. جزئیات سفارشات با پرداخت": """
        SELECT o.order_id, p.name AS product_name, oi.quantity, pay.amount
        FROM [Order] o
        JOIN OrderItem oi ON o.order_id = oi.order_id
        JOIN Product p ON oi.product_id = p.product_id
        JOIN Payment pay ON o.order_id = pay.order_id
    """
}

class App:
    def __init__(self, root):
        self.root = root
        root.title("فاز ۴ - فروشگاه اینترنتی InstaOnlineShop | سید دانیال میررشیدی - 40271954")
        root.geometry("900x500")

        tk.Label(root, text="انتخاب پرس‌وجو:", font=("Arial", 12)).pack(pady=5)
        self.query_var = tk.StringVar()
        self.combo = ttk.Combobox(root, textvariable=self.query_var, values=list(QUERIES.keys()), state="readonly", width=80)
        self.combo.pack(pady=5)
        self.combo.bind("<<ComboboxSelected>>", self.execute)

        self.tree = ttk.Treeview(root, show="headings")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.place(x=880, y=100, height=350)

    def execute(self, event=None):
        name = self.query_var.get()
        sql = QUERIES[name]

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = ()

        result = run_query(sql)
        if isinstance(result, str):
            messagebox.showerror("خطا", result)
            return

        if not result:
            messagebox.showinfo("نتیجه", "هیچ داده‌ای یافت نشد.")
            return

        cols = list(result[0].keys())
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        for row in result:
            self.tree.insert("", "end", values=list(row.values()))

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()