import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
from db import create_tables, add_user, user_exists
from auth import login
from excel_generator import generate_excel, load_user_data
import os

# Tworzenie bazy danych
create_tables()

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikacja do raportów")
        self.root.geometry("1200x700")  # Ustawienie rozmiaru okna na 1200x700

        self.logged_in_user = None
        self.create_login_ui()

    def create_login_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        tk.Label(frame, text="Użytkownik:").grid(row=0, column=0, padx=5, pady=5)
        self.username_entry = tk.Entry(frame)
        self.username_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame, text="Hasło:").grid(row=1, column=0, padx=5, pady=5)
        self.password_entry = tk.Entry(frame, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Button(frame, text="Zaloguj", command=self.handle_login).grid(row=2, column=0, pady=10)
        tk.Button(frame, text="Rejestracja", command=self.create_register_ui).grid(row=2, column=1, pady=10)

    def handle_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        user = login(username, password)
        if user:
            self.logged_in_user = user
            self.show_main_ui()
        else:
            messagebox.showerror("Błąd", "Nieprawidłowe dane logowania")

    def create_register_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        tk.Label(frame, text="Użytkownik:").grid(row=0, column=0, padx=5, pady=5)
        username_entry = tk.Entry(frame)
        username_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame, text="Hasło:").grid(row=1, column=0, padx=5, pady=5)
        password_entry = tk.Entry(frame, show="*")
        password_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame, text="Rola:").grid(row=2, column=0, padx=5, pady=5)
        role_combobox = ttk.Combobox(frame, values=["user", "manager"], state="readonly")
        role_combobox.current(0)
        role_combobox.grid(row=2, column=1, padx=5, pady=5)

        tk.Button(frame, text="Zarejestruj", command=lambda: self.handle_register(
            username_entry.get(), password_entry.get(), role_combobox.get())).grid(row=3, column=0, columnspan=2, pady=10)

        tk.Button(frame, text="Powrót", command=self.create_login_ui).grid(row=4, column=0, columnspan=2, pady=5)

    def handle_register(self, username, password, role):
        if not username or not password:
            messagebox.showerror("Błąd", "Uzupełnij wszystkie pola!")
            return

        if user_exists(username):
            messagebox.showerror("Błąd", "Użytkownik o podanej nazwie już istnieje!")
            return

        add_user(username, password, role)
        messagebox.showinfo("Sukces", "Rejestracja zakończona powodzeniem!")
        self.create_login_ui()

    def show_main_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        tk.Label(self.root, text=f"Witaj, {self.logged_in_user[1]}!").pack()
        tk.Button(self.root, text="Dodaj raport", command=self.add_report_ui).pack(pady=5)
        tk.Button(self.root, text="Podgląd raportów", command=self.view_reports).pack(pady=5)

    def add_report_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        tk.Label(frame, text="Opis raportu:").grid(row=0, column=0, padx=5, pady=5)
        description_text = tk.Text(frame, height=5, width=40)
        description_text.grid(row=0, column=1, padx=5, pady=5)

        # Kategorie
        categories = ["Naukowe", "Artykuł", "Seminarium", "Firma", "Własne"]
        tk.Label(frame, text="Kategoria:").grid(row=1, column=0, padx=5, pady=5)
        category_combobox = ttk.Combobox(frame, values=categories, state="readonly")
        category_combobox.grid(row=1, column=1, padx=5, pady=5)

        # Podkategorie
        subcategories = {
            "Naukowe": ["Badania", "Prace magisterskie", "Publikacje"],
            "Artykuł": ["Blog", "Prasa", "Magazyn"],
            "Seminarium": ["Konferencje", "Warsztaty", "Wykłady"],
            "Firma": ["Projekt", "Spotkania", "Prezentacje"],
            "Własne": []
        }

        tk.Label(frame, text="Podkategoria:").grid(row=3, column=0, padx=5, pady=5)
        self.subcategory_combobox = ttk.Combobox(frame, values=[], state="readonly")
        self.subcategory_combobox.grid(row=3, column=1, padx=5, pady=5)

        # Ustawienie widoczności dla własnej kategorii
        def update_subcategories(event):
            category = category_combobox.get()
            if category == "Własne":
                self.custom_category_entry.grid(row=2, column=1, padx=5, pady=5)
                self.custom_subcategory_entry.grid(row=4, column=1, padx=5, pady=5)
            else:
                self.custom_category_entry.grid_forget()
                self.custom_subcategory_entry.grid_forget()
                self.subcategory_combobox['values'] = subcategories.get(category, [])
                self.subcategory_combobox.set('')  # Resetowanie wybranej podkategorii

        category_combobox.bind("<<ComboboxSelected>>", update_subcategories)

        # Pole na własną kategorię i podkategorię
        self.custom_category_entry = tk.Entry(frame)
        self.custom_category_entry.grid_forget()  # Ukrywamy na początku

        self.custom_subcategory_entry = tk.Entry(frame)
        self.custom_subcategory_entry.grid_forget()  # Ukrywamy na początku

        # Zapisz raport
        tk.Button(frame, text="Zapisz raport", command=lambda: self.save_report(description_text.get("1.0", "end-1c"), category_combobox.get(), self.subcategory_combobox.get())).grid(row=5, column=0, columnspan=2, pady=10)

        # Refresh button
        tk.Button(frame, text="Refresh", command=self.refresh_ui).grid(row=6, column=0, columnspan=2, pady=10)

    def save_report(self, description, category, subcategory):
        # Jeśli użytkownik wybrał "Własne", używamy własnej kategorii i podkategorii
        if category == "Własne":
            category = self.custom_category_entry.get()
            subcategory = self.custom_subcategory_entry.get()

        # Zapisanie raportu
        report = {
            "description": description,
            "category": category,
            "subcategory": subcategory,
            "author": self.logged_in_user[1]
        }

        filename = f"raport_{self.logged_in_user[1]}.xlsx"
        generate_excel(report, filename)

        messagebox.showinfo("Sukces", "Raport zapisany do Excela!")

        # Odświeżenie podglądu po zapisaniu raportu
        self.view_reports()

    def view_reports(self):
        # Usunięcie istniejących widżetów (np. tabeli) w oknie
        for widget in self.root.winfo_children():
            widget.grid_forget()

        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        # Ładowanie danych z Excela
        filename = f"raport_{self.logged_in_user[1]}.xlsx"
        if not os.path.exists(filename):
            messagebox.showerror("Błąd", "Brak danych do wyświetlenia")
            return

        data = load_user_data(filename, self.logged_in_user[1])

        # Wyświetlenie danych w tabeli
        tree = ttk.Treeview(frame, columns=("LP", "Opis", "Kategoria", "Podkategoria", "Autor"), show="headings")
        tree.heading("LP", text="LP")
        tree.heading("Opis", text="Opis")
        tree.heading("Kategoria", text="Kategoria")
        tree.heading("Podkategoria", text="Podkategoria")
        tree.heading("Autor", text="Autor")

        for row in data:
            tree.insert("", "end", values=row)

        tree.pack()

    def refresh_ui(self):
        for widget in self.root.winfo_children():
            widget.grid_forget()

        # Ponownie wyświetlamy główny interfejs po kliknięciu Refresh
        self.show_main_ui()

root = tk.Tk()
app = App(root)
root.mainloop()
