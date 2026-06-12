import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

import pandas as pd


class ExcelAnalyzerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Analyzer")
        self.root.geometry("1100x750")
        self.root.minsize(900, 650)

        # ---------- COLORS ----------
        self.bg_main = "#2b2b2b"
        self.bg_frame = "#3c3f41"
        self.bg_input = "#4a4a4a"
        self.text_color = "#e0e0e0"
        self.red_main = "#9e3a3a"
        self.red_hover = "#c04e4e"
        self.red_dark = "#8b3a3a"
        self.border_red = "#b34b4b"
        self.green_main = "#3a6e3a"
        self.green_hover = "#4e8a4e"

        self.root.configure(bg=self.bg_main)

        self.df = None
        self.file_path = None
        self.custom_column_name = "ДКП (Смерти T4\u00d730 + Смерти T5\u00d730 + Убийства T4\u00d710 + Убийства T5\u00d720)"

        # ---------- AVAILABLE COLUMNS ----------
        self.available_columns = [
            "Мощь",
            "Макс. мощь",
            "Имя пользователя",
            "Ник",
            "Игрок",

            # Общие PvP
            "Суммарные очки убийств",

            # Убийства
            "Убийства T5",
            "Убийства T4",
            "Убийства T3",
            "Убийства T2",
            "Убийства T1",

            # Смерти
            "Смерти T5",
            "Смерти T4",
            "Смерти T3",
            "Смерти T2",
            "Смерти T1",

            # Автоматические суммы
            "Общие убийства",
            "Общие смерти",

            # Прочее
            "Собранные ресурсы",
            "Помощь альянсу",
            "Вылечено"
        ]

        # ---------- STYLE ----------
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Dark.TFrame", background=self.bg_main)

        style.configure("Dark.TLabel", background=self.bg_main, foreground=self.text_color, font=("Segoe UI", 10))

        style.configure("Dark.Treeview",
            background=self.bg_frame,
            foreground=self.text_color,
            fieldbackground=self.bg_frame,
            bordercolor=self.border_red,
            borderwidth=2,
            font=("Segoe UI", 10),
            rowheight=26
        )
        style.configure("Dark.Treeview.Heading",
            background=self.red_dark,
            foreground="white",
            font=("Segoe UI", 10, "bold"),
            borderwidth=1,
            relief="solid"
        )
        style.map("Dark.Treeview.Heading",
            background=[("active", self.red_hover)]
        )
        style.map("Dark.Treeview",
            background=[("selected", self.red_main)],
            foreground=[("selected", "white")]
        )

        style.configure("Dark.TCheckbutton", background=self.bg_main, foreground=self.text_color, font=("Segoe UI", 10))
        style.map("Dark.TCheckbutton", background=[("active", self.bg_main)], foreground=[("active", self.text_color)])

        style.configure("Dark.TCombobox",
            fieldbackground=self.bg_input,
            background=self.bg_input,
            foreground=self.text_color,
            arrowcolor=self.text_color,
            bordercolor=self.red_dark,
            lightcolor=self.red_dark,
            darkcolor=self.red_dark
        )

        # ---------- VARIABLES ----------
        self.column_var = tk.StringVar(value="Мощь")
        self.calc_mode_var = tk.StringVar(value="simple")
        self.export_enabled = tk.BooleanVar(value=False)
        self.multiplier_var = tk.IntVar(value=1)

        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()

        self.min_enabled = tk.BooleanVar(value=False)
        self.max_enabled = tk.BooleanVar(value=False)

        self.min_var = tk.StringVar()
        self.max_var = tk.StringVar()

        self.sort_enabled = tk.BooleanVar(value=True)

        self.sum_var = tk.StringVar(value="")

        # ---------- VALIDATION ----------
        vcmd = (self.root.register(self.validate_int_input), "%P")

        # ---------- MAIN FRAME ----------
        main_frame = ttk.Frame(self.root, style="Dark.TFrame", padding=15)
        main_frame.pack(fill="both", expand=True)

        # ---------- ROW 1 ----------
        row1 = ttk.Frame(main_frame, style="Dark.TFrame")
        row1.pack(fill="x", pady=(0, 12))

        self.open_btn = tk.Button(row1, text="Открыть Excel", command=self.load_file,
            bg=self.red_main, fg="white", activebackground=self.red_hover, activeforeground="white",
            relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=6, cursor="hand2")
        self.open_btn.pack(side="left")

        self.file_label = ttk.Label(row1, text="Файл не выбран", style="Dark.TLabel")
        self.file_label.pack(side="left", padx=15)

        # ---------- ROW 2 ----------
        row2 = ttk.Frame(main_frame, style="Dark.TFrame")
        row2.pack(fill="x", pady=(0, 12))

        ttk.Label(row2, text="Режим расчета:", style="Dark.TLabel").pack(side="left")

        self.simple_radio = tk.Radiobutton(row2, text="Стандартный (выбор столбца)",
            variable=self.calc_mode_var, value="simple",
            bg=self.bg_main, fg=self.text_color, selectcolor=self.bg_main,
            activebackground=self.bg_main, activeforeground=self.text_color,
            command=self.toggle_column_combo)
        self.simple_radio.pack(side="left", padx=(10, 5))

        self.formula_radio = tk.Radiobutton(row2, text="Расчет ДКП (Смерти T4\u00d730 + Смерти T5\u00d730 + Убийства T4\u00d710 + Убийства T5\u00d720)",
            variable=self.calc_mode_var, value="formula",
            bg=self.bg_main, fg=self.text_color, selectcolor=self.bg_main,
            activebackground=self.bg_main, activeforeground=self.text_color,
            command=self.toggle_column_combo)
        self.formula_radio.pack(side="left", padx=(10, 0))

        # ---------- ROW 3 ----------
        row3 = ttk.Frame(main_frame, style="Dark.TFrame")
        row3.pack(fill="x", pady=(0, 12))

        ttk.Label(row3, text="Столбец для анализа:", style="Dark.TLabel").pack(side="left")

        self.column_combo = ttk.Combobox(row3, textvariable=self.column_var,
            values=self.available_columns, state="readonly", width=40, style="Dark.TCombobox")
        self.column_combo.pack(side="left", padx=12)

        # ---------- ROW 4 ----------
        row4 = ttk.Frame(main_frame, style="Dark.TFrame")
        row4.pack(fill="x", pady=(0, 12))

        self.sort_check = ttk.Checkbutton(row4, text="Сортировать по убыванию",
            variable=self.sort_enabled, style="Dark.TCheckbutton")
        self.sort_check.pack(side="left", padx=(15, 0))

        ttk.Label(row4, text="   Множитель нормы ДКП:", style="Dark.TLabel").pack(side="left", padx=(20, 5))

        self.multiplier_spin = tk.Spinbox(row4, from_=1, to=20, textvariable=self.multiplier_var, width=5,
            bg=self.bg_input, fg=self.text_color, insertbackground=self.text_color,
            buttonbackground=self.red_dark, relief="flat", font=("Segoe UI", 10),
            justify="center")
        self.multiplier_spin.pack(side="left")

        # ---------- ROW 5 ----------
        row5 = ttk.Frame(main_frame, style="Dark.TFrame")
        row5.pack(fill="x", pady=(0, 12))

        ttk.Label(row5, text="Диапазон мест:", style="Dark.TLabel").pack(side="left")

        ttk.Label(row5, text="С места", style="Dark.TLabel").pack(side="left", padx=(20, 5))

        self.from_entry = tk.Entry(row5, textvariable=self.from_var, validate="key", validatecommand=vcmd,
            bg=self.bg_input, fg=self.text_color, insertbackground=self.text_color, relief="flat", width=12)
        self.from_entry.pack(side="left")

        ttk.Label(row5, text="По место", style="Dark.TLabel").pack(side="left", padx=(20, 5))

        self.to_entry = tk.Entry(row5, textvariable=self.to_var, validate="key", validatecommand=vcmd,
            bg=self.bg_input, fg=self.text_color, insertbackground=self.text_color, relief="flat", width=12)
        self.to_entry.pack(side="left")

        # ---------- ROW 6 ----------
        row6 = ttk.Frame(main_frame, style="Dark.TFrame")
        row6.pack(fill="x", pady=(0, 12))

        self.min_check = ttk.Checkbutton(row6, text="Нижняя граница (\u2265)",
            variable=self.min_enabled, command=self.toggle_min_entry, style="Dark.TCheckbutton")
        self.min_check.pack(side="left", padx=(15, 10))

        self.min_entry = tk.Entry(row6, textvariable=self.min_var, validate="key", validatecommand=vcmd,
            bg=self.bg_input, fg=self.text_color, insertbackground=self.text_color, relief="flat", width=20, state="disabled")
        self.min_entry.pack(side="left")

        # ---------- ROW 7 ----------
        row7 = ttk.Frame(main_frame, style="Dark.TFrame")
        row7.pack(fill="x", pady=(0, 18))

        self.max_check = ttk.Checkbutton(row7, text="Верхняя граница (\u2264)",
            variable=self.max_enabled, command=self.toggle_max_entry, style="Dark.TCheckbutton")
        self.max_check.pack(side="left", padx=(15, 10))

        self.max_entry = tk.Entry(row7, textvariable=self.max_var, validate="key", validatecommand=vcmd,
            bg=self.bg_input, fg=self.text_color, insertbackground=self.text_color, relief="flat", width=20, state="disabled")
        self.max_entry.pack(side="left")

        # ---------- ROW 8 ----------
        row8 = ttk.Frame(main_frame, style="Dark.TFrame")
        row8.pack(fill="x", pady=(0, 12))

        self.export_check = ttk.Checkbutton(row8, text="Экспортировать данные по каждому игроку в TXT (в папку с программой)",
            variable=self.export_enabled, style="Dark.TCheckbutton")
        self.export_check.pack(side="left", padx=(15, 0))

        # ---------- ROW 9 ----------
        row9 = ttk.Frame(main_frame, style="Dark.TFrame")
        row9.pack(fill="x", pady=(0, 20))

        self.calc_btn = tk.Button(row9, text="Рассчитать сумму",
            command=self.calculate,
            bg=self.red_main, fg="white", activebackground=self.red_hover, activeforeground="white",
            relief="flat", font=("Segoe UI", 11, "bold"), padx=20, pady=10, cursor="hand2")
        self.calc_btn.pack()

        self.calc_btn.bind("<Enter>", self.on_calc_hover)
        self.calc_btn.bind("<Leave>", self.on_calc_leave)

        # ---------- ROW 10 ----------
        row10 = ttk.Frame(main_frame, style="Dark.TFrame")
        row10.pack(fill="both", expand=True)

        table_container = tk.Frame(row10, bg=self.border_red, bd=2)
        table_container.pack(fill="both", expand=True)

        columns = ("place", "player", "power", "dkp", "norm", "percent")
        self.tree = ttk.Treeview(table_container, columns=columns, show="headings",
            style="Dark.Treeview", height=20)

        self.tree.heading("place", text="\u2116")
        self.tree.heading("player", text="Игрок")
        self.tree.heading("power", text="Мощь")
        self.tree.heading("dkp", text="ДКП")
        self.tree.heading("norm", text="Норма (\u041c\u043e\u0449\u044c\u00d7N)")
        self.tree.heading("percent", text="% выполнения")

        self.tree.column("place", width=50, anchor="center")
        self.tree.column("player", width=200, anchor="w")
        self.tree.column("power", width=130, anchor="center")
        self.tree.column("dkp", width=130, anchor="center")
        self.tree.column("norm", width=130, anchor="center")
        self.tree.column("percent", width=110, anchor="center")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Sum label below table
        self.sum_label = ttk.Label(row10, textvariable=self.sum_var, style="Dark.TLabel")
        self.sum_label.pack(anchor="e", pady=(8, 0))

        self.root.mainloop()

    # ---------- UI ----------
    def on_calc_hover(self, event):
        self.calc_btn.configure(bg=self.red_hover)

    def on_calc_leave(self, event):
        self.calc_btn.configure(bg=self.red_main)

    def toggle_min_entry(self):
        if self.min_enabled.get():
            self.min_entry.config(state="normal")
        else:
            self.min_entry.config(state="disabled")
            self.min_var.set("")

    def toggle_max_entry(self):
        if self.max_enabled.get():
            self.max_entry.config(state="normal")
        else:
            self.max_entry.config(state="disabled")
            self.max_var.set("")

    def toggle_column_combo(self):
        if self.calc_mode_var.get() == "formula":
            self.column_combo.config(state="disabled")
        else:
            self.column_combo.config(state="readonly")

    # ---------- HELPERS ----------
    def validate_int_input(self, value):
        if value == "":
            return True
        return value.isdigit()

    def format_number(self, number):
        try:
            return f"{int(number):,}".replace(",", " ")
        except Exception:
            return "0"
    
    def format_number_raw(self, number):
        try:
            return f"{int(number):,}".replace(",", "")
        except Exception:
            return "0"
    
    def get_player_name(self, row):
        name_cols = ["Имя пользователя", "Ник", "Игрок", "Nick", "Player", "Name", "Имя"]
        for col in name_cols:
            if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
                return str(row[col]).strip()
        return "Неизвестный"

    # ---------- LOAD FILE ----------
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)

            if df.empty:
                messagebox.showerror("Ошибка", "Excel файл пуст.")
                return

            self.df = df
            self.file_path = file_path

            rows_count = len(df.index)

            cols_found = ", ".join(list(df.columns)[:10])
            if len(df.columns) > 10:
                cols_found += f" и еще {len(df.columns)-10}"

            self.file_label.config(text=f"{os.path.basename(file_path)} | Строк данных: {rows_count}")

            self.update_column_list()

            messagebox.showinfo("Успешно", f"Файл успешно загружен.\n\nНайдены столбцы:\n{cols_found}")

        except Exception as e:
            messagebox.showerror("Ошибка загрузки", f"Не удалось открыть файл:\n{e}")

    # ---------- UPDATE COLUMNS ----------
    def update_column_list(self):
        if self.df is None:
            return

        death_cols = ["Смерти T5", "Смерти T4", "Смерти T3", "Смерти T2", "Смерти T1"]
        kill_cols = ["Убийства T5", "Убийства T4", "Убийства T3", "Убийства T2", "Убийства T1"]

        existing_death_cols = [col for col in death_cols if col in self.df.columns]
        existing_kill_cols = [col for col in kill_cols if col in self.df.columns]

        if existing_death_cols:
            self.df["Общие смерти"] = (
                self.df[existing_death_cols]
                .apply(pd.to_numeric, errors="coerce")
                .fillna(0)
                .sum(axis=1)
            )

        if existing_kill_cols:
            self.df["Общие убийства"] = (
                self.df[existing_kill_cols]
                .apply(pd.to_numeric, errors="coerce")
                .fillna(0)
                .sum(axis=1)
            )

        existing_columns = [col for col in self.available_columns if col in self.df.columns]

        self.column_combo["values"] = existing_columns

        if "Мощь" in existing_columns:
            self.column_var.set("Мощь")
        elif existing_columns:
            self.column_var.set(existing_columns[0])

    # ---------- CALCULATE FORMULA COLUMN ----------
    def calculate_formula_column(self, df):
        required_cols = ["Смерти T4", "Смерти T5", "Убийства T4", "Убийства T5"]
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            raise Exception(f"Для расчета ДКП требуются столбцы: {', '.join(missing_cols)}")

        deaths_t4 = pd.to_numeric(df["Смерти T4"], errors="coerce").fillna(0)
        deaths_t5 = pd.to_numeric(df["Смерти T5"], errors="coerce").fillna(0)
        kills_t4 = pd.to_numeric(df["Убийства T4"], errors="coerce").fillna(0)
        kills_t5 = pd.to_numeric(df["Убийства T5"], errors="coerce").fillna(0)

        result = (deaths_t4 + deaths_t5) * 30 + kills_t4 * 10 + kills_t5 * 20
        return result

    # ---------- EXPORT TO TXT ----------
    def export_players_data(self, data, column, start_place, end_place, multiplier, has_power):
        try:
            program_dir = os.path.dirname(os.path.abspath(__file__))
            filename = f"DKP_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            filepath = os.path.join(program_dir, filename)
            
            data_sorted = data.sort_values(by=column, ascending=False).reset_index(drop=True)
            
            lines = []
            lines.append(f"=== ЭКСПОРТ ДКП | {datetime.now().strftime('%d.%m.%Y %H:%M:%S')} ===")
            lines.append(f"Диапазон мест: с {start_place} по {end_place}")
            lines.append(f"Всего игроков: {len(data_sorted)}")
            lines.append(f"Множитель нормы ДКП: \u00d7{multiplier}")
            lines.append("=" * 180)
            lines.append("")
            
            for idx, row in data_sorted.iterrows():
                player_name = self.get_player_name(row)
                
                kills_t4 = 0
                if "Убийства T4" in row.index and pd.notna(row["Убийства T4"]):
                    kills_t4 = int(float(row["Убийства T4"]))
                
                kills_t5 = 0
                if "Убийства T5" in row.index and pd.notna(row["Убийства T5"]):
                    kills_t5 = int(float(row["Убийства T5"]))
                
                kills_total = kills_t4 + kills_t5
                
                deaths_t4 = 0
                if "Смерти T4" in row.index and pd.notna(row["Смерти T4"]):
                    deaths_t4 = int(float(row["Смерти T4"]))
                
                deaths_t5 = 0
                if "Смерти T5" in row.index and pd.notna(row["Смерти T5"]):
                    deaths_t5 = int(float(row["Смерти T5"]))
                
                deaths_total = deaths_t4 + deaths_t5
                
                healed = ""
                for col_name in ["Вылечено", "Healed", "Лечение"]:
                    if col_name in row.index and pd.notna(row[col_name]):
                        healed_val = int(float(row[col_name]))
                        healed = self.format_number_raw(healed_val)
                        break
                
                dkp_value = 0
                if column in row.index and pd.notna(row[column]):
                    dkp_value = int(float(row[column]))
                
                power_value = 0
                if has_power and "Мощь" in row.index and pd.notna(row["Мощь"]):
                    power_value = int(float(row["Мощь"]))
                
                norm_value = power_value * multiplier if has_power else 0
                percent = (dkp_value / norm_value * 100) if has_power and norm_value > 0 else 0.0
                
                kills_t4_str = self.format_number(kills_t4)
                kills_t5_str = self.format_number(kills_t5)
                kills_total_str = self.format_number(kills_total)
                deaths_t4_str = self.format_number(deaths_t4)
                deaths_t5_str = self.format_number(deaths_t5)
                deaths_total_str = self.format_number(deaths_total)
                dkp_str = self.format_number(dkp_value)
                power_str = self.format_number(power_value)
                norm_str = self.format_number(norm_value)
                
                place_num = idx + 1
                line = (
                    f"# {place_num} | Ник: {player_name} | Мощь: {power_str} | "
                    f"Убийства T4: {kills_t4_str} | Убийства T5: {kills_t5_str} | "
                    f"Всего убийств (T4+T5): {kills_total_str} | "
                    f"Смерти T4: {deaths_t4_str} | Смерти T5: {deaths_t5_str} | "
                    f"Всего смертей (T4+T5): {deaths_total_str} | "
                    f"Вылечено: {healed} | ДКП: {dkp_str} | "
                    f"Норма (Мощь\u00d7{multiplier}): {norm_str} | % выполнения: {percent:.2f}%"
                )
                lines.append(line)
            
            lines.append("")
            lines.append("=" * 180)
            lines.append(f"Сумма ДКП: {self.format_number(data_sorted[column].sum())}")
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            
            return filepath, len(data_sorted)
            
        except Exception as e:
            raise Exception(f"Ошибка при экспорте: {e}")

    # ---------- CALCULATE ----------
    def calculate(self):
        if self.df is None:
            messagebox.showwarning("Ошибка", "Сначала выберите Excel файл.")
            return

        try:
            data = self.df.copy()

            # ---------- ВЫБОР РЕЖИМА РАСЧЕТА ----------
            if self.calc_mode_var.get() == "formula":
                try:
                    formula_series = self.calculate_formula_column(data)
                    data[self.custom_column_name] = formula_series
                    column = self.custom_column_name
                except Exception as e:
                    messagebox.showerror("Ошибка формулы",
                        f"Не удалось вычислить ДКП:\n{e}\n\nУбедитесь, что в файле есть столбцы:\n- Смерти T4\n- Смерти T5\n- Убийства T4\n- Убийства T5")
                    return
            else:
                column = self.column_var.get()
                if column not in data.columns:
                    messagebox.showerror("Ошибка", f"Столбец '{column}' не найден.")
                    return

            # ---------- ПОДГОТОВКА ДАННЫХ ----------
            numeric_series = pd.to_numeric(data[column], errors="coerce")

            if numeric_series.isna().all():
                self.sum_var.set("Нет числовых данных")
                return

            data[column] = numeric_series
            data = data.dropna(subset=[column])

            # ---------- СОРТИРОВКА ----------
            if self.sort_enabled.get():
                data = data.sort_values(by=column, ascending=False)

            data = data.reset_index(drop=True)
            total_rows = len(data)

            if total_rows == 0:
                messagebox.showinfo("Результат", "Нет данных")
                return

            # ---------- ДИАПАЗОН МЕСТ ----------
            try:
                from_place = int(self.from_var.get()) if self.from_var.get() else 1
            except Exception:
                from_place = 1

            try:
                to_place = int(self.to_var.get()) if self.to_var.get() else total_rows
            except Exception:
                to_place = total_rows

            if from_place <= 0:
                from_place = 1
            if to_place <= 0:
                to_place = total_rows
            if from_place > total_rows:
                from_place = total_rows
            if to_place > total_rows:
                to_place = total_rows
            if from_place > to_place:
                from_place, to_place = to_place, from_place

            start_idx = from_place - 1
            end_idx = to_place

            filtered_data = data.iloc[start_idx:end_idx].copy()

            # ---------- ФИЛЬТР ПО ГРАНИЦАМ ----------
            if self.min_enabled.get():
                try:
                    min_value = int(self.min_var.get()) if self.min_var.get() else 0
                    filtered_data = filtered_data[filtered_data[column] >= min_value]
                except Exception:
                    pass

            if self.max_enabled.get():
                try:
                    max_value = int(self.max_var.get()) if self.max_var.get() else 0
                    filtered_data = filtered_data[filtered_data[column] <= max_value]
                except Exception:
                    pass

            if filtered_data.empty:
                messagebox.showinfo("Результат", "Нет данных в выбранном диапазоне")
                return

            # ---------- ЗАПОЛНЕНИЕ ТАБЛИЦЫ ----------
            multiplier = self.multiplier_var.get()
            has_power = "Мощь" in data.columns
            result_sum = filtered_data[column].sum()

            for item in self.tree.get_children():
                self.tree.delete(item)

            for idx, row in filtered_data.iterrows():
                place = idx + 1
                player_name = self.get_player_name(row)
                
                dkp_value = 0
                if column in row.index and pd.notna(row[column]):
                    dkp_value = int(float(row[column]))
                
                power_value = 0
                if has_power and "Мощь" in row.index and pd.notna(row["Мощь"]):
                    power_value = int(float(row["Мощь"]))
                
                norm_value = power_value * multiplier if has_power else 0
                percent = (dkp_value / norm_value * 100) if has_power and norm_value > 0 else 0.0
                
                self.tree.insert("", "end", values=(
                    place,
                    player_name,
                    self.format_number(power_value) if has_power else "N/A",
                    self.format_number(dkp_value),
                    self.format_number(norm_value) if has_power else "N/A",
                    f"{percent:.2f}%" if has_power else "N/A"
                ))

            self.sum_var.set(
                f"Сумма ДКП: {self.format_number(result_sum)}   |   Всего игроков: {len(filtered_data)}"
            )

            # ---------- ЭКСПОРТ В TXT ----------
            if self.export_enabled.get() and self.calc_mode_var.get() == "formula":
                try:
                    export_file, exported_count = self.export_players_data(
                        filtered_data, column, from_place, to_place, multiplier, has_power
                    )
                    messagebox.showinfo("Экспорт выполнен",
                        f"Экспортировано {exported_count} игроков\n\nФайл: {export_file}")
                except Exception as e:
                    messagebox.showwarning("Ошибка экспорта",
                        f"Не удалось экспортировать данные:\n{e}")
            elif self.export_enabled.get() and self.calc_mode_var.get() != "formula":
                messagebox.showinfo("Информация",
                    "Экспорт в TXT доступен только в режиме 'Расчет ДКП'")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при расчёте:\n{e}")


if __name__ == "__main__":
    ExcelAnalyzerApp()
