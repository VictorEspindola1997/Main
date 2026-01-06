import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import json
import os
from datetime import datetime, timedelta

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FOCO TOTAL 2026")
        self.withdraw() # Hide main window until it's centered

        # --- Style ---
        self.style = ttk.Style(self)
        self.style.configure("Bold.TButton", font=("Helvetica", 10, "bold"))

        self.create_main_menu()
        self._center_window(self, 900, 700)
        self.deiconify() # Show window after centering

    def _center_window(self, window, width, height):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
        window.resizable(False, False)

    def create_main_menu(self):
        for widget in self.winfo_children():
            widget.destroy()

        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True, padx=200, fill='x')

        enter_button = ttk.Button(main_frame, text="ENTRAR NO SISTEMA", style="Bold.TButton", command=self.open_daily_tracker)
        enter_button.pack(pady=20, fill='x')

        history_button = ttk.Button(main_frame, text="HISTÓRICO", style="Bold.TButton", command=self.open_history)
        history_button.pack(pady=20, fill='x')

        exit_button = ttk.Button(main_frame, text="SAIR", style="Bold.TButton", command=self.destroy)
        exit_button.pack(pady=20, fill='x')

    def open_daily_tracker(self):
        tracker_window = tk.Toplevel(self)
        tracker_window.title("FOCO TOTAL 2026")
        self._center_window(tracker_window, 900, 700)

        notebook = ttk.Notebook(tracker_window)
        notebook.pack(expand=True, fill='both')

        self.days = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira"]
        days_short = ["SEG", "TER", "QUA", "QUI", "SEX"]
        today = datetime.now()
        start_of_week = today - timedelta(days=today.weekday())

        self.day_tabs = {}
        self.widgets = {}
        self.tab_dates = {} # Store date for each tab

        for i, day in enumerate(self.days):
            current_day_date = start_of_week + timedelta(days=i)
            date_str_display = current_day_date.strftime("%d/%m")
            date_str_save = current_day_date.strftime("%Y-%m-%d")
            tab_title = f"{date_str_display} - {days_short[i]}"

            # Create a canvas and a scrollbar for each day
            canvas = tk.Canvas(notebook, highlightthickness=0)
            scrollbar = ttk.Scrollbar(notebook, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Mouse wheel scrolling
            canvas.bind("<MouseWheel>", lambda event, c=canvas: c.yview_scroll(int(-1*(event.delta/120)), "units"))

            tab_id = notebook.add(canvas, text=tab_title)
            self.tab_dates[tab_id] = date_str_save # Store the YYYY-MM-DD date

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            self.day_tabs[day] = scrollable_frame
            self.widgets[day] = {}
            self._create_common_widgets(day)
            self._create_diet_widgets(day) # Add diet section to all days
            if i in [0, 2, 4]: # Monday, Wednesday, Friday
                self._create_training_widgets(day)
            if i == 4: # Friday
                self._create_friday_widgets(day)

        # Select previous working day
        today_index = datetime.now().weekday()
        # Monday (0) -> Friday (4)
        # Tuesday (1) -> Monday (0)
        # ...
        # Friday (4) -> Thursday (3)
        # Saturday (5) -> Friday (4)
        # Sunday (6) -> Friday (4)
        if today_index == 0 or today_index > 4: # If Monday or weekend
            previous_day_index = 4 # Friday
        else:
            previous_day_index = today_index - 1

        if 0 <= previous_day_index < 5:
            notebook.select(previous_day_index)

        # --- Action Buttons ---
        button_frame = ttk.Frame(tracker_window)
        button_frame.pack(pady=10)

        save_button = ttk.Button(button_frame, text="SALVAR", style="Bold.TButton", command=lambda nb=notebook: self._save_data(nb))
        save_button.pack(side='left', padx=10)

        extras_button = ttk.Button(button_frame, text="EXTRAS", style="Bold.TButton", command=self.open_extras)
        extras_button.pack(side='left', padx=10)

        exit_button = ttk.Button(button_frame, text="SAIR", style="Bold.TButton", command=tracker_window.destroy)
        exit_button.pack(side='left', padx=10)


    def open_extras(self):
        messagebox.showinfo("Extras", "Esta funcionalidade será implementada em breve!")


    def _update_int_scale(self, value, label, variable, suffix=""):
        rounded_value = round(float(value))
        variable.set(rounded_value)
        label.config(text=f"{rounded_value} {suffix}")

    def _update_float_scale(self, value, label, variable, suffix=""):
        rounded_value = round(float(value) * 2) / 2
        variable.set(rounded_value)
        label.config(text=f"{rounded_value:.1f} {suffix}")


    def _create_common_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        # --- Qualidade do sono ---
        ttk.Label(frame, text="Qualidade do sono:").pack(pady=(10,0), anchor='w')
        widgets['sono'] = tk.Text(frame, height=3, width=50)
        widgets['sono'].pack(pady=5, anchor='w')
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)


        # --- Meta de água ---
        ttk.Label(frame, text="Meta diária de água: 4 litros").pack(pady=(10,0), anchor='w')
        agua_frame = ttk.Frame(frame)
        agua_frame.pack(anchor='w', fill='x')
        agua_var = tk.DoubleVar(value=0.0)
        widgets['agua_var'] = agua_var
        agua_label = ttk.Label(agua_frame, text="0.0 L", width=6)
        agua_scale = ttk.Scale(agua_frame, from_=0, to=4, orient='horizontal', variable=agua_var,
                               command=lambda v, lbl=agua_label, var=agua_var: self._update_float_scale(v, lbl, var, "L"))
        agua_scale.pack(side='left', pady=5, fill='x', expand=True)
        agua_label.pack(side='left', padx=5)
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        # --- Meta de chá ---
        ttk.Label(frame, text="Meta diária de chá de hibisco: 1 litro").pack(pady=(10,0), anchor='w')
        cha_frame = ttk.Frame(frame)
        cha_frame.pack(anchor='w', fill='x')
        cha_var = tk.DoubleVar(value=0.0)
        widgets['cha_var'] = cha_var
        cha_label = ttk.Label(cha_frame, text="0.0 L", width=6)
        cha_scale = ttk.Scale(cha_frame, from_=0, to=1, orient='horizontal', variable=cha_var,
                              command=lambda v, lbl=cha_label, var=cha_var: self._update_float_scale(v, lbl, var, "L"))
        cha_scale.pack(side='left', pady=5, fill='x', expand=True)
        cha_label.pack(side='left', padx=5)
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)


        # --- 100 flexões ---
        ttk.Label(frame, text="E as 100 flexões?").pack(pady=(10,0), anchor='w')
        flexoes_var = tk.StringVar()
        widgets['flexoes_var'] = flexoes_var

        justificativa_frame = ttk.Frame(frame)

        def toggle_justificativa():
            if flexoes_var.get() == "nao":
                justificativa_frame.pack(pady=5, anchor='w', fill='x')
            else:
                if 'flexoes_justificativa' in widgets:
                    widgets['flexoes_justificativa'].delete('1.0', tk.END)
                justificativa_frame.pack_forget()

        ttk.Radiobutton(frame, text="Opa, claro que fiz!!!", variable=flexoes_var, value="sim", command=toggle_justificativa).pack(anchor='w')
        ttk.Radiobutton(frame, text="Putz, hoje não deu...", variable=flexoes_var, value="nao", command=toggle_justificativa).pack(anchor='w')
        flexoes_var.set("sim")

        ttk.Label(justificativa_frame, text="Justifique este absurdo:").pack(anchor='w')
        widgets['flexoes_justificativa'] = tk.Text(justificativa_frame, height=2, width=50)
        widgets['flexoes_justificativa'].pack(pady=5, anchor='w')


    def _create_training_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=20)

        # --- Relato do treino ---
        ttk.Label(frame, text="Relate como foi o seu treino com o personal:").pack(anchor='w')
        widgets['treino_relato'] = tk.Text(frame, height=4, width=50)
        widgets['treino_relato'].pack(pady=5, anchor='w')
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        # --- Nível de cansaço ---
        ttk.Label(frame, text="De 0 a 10, qual foi o nível de cansaço no treino de musculação?").pack(anchor='w')
        cansaco_frame = ttk.Frame(frame)
        cansaco_frame.pack(anchor='w', pady=5, fill='x')

        cansaco_var = tk.IntVar(value=0)
        widgets['treino_cansaco'] = cansaco_var

        cansaco_label = ttk.Label(cansaco_frame, text="Selecionado: 0", width=15)

        def update_cansaco_label():
            cansaco_label.config(text=f"Selecionado: {cansaco_var.get()}")

        radio_frame = ttk.Frame(cansaco_frame)
        for i in range(11):
            rb = ttk.Radiobutton(radio_frame, text=str(i), variable=cansaco_var, value=i, command=update_cansaco_label)
            rb.pack(side='left', padx=5) # Added padx for spacing

        radio_frame.pack(side='left')
        cansaco_label.pack(side='left', padx=10)
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        # --- Progressão de carga ---
        ttk.Label(frame, text="Houve progressão de carga?").pack(anchor='w')
        carga_var = tk.StringVar()
        widgets['treino_carga'] = carga_var
        ttk.Radiobutton(frame, text="Sim", variable=carga_var, value="sim").pack(anchor='w')
        ttk.Radiobutton(frame, text="Não", variable=carga_var, value="nao").pack(anchor='w')


    def _create_friday_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=20)

        # --- Peso em jejum ---
        ttk.Label(frame, text="Qual foi o peso em jejum?").pack(anchor='w')
        peso_frame = ttk.Frame(frame)
        peso_frame.pack(anchor='w', pady=5)

        # Simple validation to allow only numbers and a comma
        vcmd = (frame.register(self.validate_peso), '%P')
        widgets['peso'] = ttk.Entry(peso_frame, validate='key', validatecommand=vcmd, width=10)
        widgets['peso'].pack(side='left')
        ttk.Label(peso_frame, text="kg").pack(side='left', padx=5)

    def validate_peso(self, P):
        if len(P) > 5:
            return False
        if all(c in "0123456789," for c in P) and P.count(',') <= 1:
            return True
        return P == ""

    def _save_data(self, notebook):
        selected_day_index = notebook.index(notebook.select())
        day = self.days[selected_day_index] # Use the stored full day names

        data = {}
        day_widgets = self.widgets[day]

        # Common widgets
        data['sono'] = day_widgets['sono'].get('1.0', tk.END).strip()
        data['agua'] = day_widgets['agua_var'].get()
        data['cha'] = day_widgets['cha_var'].get()
        data['flexoes'] = day_widgets['flexoes_var'].get()
        if data['flexoes'] == 'nao':
            data['flexoes_justificativa'] = day_widgets['flexoes_justificativa'].get('1.0', tk.END).strip()

        # Training widgets
        if selected_day_index in [0, 2, 4]:
            data['treino_relato'] = day_widgets['treino_relato'].get('1.0', tk.END).strip()
            data['treino_cansaco'] = day_widgets['treino_cansaco'].get()
            data['treino_carga'] = day_widgets['treino_carga'].get()

        # Friday widgets
        if selected_day_index == 4:
            data['peso'] = day_widgets['peso'].get()

        # Diet widgets
        data['dieta'] = {}
        if 'dieta' in day_widgets:
            for meal, options in day_widgets['dieta'].items():
                data['dieta'][meal] = {}
                for option, combo in options.items():
                    data['dieta'][meal][option] = combo.get()

        # Save data
        file_path = '.foco_data.json'
        all_data = {}
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                try:
                    all_data = json.load(f)
                except json.JSONDecodeError:
                    pass # file is empty or corrupted

        selected_tab_id = notebook.select()
        save_date_str = self.tab_dates[selected_tab_id]

        all_data[save_date_str] = data

        with open(file_path, 'w') as f:
            json.dump(all_data, f, indent=4)

        # Hide the file on Windows
        if os.name == 'nt':
            try:
                os.system(f'attrib +h "{file_path}"')
            except Exception as e:
                print(f"Não foi possível ocultar o arquivo: {e}")

        messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")


    def _create_diet_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)

        diet_title_font = ("Helvetica", 12, "bold")
        ttk.Label(frame, text="DIETA:", font=diet_title_font).pack(anchor='w', pady=(10, 5))

        diet_options = {
            "Café da manhã": [
                ("Pão:", ["2 fatias de pão de forma", "1 pão francês"]),
                ("Proteína:", ["1 ovo", "2 fatias de muçarela"]),
                ("Fruta:", ["100g de uva", "100g de mamão", "100g de abacaxi", "150g de melancia", "150g de melão", "150g de morango", "100g de jabuticaba"])
            ],
            "Almoço": [
                ("Carboidrato:", ["50g de arroz", "50g de mandioca", "120g de batata"]),
                ("Proteína:", ["200g de peito de frango", "150g de patinho", "180g de atum"]),
                ("Leguminosa:", ["100g de feijão"]),
                ("Salada:", ["100g de vegetais e folhas"]),
                ("Sobremesa:", ["20g de paçoca"])
            ],
            "Lanche da tarde": [
                ("Base:", ["200g de iogurte natural desnatado", "220ml de leite desnatado"]),
                ("Fibra:", ["10g de psyllium"]),
                ("Complemento:", ["100g de morango", "70g de uva", "50g de whey"])
            ],
            "Jantar": [
                ("Carboidrato:", ["50g de arroz", "50g de macarrão", "120g de batata"]),
                ("Proteína:", ["200g de patinho", "200g de atum"]),
                ("Salada:", ["100g de vegetais e folhas"])
            ]
        }

        widgets['dieta'] = {}
        for meal, options in diet_options.items():
            meal_frame = ttk.LabelFrame(frame, text=meal, padding=10)
            meal_frame.pack(fill='x', expand=True, padx=5, pady=5)

            widgets['dieta'][meal] = {}
            for i, (label_text, option_list) in enumerate(options):
                ttk.Label(meal_frame, text=label_text).grid(row=i, column=0, sticky='w', padx=5, pady=5)
                combo = ttk.Combobox(meal_frame, values=option_list, state="readonly")
                combo.grid(row=i, column=1, sticky='ew', padx=5, pady=5)
                meal_frame.grid_columnconfigure(1, weight=1)
                widgets['dieta'][meal][label_text.replace(":", "")] = combo


    def open_history(self):
        file_path = '.foco_data.json'
        if not os.path.exists(file_path):
            messagebox.showinfo("Histórico Vazio", "Ainda não há nenhum registro salvo.")
            return

        with open(file_path, 'r') as f:
            try:
                all_data = json.load(f)
            except json.JSONDecodeError:
                messagebox.showinfo("Histórico Vazio", "O arquivo de histórico está vazio ou corrompido.")
                return

        history_window = tk.Toplevel(self)
        history_window.title("FOCO TOTAL 2026")
        self._center_window(history_window, 900, 700)

        # Frame for buttons and a potential scrollbar
        canvas = tk.Canvas(history_window)
        scrollbar = ttk.Scrollbar(history_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        sorted_dates = sorted(all_data.keys(), reverse=True)

        for date_str in sorted_dates:
            btn = ttk.Button(scrollable_frame, text=date_str, style="Bold.TButton",
                             command=lambda d=date_str, data=all_data[date_str]: self._show_history_details(d, data))
            btn.pack(pady=5, padx=10, fill='x')

        exit_button = ttk.Button(history_window, text="SAIR", style="Bold.TButton", command=history_window.destroy)
        exit_button.pack(pady=10)


    def _show_history_details(self, date_str, data):
        details_window = tk.Toplevel(self)
        details_window.title("FOCO TOTAL 2026")
        self._center_window(details_window, 900, 700)

        # Create a canvas and a scrollbar
        canvas = tk.Canvas(details_window, highlightthickness=0)
        scrollbar = ttk.Scrollbar(details_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.bind("<MouseWheel>", lambda event, c=canvas: c.yview_scroll(int(-1*(event.delta/120)), "units"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        display_text = ""
        # Using a dictionary to map keys to user-friendly labels
        labels = {
            "sono": "Qualidade do Sono",
            "dentes": "Escovações de Dentes",
            "agua": "Meta de Água (Litros)",
            "cha": "Meta de Chá (Litros)",
            "flexoes": "Fez as 100 flexões?",
            "flexoes_justificativa": "Justificativa (Flexões)",
            "treino_relato": "Relato do Treino",
            "treino_cansaco": "Nível de Cansaço (Treino)",
            "treino_carga": "Houve Progressão de Carga?",
            "peso": "Peso em Jejum (kg)"
        }

        # General data first
        for key, value in data.items():
            if key == 'dieta': continue # Skip diet for now

            label_text = labels.get(key, key.replace("_", " ").title())

            if key in ['flexoes', 'treino_carga']:
                value = "Sim" if value == 'sim' else "Não"

            ttk.Label(scrollable_frame, text=f"{label_text}:", font=("Helvetica", 10, "bold")).pack(anchor='w', pady=(10, 0))
            ttk.Label(scrollable_frame, text=f"  {value}").pack(anchor='w')

        # Diet data
        if 'dieta' in data and data['dieta']:
            ttk.Label(scrollable_frame, text="DIETA:", font=("Helvetica", 11, "bold")).pack(anchor='w', pady=(10, 5))
            for meal, options in data['dieta'].items():
                ttk.Label(scrollable_frame, text=f"{meal}:").pack(anchor='w', pady=(5, 0))
                for item, choice in options.items():
                    if choice:
                        ttk.Label(scrollable_frame, text=f"  - {item}: {choice}").pack(anchor='w')

        exit_button = ttk.Button(details_window, text="SAIR", style="Bold.TButton", command=details_window.destroy)
        exit_button.pack(pady=10)


if __name__ == "__main__":
    app = App()
    app.mainloop()
