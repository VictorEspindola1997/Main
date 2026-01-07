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
        self.style.configure("Error.TCombobox", fieldbackground="red")
        self.style.map("Error.TCombobox", fieldbackground=[("readonly", "red")])
        self.style.configure("Error.TEntry", fieldbackground="red")

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
            date_str_save = current_day_date.strftime("%Y-%m-%d")
            date_str_display = current_day_date.strftime("%d/%m")
            tab_title = f"{date_str_display} - {days_short[i]}"

            # Create a container frame for each tab
            tab_frame = ttk.Frame(notebook)
            notebook.add(tab_frame, text=tab_title)

            # Store the date against the tab's ID (which is the frame itself)
            self.tab_dates[str(tab_frame)] = date_str_save

            if current_day_date.date() >= today.date():
                notebook.tab(tab_frame, state='disabled')

            # Create a canvas and a scrollbar inside the container frame
            canvas = tk.Canvas(tab_frame, highlightthickness=0)
            scrollbar = ttk.Scrollbar(tab_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas, padding=(20, 10))

            scrollable_frame.bind(
                "<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Mouse wheel scrolling - bind to the canvas of this tab
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            self.day_tabs[day] = scrollable_frame
            self.widgets[day] = {}

            # Call all widget creation functions in the correct order
            self._create_sleep_widgets(day)
            if day == "Sexta-feira":
                self._create_friday_widgets(day)
            self._create_common_widgets(day)
            self._create_diet_widgets(day)
            is_training_day = day in ["Segunda-feira", "Quarta-feira", "Sexta-feira"]
            is_tue_thu = day in ["Terça-feira", "Quinta-feira"]
            self._create_training_and_flex_widgets(day, is_training_day, is_tue_thu)

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

        # Universal scroll for all widgets inside the notebook
        def _universal_scroll(event):
            try:
                # Check if notebook still exists to prevent errors on window close
                if not notebook.winfo_exists():
                    return "break"
                active_tab_widget = notebook.nametowidget(notebook.select())
                if active_tab_widget.winfo_children():
                    canvas = active_tab_widget.winfo_children()[0]
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                return "break"
            except (tk.TclError, AttributeError):
                # This can happen if the widget is in the process of being destroyed
                return "break"

        # Bind globally, but we will manage it with on_close
        self.bind_all("<MouseWheel>", _universal_scroll)

        def on_close():
            # Unbind the global scroll event before destroying the window
            self.unbind_all("<MouseWheel>")
            tracker_window.destroy()

        tracker_window.protocol("WM_DELETE_WINDOW", on_close)

        # --- Action Buttons ---
        button_frame = ttk.Frame(tracker_window)
        button_frame.pack(pady=10)

        save_button = ttk.Button(button_frame, text="SALVAR", style="Bold.TButton", command=lambda nb=notebook: self._save_data(nb))
        save_button.pack(side='left', padx=10)

        extras_button = ttk.Button(button_frame, text="EXTRAS", style="Bold.TButton", command=self.open_extras)
        extras_button.pack(side='left', padx=10)

        exit_button = ttk.Button(button_frame, text="SAIR", style="Bold.TButton", command=on_close)
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


    def _create_sleep_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        title_font = ("Helvetica", 12, "bold")
        title_label = ttk.Label(frame, text="SONO", font=title_font)
        title_label.pack(anchor='w', pady=(10, 5))

        question_label = ttk.Label(frame, text="Qualidade do sono:")
        question_label.pack(pady=(10,0), anchor='w')

        widgets['sono'] = tk.Text(frame, height=3, width=50)
        widgets['sono'].pack(pady=5, anchor='w')

        separator = ttk.Separator(frame, orient='horizontal')
        separator.pack(fill='x', pady=10)


    def _create_common_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]
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


    def _create_training_and_flex_widgets(self, day, is_training_day, is_tue_thu):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        title_font = ("Helvetica", 12, "bold")
        ttk.Label(frame, text="TREINO", font=title_font).pack(anchor='w', pady=(10, 5))

        if is_training_day:
            self._create_training_widgets(day)

        if is_tue_thu:
            self._create_tue_thu_training_widgets(day)

        self._create_flex_widgets(day)

    def _create_tue_thu_training_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        ttk.Label(frame, text="Treinou hoje?").pack(anchor='w', pady=(10, 0))
        treino_var = tk.StringVar()
        widgets['tue_thu_treino_var'] = treino_var

        sim_frame = ttk.Frame(frame)
        nao_frame = ttk.Frame(frame)

        def toggle_treino_details():
            if treino_var.get() == "sim":
                sim_frame.pack(pady=5, anchor='w', fill='x')
                nao_frame.pack_forget()
            elif treino_var.get() == "nao":
                nao_frame.pack(pady=5, anchor='w', fill='x')
                sim_frame.pack_forget()

        ttk.Radiobutton(frame, text="Sim", variable=treino_var, value="sim", command=toggle_treino_details).pack(anchor='w')
        ttk.Radiobutton(frame, text="Não", variable=treino_var, value="nao", command=toggle_treino_details).pack(anchor='w')

        # --- Sim Frame ---
        ttk.Label(sim_frame, text="Descreva como foi:").pack(anchor='w')
        widgets['tue_thu_treino_desc'] = tk.Text(sim_frame, height=3, width=50)
        widgets['tue_thu_treino_desc'].pack(pady=5, anchor='w')

        # --- Não Frame ---
        ttk.Label(nao_frame, text="Justifique-se...").pack(anchor='w')
        widgets['tue_thu_treino_just'] = tk.Text(nao_frame, height=3, width=50)
        widgets['tue_thu_treino_just'].pack(pady=5, anchor='w')


    def _create_flex_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

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

        ttk.Label(justificativa_frame, text="Justifique este absurdo:").pack(anchor='w')
        widgets['flexoes_justificativa'] = tk.Text(justificativa_frame, height=2, width=50)
        widgets['flexoes_justificativa'].pack(pady=5, anchor='w')

    def _create_training_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        # --- Relato do treino ---
        ttk.Label(frame, text="Relate como foi o seu treino com o personal:").pack(anchor='w')
        widgets['treino_relato'] = tk.Text(frame, height=4, width=50)
        widgets['treino_relato'].pack(pady=5, anchor='w')
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        # --- Nível de cansaço ---
        ttk.Label(frame, text="De 0 a 10, qual foi o nível de cansaço no treino de musculação?").pack(anchor='w')
        cansaco_frame = ttk.Frame(frame)
        cansaco_frame.pack(anchor='w', pady=5, fill='x')

        cansaco_var = tk.IntVar(value=-1)
        widgets['treino_cansaco'] = cansaco_var

        cansaco_label = ttk.Label(cansaco_frame, text="Selecionado: --", width=15)

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

        title_font = ("Helvetica", 12, "bold")
        ttk.Label(frame, text="PESO", font=title_font).pack(anchor='w', pady=(10, 5))

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
        day = self.days[selected_day_index]
        day_widgets = self.widgets[day]

        # --- Reset all styles first ---
        day_widgets['sono'].config(highlightbackground="grey", highlightcolor="grey", highlightthickness=1)
        day_widgets['treino_relato'].config(highlightbackground="grey", highlightcolor="grey", highlightthickness=1) if 'treino_relato' in day_widgets else None
        day_widgets['flexoes_justificativa'].config(highlightbackground="grey", highlightcolor="grey", highlightthickness=1) if 'flexoes_justificativa' in day_widgets else None
        if day == "Sexta-feira":
            day_widgets['peso'].config(style="TEntry")
        for meal_data in day_widgets['dieta'].values():
            for item_widgets in meal_data['items'].values():
                item_widgets['combo'].config(style="TCombobox")


        # --- Validation ---
        is_valid = True
        error_messages = []

        # Sono
        if not day_widgets['sono'].get('1.0', tk.END).strip():
            is_valid = False
            error_messages.append("- Qualidade do Sono")
            day_widgets['sono'].config(highlightbackground="red", highlightcolor="red", highlightthickness=1)

        # Metas (Água, Chá) - Assuming 0 is not a valid entry
        if day_widgets['agua_var'].get() == 0.0:
            is_valid = False
            error_messages.append("- Meta de Água")
        if day_widgets['cha_var'].get() == 0.0:
            is_valid = False
            error_messages.append("- Meta de Chá")

        # Dieta
        for meal, meal_data in day_widgets['dieta'].items():
            if meal_data['pulei_var'].get():
                if not meal_data['justificativa_pulei_widget'].get('1.0', tk.END).strip():
                    is_valid = False
                    error_messages.append(f"- {meal}: Justificativa (Pulei)")
            else:
                for item_name, item_widgets in meal_data['items'].items():
                    selection = item_widgets['combo'].get()
                    if not selection:
                        is_valid = False
                        error_messages.append(f"- {meal}: {item_name}")
                        item_widgets['combo'].config(style="Error.TCombobox")
                    elif selection == "Outra":
                        if not item_widgets['outra_widget'].get('1.0', tk.END).strip():
                            is_valid = False
                            error_messages.append(f"- {meal}: {item_name} (Descreva)")
                    elif selection == "Nenhum":
                        if not item_widgets['nenhum_widget'].get('1.0', tk.END).strip():
                            is_valid = False
                            error_messages.append(f"- {meal}: {item_name} (Justifique-se)")

        # Treino (if applicable)
        if day in ["Segunda-feira", "Quarta-feira", "Sexta-feira"]:
            if not day_widgets['treino_relato'].get('1.0', tk.END).strip():
                is_valid = False
                error_messages.append("- Relato do Treino")
                day_widgets['treino_relato'].config(highlightbackground="red", highlightcolor="red", highlightthickness=1)
            if not day_widgets['treino_carga'].get():
                 is_valid = False
                 error_messages.append("- Progressão de Carga") # No visual feedback for radio buttons, msg is enough

        # Tue/Thu Training
        if day in ["Terça-feira", "Quinta-feira"]:
            if not day_widgets['tue_thu_treino_var'].get():
                is_valid = False
                error_messages.append("- Treinou hoje?")
            elif day_widgets['tue_thu_treino_var'].get() == 'sim' and not day_widgets['tue_thu_treino_desc'].get('1.0', tk.END).strip():
                is_valid = False
                error_messages.append("- Descrição do Treino")
            elif day_widgets['tue_thu_treino_var'].get() == 'nao' and not day_widgets['tue_thu_treino_just'].get('1.0', tk.END).strip():
                is_valid = False
                error_messages.append("- Justificativa do Treino")

        # Flexões Justificativa (if applicable)
        if day_widgets['flexoes_var'].get() == 'nao' and not day_widgets['flexoes_justificativa'].get('1.0', tk.END).strip():
            is_valid = False
            error_messages.append("- Justificativa das Flexões")
            day_widgets['flexoes_justificativa'].config(highlightbackground="red", highlightcolor="red", highlightthickness=1)

        # Peso (if applicable)
        if day == "Sexta-feira":
            if not day_widgets['peso'].get().strip():
                is_valid = False
                error_messages.append("- Peso em Jejum")
                day_widgets['peso'].config(style="Error.TEntry")

        if not is_valid:
            # Pass the tracker_window as parent to keep it in front
            tracker_window = notebook.winfo_toplevel()
            messagebox.showerror("Campos Obrigatórios", "Faltam informações a serem preenchidas:\n\n" + "\n".join(error_messages), parent=tracker_window)
            return

        # --- Data Collection ---
        data = {}
        # Common widgets
        data['sono'] = day_widgets['sono'].get('1.0', tk.END).strip()
        data['agua'] = day_widgets['agua_var'].get()
        data['cha'] = day_widgets['cha_var'].get()
        data['flexoes'] = day_widgets['flexoes_var'].get()
        if data['flexoes'] == 'nao':
            data['flexoes_justificativa'] = day_widgets['flexoes_justificativa'].get('1.0', tk.END).strip()

        # Training widgets
        if day in ["Segunda-feira", "Quarta-feira", "Sexta-feira"]:
            data['treino_relato'] = day_widgets['treino_relato'].get('1.0', tk.END).strip()
            data['treino_cansaco'] = day_widgets['treino_cansaco'].get()
            data['treino_carga'] = day_widgets['treino_carga'].get()

        if day in ["Terça-feira", "Quinta-feira"]:
            data['tue_thu_treino'] = day_widgets['tue_thu_treino_var'].get()
            if data['tue_thu_treino'] == 'sim':
                data['tue_thu_treino_desc'] = day_widgets['tue_thu_treino_desc'].get('1.0', tk.END).strip()
            else:
                data['tue_thu_treino_just'] = day_widgets['tue_thu_treino_just'].get('1.0', tk.END).strip()

        # Friday widgets
        if day == "Sexta-feira":
            data['peso'] = day_widgets['peso'].get()

        # Diet widgets
        data['dieta'] = {}
        if 'dieta' in day_widgets:
            for meal, meal_data in day_widgets['dieta'].items():
                data['dieta'][meal] = {'pulei': meal_data['pulei_var'].get()}
                if meal_data['pulei_var'].get():
                    data['dieta'][meal]['justificativa_pulei'] = meal_data['justificativa_pulei_widget'].get('1.0', tk.END).strip()
                else:
                    data['dieta'][meal]['items'] = {}
                    for item_name, item_widgets in meal_data['items'].items():
                        selection = item_widgets['combo'].get()
                        data['dieta'][meal]['items'][item_name] = {'selection': selection}
                        if selection == "Outra":
                            data['dieta'][meal]['items'][item_name]['outra'] = item_widgets['outra_widget'].get('1.0', tk.END).strip()
                        elif selection == "Nenhum":
                            data['dieta'][meal]['items'][item_name]['nenhum'] = item_widgets['nenhum_widget'].get('1.0', tk.END).strip()
        # Save data
        file_path = '.foco_data.json'
        all_data = {}

        # On Windows, un-hide the file before writing
        if os.name == 'nt' and os.path.exists(file_path):
            try:
                os.system(f'attrib -h "{file_path}"')
            except Exception as e:
                print(f"Error removing hidden attribute: {e}")

        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                try:
                    all_data = json.load(f)
                except json.JSONDecodeError:
                    pass # file is empty or corrupted

        selected_tab_id = str(notebook.nametowidget(notebook.select()))
        save_date_str = self.tab_dates[selected_tab_id]

        all_data[save_date_str] = data

        with open(file_path, 'w') as f:
            json.dump(all_data, f, indent=4)

        # Re-hide the file on Windows after writing
        if os.name == 'nt':
            try:
                os.system(f'attrib +h "{file_path}"')
            except Exception as e:
                print(f"Não foi possível ocultar o arquivo após escrita: {e}")

        messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")


    def _create_diet_widgets(self, day):
        frame = self.day_tabs[day]
        widgets = self.widgets[day]

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)

        diet_title_font = ("Helvetica", 12, "bold")
        ttk.Label(frame, text="DIETA:", font=diet_title_font).pack(anchor='w', pady=(10, 5))

        diet_options = {
            # ... (diet options will be defined inside)
        }

        widgets['dieta'] = {}

        diet_options = {
            "Café da manhã": [
                ("Carboidrato:", ["2 fatias de pão de forma", "1 pão francês"]),
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
                ("Proteína:", ["50g de whey"]),
                ("Fruta:", ["100g de morango", "70g de uva"])
            ],
            "Jantar": [
                ("Carboidrato:", ["50g de arroz", "50g de macarrão", "120g de batata"]),
                ("Proteína:", ["200g de patinho", "200g de atum"]),
                ("Salada:", ["100g de vegetais e folhas"])
            ]
        }

        for meal, options in diet_options.items():
            meal_container = ttk.Frame(frame)
            meal_container.pack(fill='x', expand=True, padx=5, pady=5)

            meal_frame = ttk.LabelFrame(meal_container, text=meal, padding=10)

            pulei_var = tk.BooleanVar()
            pulei_check = ttk.Checkbutton(meal_frame, text="Pulei esta refeição.", variable=pulei_var)
            pulei_check.pack(anchor='w', pady=5, fill='x')

            justificativa_pulei_frame = ttk.Frame(meal_container)
            ttk.Label(justificativa_pulei_frame, text="Justifique-se:").pack(anchor='w')
            justificativa_pulei_text = tk.Text(justificativa_pulei_frame, height=2, width=40)
            justificativa_pulei_text.pack(pady=5, anchor='w', fill='x')

            widgets['dieta'][meal] = {
                'frame': meal_frame,
                'pulei_var': pulei_var,
                'justificativa_pulei_widget': justificativa_pulei_text,
                'items': {}
            }

            def toggle_meal_frame(p_var, m_frame, j_frame):
                if p_var.get():
                    m_frame.pack_forget()
                    j_frame.pack(fill='x', expand=True)
                else:
                    m_frame.pack(fill='x', expand=True)
                    j_frame.pack_forget()

            pulei_check.config(command=lambda p=pulei_var, m=meal_frame, j=justificativa_pulei_frame: toggle_meal_frame(p, m, j))

            meal_frame.pack(fill='x', expand=True) # Initially visible

            for i, (label_text, option_list) in enumerate(options):
                option_list_ext = option_list + ["Nenhum", "Outra"]

                item_frame = ttk.Frame(meal_frame)
                item_frame.pack(fill='x', expand=True, pady=5)

                ttk.Label(item_frame, text=label_text).pack(anchor='w')
                combo = ttk.Combobox(item_frame, values=option_list_ext, state="readonly", width=30)
                combo.pack(fill='x', expand=True)

                # Conditional text boxes
                outra_frame = ttk.Frame(item_frame)
                ttk.Label(outra_frame, text="Descreva:").pack(anchor='w')
                outra_text = tk.Text(outra_frame, height=2, width=40)
                outra_text.pack(pady=5, anchor='w', fill='x')

                nenhum_frame = ttk.Frame(item_frame)
                ttk.Label(nenhum_frame, text="Justifique-se:").pack(anchor='w')
                nenhum_text = tk.Text(nenhum_frame, height=2, width=40)
                nenhum_text.pack(pady=5, anchor='w', fill='x')

                widgets['dieta'][meal]['items'][label_text.replace(":", "")] = {
                    'combo': combo,
                    'outra_widget': outra_text,
                    'nenhum_widget': nenhum_text
                }

                def on_combo_select(event, c=combo, o_frame=outra_frame, n_frame=nenhum_frame):
                    selection = c.get()
                    if selection == "Outra":
                        o_frame.pack(fill='x', expand=True, pady=5)
                    else:
                        o_frame.pack_forget()

                    if selection == "Nenhum":
                        n_frame.pack(fill='x', expand=True, pady=5)
                    else:
                        n_frame.pack_forget()

                combo.bind("<<ComboboxSelected>>", on_combo_select)


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

        day_map = {0: "SEG", 1: "TER", 2: "QUA", 3: "QUI", 4: "SEX", 5: "SÁB", 6: "DOM"}

        for date_str in sorted_dates:
            try:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                day_of_week = day_map[date_obj.weekday()]
                formatted_date = date_obj.strftime('%d/%m/%Y')
                button_text = f"{day_of_week} - {formatted_date}"
            except ValueError:
                button_text = date_str # Fallback for old format

            btn = ttk.Button(scrollable_frame, text=button_text, style="Bold.TButton",
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

        def _details_scroll(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"

        details_window.bind_all("<MouseWheel>", _details_scroll)

        def _on_details_close():
            details_window.unbind_all("<MouseWheel>")
            details_window.destroy()

        details_window.protocol("WM_DELETE_WINDOW", _on_details_close)

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
            "peso": "Peso em Jejum (kg)",
            "tue_thu_treino": "Treinou Hoje?",
            "tue_thu_treino_desc": "Descrição do Treino",
            "tue_thu_treino_just": "Justificativa (Treino)"
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
            for meal, meal_data in data['dieta'].items():
                if meal_data.get('pulei'):
                    ttk.Label(scrollable_frame, text=f"{meal}: Pulei").pack(anchor='w', pady=(5, 0))
                    justificativa = meal_data.get('justificativa_pulei', 'N/A')
                    ttk.Label(scrollable_frame, text=f"  - Justificativa: {justificativa}").pack(anchor='w')
                elif 'items' in meal_data:
                    ttk.Label(scrollable_frame, text=f"{meal}:").pack(anchor='w', pady=(5, 0))
                    for item_name, item_data in meal_data['items'].items():
                        selection = item_data.get('selection')
                        if selection:
                            if selection == "Outra":
                                desc = item_data.get('outra', 'N/A')
                                ttk.Label(scrollable_frame, text=f"  - {item_name}: Outra ({desc})").pack(anchor='w')
                            elif selection == "Nenhum":
                                just = item_data.get('nenhum', 'N/A')
                                ttk.Label(scrollable_frame, text=f"  - {item_name}: Nenhum ({just})").pack(anchor='w')
                            else:
                                ttk.Label(scrollable_frame, text=f"  - {item_name}: {selection}").pack(anchor='w')

        exit_button = ttk.Button(details_window, text="SAIR", style="Bold.TButton", command=_on_details_close)
        exit_button.pack(pady=10)


if __name__ == "__main__":
    app = App()
    app.mainloop()
