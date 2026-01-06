import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import json
import os
from datetime import datetime

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

    def create_main_menu(self):
        for widget in self.winfo_children():
            widget.destroy()

        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True)

        enter_button = ttk.Button(main_frame, text="ENTRAR NO SISTEMA", style="Bold.TButton", command=self.open_daily_tracker)
        enter_button.pack(pady=20)

        history_button = ttk.Button(main_frame, text="HISTÓRICO", style="Bold.TButton", command=self.open_history)
        history_button.pack(pady=20)

        exit_button = ttk.Button(main_frame, text="SAIR", style="Bold.TButton", command=self.destroy)
        exit_button.pack(pady=20)

    def open_daily_tracker(self):
        tracker_window = tk.Toplevel(self)
        tracker_window.title("FOCO TOTAL 2026")
        self._center_window(tracker_window, 900, 700)

        notebook = ttk.Notebook(tracker_window)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)

        days = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira"]
        self.day_tabs = {}
        self.widgets = {}

        for i, day in enumerate(days):
            frame = ttk.Frame(notebook, padding="10")
            notebook.add(frame, text=day)
            self.day_tabs[day] = frame
            self.widgets[day] = {}
            self._create_common_widgets(day)
            if i in [0, 2, 4]: # Monday, Wednesday, Friday
                self._create_training_widgets(day)
            if i == 4: # Friday
                self._create_friday_widgets(day)

        # Select current day
        current_day_index = datetime.now().weekday()
        if 0 <= current_day_index < 5:
            notebook.select(current_day_index)

        # --- Action Buttons ---
        button_frame = ttk.Frame(tracker_window)
        button_frame.pack(pady=10)

        save_button = ttk.Button(button_frame, text="SALVAR", style="Bold.TButton", command=lambda: self._save_data(notebook, days))
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

        # --- Escovou os dentes ---
        ttk.Label(frame, text="Quantas vezes escovou os dentes?").pack(pady=(10,0), anchor='w')
        dentes_frame = ttk.Frame(frame)
        dentes_frame.pack(anchor='w', fill='x')
        dentes_var = tk.IntVar(value=0)
        widgets['dentes_var'] = dentes_var
        dentes_label = ttk.Label(dentes_frame, text="0 ", width=4)
        dentes_scale = ttk.Scale(dentes_frame, from_=0, to=4, orient='horizontal', variable=dentes_var,
                                 command=lambda v, lbl=dentes_label, var=dentes_var: self._update_int_scale(v, lbl, var))
        dentes_scale.pack(side='left', pady=5, fill='x', expand=True)
        dentes_label.pack(side='left', padx=5)


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
        ttk.Label(frame, text="Relate como foi o seu treino:").pack(anchor='w')
        widgets['treino_relato'] = tk.Text(frame, height=4, width=50)
        widgets['treino_relato'].pack(pady=5, anchor='w')

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
            ttk.Radiobutton(radio_frame, text=str(i), variable=cansaco_var, value=i, command=update_cansaco_label).pack(side='left')

        radio_frame.pack(side='left')
        cansaco_label.pack(side='left', padx=10)

        # --- Progressão de carga ---
        ttk.Label(frame, text="Houve progressão de carga?").pack(anchor='w')
        carga_var = tk.StringVar()
        widgets['treino_carga'] = carga_var
        ttk.Radiobutton(frame, text="SIM", variable=carga_var, value="sim").pack(anchor='w')
        ttk.Radiobutton(frame, text="NÃO", variable=carga_var, value="nao").pack(anchor='w')


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

    def _save_data(self, notebook, days):
        selected_day_index = notebook.index(notebook.select())
        day = days[selected_day_index]

        data = {}
        day_widgets = self.widgets[day]

        # Common widgets
        data['sono'] = day_widgets['sono'].get('1.0', tk.END).strip()
        data['dentes'] = day_widgets['dentes_var'].get()
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

        # Save data
        file_path = '.foco_data.json'
        all_data = {}
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                try:
                    all_data = json.load(f)
                except json.JSONDecodeError:
                    pass # file is empty or corrupted

        today_str = datetime.now().strftime('%Y-%m-%d')
        all_data[today_str] = data

        with open(file_path, 'w') as f:
            json.dump(all_data, f, indent=4)

        # Hide the file on Windows
        if os.name == 'nt':
            try:
                os.system(f'attrib +h "{file_path}"')
            except Exception as e:
                print(f"Não foi possível ocultar o arquivo: {e}")

        messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")


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

        text_widget = tk.Text(details_window, wrap='word', font=("Helvetica", 10), spacing1=5, spacing2=5, spacing3=5)
        text_widget.pack(expand=True, fill='both', padx=10, pady=10)

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

        for key, value in data.items():
            label = labels.get(key, key.replace("_", " ").title()) # Get friendly label or format the key

            # Format some values for better readability
            if key == 'flexoes' or key == 'treino_carga':
                value = "Sim" if value == 'sim' else "Não"

            display_text += f"{label}:\n"
            # Indent the value for clarity
            display_text += f"  {value}\n\n"

        text_widget.insert(tk.END, display_text)
        text_widget.config(state='disabled') # Make it read-only

        exit_button = ttk.Button(details_window, text="SAIR", style="Bold.TButton", command=details_window.destroy)
        exit_button.pack(pady=10)


if __name__ == "__main__":
    app = App()
    app.mainloop()
