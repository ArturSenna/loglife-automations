from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
import tkinter
import tkinter.messagebox
import customtkinter as ctk
import datetime as dt
import tempfile
import os
from functions import *
from billingHIAE import hiae_billing


# noinspection PyTypeChecker
class FinanceApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # congigure window
        self.title('Faturamento HIAE')
        self.resizable(True, True)
        self.geometry(f"{580}x{260}")
        self.iconbitmap("my_icon.ico")

        # start threads
        self.thread_0 = Start(self)
        self.thread_1 = Start(self)

        # create tabs
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(expand=1, fill="both")
        self.tab1 = self.tabs.add("Faturamento")
        self.tab2 = self.tabs.add("Arquivos")

        # set styling
        self.style = ThemedStyle()
        self.style.theme_use("arc")
        self.style.configure('my.DateEntry',
                             fieldbackground='dark blue',
                             background='black',
                             foreground='black',
                             arrowcolor='dark blue')

        self.create_frames()

    def create_frames(self):

        tab1_frame = ctk.CTkFrame(self.tab1)
        tab2_frame = ctk.CTkFrame(self.tab2)

        tab1_frame.pack()
        tab2_frame.pack()

        initial_date = DateEntry(tab1_frame, style='my.DateEntry', **cal_config)
        initial_date.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")

        final_date = DateEntry(tab1_frame, style='my.DateEntry', **cal_config)
        final_date.grid(column=1, row=1, padx=30, pady=10, sticky="E, W")

        os.makedirs(f'{temp_folder}/filepaths', exist_ok=True)

        folderpath = StringVar()

        try:
            with open(f'{temp_folder}/filepaths/folderpath.txt') as m:
                text = m.read()
            lines = text.split('\n')
            folderpath.set(lines[0])
        except FileNotFoundError:
            folderpath.set('Clique em "Procurar" e selecione a pasta onde os relatórios ficarão')

        folder_label = ctk.CTkLabel(tab2_frame, width=20, text=folderpath.get(), wraplength=140)
        folder_label.grid(column=1, row=1, sticky="E, W", padx=20, pady=10)

        filename = StringVar()

        try:
            with open(f'{temp_folder}/filepaths/sheet_name.txt') as m:
                text = m.read()
            lines = text.split('\n')
            filename.set(lines[0])
        except FileNotFoundError:
            filename.set(r'Clique em "Procurar" e selecione o arquivo')

        filename_label = ctk.CTkLabel(tab2_frame, width=20, text=filename.get(), wraplength=140)
        filename_label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)

        ctk.CTkLabel(tab1_frame, text="      Data Inicial:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
        ctk.CTkLabel(tab1_frame, text="       Data Final:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

        ctk.CTkLabel(tab2_frame, text="Arquivo:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
        ctk.CTkLabel(tab2_frame, text="Pasta do relatório:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

        browse1 = Browse(filename_label)
        browse2 = Browse(folder_label)

        ctk.CTkButton(
            tab1_frame,
            text="Atualizar dados",
            command=lambda: self.thread_0.start_thread(hiae_billing, progressbar, arguments=[initial_date,
                                                                                             final_date])
        ).grid(column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

        ctk.CTkButton(tab1_frame, text="Limpar dados"
                      #               self.thread_0.start_thread(lamdba: pass, progressbar, arguments=[]
                      ).grid(column=2, row=1, padx=10, pady=10, sticky="N, S, E, W")

        ctk.CTkButton(tab2_frame, text="Procurar",
                      command=lambda: browse1.browse_files(filename,
                                                           f'{temp_folder}/filepaths/sheet_name.txt',
                                                           master=tab2_frame,
                                                           label_config={'width': 20, 'wraplength': 140},
                                                           grid_config={'column': 1, 'row': 0, 'sticky': "E, W",
                                                                        'padx': 20,
                                                                        'pady': 10})
                      ).grid(column=2, row=0, padx=20, pady=10, sticky="E, W")

        ctk.CTkButton(tab2_frame, text="Procurar",
                      command=lambda: browse2.browse_folder(folderpath,
                                                            f'{temp_folder}/filepaths/folderpath.txt',
                                                            master=tab2_frame,
                                                            label_config={'width': 20, 'wraplength': 140},
                                                            grid_config={'column': 1, 'row': 1, 'sticky': "E, W",
                                                                         'padx': 20,
                                                                         'pady': 10})
                      ).grid(column=2, row=1, padx=20, pady=10, sticky="E, W")

        # Progress Bar

        progressbar = ctk.CTkProgressBar(self.tab1, mode='indeterminate')
        progressbar.pack(side=BOTTOM, fill='x')

        progressbar2 = ctk.CTkProgressBar(self.tab2, mode='indeterminate')
        progressbar2.pack(side=BOTTOM, fill='x')

    @staticmethod
    def change_scaling_event(new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

today = dt.datetime.today()
dia = today.day
mes = today.month
ano = today.year
cal_config = {'selectmode': 'day',
              'day': dia,
              'month': mes,
              'year': ano,
              'locale': 'pt_BR',
              'firstweekday': 'sunday',
              'showweeknumbers': False,
              'bordercolor': "white",
              'background': "white",
              'disabledbackground': "white",
              'headersbackground': "white",
              'normalbackground': "white",
              'normalforeground': 'black',
              'headersforeground': 'black',
              'selectbackground': '#00a5e7',
              'selectforeground': 'white',
              'weekendbackground': 'white',
              'weekendforeground': 'black',
              'othermonthforeground': 'black',
              'othermonthbackground': '#E8E8E8',
              'othermonthweforeground': 'black',
              'othermonthwebackground': '#E8E8E8',
              'foreground': "black"}

temp_folder = f'{tempfile.gettempdir()}/Financeiro LogLife'

if __name__ == "__main__":
    app = FinanceApp()
    app.mainloop()