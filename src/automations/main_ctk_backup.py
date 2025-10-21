from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
import customtkinter as ctk
from operations import *
import tempfile
import os
import datetime as dt
from pathlib import Path

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

download_path = str(Path.home() / "Downloads")

print(download_path)

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


class LogLifeApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # configure window
        self.title('Operação LogLife')
        self.resizable(True, True)
        self.geometry(f"{580}x{340}")
        
        # Use absolute path for icon
        icon_path = Path(__file__).parent.parent / 'assets' / 'my_icon.ico'
        if icon_path.exists():
            self.iconbitmap(str(icon_path))

        # start threads
        self.thread_0 = Start(self)
        self.thread_1 = Start(self)
        self.thread_2 = Start(self)

        # create tabs
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(expand=1, fill="both")
        self.tab1 = self.tabs.add("Serviços")
        self.tab2 = self.tabs.add("Minutas")
        self.tab3 = self.tabs.add("Relação de cargas")
        self.tab4 = self.tabs.add("Arquivos")

        # set styling
        self.style = ThemedStyle()
        self.style.theme_use('arc')
        self.style.configure('my.DateEntry',
                             fieldbackground='dark blue',
                             background='black',
                             foreground='black',
                             arrowcolor='dark blue')

        self.create_frames()

    def create_frames(self):

        tab1_frame = ctk.CTkFrame(self.tab1)
        tab2_frame_a = ctk.CTkFrame(self.tab2)
        tab2_frame_b = ctk.CTkFrame(self.tab2)
        tab2_frame_c = ctk.CTkFrame(self.tab2)
        tab3_frame = ctk.CTkFrame(self.tab3)
        tab4_frame = ctk.CTkFrame(self.tab4)

        # pack frames
        tab1_frame.pack()
        tab2_frame_a.pack()
        tab2_frame_b.pack()
        tab2_frame_c.pack()
        tab3_frame.pack(anchor=CENTER, padx=50)
        tab4_frame.pack()

        cal = DateEntry(tab1_frame, style='my.DateEntry', **cal_config)
        cal1 = DateEntry(tab1_frame, style='my.DateEntry', **cal_config)

        cal.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
        cal1.grid(column=1, row=1, padx=30, pady=10, sticky="E, W")

        #############################

        # scaling_optionemenu = ctk.CTkOptionMenu(tab3_frame,
        #                                         values=["80%", "90%", "100%", "110%", "120%"],
        #                                         command=self.change_scaling_event)
        # scaling_optionemenu.pack()

        os.makedirs(f'{temp_folder}/filepaths', exist_ok=True)

        filename = StringVar()

        try:
            with open(f'{temp_folder}/filepaths/sheet_name.txt') as m:
                text = m.read()
            lines = text.split('\n')
            filename.set(lines[0])
        except FileNotFoundError:
            filename.set(r'Clique em "Procurar" e selecione a Planilha de Serviços')

        filename_label = ctk.CTkLabel(tab4_frame, width=20, text=filename.get(), wraplength=140)
        filename_label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)

        filename2 = StringVar()

        try:
            with open(f'{temp_folder}/filepaths/filename2.txt') as m:
                text = m.read()
            lines = text.split('\n')
            filename2.set(lines[0])
        except FileNotFoundError:
            filename2.set(r'Clique em "Procurar" e selecione a Relação de Cargas')

        filename_label2 = ctk.CTkLabel(tab4_frame, width=20, text=filename2.get(), wraplength=140)
        filename_label2.grid(column=1, row=1, sticky="E, W", padx=20, pady=10)

        folderpath = StringVar()

        fleury_sheet_name = StringVar()

        try:
            with open(f'{temp_folder}/filepaths/fleury_sheet_name.txt') as m:
                text = m.read()
            lines = text.split('\n')
            fleury_sheet_name.set(lines[0])
        except FileNotFoundError:
            fleury_sheet_name.set(r'Clique em "Procurar" e selecione a Planilha Fleury')

        fleury_sheet_label = ctk.CTkLabel(tab4_frame, width=20, text=fleury_sheet_name.get(), wraplength=140)
        fleury_sheet_label.grid(column=1, row=3, sticky="E, W", padx=20, pady=10)

        try:
            with open(f'{temp_folder}/filepaths/folderpath.txt') as m:
                text = m.read()
            lines = text.split('\n')
            folderpath.set(lines[0])
        except FileNotFoundError:
            folderpath.set('Clique em "Procurar" e selecione a pasta onde as minutas ficarão')

        folder_label = ctk.CTkLabel(tab4_frame, width=20, text=folderpath.get(), wraplength=140)
        folder_label.grid(column=1, row=2, sticky="E, W", padx=20, pady=10)

        downloadpath = StringVar()

        try:
            with open(f'{temp_folder}/filepaths/downloadpath.txt') as m:
                text = m.read()
            lines = text.split('\n')
            downloadpath.set(lines[0])
        except FileNotFoundError:
            downloadpath.set(download_path)

        # Entry Text

        dispatch_prot = StringVar()
        ctk.CTkEntry(tab2_frame_a, textvariable=dispatch_prot, width=80).grid(column=1, row=0, padx=10, pady=10)

        # Labels

        ctk.CTkLabel(tab1_frame, text="      Data Inicial:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
        ctk.CTkLabel(tab1_frame, text="       Data Final:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

        ctk.CTkLabel(tab2_frame_a, text="Protocolo:").grid(column=0, row=0, pady=15, sticky="W,E")
        ctk.CTkLabel(tab2_frame_b, text="Tipo de Material:").grid(column=0, row=0, sticky="W,E")
        ctk.CTkLabel(tab2_frame_b, text="Serviço:").grid(column=1, row=0, sticky="W,E")
        ctk.CTkLabel(tab2_frame_b, text="Vols.:").grid(column=3, row=0, sticky="W,E", padx=21)
        ctk.CTkLabel(tab2_frame_b, text="Peso:").grid(column=3, row=2, sticky="W,E", padx=21)

        ctk.CTkLabel(tab4_frame, text="Planilha de Serviços:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
        ctk.CTkLabel(tab4_frame, text="Relação de Cargas:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')
        ctk.CTkLabel(tab4_frame, text="Pasta das minutas:").grid(column=0, row=2, padx=10, pady=10, sticky='W, E')
        ctk.CTkLabel(tab4_frame, text="Planilha Fleury:").grid(column=0, row=3, padx=10, pady=10, sticky='W, E')

        browse1 = Browse(filename_label)
        browse2 = Browse(filename_label2)
        browse3 = Browse(folder_label)
        browse4 = Browse(None)
        browse5 = Browse(fleury_sheet_label)

        # adding buttons
        ctk.CTkButton(
            tab1_frame,
            text="Atualizar dados",
            command=lambda: self.thread_0.start_thread(services_api, progressbar, arguments=[self,
                                                                                             cal,
                                                                                             cal1,
                                                                                             tab1_frame,
                                                                                             date_filter.get(),
                                                                                             reg_filter.get(),
                                                                                             reg.get(),
                                                                                             filename.get()])
        ).grid(column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

        ctk.CTkButton(tab1_frame, text="Limpar dados",
                      command=lambda: self.thread_0.start_thread(clear_data, progressbar, arguments=[
                          filename.get(),
                          "COLETAS;A2:P500",
                          "EMBARQUES;A2:N500",
                          "ENTREGAS;A2:L1000",
                          "CARGAS;A2:S2000"
                      ])).grid(column=2, row=1, padx=10, pady=10, sticky="N, S, E, W")

        ctk.CTkButton(
            tab2_frame_a,
            text="Emitir minuta",
            command=lambda: self.thread_1.start_thread(minutas_api, progressbar2, arguments=[dispatch_prot.get(),
                                                                                             False,
                                                                                             "",
                                                                                             flight_service.get(),
                                                                                             material_type.get(),
                                                                                             vols.get(),
                                                                                             kg_record.get(),
                                                                                             folderpath.get(),
                                                                                             downloadpath.get(),
                                                                                             dispatch_prot
                                                                                             ])).grid(
            column=2, row=0, padx=10, pady=10, ipadx=11, sticky="N, S, E, W"
        )

        ctk.CTkButton(
            tab2_frame_c,
            text="Emitir minutas aéreas",
            command=lambda: self.thread_1.start_thread(minutas_all_api, progressbar2, arguments=[cal,
                                                                                                 cal1,
                                                                                                 folderpath.get(),
                                                                                                 downloadpath.get(),
                                                                                                 dispatch_prot,
                                                                                                 self
                                                                                                 ])).grid(
            column=0, row=0, padx=10, pady=10, ipadx=15, sticky="N, S, E, W"
        )

        ctk.CTkButton(
            tab2_frame_c,
            text="..",
            width=2,
            command=lambda: browse4.browse_folder(
                downloadpath,
                f'{temp_folder}/filepaths/downloadpath.txt')
        ).grid(column=1, row=0)

        ctk.CTkButton(
            tab3_frame,
            text="Atualizar Relação de Cargas",
            command=lambda: self.thread_2.start_thread(cargas_api, progressbar3, arguments=[filename2.get()])
        ).pack(fill="both", anchor=CENTER, pady=20)

        ctk.CTkButton(
            tab3_frame,
            text="Atualizar Planilha Fleury",
            command=lambda: self.thread_2.start_thread(fleury_sheet, progressbar3, arguments=[cal,
                                                                                              fleury_sheet_name.get()])
        ).pack(fill="both", anchor=CENTER, pady=20)

        ctk.CTkButton(tab4_frame, text="Procurar",
                      command=lambda: browse1.browse_files(filename,
                                                           f'{temp_folder}/filepaths/sheet_name.txt',
                                                           master=tab4_frame,
                                                           label_config={'width': 20, 'wraplength': 140},
                                                           grid_config={'column': 1, 'row': 0, 'sticky': "E, W",
                                                                        'padx': 20,
                                                                        'pady': 10})
                      ).grid(column=2, row=0, padx=20, pady=10, sticky="E, W")

        ctk.CTkButton(tab4_frame, text="Procurar",
                      command=lambda: browse2.browse_files(filename2,
                                                           f'{temp_folder}/filepaths/filename2.txt',
                                                           master=tab4_frame,
                                                           label_config={'width': 20, 'wraplength': 140},
                                                           grid_config={'column': 1, 'row': 1, 'sticky': "E, W",
                                                                        'padx': 20,
                                                                        'pady': 10})
                      ).grid(column=2, row=1, padx=20, pady=10, sticky="E, W")

        ctk.CTkButton(tab4_frame, text="Procurar",
                      command=lambda: browse3.browse_folder(folderpath,
                                                            f'{temp_folder}/filepaths/folderpath.txt',
                                                            master=tab4_frame,
                                                            label_config={'width': 20, 'wraplength': 140},
                                                            grid_config={'column': 1, 'row': 2, 'sticky': "E, W",
                                                                         'padx': 20,
                                                                         'pady': 10})
                      ).grid(column=2, row=2, padx=20, pady=10, sticky="E, W")

        ctk.CTkButton(tab4_frame, text="Procurar",
                      command=lambda: browse5.browse_files(fleury_sheet_name,
                                                           f'{temp_folder}/filepaths/fleury_sheet_name.txt',
                                                            master=tab4_frame,
                                                            label_config={'width': 20, 'wraplength': 140},
                                                            grid_config={'column': 1, 'row': 3, 'sticky': "E, W",
                                                                         'padx': 20,
                                                                         'pady': 10})
                      ).grid(column=2, row=3, padx=20, pady=10, sticky="E, W")

        # Radio Buttons

        material_type = IntVar(tab2_frame_b, 0)
        flight_service = IntVar(tab2_frame_b, 0)

        ctk.CTkRadioButton(
            tab2_frame_b, text="Biológico", value=0, variable=material_type).grid(column=0, row=1, sticky="W,E")
        ctk.CTkRadioButton(
            tab2_frame_b, text="Carga Geral", value=1, variable=material_type).grid(column=0, row=2, sticky="W,E")
        ctk.CTkRadioButton(
            tab2_frame_b, text="Infectante", value=2, variable=material_type).grid(column=0, row=3, sticky="W,E")
        ctk.CTkRadioButton(
            tab2_frame_b, text="Med. e Vacinas", value=3, variable=material_type).grid(column=0, row=4, sticky="W,E")

        ctk.CTkRadioButton(
            tab2_frame_b, text="Próximo Voo", value=0, variable=flight_service).grid(column=1, row=1, sticky="W,E")
        ctk.CTkRadioButton(
            tab2_frame_b, text="Convencional", value=1, variable=flight_service).grid(column=1, row=2, sticky="W,E")

        # Checkbox

        date_filter = IntVar(tab1_frame, 1)
        date_filter_chk = ctk.CTkCheckBox(tab1_frame, text='Filtrar por Data', variable=date_filter, onvalue=1,
                                          offvalue=0)
        date_filter_chk.grid(column=0, row=2, padx=10, pady=10, sticky='W, E')

        reg_filter = IntVar(tab1_frame, 0)
        reg_filter_chk = ctk.CTkCheckBox(tab1_frame, text='Filtrar por Regional', variable=reg_filter, onvalue=1,
                                         offvalue=0)
        reg_filter_chk.grid(column=1, row=2, padx=10, pady=10, sticky='W, E')

        # Spinbox

        reg = StringVar(tab1_frame, "1")
        vol_spinbox = ctk.CTkOptionMenu(tab1_frame, variable=reg, values=["1", "2", "3", "4", "5", "6"], width=50)
        vol_spinbox.grid(column=2, row=2, padx=10, pady=10)

        vols = IntVar(tab2_frame_b, 1)
        vol_spinbox = ttk.Spinbox(tab2_frame_b, textvariable=vols, from_=1, to=50, wrap=True, width=6)
        vol_spinbox.grid(column=3, row=1, padx=21)

        kg_record = IntVar(tab2_frame_b, 1)
        kg_spinbox = ttk.Spinbox(tab2_frame_b, textvariable=kg_record, from_=1, to=1000, wrap=True, width=6)
        kg_spinbox.grid(column=3, row=3, padx=21)

        # Progress Bar

        progressbar = ctk.CTkProgressBar(self.tab1, mode='indeterminate')
        progressbar.pack(side=BOTTOM, fill='x')

        progressbar2 = ctk.CTkProgressBar(self.tab2, mode='indeterminate')
        progressbar2.pack(side=BOTTOM, fill='x')

        progressbar3 = ctk.CTkProgressBar(self.tab3, mode='indeterminate')
        progressbar3.pack(side=BOTTOM, fill='x')

    @staticmethod
    def change_scaling_event(new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)


if __name__ == "__main__":
    app = LogLifeApp()
    app.mainloop()
