from tkcalendar import DateEntry
from ttkthemes import ThemedStyle

from operations import *

# Create Object

root = Tk()
root.title('Operação LogLife')
root.resizable(False, False)
root.iconbitmap('my_icon.ico')
thread_0 = Start(root)
thread_1 = Start(root)
thread_2 = Start(root)
tabs = ttk.Notebook(root)

# Setting tabs

tab1 = ttk.Frame(tabs)
tab2 = ttk.Frame(tabs)
tab3 = ttk.Frame(tabs)
tab4 = ttk.Frame(tabs)

tabs.add(tab1, text='Serviços')
tabs.add(tab2, text='Minutas')
tabs.add(tab3, text='Relação de cargas')
tabs.add(tab4, text='Arquivos')

tab1_frame = Frame(tab1)
f2 = Frame(tab2)
f2_1 = Frame(tab2)
f2_2 = Frame(tab2)
tab3_frame = Frame(tab3)
tab4_frame = Frame(tab4)
tab1_frame.pack()
f2.pack()
f2_1.pack()
f2_2.pack()
tab3_frame.pack(anchor=CENTER, padx=50)
tab4_frame.pack()

tabs.pack(expand=1, fill="both")

# Set geometry
root.geometry("")

today = dt.datetime.today()

dia = today.day
mes = today.month
ano = today.year

style = ThemedStyle(root)
style.theme_use('breeze')

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

cal = DateEntry(tab1_frame, **cal_config)
cal1 = DateEntry(tab1_frame, **cal_config)

cal.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
cal1.grid(column=1, row=1, padx=30, pady=10, sticky="E, W")

# Read file name/path

filename = StringVar()

try:
    with open('sheet_name.txt') as m:
        text = m.read()
    lines = text.split('\n')
    filename.set(lines[0])
except FileNotFoundError:
    filename.set(r'Serviços.xlsx')

filename_Label = ttk.Label(tab4_frame, width=20, text=filename.get(), wraplength=140)
filename_Label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)

filename2 = StringVar()

try:
    with open('filename2.txt') as m:
        text = m.read()
    lines = text.split('\n')
    filename2.set(lines[0])
except FileNotFoundError:
    filename2.set(r'Serviços.xlsx')

filename_Label2 = ttk.Label(tab4_frame, width=20, text=filename2.get(), wraplength=140)
filename_Label2.grid(column=1, row=1, sticky="E, W", padx=20, pady=10)

folderpath = StringVar()

try:
    with open('folderpath.txt') as m:
        text = m.read()
    lines = text.split('\n')
    folderpath.set(lines[0])
except FileNotFoundError:
    folderpath.set('')

folder_Label = ttk.Label(tab4_frame, width=20, text=folderpath.get(), wraplength=140)
folder_Label.grid(column=1, row=2, sticky="E, W", padx=20, pady=10)

downloadpath = StringVar()

try:
    with open('downloadpath.txt') as m:
        text = m.read()
    lines = text.split('\n')
    downloadpath.set(lines[0])
except FileNotFoundError:
    downloadpath.set('')

# Entry Text

dispatch_prot = StringVar()
ttk.Entry(f2, textvariable=dispatch_prot, width=15).grid(column=1, row=0, padx=10, pady=10)

# Labels

ttk.Label(tab1_frame, text="      Data Inicial:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab1_frame, text="       Data Final:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

ttk.Label(f2, text="Protocolo:").grid(column=0, row=0, pady=15, sticky="W,E")
ttk.Label(f2_1, text="Tipo de Material:").grid(column=0, row=0, sticky="W,E")
ttk.Label(f2_1, text="Serviço:").grid(column=1, row=0, sticky="W,E")
ttk.Label(f2_1, text="Vols.:").grid(column=3, row=0, sticky="W,E", padx=21)
ttk.Label(f2_1, text="Peso:").grid(column=3, row=2, sticky="W,E", padx=21)

ttk.Label(tab4_frame, text="Planilha de Serviços:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab4_frame, text="Relação de Cargas:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')
ttk.Label(tab4_frame, text="Pasta das minutas:").grid(column=0, row=2, padx=10, pady=10, sticky='W, E')

browse1 = Browse(filename_Label)
browse2 = Browse(filename_Label2)
browse3 = Browse(folder_Label)
browse4 = Browse(None)

# Add Button and Label
ttk.Button(
    tab1_frame,
    text="Atualizar dados",
    command=lambda: thread_0.start_thread(services_api, progressbar, arguments=[root,
                                                                                cal,
                                                                                cal1,
                                                                                tab1_frame,
                                                                                date_filter.get(),
                                                                                reg_filter.get(),
                                                                                reg.get(),
                                                                                filename.get()])
).grid(column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(tab1_frame, text="Limpar dados",
           command=lambda: thread_0.start_thread(clear_data, progressbar, arguments=[filename.get(),
                                                                                     "COLETAS;A2:P500",
                                                                                     "EMBARQUES;A2:N500",
                                                                                     "ENTREGAS;A2:L1000",
                                                                                     "CARGAS;A2:S2000"])).grid(
    column=2, row=1, padx=10, pady=10, sticky="N, S, E, W"
)

ttk.Button(
    f2,
    text="Emitir minuta",
    command=lambda: thread_1.start_thread(minutas_api, progressbar2, arguments=[dispatch_prot.get(),
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

ttk.Button(
    f2_2,
    text="Emitir minutas aéreas",
    command=lambda: thread_1.start_thread(minutas_all_api, progressbar2, arguments=[cal,
                                                                                    cal1,
                                                                                    folderpath.get(),
                                                                                    downloadpath.get(),
                                                                                    dispatch_prot,
                                                                                    root
                                                                                    ])).grid(
    column=0, row=0, padx=10, pady=10, ipadx=15, sticky="N, S, E, W"
)

ttk.Button(
    f2_2,
    text="..",
    width=2,
    command=lambda: browse4.browse_folder(downloadpath, 'downloadpath.txt')
).grid(column=1, row=0)

ttk.Button(
    tab3_frame,
    text="Atualizar Relação de Cargas",
    command=lambda: thread_2.start_thread(cargas_api, progressbar3, arguments=[filename2.get()])
).pack(fill="both", anchor=CENTER, pady=20)

ttk.Button(
    tab3_frame,
    text="Atualizar Planilha Fleury",
    command=lambda: thread_2.start_thread(fleury_sheet, progressbar3, arguments=[cal,
                                                                                 filename2.get()])
).pack(fill="both", anchor=CENTER, pady=20)

ttk.Button(tab4_frame, text="Procurar",
           command=lambda: browse1.browse_files(filename, 'sheet_name.txt', master=tab4_frame,
                                                label_config={'width': 20, 'wraplength': 140},
                                                grid_config={'column': 1, 'row': 0, 'sticky': "E, W", 'padx': 20,
                                                             'pady': 10})
           ).grid(column=2, row=0, padx=20, pady=10, sticky="E, W")

ttk.Button(tab4_frame, text="Procurar",
           command=lambda: browse2.browse_files(filename2, 'filename2.txt', master=tab4_frame,
                                                label_config={'width': 20, 'wraplength': 140},
                                                grid_config={'column': 1, 'row': 1, 'sticky': "E, W", 'padx': 20,
                                                             'pady': 10})
           ).grid(column=2, row=1, padx=20, pady=10, sticky="E, W")

ttk.Button(tab4_frame, text="Procurar",
           command=lambda: browse3.browse_folder(folderpath, 'folderpath.txt', master=tab4_frame,
                                                 label_config={'width': 20, 'wraplength': 140},
                                                 grid_config={'column': 1, 'row': 2, 'sticky': "E, W", 'padx': 20,
                                                              'pady': 10})
           ).grid(column=2, row=2, padx=20, pady=10, sticky="E, W")

# Radio Buttons

material_type = IntVar(f2_1, 0)
flight_service = IntVar(f2_1, 0)

ttk.Radiobutton(f2_1, text="Biológico", value=0, variable=material_type).grid(column=0, row=1, sticky="W,E")
ttk.Radiobutton(f2_1, text="Carga Geral", value=1, variable=material_type).grid(column=0, row=2, sticky="W,E")
ttk.Radiobutton(f2_1, text="Med. e Vacinas", value=2, variable=material_type).grid(column=0, row=3, sticky="W,E")

ttk.Radiobutton(f2_1, text="Próximo Voo", value=0, variable=flight_service).grid(column=1, row=1, sticky="W,E")
ttk.Radiobutton(f2_1, text="Convencional", value=1, variable=flight_service).grid(column=1, row=2, sticky="W,E")

# Checkbox

date_filter = IntVar(tab1_frame, 1)
date_filter_chk = ttk.Checkbutton(tab1_frame, text='Filtrar por Data', variable=date_filter, onvalue=1, offvalue=0)
date_filter_chk.grid(column=0, row=2, padx=10, pady=10, sticky='W, E')

reg_filter = IntVar(tab1_frame, 0)
reg_filter_chk = ttk.Checkbutton(tab1_frame, text='Filtrar por Regional', variable=reg_filter, onvalue=1, offvalue=0)
reg_filter_chk.grid(column=1, row=2, padx=10, pady=10, sticky='W, E')

# Spinbox

reg = IntVar(tab1_frame, 1)
vol_spinbox = ttk.Spinbox(tab1_frame, textvariable=reg, from_=1, to=6, wrap=True, width=5)
vol_spinbox.grid(column=2, row=2, padx=10, pady=10)

vols = IntVar(f2_1, 1)
vol_spinbox = ttk.Spinbox(f2_1, textvariable=vols, from_=1, to=50, wrap=True, width=6)
vol_spinbox.grid(column=3, row=1, padx=21)

kg_record = IntVar(f2_1, 1)
kg_spinbox = ttk.Spinbox(f2_1, textvariable=kg_record, from_=1, to=1000, wrap=True, width=6)
kg_spinbox.grid(column=3, row=3, padx=21)

# Progress Bar

progressbar = ttk.Progressbar(tab1, mode='indeterminate')
progressbar.pack(side=BOTTOM, fill='x')

progressbar2 = ttk.Progressbar(tab2, mode='indeterminate')
progressbar2.pack(side=BOTTOM, fill='x')

progressbar3 = ttk.Progressbar(tab3, mode='indeterminate')
progressbar3.pack(side=BOTTOM, fill='x')

# Execute Tkinter
root.mainloop()
