import requests
import threading
import numpy as np
import pandas as pd
import xlwings as xw
import datetime as dt
import os
from tkinter import filedialog as fd
from tkinter import *
from tkinter import ttk
from json import loads
from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
from pythoncom import CoInitialize


def budget():

    def get_address_city(add):
        add_lenght = len(add)
        add_total = []
        count = 0
        cities = ''
        for i in range(add_lenght):
            add_city = address.loc[address['id'] == add[i], 'cityIDAddress.name'].values.item()
            add_total.append(add_city)
        add_total = np.unique(add_total)
        for p in add_total:
            if count == 0:
                cities = p
            elif count == len(add_total) - 1:
                cities += ', ' + p
            else:
                cities += ', ' + p
            count += 1

        return cities

    def excel_date(date1: datetime) -> float:
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days)

    def excel_datetime(date1: datetime) -> float:
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days) + float(delta.seconds) / 86400

    def get_branch(id_t: str) -> str:
        try:
            branch = bases.loc[bases['id'] == id_t, 'nickname'].values.item()
        except ValueError:
            branch = ''
        return branch

    def get_transp(id_t):
        try:
            if id_t is not None:
                transp = bases.loc[bases['id'] == id_t, 'shippingIDBranch.company_name'].values.item()
            else:
                transp = ""
        except ValueError:
            transp = ""
        return transp

    di = cal.get_date()
    df = cal1.get_date()

    di = dt.datetime.strftime(di, '%d/%m/%Y')
    df = dt.datetime.strftime(df, '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    di_temp = di_dt - dt.timedelta(days=5)
    df_temp = df_dt + dt.timedelta(days=5)

    di = dt.datetime.strftime(di_temp, '%d/%m/%Y')
    df = dt.datetime.strftime(df_temp, '%d/%m/%Y')

    # Variable return

    user_fb = StringVar()
    feedback = ttk.Label(tab1, textvariable=user_fb)
    feedback.grid(column=0, row=2, padx=10, pady=10, columnspan=3, sticky='W, E')

    user_fb.set('Requisitando dados...')

    a = requests.get('https://transportebiologico.com.br/api/public/address', headers=header)
    b = requests.get('https://transportebiologico.com.br/api/public/branch', headers=header)
    c = requests.get('https://transportebiologico.com.br/api/public/collector', headers=header)
    s = requests.get('https://transportebiologico.com.br/api/public/service', headers=header)
    f = requests.post(
        f'https://transportebiologico.com.br/api/public/service/finalized/?startFilter={di}&endFilter={df}',
        headers=header)
    a_json = loads(a.text)
    b_json = loads(b.text)
    c_json = loads(c.text)
    s_json = loads(s.text)
    f_json = loads(f.text)
    address = pd.json_normalize(a_json)
    bases = pd.json_normalize(b_json)
    collector = pd.json_normalize(c_json)
    services_ongoing = pd.json_normalize(s_json)
    services_finalized = pd.json_normalize(f_json)

    services = pd.concat([services_ongoing, services_finalized], ignore_index=True)

    services['length'] = services['serviceIDBoard'].str.len()

    services.drop(services[services['length'] == 0].index, inplace=True)
    services.drop(services[services['step'] == 'toBoardValidate'].index, inplace=True)

    services = services.explode('serviceIDBoard')

    user_fb.set('Organizando informações')

    cargas = pd.DataFrame(columns=[])

    services['ETAPA'] = np.select(
        condlist=[
            services['step'] == 'availableService',
            services['step'] == 'toAllocateService',
            services['step'] == 'toDeliveryService',
            services['step'] == 'deliveringService',
            services['step'] == 'toLandingService',
            services['step'] == 'landingService',
            services['step'] == 'toBoardValidate',
            services['step'] == 'toCollectService',
            services['step'] == 'collectingService',
            services['step'] == 'toBoardService',
            services['step'] == 'boardingService',
            services['step'] == 'finishedService'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE', 'AGENDADO', 'COLETANDO',
            'EM ROTA DE EMBARQUE', 'EMBARCANDO SERVIÇO', 'FINALIZADO'],
        default=0)

    col = collector.rename(columns={
        'id': 'serviceIDRequested.crossdocking_collector_id', 'trading_name': 'Coletador intermediário'})

    services = pd.merge(services, col[['serviceIDRequested.crossdocking_collector_id', 'Coletador intermediário']],
                        on='serviceIDRequested.crossdocking_collector_id', how='left')

    services['CROSSDOCKING'] = np.where(services['Coletador intermediário'].str.len() > 0, 1, 0)

    services['Identificador'] = np.select(
        condlist=[(services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] != 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] == 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 2)],
        choicelist=[0, 1, 1],
        default=0
    )

    services['dataColeta'] = pd.to_datetime(services['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    services['dataColeta'] = services['dataColeta'].dt.tz_localize(None)

    services['BASE ORIGEM'] = services['serviceIDBoard'].map(
        lambda x: x.get('branch_id', np.nan) if isinstance(x, dict) else np.nan)

    services['VOLUMES'] = services['serviceIDBoard'].map(
        lambda x: x.get('board_volume', np.nan) if isinstance(x, dict) else np.nan)

    services['RASTREADOR'] = services['serviceIDBoard'].map(
        lambda x: x.get('operational_number', np.nan) if isinstance(x, dict) else np.nan)

    services['CTE'] = services['serviceIDBoard'].map(
        lambda x: x.get('cte', np.nan) if isinstance(x, dict) else np.nan)

    services['VALOR'] = services['serviceIDBoard'].map(
        lambda x: x.get('cte_transfer_cost', np.nan) if isinstance(x, dict) else np.nan)

    services['PESO'] = services['serviceIDBoard'].map(
        lambda x: x.get('taxed_weight', np.nan) if isinstance(x, dict) else np.nan)

    services['realBoardDate'] = services['serviceIDBoard'].map(
        lambda x: x.get('real_board_date', np.nan) if isinstance(x, dict) else np.nan)

    services.to_excel('debugtransferência.xlsx')

    services['boardDate'] = pd.to_datetime(services['realBoardDate']) - dt.timedelta(hours=3)
    services['boardDate'] = services['boardDate'].dt.tz_localize(None)

    services.to_excel('debugtransferência.xlsx')

    services['TRANSPORTADORA'] = services['BASE ORIGEM'].map(get_transp)

    services['baseOrig'] = services['serviceIDRequested.source_branch_id'].map(get_branch)
    services['baseDestCD'] = services['serviceIDRequested.destination_crossdocking_branch_id'].map(get_branch)
    services['baseOrigCD'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_branch)
    services['baseDest'] = services['serviceIDRequested.destination_branch_id'].map(get_branch)

    services['ROTA'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) &
                  (services['BASE ORIGEM'] == services['serviceIDRequested.source_branch_id']),
                  (services['CROSSDOCKING'] == 1) &
                  (services['BASE ORIGEM'] == services['serviceIDRequested.source_crossdocking_branch_id']),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['baseOrig'] + '-' + services['baseDestCD'],
                    services['baseOrigCD'] + '-' + services['baseDest'],
                    services['baseOrig'] + '-' + services['baseDest'],
                    '-'],
        default='')

    services['BUDGET'] = np.select(
        condlist=[services['serviceIDRequested.budgetIDService.transfer_budget'] <= 5,
                  services['PESO'] > services['serviceIDRequested.budgetIDService.franchising']],
        choicelist=[services['VALOR'],
                    np.ceil(services['PESO'] - services['serviceIDRequested.budgetIDService.franchising']) *
                    (services['serviceIDRequested.budgetIDService.price_kg_extra'] * 0.8275) +
                    services['serviceIDRequested.budgetIDService.transfer_budget']],
        default=services['serviceIDRequested.budgetIDService.transfer_budget'])

    services['DESVIO'] = services['BUDGET'] - services['VALOR']

    services.sort_values(by="boardDate", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')

    df_dt += dt.timedelta(days=1)  # Correção para poder buscar somente um dia

    services.drop(services[services['boardDate'] < di_dt].index, inplace=True)
    services.drop(services[services['boardDate'] > df_dt].index, inplace=True)

    user_fb.set('Atualizando dados...')

    cargas['RASTREADOR'] = services['RASTREADOR']
    cargas['PROTOCOLO'] = services['protocol']
    cargas['CLIENTE'] = services['customerIDService.trading_firstname']
    cargas['DATA DA COLETA'] = services['dataColeta'].map(excel_date)
    cargas['DATA DO EMBARQUE'] = services['boardDate'].map(excel_date)
    cargas['TRANSPORTADORA'] = services['TRANSPORTADORA']
    cargas['ROTA'] = services['ROTA']
    cargas['CIDADE ORIGEM'] = services['serviceIDRequested.source_address_id'].map(get_address_city)
    cargas['CIDADE DESTINO'] = services['serviceIDRequested.destination_address_id'].map(get_address_city)
    cargas['CTE'] = services['CTE']
    cargas['VOLUMES'] = services['VOLUMES']
    cargas['PESO FRANQUIA'] = services['serviceIDRequested.budgetIDService.franchising']
    cargas['PESO REAL'] = services['PESO']
    cargas['VALOR CTE'] = services['VALOR']
    cargas['BUDGET'] = services['BUDGET']
    cargas['DESVIO'] = services['DESVIO']
    cargas['CROSSDOCKING'] = services['CROSSDOCKING']
    cargas['order'] = services['boardDate'].map(excel_datetime)

    cargas.sort_values(
        by="order", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')

    # cargas.drop(['order'], axis=1, inplace=True)
    cargas = cargas.loc[cargas['DATA DO EMBARQUE'].notnull()]

    cargas.sort_values(
        by="DESVIO", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

    unique = pd.DataFrame(columns=[])
    unique['TRANSPORTADORA'] = cargas['TRANSPORTADORA']
    unique['RASTREADOR'] = cargas['RASTREADOR']
    unique['CTE'] = cargas['CTE']
    unique['ROTA'] = cargas['ROTA']
    unique['DATA DO EMBARQUE'] = cargas['DATA DO EMBARQUE']

    unique.drop_duplicates(keep="first", inplace=True)

    CoInitialize()

    user_fb.set('Exportando planilhas...')

    CoInitialize()

    file_name = filename.get().replace('/', '\\')

    latam = cargas.loc[cargas['TRANSPORTADORA'] == 'LATAM CARGO']
    gollog = cargas.loc[cargas['TRANSPORTADORA'] == 'GOLLOG']

    unique_latam = pd.DataFrame()
    unique_gollog = pd.DataFrame()

    unique_latam['RASTREADOR'] = np.unique(latam['RASTREADOR'])
    unique_gollog['RASTREADOR'] = np.unique(gollog['RASTREADOR'])

    export_to_excel(cargas, file_name, 'Embarques', 'A2:R5000', autofit=False, change_header=True)
    export_to_excel(unique, file_name, 'Resumo', 'A1:E2000', autofit=False, change_header=True)
    export_to_excel(unique_latam, file_name, 'Desvio_LATAM', 'A2:A400', autofit=False, change_header=False)
    export_to_excel(unique_gollog, file_name, 'Desvio_GOLLOG', 'A2:A400', autofit=False, change_header=False)

    root.after(10, feedback.destroy())


def export_to_excel(df, excel_name, sheet, clear_range, autofit=True, change_header=True):
    app = xw.App(visible=False)
    wb = xw.Book(f'{excel_name}')
    ws = wb.sheets[f'{sheet}']
    app.kill()
    if wb.sheets[f'{sheet}'].api.AutoFilter:
        wb.sheets[f'{sheet}'].api.AutoFilter.ShowAllData()
    ws.range(clear_range).clear_contents()

    if change_header:
        start_write = "A1"
        header_config = 1
    else:
        start_write = "A2"
        header_config = 0

    # Inserção do DataFrame na planilha
    ws[f"{start_write}"].options(pd.DataFrame, header=header_config, index=False, expand='table').value = df

    if autofit:
        ws.autofit('r')


def clear_data(*sheet):

    file_name = filename.get().replace('/', '\\')

    CoInitialize()

    for value in sheet:
        app = xw.App(visible=False)
        wb = xw.Book(f"{file_name}")
        terms = value.split(';')
        ws = wb.sheets[f'{terms[0]}']
        app.kill()
        if wb.sheets[f'{terms[0]}'].api.AutoFilter:
            wb.sheets[f'{terms[0]}'].api.AutoFilter.ShowAllData()
        ws.range(terms[1]).clear_contents()


# noinspection PyGlobalUndefined
def start_thread(target_function, progress_bar_func, arguments=()):

    def check_thread():
        if submit_thread.is_alive():
            root.after(20, check_thread)
        else:
            nonlocal progress_bar_func
            progress_bar_func.stop()

    global submit_thread
    submit_thread = threading.Thread(target=target_function, args=arguments)
    submit_thread.daemon = True
    progress_bar_func.start()
    submit_thread.start()
    root.after(20, check_thread)


def browse_files():
    filetypes = (
        ('Excel', '*.xlsx'),
        ('Excel habilitado para macro', '*.xlsm')
    )

    file_name = fd.askopenfilename(
        title='Selecione o arquivo',
        initialdir='cd',
        filetypes=filetypes
    )

    global filename
    with open("filename.txt", 'w') as w:
        w.write(file_name)
    filename.set(file_name)

    global filename_Label
    filename_Label.destroy()
    filename_Label = ttk.Label(tab3, width=20, text=filename.get(), wraplength=140)
    filename_Label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


header = {"xtoken": "myqhF6Nbzx"}

# Create Object
root = Tk()
root.iconbitmap(resource_path("my_icon.ico"))
root.title('Budget LogLife')
root.resizable(False, False)
# root.iconbitmap('my_icon.ico')
tabs = ttk.Notebook(root)

# Setting tabs

tab1 = ttk.Frame(tabs)
tab3 = ttk.Frame(tabs)

tabs.add(tab1, text='Budget')
tabs.add(tab3, text='Arquivos')

tabs.pack(expand=1, fill="both")

# Set geometry
root.geometry("")

today = dt.datetime.today()

dia = today.day
mes = today.month
ano = today.year

style = ThemedStyle(root)
style.theme_use('breeze')

cal = DateEntry(tab1, selectmode='day', day=dia, month=mes, year=ano, locale='pt_BR',
                 firstweekday='sunday',
                 showweeknumbers=False,
                 bordercolor="white",
                 background="white",
                 disabledbackground="white",
                 headersbackground="white",
                 normalbackground="white",
                 normalforeground='black',
                 headersforeground='black',
                 selectbackground='#00a5e7',
                 selectforeground='white',
                 weekendbackground='white',
                 weekendforeground='black',
                 othermonthforeground='black',
                 othermonthbackground='#E8E8E8',
                 othermonthweforeground='black',
                 othermonthwebackground='#E8E8E8',
                 foreground="black")

cal1 = DateEntry(tab1, selectmode='day', day=dia, month=mes, year=ano, locale='pt_BR',
                 firstweekday='sunday',
                 showweeknumbers=False,
                 bordercolor="white",
                 background="white",
                 disabledbackground="white",
                 headersbackground="white",
                 normalbackground="white",
                 normalforeground='black',
                 headersforeground='black',
                 selectbackground='#00a5e7',
                 selectforeground='white',
                 weekendbackground='white',
                 weekendforeground='black',
                 othermonthforeground='black',
                 othermonthbackground='#E8E8E8',
                 othermonthweforeground='black',
                 othermonthwebackground='#E8E8E8',
                 foreground="black")

cal.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
cal1.grid(column=1, row=1, padx=30, pady=10, sticky="E, W")

# Add Button and Label
ttk.Button(
    tab1,
    text="Atualizar dados",
    command=lambda: start_thread(budget, progressbar)).grid(
    column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(
    tab1,
    text="Limpar dados",
    command=lambda: start_thread(clear_data, progressbar, arguments=["Embarques;A2:Q10000",
                                                                     "Desvio_LATAM;A2:A400",
                                                                     "Desvio_GOLLOG;A2:A400"])).grid(
    column=2, row=1, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(tab3, text="Procurar",
           command=browse_files).grid(column=2, row=0, padx=10, pady=10, sticky="E, W")


# Read file name

filename = StringVar()

try:
    with open('filename.txt') as m:
        text = m.read()
    lines = text.split('\n')
    filename.set(lines[0])
except FileNotFoundError:
    filename.set('SLA.xlsx')

filename_Label = ttk.Label(tab3, width=20, text=filename.get(), wraplength=140)
filename_Label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)

# Labels

ttk.Label(tab1, text="      Data Inicial:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab1, text="       Data Final:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

ttk.Label(tab3, text="Nome do Arquivo:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')

# Progress Bar

progressbar = ttk.Progressbar(tab1, mode='indeterminate')
progressbar.grid(column=0, row=5, sticky='W, E', columnspan=3)

# Auto resize tabs

tab1.rowconfigure(2, weight=2)
tab1.columnconfigure(0, weight=1)
tab1.columnconfigure(1, weight=1)
tab1.columnconfigure(2, weight=1)

tab3.rowconfigure(0, weight=1)
tab3.columnconfigure(0, weight=1)
tab3.columnconfigure(1, weight=1)
tab3.columnconfigure(2, weight=1)

# Execute Tkinter
root.mainloop()
