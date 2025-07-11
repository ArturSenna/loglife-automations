import pandas as pd
import numpy as np
import datetime as dt
import xlwings as xw
import requests
from json import loads


def get_collector(col_id):
    try:
        col_name = collector.loc[collector['id'] == col_id, 'trading_name'].values.item()
    except ValueError:
        col_name = ""
    return col_name


def get_transp(id_t):
    try:
        transp = bases.loc[bases['id'] == id_t, 'transportadora'].values.item()
    except ValueError:
        transp = ""
    return transp


def get_branch(id_t):
    try:
        branch = bases.loc[bases['id'] == id_t, 'nome'].values.item()
    except ValueError:
        branch = ''
    return branch


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


header = {"xtoken": "myqhF6Nbzx"}

print("R...")

a = requests.get('https://transportebiologico.com.br/api/public/address', headers=header)
b = requests.get('https://transportebiologico.com.br/api/public/branch', headers=header)
c = requests.get('https://transportebiologico.com.br/api/public/collector', headers=header)
s = requests.get('https://transportebiologico.com.br/api/public/service', headers=header)
a_json = loads(a.text)
b_json = loads(b.text)
c_json = loads(c.text)
s_json = loads(s.text)
address = pd.json_normalize(a_json)
bases = pd.json_normalize(b_json)
collector = pd.json_normalize(c_json)
services = pd.json_normalize(s_json)

services.drop(services[services['step'] == 'toCollectService'].index, inplace=True)
services.drop(services[services['step'] == 'collectingService'].index, inplace=True)
services.drop(services[services['step'] == 'toBoardService'].index, inplace=True)
services.drop(services[services['step'] == 'boardingService'].index, inplace=True)

print("O...")

services['ETAPA'] = np.select(
    condlist=[
        services['step'] == 'availableService',
        services['step'] == 'toAllocateService',
        services['step'] == 'toDeliveryService',
        services['step'] == 'deliveringService',
        services['step'] == 'toLandingService',
        services['step'] == 'landingService',
        services['step'] == 'toBoardValidate'],
    choicelist=[
        'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
        'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE'],
    default=0)

services['CROSSDOCKING'] = np.where(services['serviceIDRequested.crossdocking_collector_id'].str.len() > 0, 1, 0)

services['Identificador'] = np.select(
    condlist=[(services['serviceIDBoard'].str.len() == 1),
              (services['serviceIDBoard'].str.len() == 2)],
    choicelist=[1, 2],
    default=0
)

services['Coletador de Origem'] = services['serviceIDRequested.source_collector_id'].map(get_collector)
services['Coletador intermediário'] = services['serviceIDRequested.crossdocking_collector_id'].map(get_collector)
services['Coletador de destino'] = services['serviceIDRequested.destination_collector_id'].map(get_collector)

services['transportadora1'] = services['serviceIDRequested.source_branch_id'].map(get_transp)
services['transportadora2'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_transp)

services['Base_origem'] = services['serviceIDRequested.source_branch_id'].map(get_branch)
services['cd_Base_destino'] = services['serviceIDRequested.destination_crossdocking_branch_id'].map(get_branch)
services['cd_Base_origem'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_branch)
services['Base_destino'] = services['serviceIDRequested.destination_branch_id'].map(get_branch)

services['EMBARQUE'] = services['serviceIDBoard'].str[0]
services['EMBARQUE2'] = services['serviceIDBoard'].str[-1]

services['VOLUMES1'] = services['EMBARQUE'].map(
    lambda x: x.get('board_volume', np.nan) if isinstance(x, dict) else np.nan)

services['RASTREADOR1'] = services['EMBARQUE'].map(
    lambda x: x.get('operational_number', np.nan) if isinstance(x, dict) else np.nan)

services['CTE1'] = services['EMBARQUE'].map(
    lambda x: x.get('cte', np.nan) if isinstance(x, dict) else np.nan)

services['VOLUMES2'] = services['EMBARQUE2'].map(
    lambda x: x.get('board_volume', np.nan) if isinstance(x, dict) else np.nan)

services['RASTREADOR2'] = services['EMBARQUE2'].map(
    lambda x: x.get('operational_number', np.nan) if isinstance(x, dict) else np.nan)

services['CTE2'] = services['EMBARQUE2'].map(
    lambda x: x.get('cte', np.nan) if isinstance(x, dict) else np.nan)

services['tempo1'] = services['EMBARQUE'].map(
    lambda x: x.get('createdAt', np.nan) if isinstance(x, dict) else np.nan)

services['tempo2'] = services['EMBARQUE2'].map(
    lambda x: x.get('createdAt', np.nan) if isinstance(x, dict) else np.nan)

services['VOLUMES'] = np.where(services['tempo1'] >= services['tempo2'], services['VOLUMES1'], services['VOLUMES2'])
services['CTE'] = np.where(services['tempo1'] >= services['tempo2'], services['CTE1'], services['CTE2'])
services['RASTREADOR'] = np.where(
    services['tempo1'] >= services['tempo2'], services['RASTREADOR1'], services['RASTREADOR2'])

services['COLETADOR DESTINO'] = np.select(
    condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
              (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 2),
              (services['CROSSDOCKING'] == 0),
              (services['serviceIDRequested.service_type'] == "DEDICADO")],
    choicelist=[services['Coletador intermediário'],
                services['Coletador de destino'],
                services['Coletador de destino'],
                '-'],
    default=0)

services['BASE ORIGEM'] = np.select(
    condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
              (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 2),
              (services['CROSSDOCKING'] == 0),
              (services['serviceIDRequested.service_type'] == "DEDICADO")],
    choicelist=[services['Base_origem'],
                services['cd_Base_origem'],
                services['Base_origem'],
                '-'],
    default=0)

services['BASE DESTINO'] = np.select(
    condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
              (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 2),
              (services['CROSSDOCKING'] == 0),
              (services['serviceIDRequested.service_type'] == "DEDICADO")],
    choicelist=[services['cd_Base_destino'],
                services['Base_destino'],
                services['Base_destino'],
                '-'],
    default=0)

services['Transportadora'] = np.select(
    condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
              (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 2),
              (services['CROSSDOCKING'] == 0),
              (services['serviceIDRequested.service_type'] == "DEDICADO")],
    choicelist=[services['transportadora1'],
                services['transportadora2'],
                services['transportadora1'],
                '-'],
    default=0)

services['DATA DA COLETA'] = pd.to_datetime(services['collect_date']) - dt.timedelta(hours=3)
services['DATA DA COLETA'] = services['DATA DA COLETA'].dt.tz_localize(None)

services['DATA DE ENTREGA'] = pd.to_datetime(
    services['serviceIDRequested.delivery_date'], dayfirst=True
) - dt.timedelta(hours=3)
services['DATA DE ENTREGA'] = services['DATA DE ENTREGA'].dt.tz_localize(None)

services['HORÁRIO LIMITE DE ENTREGA'] = pd.to_datetime(
    services['serviceIDRequested.delivery_hour'], dayfirst=True
) - dt.timedelta(hours=3)
services['HORÁRIO LIMITE DE ENTREGA'] = services['DATA DE ENTREGA'].dt.tz_localize(None)

services['CIDADES ORIGEM'] = services['serviceIDRequested.source_address_id'].map(get_address_city)
services['CIDADES DESTINO'] = services['serviceIDRequested.destination_address_id'].map(get_address_city)

print("E...")

cg = services[[
    'protocol',
    'customerIDService.trading_firstname',
    'DATA DA COLETA',
    'CIDADES ORIGEM',
    'CIDADES DESTINO',
    'RASTREADOR',
    'CTE',
    'serviceIDRequested.vehicle',
    'VOLUMES',
    'Transportadora',
    'BASE ORIGEM',
    'BASE DESTINO',
    'COLETADOR DESTINO',
    'DATA DE ENTREGA',
    'HORÁRIO LIMITE DE ENTREGA',
    'serviceIDRequested.gelo_seco',
    'ETAPA',
    'serviceIDRequested.service_type',
    'CROSSDOCKING'
]].copy()

cg.sort_values(by="HORÁRIO LIMITE DE ENTREGA", axis=0, ascending=True, inplace=True, kind='quicksort',
               na_position='last')
cg.sort_values(by="DATA DE ENTREGA", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')
cg.sort_values(by="RASTREADOR", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')
cg.sort_values(by="Transportadora", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

cg1 = cg.loc[(cg['ETAPA'] != 'EM ROTA DE ENTREGA') & (cg['ETAPA'] != 'ENTREGANDO')]
cg2 = cg.loc[(cg['ETAPA'] == 'EM ROTA DE ENTREGA') | (cg['ETAPA'] == 'ENTREGANDO')]

filename = "RELAÇÃO DE CARGAS.XLSX"
app = xw.App(visible=False)
wb = xw.Book(filename)
ws1 = wb.sheets['Desembarque']
ws2 = wb.sheets['Entrega']
if wb.sheets['Desembarque'].api.AutoFilter:
    wb.sheets['Desembarque'].api.AutoFilter.ShowAllData()
if wb.sheets['Entrega'].api.AutoFilter:
    wb.sheets['Entrega'].api.AutoFilter.ShowAllData()

ws1.range('B4:T5000').clear_contents()
ws2.range('B4:T5000').clear_contents()

# Inserção do DataFrame na planilha
ws1["B4"].options(pd.DataFrame, header=0, index=False, expand='table').value = cg1
ws2["B4"].options(pd.DataFrame, header=0, index=False, expand='table').value = cg2
