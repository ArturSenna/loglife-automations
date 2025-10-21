import time
import requests
import threading
import numpy as np
import pandas as pd
import xlwings as xw
import datetime as dt
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from json import loads
from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
from pythoncom import CoInitialize
import os


class RequestDataFrame:

    def __init__(self):
        self.headers = {"xtoken": "myqhF6Nbzx"}
        details = {"email": "ARTURSENNA@LOGLIFELOGISTICA.COM.BR", "password": "A3928024854c#"}
        key = requests.post("https://transportebiologico.com.br/api/sessions", json=details)
        key_json = loads(key.text)
        self.auth = {"authorization": key_json["token"]}

    def request_public(self, link, request_type="get"):

        if request_type == "get":
            response = requests.get(link, headers=self.headers)
        elif request_type == "post":
            response = requests.post(link, headers=self.headers)
        else:
            response = requests.get(link, headers=self.headers)  # CHANGE LATER

        response_json = loads(response.text)
        dataframe = pd.json_normalize(response_json)

        return dataframe

    def request_private(self, link: str, request_type="get", payload: dict = None) -> pd.DataFrame:

        if request_type == "get":
            response = requests.get(link, headers=self.auth, data=payload)
        elif request_type == "post":
            response = requests.post(link, headers=self.auth, data=payload)
        else:
            response = requests.get(link, headers=self.auth)  # CHANGE LATER

        print(response.text)

        response_json = loads(response.text)
        dataframe = pd.json_normalize(response_json)

        return dataframe

    def post_file(self, link, file, file_type="csv_file", file_format="text/csv", upload_type=None):

        payload = {
            file_type: (
                file,
                open(file, "rb"),
                file_format,
                {"Expires": "0"},
            )
        }

        if upload_type is not None:
            upload_type = {
                "type": upload_type
            }

        response = requests.post(url=link, headers=self.auth, files=payload, data=upload_type)

        return response

    def post_private(self, link, payload):

        response = requests.post(url=link, headers=self.auth, data=payload)

        return response


def sla_last_mile():
    def get_address_city(add):
        try:
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
        except ValueError:
            erro = "ERRO"
            return erro

    def get_address(add):

        try:
            add_name = address.loc[address['id'] == add, 'trading_name'].values.item()
            add_branch = address.loc[address['id'] == add, 'branch'].values.item()
            # add_city = address.loc[address['id'] == add, 'cityIDAddress.name'].values.item()
            # return add_name + ' - ' + add_branch
            return add_name + ' - ' + add_branch
        except ValueError:
            erro = 'ERRO'
            return erro

    def get_collector(col):
        col_name = collector.loc[collector['id'] == col, 'trading_name'].values.item()
        return col_name

    def excel_date(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days)

    def excel_time(date1):
        try:
            temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
            delta = date1 - temp
            return float(delta.seconds) / 86400
        except Error as e:
            print(e)
            return 0

    def excel_datetime(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days) + float(delta.seconds) / 86400

    def get_occurence_time(dict_list, oc_list):

        if type(dict_list) is float or len(dict_list) == 0:
            pass
            return None
        else:
            oc_creation_time = []
            oc_datetime = []

            for ocurrence in dict_list:

                oc_type = ocurrence.get('intercurrence', np.nan) if isinstance(ocurrence, dict) else np.nan

                if oc_type not in oc_list:
                    pass
                    return None

                else:
                    oc_time = ocurrence.get('createdAt', np.nan) if isinstance(ocurrence, dict) else np.nan
                    oc_hour = ocurrence.get('occurrence_hour', np.nan) if isinstance(ocurrence, dict) else np.nan
                    oc_date = ocurrence.get('occurrence_date', np.nan) if isinstance(ocurrence, dict) else np.nan

                    try:
                        oc_dt = excel_date(pd.to_datetime(oc_date).tz_localize(None)) + \
                                excel_time(pd.to_datetime(oc_hour).tz_localize(None))
                    except AttributeError:
                        oc_dt = excel_time(pd.to_datetime(oc_hour).tz_localize(None))

                    oc_time = excel_date(pd.to_datetime(oc_time).tz_localize(None)) + \
                              excel_time(pd.to_datetime(oc_time).tz_localize(None))

                    oc_creation_time.append(oc_time)
                    oc_datetime.append(oc_dt)

                    position = oc_creation_time.index(max(oc_creation_time))

                    return oc_datetime[position]

    def get_hub_state(hub_id):
        try:
            state = hubs.loc[hubs['id'] == hub_id, 'state'].values.item()
        except TypeError:
            state = "ERRO"
        # except ValueError:
        #     regional = "VAZIO"
        return state

    def get_uf(state):
        try:
            uf = states.loc[states['ESTADO'] == state, 'UF'].values.item()
        except ValueError:
            uf = ''
        return uf

    def get_hub_regional(hub_id):
        try:
            _regional = hubs.loc[hubs['id'] == hub_id, 'regional'].values.item()
        except TypeError:
            _regional = "ERRO"
        # except ValueError:
        #     regional = "VAZIO"
        return _regional

    def get_driver_name(driver_id):
        try:
            driver_name = driver.loc[driver['id'] == driver_id, 'firstname'].values.item()
            driver_lastname = driver.loc[driver['id'] == driver_id, 'lastname'].values.item()
        except ValueError:
            driver_name = ''
            driver_lastname = ''
        return driver_name + ' ' + driver_lastname

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
    feedback.grid(column=1, row=2, padx=10, pady=4, columnspan=2, sticky='W, E')

    user_fb.set('Requisitando dados...')

    # Requesting data from API

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')
    hubs = r.request_private('https://transportebiologico.com.br/api/hub')
    driver = r.request_private('https://transportebiologico.com.br/api/driver')
    services_ongoing = r.request_public('https://transportebiologico.com.br/api/public/service')
    services_finalized = r.request_public(
        f'https://transportebiologico.com.br/api/public/service/finalized/?startFilter={di}&endFilter={df}', 'post')

    coletas_sla = pd.DataFrame(columns=[])
    entregas_sla = pd.DataFrame(columns=[])

    services = pd.concat([services_ongoing, services_finalized], ignore_index=True)
    services = services.loc[services['is_business'] == False]
    # services.drop(
    #     services[services['customerIDService.trading_firstname'] != 'LOCUS - ANATOMIA PATOLOGICA E CITOLOGIA'].index,
    #     inplace=True)

    services.to_excel('ServicesAPI.xlsx', index=False)

    coletas = services
    entregas = services

    tolerance = excel_time(pd.to_datetime('00:30'))  # Tolerância

    df_dt += dt.timedelta(days=1)  # Correção para poder buscar somente um dia

    user_fb.set('Organizando informações das coletas...')  # Tratando dados da coleta

    coletas['length'] = coletas['serviceIDCollect'].str.len()

    coletas.drop(coletas[coletas['length'] == 0].index, inplace=True)
    coletas.drop(coletas[coletas['step'] == 'collectingService'].index, inplace=True)

    coletas['REGIONAL'] = coletas['serviceIDRequested.budgetIDService.source_hub_id'].map(get_hub_regional)

    coletas['RESPONSÁVEL'] = np.select(
        condlist=[coletas['REGIONAL'] == 1,
                  coletas['REGIONAL'] == 2,
                  coletas['REGIONAL'] == 3,
                  coletas['REGIONAL'] == 4,
                  coletas['REGIONAL'] == 5],
        choicelist=['Filial SP', 'Filial DF', 'Filial RJ', 'Brasil', 'Brasil'],
        default=''
    )

    coletas['OC_HORA_COLETA'] = coletas['occurrenceIDService'].map(
        lambda x: get_occurence_time(x, ['ATRASO NA COLETA'])
    )
    coletas['OC_HORA_COLETA'] = pd.to_datetime(coletas['OC_HORA_COLETA']) - dt.timedelta(hours=3)

    coletas['ESTADO'] = services['serviceIDRequested.budgetIDService.source_hub_id'].map(get_hub_state)
    coletas['UF'] = coletas['ESTADO'].map(get_uf)

    coletas = coletas.explode('serviceIDCollect')

    coletas['collectDatetime'] = coletas['serviceIDCollect'].map(
        lambda x: x.get('real_arrival_timestamp', np.nan) if isinstance(x, dict) else np.nan)

    coletas['registeredDateTime'] = coletas['serviceIDCollect'].map(
        lambda x: x.get('arrival_timestamp', np.nan) if isinstance(x, dict) else np.nan)

    coletas['driverID'] = coletas['serviceIDCollect'].map(
        lambda x: x.get('driver_id', np.nan) if isinstance(x, dict) else np.nan)

    coletas['collectDatetime'] = np.where(coletas['collectDatetime'].isnull(),
                                          coletas['registeredDateTime'],
                                          coletas['collectDatetime'])

    coletas['collectDatetime'] = pd.to_datetime(coletas['collectDatetime']) - dt.timedelta(hours=3)
    coletas['collectDatetime'] = coletas['collectDatetime'].dt.tz_localize(None)

    coletas.drop(coletas[coletas['collectDatetime'] < di_dt].index, inplace=True)
    coletas.drop(coletas[coletas['collectDatetime'] > df_dt].index, inplace=True)

    coletas['idAddress'] = coletas['serviceIDCollect'].map(
        lambda x: x.get('address_id', np.nan) if isinstance(x, dict) else np.nan)

    coletas['idAddress'].replace('', np.nan, inplace=True)
    coletas.dropna(subset=['idAddress'], inplace=True)

    coletas['LOCAL DE COLETA'] = coletas['idAddress'].map(get_address)

    coletas['dataReal'] = coletas['collectDatetime'].map(excel_date)
    coletas['horaReal'] = coletas['collectDatetime'].map(excel_time)

    coletas['dataColeta'] = pd.to_datetime(coletas['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    coletas['dataColeta'] = coletas['dataColeta'].dt.tz_localize(None)
    coletas['dataColeta'] = coletas['dataColeta'].map(excel_date)

    coletas['hourEnd'] = pd.to_datetime(coletas['serviceIDRequested.collect_hour_end']) - dt.timedelta(hours=3)
    coletas['hourEnd'] = coletas['hourEnd'].dt.tz_localize(None)
    coletas['hourEnd'] = coletas['hourEnd'].map(excel_time)

    coletas['OC_HORA_COLETA_EXCEL'] = coletas['OC_HORA_COLETA'].map(excel_time)

    coletas['datahoraLimite'] = np.where(coletas['OC_HORA_COLETA'].isnull(),
                                         coletas['dataColeta'] + coletas['hourEnd'],
                                         coletas['dataColeta'] + coletas['OC_HORA_COLETA_EXCEL'])

    coletas['datahoraReal'] = coletas['dataReal'] + coletas['horaReal']

    coletas['datahoraReal'] = np.select(
        condlist=[coletas['UF'] == 'AC',
                  (coletas['UF'] == 'AM') | (coletas['UF'] == 'RR') | (coletas['UF'] == 'RO') | (coletas['UF'] == 'MT')
                  | (coletas['UF'] == 'MS')],
        choicelist=[coletas['datahoraReal'] - excel_time(pd.to_datetime('02:00')),
                    coletas['datahoraReal'] - excel_time(pd.to_datetime('01:00'))],
        default=coletas['datahoraReal']
    )

    coletas['ATRASO'] = np.where(
        coletas['datahoraReal'] <= (coletas['datahoraLimite'] + tolerance),
        'NÃO',
        'SIM'
    )

    coletas.sort_values(
        by="collectDatetime", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')

    user_fb.set('Organizando informações das entregas...')  # Tratando dados da entrega

    entregas['OC_HORA_ENTREGA'] = entregas['occurrenceIDService'].map(
        lambda x: get_occurence_time(x, ['CORTE DE VOO (NÃO ALOCADO VOO PLANEJADO)',
                                         'CANCELAMENTO DE VOO',
                                         'ATRASO NA ENTREGA'])
    )

    entregas['REGIONAL'] = entregas['serviceIDRequested.budgetIDService.destination_hub_id'].map(get_hub_regional)

    entregas['RESPONSÁVEL'] = np.select(
        condlist=[entregas['REGIONAL'] == 1,
                  entregas['REGIONAL'] == 2,
                  entregas['REGIONAL'] == 3,
                  entregas['REGIONAL'] == 4,
                  entregas['REGIONAL'] == 5],
        choicelist=['Filial SP', 'Filial DF', 'Filial RJ', 'Brasil', 'Brasil'],
        default=0
    )

    entregas['ESTADO'] = services['serviceIDRequested.budgetIDService.destination_hub_id'].map(get_hub_state)
    entregas['UF'] = entregas['ESTADO'].map(get_uf)

    entregas = entregas.explode('serviceIDDelivery')

    entregas['deliveryDatetime'] = entregas['serviceIDDelivery'].map(
        lambda x: x.get('real_arrival_timestamp', np.nan) if isinstance(x, dict) else np.nan)

    entregas['registeredDateTime'] = entregas['serviceIDDelivery'].map(
        lambda x: x.get('arrival_timestamp', np.nan) if isinstance(x, dict) else np.nan)

    entregas['driverID'] = entregas['serviceIDDelivery'].map(
        lambda x: x.get('driver_id', np.nan) if isinstance(x, dict) else np.nan)

    # entregas['RESPONSÁVEL'] = entregas['serviceIDDelivery'].map(
    #     lambda x: x.get('responsible_name', np.nan) if isinstance(x, dict) else np.nan)

    entregas['deliveryDatetime'] = np.where(entregas['deliveryDatetime'].isnull(),
                                            entregas['registeredDateTime'],
                                            entregas['deliveryDatetime'])

    entregas['deliveryDatetime'] = pd.to_datetime(entregas['deliveryDatetime']) - dt.timedelta(hours=3)
    entregas['deliveryDatetime'] = entregas['deliveryDatetime'].dt.tz_localize(None)

    entregas.drop(entregas[entregas['deliveryDatetime'] < di_dt].index, inplace=True)
    entregas.drop(entregas[entregas['deliveryDatetime'] > df_dt].index, inplace=True)

    entregas.drop(entregas[entregas['step'] != 'finishedService'].index, inplace=True)

    entregas['idAddress'] = entregas['serviceIDDelivery'].map(
        lambda x: x.get('address_id', np.nan) if isinstance(x, dict) else np.nan)

    entregas['idAddress'].replace('', np.nan, inplace=True)
    entregas.dropna(subset=['idAddress'], inplace=True)

    entregas['LOCAL DE ENTREGA'] = entregas['idAddress'].map(get_address)

    entregas['dataReal'] = entregas['deliveryDatetime'].map(excel_date)
    entregas['horaReal'] = entregas['deliveryDatetime'].map(excel_time)

    entregas['dataEntrega'] = pd.to_datetime(entregas['serviceIDRequested.delivery_date']) - dt.timedelta(hours=3)
    entregas['dataEntrega'] = entregas['dataEntrega'].dt.tz_localize(None)
    entregas['dataEntrega'] = entregas['dataEntrega'].map(excel_date)

    entregas['hourEnd'] = pd.to_datetime(entregas['serviceIDRequested.delivery_hour']) - dt.timedelta(hours=3)
    entregas['hourEnd'] = entregas['hourEnd'].dt.tz_localize(None)
    entregas['hourEnd'] = entregas['hourEnd'].map(excel_time)

    entregas['datahoraLimite'] = np.where(entregas['OC_HORA_ENTREGA'].isnull(),
                                           entregas['dataEntrega'] + entregas['hourEnd'],
                                           entregas['OC_HORA_ENTREGA'])

    entregas['datahoraReal'] = entregas['dataReal'] + entregas['horaReal']

    entregas['datahoraReal'] = np.select(
        condlist=[entregas['UF'] == 'AC',
                  (entregas['UF'] == 'AM') | (entregas['UF'] == 'RR') | (entregas['UF'] == 'RO') |
                  (entregas['UF'] == 'MT') | (entregas['UF'] == 'MS')],
        choicelist=[entregas['datahoraReal'] - excel_time(pd.to_datetime('02:00')),
                    entregas['datahoraReal'] - excel_time(pd.to_datetime('01:00'))],
        default=entregas['datahoraReal']
    )

    entregas['ATRASO'] = np.where(
        entregas['datahoraReal'] <= (entregas['datahoraLimite'] + tolerance),
        'NÃO',
        'SIM'
    )

    entregas.sort_values(
        by="deliveryDatetime", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')

    user_fb.set('Atualizando dados...')

    coletas_sla['PROTOCOLO'] = coletas['protocol']
    coletas_sla['CLIENTE'] = coletas['customerIDService.trading_firstname']
    coletas_sla['CIDADE ORIGEM'] = coletas['serviceIDRequested.source_address_id'].map(get_address_city)
    coletas_sla['CIDADE DESTINO'] = coletas['serviceIDRequested.destination_address_id'].map(get_address_city)
    coletas_sla['COLETADOR ORIGEM'] = coletas['serviceIDRequested.source_collector_id'].map(get_collector)
    coletas_sla['LOCAL'] = coletas['LOCAL DE COLETA']
    coletas_sla['HORA LIMITE DE COLETA'] = coletas['datahoraLimite']
    coletas_sla['HORA COLETADO'] = coletas['datahoraReal']
    coletas_sla['ATRASO'] = coletas['ATRASO']
    coletas_sla['UF'] = coletas['UF']
    coletas_sla['RESPONSÁVEL'] = coletas['RESPONSÁVEL']
    coletas_sla['VEÍCULO'] = coletas['serviceIDRequested.vehicle']
    coletas_sla['MOTORISTA'] = coletas['driverID'].map(get_driver_name)

    entregas_sla['PROTOCOLO'] = entregas['protocol']
    entregas_sla['CLIENTE'] = entregas['customerIDService.trading_firstname']
    entregas_sla['CIDADE ORIGEM'] = entregas['serviceIDRequested.source_address_id'].map(get_address_city)
    entregas_sla['CIDADE DESTINO'] = entregas['serviceIDRequested.destination_address_id'].map(get_address_city)
    entregas_sla['COLETADOR DESTINO'] = entregas['serviceIDRequested.destination_collector_id'].map(get_collector)
    entregas_sla['LOCAL'] = entregas['LOCAL DE ENTREGA']
    entregas_sla['HORA LIMITE DE ENTREGA'] = entregas['datahoraLimite']
    entregas_sla['HORA ENTREGUE'] = entregas['datahoraReal']
    entregas_sla['ATRASO'] = entregas['ATRASO']
    entregas_sla['UF'] = entregas['UF']
    entregas_sla['RESPONSÁVEL'] = entregas['RESPONSÁVEL']
    entregas_sla['VEÍCULO'] = entregas['serviceIDRequested.vehicle']
    entregas_sla['MOTORISTA'] = entregas['driverID'].map(get_driver_name)
    # entregas_sla['RESPONSÁVEL'] = entregas['RESPONSÁVEL']

    user_fb.set('Exportando planilhas...')

    CoInitialize()

    file_name = filename.get().replace('/', '\\')

    export_to_excel(coletas_sla, file_name, 'Coletas', 'A1:M20000')
    time.sleep(0.1)
    export_to_excel(entregas_sla, file_name, 'Entregas', 'A1:M20000')

    root.after(10, feedback.destroy())


def sla_transferencia():
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

    def get_branch(id_t):
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

    def excel_date(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days)

    def excel_time(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.seconds) / 86400

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
    feedback.grid(column=1, row=3, padx=10, pady=4, columnspan=2, sticky='W, E')

    user_fb.set('Requisitando dados...')

    # Requesting data from API

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    bases = r.request_public('https://transportebiologico.com.br/api/public/branch')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')
    services_ongoing = r.request_public('https://transportebiologico.com.br/api/public/service')
    services_finalized = r.request_public(
        f'https://transportebiologico.com.br/api/public/service/finalized/?startFilter={di}&endFilter={df}', 'post')

    transf_sla = pd.DataFrame(columns=[])

    services = pd.concat([services_ongoing, services_finalized], ignore_index=True)
    services = services.loc[services['is_business'] == False]

    transf = services

    transf['ETAPA'] = np.select(
        condlist=[
            transf['step'] == 'availableService',
            transf['step'] == 'toAllocateService',
            transf['step'] == 'toDeliveryService',
            transf['step'] == 'deliveringService',
            transf['step'] == 'toLandingService',
            transf['step'] == 'landingService',
            transf['step'] == 'toBoardValidate',
            transf['step'] == 'toCollectService',
            transf['step'] == 'collectingService',
            transf['step'] == 'toBoardService',
            transf['step'] == 'boardingService',
            transf['step'] == 'finishedService'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE', 'AGENDADO', 'COLETANDO',
            'EM ROTA DE EMBARQUE', 'EMBARCANDO SERVIÇO', 'FINALIZADO'],
        default=0)

    # Tratando dados da transferência

    user_fb.set('Organizando informações...')

    col = collector.rename(columns={
        'id': 'serviceIDRequested.crossdocking_collector_id', 'trading_name': 'Coletador intermediário'})

    transf = pd.merge(
        transf, col[['serviceIDRequested.crossdocking_collector_id', 'Coletador intermediário']],
        on='serviceIDRequested.crossdocking_collector_id', how='left')

    transf['CROSSDOCKING'] = np.where(transf['Coletador intermediário'].str.len() > 0, 1, 0)

    transf['Identificador'] = np.select(
        condlist=[(transf['serviceIDBoard'].str.len() == 1) & (transf['ETAPA'] != 'EM ROTA DE EMBARQUE'),
                  (transf['serviceIDBoard'].str.len() == 1) & (transf['ETAPA'] == 'EM ROTA DE EMBARQUE'),
                  (transf['serviceIDBoard'].str.len() == 2)],
        choicelist=[0, 1, 1],
        default=0
    )

    transf['length'] = transf['serviceIDAllocate'].str.len()

    transf.drop(transf[transf['length'] == 0].index, inplace=True)
    transf.drop(transf[transf['step'] == 'toBoardValidate'].index, inplace=True)
    transf.drop(transf[transf['step'] == 'toAllocateValidate'].index, inplace=True)

    transf['collectDatetime'] = pd.to_datetime(transf['collect_date']) - dt.timedelta(hours=3)
    transf['collectDatetime'] = transf['collectDatetime'].dt.tz_localize(None)

    df_dt += dt.timedelta(days=1)  # Correção para poder buscar somente um dia

    transf.drop(transf[transf['collectDatetime'] < di_dt].index, inplace=True)
    transf.drop(transf[transf['collectDatetime'] > df_dt].index, inplace=True)

    transf.to_excel('Transferência.xlsx', index=False)

    transf = transf.explode(['serviceIDAllocate', 'serviceIDBoard'], True)

    transf['availableDate'] = transf['serviceIDAllocate'].map(
        lambda x: x.get('allocate_availability_date', np.nan) if isinstance(x, dict) else np.nan)

    transf['availableTime'] = transf['serviceIDAllocate'].map(
        lambda x: x.get('allocate_availability_hour', np.nan) if isinstance(x, dict) else np.nan)

    transf['BASE ORIGEM'] = transf['serviceIDBoard'].map(
        lambda x: x.get('branch_id', np.nan) if isinstance(x, dict) else np.nan)

    transf['RASTREADOR'] = transf['serviceIDBoard'].map(
        lambda x: x.get('operational_number', np.nan) if isinstance(x, dict) else np.nan)

    transf['realBoardDate'] = transf['serviceIDBoard'].map(
        lambda x: x.get('real_board_date', np.nan) if isinstance(x, dict) else np.nan)

    transf['availableDate'] = pd.to_datetime(transf['availableDate']) - dt.timedelta(hours=3)
    transf['availableTime'] = pd.to_datetime(transf['availableTime']) - dt.timedelta(hours=3)
    transf['boardDate'] = pd.to_datetime(transf['realBoardDate']) - dt.timedelta(hours=3)

    transf['availableDate'] = transf['availableDate'].dt.tz_localize(None)
    transf['availableTime'] = transf['availableTime'].dt.tz_localize(None)
    transf['boardDate'] = transf['boardDate'].dt.tz_localize(None)

    transf['dataReal'] = transf['availableDate'].map(excel_date)
    transf['horaReal'] = transf['availableTime'].map(excel_time)

    transf.to_excel('Transferência.xlsx', index=False)

    transf['previewHour'] = np.select(
        condlist=[(transf['CROSSDOCKING'] == 1) & (transf['Identificador'] == 0),
                  (transf['CROSSDOCKING'] == 1) & (transf['Identificador'] == 1),
                  (transf['CROSSDOCKING'] == 0)],
        choicelist=[transf['serviceIDRequested.crossdocking_availability_forecast_time'],
                    transf['serviceIDRequested.availability_forecast_time'],
                    transf['serviceIDRequested.availability_forecast_time']],
        default=''
    )

    transf['previewDate'] = np.select(
        condlist=[(transf['CROSSDOCKING'] == 1) & (transf['Identificador'] == 0),
                  (transf['CROSSDOCKING'] == 1) & (transf['Identificador'] == 1),
                  (transf['CROSSDOCKING'] == 0)],
        choicelist=[transf['serviceIDRequested.crossdocking_availability_forecast_day'],
                    transf['serviceIDRequested.availability_forecast_day'],
                    transf['serviceIDRequested.availability_forecast_day']],
        default=''
    )

    transf['previewHour'] = pd.to_datetime(transf['previewHour']) - dt.timedelta(hours=3)
    transf['previewDate'] = pd.to_datetime(transf['previewDate']) - dt.timedelta(hours=3)
    transf['previewHour'] = transf['previewHour'].dt.tz_localize(None)
    transf['previewDate'] = transf['previewDate'].dt.tz_localize(None)
    transf['previewHour'] = transf['previewHour'].map(excel_time)
    transf['previewDate'] = transf['previewDate'].map(excel_date)

    tolerance = excel_time(pd.to_datetime('00:41'))  # Tolerância

    transf['ATRASO'] = np.where(
        (transf['dataReal'] + transf['horaReal']) <= (transf['previewDate'] + transf['previewHour'] + tolerance),
        'NÃO',
        'SIM'
    )

    transf['baseOrig'] = transf['serviceIDRequested.source_branch_id'].map(get_branch)
    transf['baseDestCD'] = transf['serviceIDRequested.destination_crossdocking_branch_id'].map(get_branch)
    transf['baseOrigCD'] = transf['serviceIDRequested.source_crossdocking_branch_id'].map(get_branch)
    transf['baseDest'] = transf['serviceIDRequested.destination_branch_id'].map(get_branch)

    transf['ROTA'] = np.select(
        condlist=[(transf['CROSSDOCKING'] == 1) &
                  (transf['BASE ORIGEM'] == transf['serviceIDRequested.source_branch_id']),
                  (transf['CROSSDOCKING'] == 1) &
                  (transf['BASE ORIGEM'] == transf['serviceIDRequested.source_crossdocking_branch_id']),
                  (transf['CROSSDOCKING'] == 0),
                  (transf['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[transf['baseOrig'] + ' → ' + transf['baseDestCD'],
                    transf['baseOrigCD'] + ' → ''-' + transf['baseDest'],
                    transf['baseOrig'] + ' → ' + transf['baseDest'],
                    '-'],
        default='')

    user_fb.set('Atualizando dados...')

    transf.sort_values(
        by="collectDatetime", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')

    transf_sla['PROTOCOLO'] = transf['protocol']
    transf_sla['CLIENTE'] = transf['customerIDService.trading_firstname']
    transf_sla['CIDADE ORIGEM'] = transf['serviceIDRequested.source_address_id'].map(get_address_city)
    transf_sla['CIDADE DESTINO'] = transf['serviceIDRequested.destination_address_id'].map(get_address_city)
    transf_sla['TRANSPORTADORA'] = transf['BASE ORIGEM'].map(get_transp)
    transf_sla['RASTREADOR'] = transf['RASTREADOR']
    transf_sla['ROTA'] = transf['ROTA']
    # transf_sla['EMBARQUE PREVISTO'] = ''
    # transf_sla['EMBARQUE REAL'] = ''
    transf_sla['DIPONIBILIDADE PREVISTA'] = transf['previewDate'] + transf['previewHour']
    transf_sla['DIPONIBILIDADE REAL'] = transf['dataReal'] + transf['horaReal']
    transf_sla['ATRASO'] = transf['ATRASO']
    # transf_sla['CROSSDOCKING'] = transf['CROSSDOCKING']
    # transf_sla['IDENTIFICADOR'] = transf['Identificador']

    # transf_sla = pd.concat([
    #     transf_sla[transf_sla['TRANSPORTADORA'] == 'LATAM CARGO'],
    #     transf_sla[transf_sla['TRANSPORTADORA'] == 'GOLLOG'],
    #     transf_sla[transf_sla['TRANSPORTADORA'] == 'AZUL CARGO']
    # ], ignore_index=True)

    user_fb.set('Exportando planilhas...')

    CoInitialize()

    file_name = filename.get().replace('/', '\\')

    export_to_excel(transf_sla, file_name, 'Transfer', 'A1:L10000', autofit=False)

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


def start_update():
    if do_transfer.get():
        start_thread(sla_transferencia, progressbar)

    if do_last_mile.get():
        start_thread(sla_last_mile, progressbar)


def start_clear():
    if do_transfer.get():
        start_thread(clear_data, progressbar, arguments=["Transfer;A2:J5000"])

    if do_last_mile.get():
        start_thread(clear_data, progressbar, arguments=["Coletas;A2:I20000", "Entregas;A2:I20000"])

def resource_path(relative_path):
    try:
        import sys
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


r = RequestDataFrame()

states = pd.DataFrame([['ACRE', 'AC'],
                       ['ALAGOAS', 'AL'],
                       ['AMAPÁ', 'AP'],
                       ['AMAZONAS', 'AM'],
                       ['BAHIA', 'BA'],
                       ['CEARÁ', 'CE'],
                       ['DISTRITO FEDERAL', 'DF'],
                       ['ESPÍRITO SANTO', 'ES'],
                       ['GOIÁS', 'GO'],
                       ['MARANHÃO', 'MA'],
                       ['MATO GROSSO', 'MT'],
                       ['MATO GROSSO DO SUL', 'MS'],
                       ['MINAS GERAIS', 'MG'],
                       ['PARÁ', 'PA'],
                       ['PARAÍBA', 'PB'],
                       ['PARANÁ', 'PR'],
                       ['PERNAMBUCO', 'PE'],
                       ['PIAUÍ', 'PI'],
                       ['RIO DE JANEIRO', 'RJ'],
                       ['RIO GRANDE DO NORTE', 'RN'],
                       ['RIO GRANDE DO SUL', 'RS'],
                       ['RONDÔNIA', 'RO'],
                       ['RORAIMA', 'RR'],
                       ['SANTA CATARINA', 'SC'],
                       ['SÃO PAULO', 'SP'],
                       ['SERGIPE', 'SE'],
                       ['TOCANTINS', 'TO']], columns=['ESTADO', 'UF'])

# Create Object
root = Tk()
root.title('SLA LogLife')
try:
    root.iconbitmap(resource_path("assets/my_icon.ico"))
except:
    pass  # Fallback if icon is not found
root.resizable(False, False)
tabs = ttk.Notebook(root)

# Setting tabs

tab1 = ttk.Frame(tabs)
tab3 = ttk.Frame(tabs)

tabs.add(tab1, text='SLA')
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
    command=start_update).grid(
    column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(
    tab1,
    text="Limpar dados",
    command=start_clear).grid(
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

# Checkboxes

do_last_mile = BooleanVar(tab1, True)
do_transfer = BooleanVar(tab1, True)
checkbox_last_mile = ttk.Checkbutton(tab1,
                                     text='Last Mile',
                                     variable=do_last_mile,
                                     onvalue=True,
                                     offvalue=False,
                                     state='1')
checkbox_transfer = ttk.Checkbutton(tab1, text='Transferência', variable=do_transfer, onvalue=True, offvalue=False)
checkbox_last_mile.grid(column=0, row=2, padx=10, pady=4, sticky='W, E')
checkbox_transfer.grid(column=0, row=3, padx=10, pady=4, sticky='W, E')

ttk.Checkbutton()

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
