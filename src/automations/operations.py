import datetime as dt
import os
import time
import base64
import numpy as np
import pandas as pd
import win32com.client
from tkinter import StringVar, IntVar, CENTER, BOTTOM
import tkinter.ttk as ttk
import customtkinter as ctk
import sys
from pathlib import Path
# Add parent directory to path to enable absolute imports
sys.path.append(str(Path(__file__).parent.parent))
from utils.twocaptcha_solve import solve_captcha
from pythoncom import CoInitialize
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, \
    StaleElementReferenceException, SessionNotCreatedException, InvalidSessionIdException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
import tempfile
from utils.functions import *


# noinspection PyTypeChecker
def services_api(root, start_date, end_date, fb_frame, date_filter, reg_filter, regional, filename):
    def get_hub(hub_id):
        try:
            hub = hubs.loc[hubs['id'] == hub_id, 'name'].values.item()
        except Exception as e:
            hub = e
        return hub

    def get_hub_regional(hub_id):
        try:
            _regional = hubs.loc[hubs['id'] == hub_id, 'regional'].values.item()
        except Exception as e:
            _regional = e
        # except ValueError:
        #     regional = "VAZIO"
        return _regional

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
                cities += ' e ' + p
            else:
                cities += ', ' + p
            count += 1
        if add_lenght == 1:
            listing = "Material de "
        else:
            listing = "Materiais de "
        return listing + cities

    def get_address_city2(add):
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

    def excel_date(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days)

    def excel_time(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.seconds) / 86400

    def int_to_string(value):
        value_int = str(value)[:-2]
        return value_int

    def materials_description(mat_list):
        mat_description = ""
        count = 0
        description_list = ['Embalagem secundária', 'Gelox', 'Isopor 3L', 'Terciária 3L', 'Isopor 7L', 'Terciária 8L',
                            'Caixa térmica', 'Kg Gelo Seco']
        for material in mat_list:
            if material == 0:
                count += 1
                pass
            else:
                mat_description += f'{material} {description_list[count]}\n'
                count += 1
        mat_description = mat_description.strip("\n")
        return mat_description

    def define_failed_service(step, collect_list, collect_quantity, gelo_seco):
        status_list = []
        if step == 'COLETANDO':
            for coleta in collect_list:
                status = coleta.get('step', np.nan) if isinstance(coleta, dict) else np.nan
                status_list.append(status)

        if gelo_seco > 0:
            ajuste_gelo = 1
        else:
            ajuste_gelo = 0

        unique_status = list(np.unique(status_list))

        for i in unique_status[:]:
            if i == 'DONE' or i == 'UNSUCCESS':
                unique_status.remove(i)

        try:
            if len(unique_status) == 1 and len(collect_quantity) == (len(collect_list) - ajuste_gelo) and \
                    unique_status[0] == 'VALIDATEUNSUCCESS':
                return 'VALIDAR SEM SUCESSO'
            else:
                return step
        except TypeError:
            return step

    def get_occurence_time(dict_list, oc_list):

        if isinstance(dict_list, float) or len(dict_list) == 0:
            pass
            return None
        else:
            oc_creation_time = []
            oc_date_list = []
            oc_time_list = []

            for ocurrence in dict_list:

                oc_type = ocurrence.get('intercurrence', np.nan) if isinstance(ocurrence, dict) else np.nan

                if oc_type not in oc_list:
                    print(oc_type, 'Não considerado')
                    pass
                    return None

                else:
                    print(oc_type, 'Considerado')
                    oc_time = ocurrence.get('createdAt', np.nan) if isinstance(ocurrence, dict) else np.nan
                    oc_hour = ocurrence.get('occurrence_hour', np.nan) if isinstance(ocurrence, dict) else np.nan
                    oc_date = ocurrence.get('occurrence_date', np.nan) if isinstance(ocurrence, dict) else np.nan

                    oc_time = excel_date(pd.to_datetime(oc_time).tz_localize(None)) + \
                              excel_time(pd.to_datetime(oc_time).tz_localize(None))

                    oc_creation_time.append(oc_time)
                    try:
                        oc_date_list.append(oc_date)
                    except Exception as e:
                        print(e, e.message)
                        oc_date_list.append(0)
                        pass
                    oc_time_list.append(oc_hour)

            position = oc_creation_time.index(max(oc_creation_time))

            return [oc_date_list[position], oc_time_list[position]]

    def optimized_get_address(add):
        # Filter the DataFrame only once for all needed addresses
        address_filtered = address[address['id'].isin(add)]
        address_filtered = address_filtered.set_index('id')  # Set 'id' as index for faster access

        # Accumulate results in a list for better performance with string concatenation
        address_list = []
        for count, i in enumerate(add, start=1):
            try:
                entry = address_filtered.loc[i]
                # Using f-string for cleaner formatting
                listing = f"{count}. {entry['trading_name']} - {entry['branch']}: {entry['street'].title()}, " \
                          f"{entry['number']} - " \
                          f"{entry['neighborhood'].title()}, {entry['cityIDAddress.name'].title()} - " \
                          f"{entry['complement'].title()} | {entry['responsible_name'].title()}"
                address_list.append(listing)
            except KeyError:
                # Handle missing IDs gracefully
                address_list.append(f"{count}. Not Available")

        # Join all listings with newline separators
        add_total = "\n".join(address_list)
        return add_total

    def refined_get_branch(ids, memo=None):
        # Check and return from memo if already processed
        if memo is None:
            memo = {}
        ids_to_fetch = [id_t for id_t in ids if id_t not in memo]
        if ids_to_fetch:
            # Batch filter for new IDs and cache in memo
            bases_filtered = bases[bases['id'].isin(ids_to_fetch)].set_index('id')
            memo.update({id_t: bases_filtered.loc[id_t, 'nickname'] if id_t in bases_filtered.index else '' for id_t in
                         ids_to_fetch})
        return {id_t: memo.get(id_t, '') for id_t in ids}

    def refined_get_transp(ids, memo=None):
        # Check and return from memo if already processed
        if memo is None:
            memo = {}
        ids_to_fetch = [id_t for id_t in ids if id_t not in memo]
        if ids_to_fetch:
            # Batch filter for new IDs and cache in memo
            bases_filtered = bases[bases['id'].isin(ids_to_fetch)].set_index('id')
            memo.update({
                id_t: bases_filtered.loc[id_t, 'shippingIDBranch.company_name'] if id_t in bases_filtered.index else ""
                for id_t in ids_to_fetch if id_t is not None
            })
        return {id_t: memo.get(id_t, "") for id_t in ids}

    def refined_get_collector(col_ids, memo=None):
        # Check and return from memo if already processed
        if memo is None:
            memo = {}
        ids_to_fetch = [col_id for col_id in col_ids if col_id not in memo]
        if ids_to_fetch:
            # Batch filter for new IDs and cache in memo
            collector_filtered = collector[collector['id'].isin(ids_to_fetch)].set_index('id')
            memo.update({
                col_id: collector_filtered.loc[col_id, 'trading_name'] if col_id in collector_filtered.index else ""
                for col_id in ids_to_fetch
            })
        return {col_id: memo.get(col_id, "") for col_id in col_ids}

    def branch_wrapper(id_t):
        # Check if the ID is already in the memo
        if id_t in branch_memo:
            return branch_memo[id_t]
        else:
            # Fetch the result for this ID and store it in memo
            result = refined_get_branch([id_t], memo=branch_memo).get(id_t, '')
            branch_memo[id_t] = result  # Cache it
            return result

    def transp_wrapper(id_t):
        # Check if the ID is already in the memo
        if id_t in transp_memo:
            return transp_memo[id_t]
        else:
            # Fetch the result for this ID and store it in memo
            result = refined_get_transp([id_t], memo=transp_memo).get(id_t, '')
            transp_memo[id_t] = result  # Cache it
            return result

    def collector_wrapper(id_t):
        # Check if the ID is already in the memo
        if id_t in collector_memo:
            return collector_memo[id_t]
        else:
            # Fetch the result for this ID and store it in memo
            result = refined_get_collector([id_t], memo=collector_memo).get(id_t, '')
            collector_memo[id_t] = result  # Cache it
            return result

    branch_memo = {}
    transp_memo = {}
    collector_memo = {}

    di = start_date.get_date()
    df = end_date.get_date()

    di = dt.datetime.strftime(di, '%d/%m/%Y')
    df = dt.datetime.strftime(df, '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    # di_temp = di_dt - dt.timedelta(days=5)
    # df_temp = df_dt + dt.timedelta(days=5)

    # di = dt.datetime.strftime(di_temp, '%Y-%m-%d')
    # df = dt.datetime.strftime(df_temp, '%Y-%m-%d')

    materials = ['embalagem_secundaria',
                 'gelox',
                 'isopor3l',
                 'terciaria3l',
                 'isopor7l',
                 'terciaria8l',
                 'caixa_termica',
                 'gelo_seco']

    # Variable return

    user_fb = StringVar()
    feedback = ctk.CTkLabel(fb_frame, textvariable=user_fb)
    feedback.grid(column=0, row=3, padx=10, pady=10, columnspan=3, sticky='W, E')

    user_fb.set('Requisitando dados...')
    initial_time = time.time()

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    bases = r.request_public('https://transportebiologico.com.br/api/public/branch')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')
    # address = r.request_private('https://transportebiologico.com.br/api/address')
    # bases = r.request_private('https://transportebiologico.com.br/api/branch')
    # collector = r.request_private('https://transportebiologico.com.br/api/collector')
    # services = r.request_public('https://transportebiologico.com.br/api/public/service')
    services = r.request_private('https://transportebiologico.com.br/api/service/?is_business=false')
    hubs = r.request_private('https://transportebiologico.com.br/api/hub')

    # services = services.loc[services['is_business'] == False]
    services.drop(services[services['step'] == 'toValidateCancelRequest'].index, inplace=True)
    services.drop(services[services['step'] == 'toValidateRequestedService'].index, inplace=True)
    # services.drop(services[services['step'] == 'cancelledService'].index, inplace=True)
    # services.drop(services[services['step'] == 'unsuccessService'].index, inplace=True)

    final_time = time.time()

    execution_time = final_time - initial_time
    print(execution_time)

    user_fb.set('Organizando informações...')

    services['ETAPA1'] = np.select(
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
            services['step'] == 'pendingGeloSecoMaintenance',
            services['step'] == 'toMaintenanceService',
            services['step'] == 'toValidateFiscalRetention'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO',
            'AGUARDANDO ALOCAÇÃO',
            'EM ROTA DE ENTREGA',
            'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA',
            'DESEMBARCANDO',
            'VALIDAR EMBARQUE',
            'AGENDADO',
            'COLETANDO',
            'EM ROTA DE EMBARQUE',
            'EMBARCANDO SERVIÇO',
            'AGUARDANDO MANUTENÇÃO',
            'AGUARDANDO MANUTENÇÃO',
            'MATERIAL RETIDO'],
        default=0)
    'serviceIDCollect'

    services['ETAPA1'] = [
        define_failed_service(*arguments) for arguments in tuple(zip(services['ETAPA1'],
                                                                     services['serviceIDCollect'],
                                                                     services['serviceIDRequested.source_address_id'],
                                                                     services['serviceIDRequested.gelo_seco']))]

    services['ETAPA'] = np.select(
        condlist=[(services['ETAPA1'] == 'AGENDADO') & (services['serviceIDRequested.is_recurrent'] == True),
                  (services['ETAPA1'] == 'COLETANDO') & (services['serviceIDRequested.is_recurrent'] == True)],
        choicelist=['AGENDADO (R)', 'COLETANDO (R)'],
        default=services['ETAPA1']
    )

    services['CROSSDOCKING'] = np.where(services['serviceIDRequested.crossdocking_collector_id'].str.len() > 0, 1, 0)

    services['Identificador'] = np.select(
        condlist=[(services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] != 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] == 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 2)],
        choicelist=[0, 1, 1],
        default=0
    )

    services['transportadora1'] = services['serviceIDRequested.source_branch_id'].map(transp_wrapper)
    services['transportadora2'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(transp_wrapper)

    services['Base_origem'] = services['serviceIDRequested.source_branch_id'].map(branch_wrapper)
    services['cd_Base_destino'] = services['serviceIDRequested.destination_crossdocking_branch_id'].map(branch_wrapper)
    services['cd_Base_origem'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(branch_wrapper)
    services['Base_destino'] = services['serviceIDRequested.destination_branch_id'].map(branch_wrapper)

    services['BASE ORIGEM'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['Base_origem'],
                    services['cd_Base_origem'],
                    services['Base_origem'],
                    '-'],
        default=None)

    services['BASE DESTINO'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['cd_Base_destino'],
                    services['Base_destino'],
                    services['Base_destino'],
                    '-'],
        default=None)

    services['Transportadora'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0) & (services['serviceIDRequested.service_type'] == "FRACIONADO"),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['transportadora1'],
                    services['transportadora2'],
                    services['transportadora1'],
                    services['serviceIDRequested.service_type']],
        default=None)

    services['Coletador de Origem'] = services['serviceIDRequested.source_collector_id'].map(
        collector_wrapper)
    services['Coletador intermediário'] = services['serviceIDRequested.crossdocking_collector_id'].map(
        collector_wrapper)
    services['Coletador de destino'] = services['serviceIDRequested.destination_collector_id'].map(
        collector_wrapper)

    services['GELO SECO'] = services['serviceIDRequested.gelo_seco']

    mt = pd.DataFrame(columns=[])

    for mat in materials:
        mt[f'{mat}'] = services[f'serviceIDRequested.{mat}']

    sv_materials_list = []

    for rows in mt.itertuples():
        sv_mat_list = [rows.embalagem_secundaria,
                       rows.gelox,
                       rows.isopor3l,
                       rows.terciaria3l,
                       rows.isopor7l,
                       rows.terciaria8l,
                       rows.caixa_termica,
                       rows.gelo_seco]
        sv_mat_list = [int(i) for i in sv_mat_list]
        sv_materials_list.append(sv_mat_list)

    services = services.assign(materials_service=sv_materials_list)

    services['INSUMOS SERVIÇO'] = services['materials_service'].map(materials_description)

    services['oc_datahora_coleta'] = services['occurrenceIDService'].map(
        lambda x: get_occurence_time(x, ['ATRASO NA COLETA', 'NO SHOW'])
    )

    services['oc_data_coleta'] = services['oc_datahora_coleta'].map(lambda x: x[0] if isinstance(x, list) else np.nan)
    services['oc_hora_coleta'] = services['oc_datahora_coleta'].map(lambda x: x[1] if isinstance(x, list) else np.nan)

    # teste = pd.DataFrame()
    # teste['protocolo'] = services['protocol']
    # teste['oc_coleta'] = services['oc_hora_coleta']
    # teste.to_excel('teste.xlsx', index=False)

    services['oc_datahora_entrega'] = services['occurrenceIDService'].map(
        lambda x: get_occurence_time(x, ['CORTE DE VOO (NÃO ALOCADO VOO PLANEJADO)',
                                         'CANCELAMENTO DE VOO',
                                         'ATRASO NA ENTREGA'])
    )
    services['oc_data_entrega'] = services['oc_datahora_entrega'].map(lambda x: x[0] if isinstance(x, list) else np.nan)
    services['oc_hora_entrega'] = services['oc_datahora_entrega'].map(lambda x: x[1] if isinstance(x, list) else np.nan)

    services['dataColeta'] = np.where(services['oc_data_coleta'].isnull(),
                                      services['serviceIDRequested.collect_date'],
                                      services['oc_data_coleta'])

    services['dataColeta'] = pd.to_datetime(services['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    services['dataColeta'] = services['dataColeta'].dt.tz_localize(None)

    services['hourStart'] = np.where(services['oc_hora_coleta'].isnull(),
                                     services['serviceIDRequested.collect_hour_start'],
                                     services['oc_hora_coleta'])

    services['HORA INÍCIO'] = pd.to_datetime(services['hourStart']) - dt.timedelta(hours=3)
    services['HORA INÍCIO'] = services['HORA INÍCIO'].dt.strftime(date_format="%H:%M")

    services['hourEnd'] = np.where(services['oc_hora_coleta'].isnull(),
                                   services['serviceIDRequested.collect_hour_end'],
                                   services['oc_hora_coleta'])
    services['hourEnd'] = pd.to_datetime(services['hourEnd']) - dt.timedelta(hours=3)
    services['HORA FIM'] = services['hourEnd'].dt.strftime(date_format="%H:%M")
    services['hourEnd'] = services['hourEnd'].dt.tz_localize(None)

    # services['dataEmbarque'] =

    services['dataEntrega'] = np.where(services['oc_datahora_entrega'].isnull(),
                                       services['serviceIDRequested.delivery_date'],
                                       services['oc_data_entrega'])
    services['dataEntrega'] = pd.to_datetime(services['dataEntrega']) - dt.timedelta(hours=3)
    services['dataEntrega'] = services['dataEntrega'].dt.tz_localize(None)
    services['horaEntrega'] = np.where(services['oc_datahora_entrega'].isnull(),
                                       services['serviceIDRequested.delivery_hour'],
                                       services['oc_hora_entrega'])
    services['horaEntrega'] = pd.to_datetime(services['horaEntrega']) - dt.timedelta(hours=3)
    services['horaEntrega'] = services['horaEntrega'].dt.tz_localize(None)

    services['HORA ENTREGA'] = services['horaEntrega'].dt.strftime(date_format="%H:%M")
    services['DATA ENTREGA'] = services['dataEntrega'].dt.strftime(date_format="%d/%m/%Y")

    hoje = dt.datetime.today()

    services['LIMITE ENTREGA'] = np.where(
        services['dataEntrega'].dt.date == hoje.date(),
        'Entrega até ' + services['HORA ENTREGA'],
        'Entrega até ' + services['HORA ENTREGA'] +
        '\ndo dia ' + services['DATA ENTREGA'])

    services['boardDate'] = pd.to_datetime(services['serviceIDRequested.board_date']) - dt.timedelta(hours=3)
    services['boardDate'] = services['boardDate'].dt.tz_localize(None)
    services['boardHour'] = pd.to_datetime(services['serviceIDRequested.board_hour']) - dt.timedelta(hours=3)
    services['boardHour'] = services['boardHour'].dt.tz_localize(None)
    services['CDboardDate'] = pd.to_datetime(
        services['serviceIDRequested.crossdocking_board_date']
    ) - dt.timedelta(hours=3)
    services['CDboardHour'] = pd.to_datetime(
        services['serviceIDRequested.crossdocking_board_hour']
    ) - dt.timedelta(hours=3)
    services['CDboardDate'] = services['CDboardDate'].dt.tz_localize(None)
    services['CDboardHour'] = services['CDboardHour'].dt.tz_localize(None)

    services['DATA LIMITE DE EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['CDboardDate'],
                    services['boardDate'],
                    services['boardDate'],
                    services['boardDate']],
        default=services['boardDate'])

    services['HORÁRIO LIMITE DE EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['CDboardHour'],
                    services['boardHour'],
                    services['boardHour'],
                    services['boardHour']],
        default=services['boardHour'])

    services['HORÁRIO LIMITE DE EMBARQUE T'] = services['HORÁRIO LIMITE DE EMBARQUE'].dt.strftime(date_format="%H:%M")
    services['HORÁRIO LIMITE DE EMBARQUE'] = services['HORÁRIO LIMITE DE EMBARQUE'].dt.tz_localize(None)

    services['EMBARQUE/ENTREGA'] = np.select(
        condlist=[services['serviceIDRequested.service_type'] == 'DEDICADO',
                  services['serviceIDRequested.service_type'] == 'FRACIONADO'],
        choicelist=[services['LIMITE ENTREGA'],
                    'Embarque ' + services['Transportadora'] + '\n' + services['BASE ORIGEM'] + ' → ' +
                    services['BASE DESTINO'] + '\n até ' +
                    services['HORÁRIO LIMITE DE EMBARQUE T']],
        default=123)

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

    services['VOLUMES'] = np.where(services['tempo1'] >= services['tempo2'],
                                   services['VOLUMES1'],
                                   services['VOLUMES2'])
    services['VOLS'] = np.where(services['VOLUMES'] == 1, 'vol.', 'vols.')
    services['VOLUMES T'] = services['VOLUMES'].map(int_to_string) + ' ' + services['VOLS']
    services['CTE'] = np.where(services['tempo1'] >= services['tempo2'], services['CTE1'], services['CTE2'])
    services['RASTREADOR'] = np.where(
        services['tempo1'] >= services['tempo2'], services['RASTREADOR1'], services['RASTREADOR2'])

    services['COLETADOR DESTINO'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['Coletador intermediário'],
                    services['Coletador de destino'],
                    services['Coletador de destino'],
                    '-'],
        default=0)

    services['destinationHubId'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.budgetIDService.crossdocking_hub_id'],
                    services['serviceIDRequested.budgetIDService.destination_hub_id'],
                    services['serviceIDRequested.budgetIDService.destination_hub_id'],
                    '-'],
        default=None)

    services['COLETADOR EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0)],
        choicelist=[services['Coletador de Origem'],
                    services['Coletador intermediário'],
                    services['Coletador de Origem']],
        default=None
    )

    services['boardCollectorId'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0)],
        choicelist=[services['serviceIDRequested.source_collector_id'],
                    services['serviceIDRequested.crossdocking_collector_id'],
                    services['serviceIDRequested.source_collector_id']],
        default=None
    )

    services['boardHubId'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0)],
        choicelist=[services['serviceIDRequested.budgetIDService.source_hub_id'],
                    services['serviceIDRequested.budgetIDService.crossdocking_hub_id'],
                    services['serviceIDRequested.budgetIDService.source_hub_id']],
        default=None
    )

    services['HUB EMBARQUE'] = services['boardHubId'].map(get_hub)
    services['R.E.'] = services['boardHubId'].map(get_hub_regional)

    services['RETIRADA'] = 'Retirada de ' + services['VOLUMES T'] + ' na\n' + \
                           services['Transportadora'] + ' ' + services['BASE DESTINO'] + '\nCTe: ' + \
                           services['CTE'].astype(str) + '\nRastreador: ' + services['RASTREADOR'].astype(str)
    services['HUB ORIGEM'] = services['serviceIDRequested.budgetIDService.source_hub_id'].map(get_hub)
    services['R.O.'] = services['serviceIDRequested.budgetIDService.source_hub_id'].map(get_hub_regional)
    services.to_excel(f'{temp_folder}/Excel/SericesAPI2.xlsx')
    services['HUB DESTINO'] = services['destinationHubId'].map(get_hub)
    services['R.D.'] = services['destinationHubId'].map(get_hub_regional)
    services['REMETENTES'] = services['serviceIDRequested.source_address_id'].map(optimized_get_address)
    services['DESTINATÁRIOS'] = services['serviceIDRequested.destination_address_id'].map(optimized_get_address)
    services['DATA DA COLETA'] = services['dataColeta'].map(excel_date)
    services['HORÁRIO DA COLETA'] = services['HORA INÍCIO'] + '\nàs ' + services['HORA FIM']
    services['HORA FIM'] = services['hourEnd'].map(excel_time)
    services['DATA LIMITE DE EMBARQUE'] = services['DATA LIMITE DE EMBARQUE'].map(excel_date)
    services['HORÁRIO LIMITE DO EMBARQUE'] = services['HORÁRIO LIMITE DE EMBARQUE'].map(excel_time)
    services['DATA LIMITE DE ENTREGA'] = services['dataEntrega'].map(excel_date)
    services['HORÁRIO LIMITE DE ENTREGA'] = services['horaEntrega'].map(excel_time)

    services['CIDADES ORIGEM'] = services['serviceIDRequested.source_address_id'].map(get_address_city2)
    services['CIDADES DESTINO'] = services['serviceIDRequested.destination_address_id'].map(get_address_city2)

    services['REMETENTES ENTREGA'] = np.select(
        condlist=[services['serviceIDRequested.service_type'] == 'DEDICADO',
                  services['serviceIDRequested.service_type'] == 'FRACIONADO'],
        choicelist=[services['serviceIDRequested.source_address_id'].map(optimized_get_address),
                    services['serviceIDRequested.source_address_id'].map(get_address_city)])
    services['RETIRADA/ENTREGA'] = np.select(
        condlist=[services['serviceIDRequested.service_type'] == 'DEDICADO',
                  services['serviceIDRequested.service_type'] == 'FRACIONADO'],
        choicelist=['SERVIÇO DEDICADO',
                    services['RETIRADA']])

    df_dt += dt.timedelta(days=1)  # Correção para poder buscar somente um dia

    user_fb.set('Atualizando dados...')

    coletas = services[[
        'ETAPA',
        'HUB ORIGEM',
        'Coletador de Origem',
        'protocol',
        'customerIDService.trading_firstname',
        'serviceIDRequested.vehicle',
        'REMETENTES',
        'DATA DA COLETA',
        'HORÁRIO DA COLETA',
        'EMBARQUE/ENTREGA',
        'INSUMOS SERVIÇO',
        'DESTINATÁRIOS',
        'serviceIDRequested.observation',
        'cte_loglife',
        'Coletador de destino',
        'HORA FIM',
        'dataColeta',
        'R.O.'
    ]].copy()

    if date_filter == 1:
        coletas.drop(coletas[coletas['dataColeta'] < di_dt].index, inplace=True)
        coletas.drop(coletas[coletas['dataColeta'] > df_dt].index, inplace=True)

    services.to_excel(f'{temp_folder}/Excel/ServicesAPI.xlsx', index=False)

    embarques = services[[
        'ETAPA',
        'HUB EMBARQUE',
        'COLETADOR EMBARQUE',
        'protocol',
        'customerIDService.trading_firstname',
        'REMETENTES',
        'DATA LIMITE DE EMBARQUE',
        'HORÁRIO LIMITE DO EMBARQUE',
        'EMBARQUE/ENTREGA',
        'DESTINATÁRIOS',
        'serviceIDRequested.observation',
        'cte_loglife',
        'Coletador de destino',
        'R.E.'
    ]].copy()

    cargas = services[[
        'protocol',
        'ETAPA',
        'COLETADOR DESTINO',
        'customerIDService.trading_firstname',
        'CIDADES ORIGEM',
        'CIDADES DESTINO',
        'RASTREADOR',
        'CTE',
        'serviceIDRequested.vehicle',
        'VOLUMES',
        'Transportadora',
        'BASE ORIGEM',
        'BASE DESTINO',
        'DATA LIMITE DE ENTREGA',
        'HORÁRIO LIMITE DE ENTREGA',
        'GELO SECO',
        'serviceIDRequested.service_type',
        'CROSSDOCKING',
        'R.D.',
        'dataEntrega'
    ]].copy()

    if date_filter == 1:
        cargas.drop(cargas[cargas['dataEntrega'] > df_dt].index, inplace=True)
    cargas.drop(columns=['dataEntrega'], inplace=True)

    entregas = services[[
        'HUB DESTINO',
        'COLETADOR DESTINO',
        'protocol',
        'customerIDService.trading_firstname',
        'serviceIDRequested.vehicle',
        'REMETENTES ENTREGA',
        'RETIRADA/ENTREGA',
        'DATA LIMITE DE ENTREGA',
        'HORÁRIO LIMITE DE ENTREGA',
        'DESTINATÁRIOS',
        'ETAPA',
        'R.D.'
    ]].copy()

    coletas.sort_values(by="HORA FIM", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')
    coletas.sort_values(
        by="DATA DA COLETA", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

    embarques.sort_values(
        by="HORÁRIO LIMITE DO EMBARQUE", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')
    embarques.sort_values(
        by="DATA LIMITE DE EMBARQUE", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

    cargas.sort_values(
        by="HORÁRIO LIMITE DE ENTREGA", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')
    cargas.sort_values(
        by="DATA LIMITE DE ENTREGA", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

    entregas.sort_values(
        by="HORÁRIO LIMITE DE ENTREGA", axis=0, ascending=True, inplace=True, kind='quicksort', na_position='last')
    entregas.sort_values(
        by="DATA LIMITE DE ENTREGA", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

    coletas.drop(coletas[coletas['ETAPA'] == 'EM ROTA DE EMBARQUE'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'EMBARCANDO SERVIÇO'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'VALIDAR EMBARQUE'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'AGUARDANDO ALOCAÇÃO'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'AGUARDANDO DISPONIBILIZAÇÃO'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'MATERIAL RETIDO'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'DISPONÍVEL PARA RETIRADA'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'DESEMBARCANDO'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'AGUARDANDO MANUTENÇÃO'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'EM ROTA DE ENTREGA'].index, inplace=True)
    coletas.drop(coletas[coletas['ETAPA'] == 'ENTREGANDO'].index, inplace=True)
    coletas.drop(columns=['HORA FIM', 'dataColeta'], inplace=True)

    embarques.drop(embarques[embarques['ETAPA'] == 'AGENDADO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'AGENDADO (R)'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'COLETANDO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'COLETANDO (R)'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'VALIDAR SEM SUCESSO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'VALIDAR EMBARQUE'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'AGUARDANDO ALOCAÇÃO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'MATERIAL RETIDO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'AGUARDANDO DISPONIBILIZAÇÃO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'DISPONÍVEL PARA RETIRADA'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'DESEMBARCANDO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'AGUARDANDO MANUTENÇÃO'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'EM ROTA DE ENTREGA'].index, inplace=True)
    embarques.drop(embarques[embarques['ETAPA'] == 'ENTREGANDO'].index, inplace=True)

    # cargas.drop(cargas[cargas['ETAPA'] == 'AGENDADO'].index, inplace=True)
    # cargas.drop(cargas[cargas['ETAPA'] == 'COLETANDO'].index, inplace=True)
    # cargas.drop(cargas[cargas['ETAPA'] == 'VALIDAR SEM SUCESSO'].index, inplace=True)
    # cargas.drop(cargas[cargas['ETAPA'] == 'EM ROTA DE EMBARQUE'].index, inplace=True)
    # cargas.drop(cargas[cargas['ETAPA'] == 'EMBARCANDO SERVIÇO'].index, inplace=True)

    entregas.drop(entregas[entregas['ETAPA'] == 'AGENDADO'].index, inplace=True)
    entregas.drop(entregas[entregas['ETAPA'] == 'AGENDADO (R)'].index, inplace=True)
    entregas.drop(entregas[entregas['ETAPA'] == 'COLETANDO'].index, inplace=True)
    entregas.drop(entregas[entregas['ETAPA'] == 'COLETANDO (R)'].index, inplace=True)
    entregas.drop(entregas[entregas['ETAPA'] == 'VALIDAR SEM SUCESSO'].index, inplace=True)
    entregas.drop(entregas[entregas['ETAPA'] == 'EM ROTA DE EMBARQUE'].index, inplace=True)
    entregas.drop(entregas[entregas['ETAPA'] == 'EMBARCANDO SERVIÇO'].index, inplace=True)

    regional = int(regional)

    if reg_filter == 1:
        coletas = coletas.loc[coletas['R.O.'] == regional]
        # coletas.drop(columns=['R.O.'], inplace=True)
        embarques = embarques.loc[embarques['R.E.'] == regional]
        # embarques.drop(columns=['REGIONAL ORIGEM'], inplace=True)
        cargas = cargas.loc[cargas['R.D.'] == regional]
        # cargas.drop(columns=['REGIONAL DESTINO'], inplace=True)
        entregas = entregas.loc[entregas['R.D.'] == regional]
        # entregas.drop(columns=['R.D.'], inplace=True)

    user_fb.set('Exportando planilhas...')

    file_name = filename.replace('/', '\\')

    CoInitialize()

    export_to_excel(coletas, file_name, 'COLETAS', 'A2:P1000', change_header=False)
    export_to_excel(embarques, file_name, 'EMBARQUES', 'A2:P1000', change_header=False)
    export_to_excel(entregas, file_name, 'ENTREGAS', 'A2:L1000', change_header=False)
    export_to_excel(cargas, file_name, 'CARGAS', 'A2:S1000', autofit=False, change_header=False)

    root.after(10, feedback.destroy())


# noinspection PyTypeChecker
def minutas_api(input_prot, multiple=False, prot_strings="", flight_service=0, material_type=0, vols="1",
                kg_record="1", folderpath="", downloadpath="", prot_entry=None, driver_fox=None):

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

    def save_to_pdf(workbook, application, sheet_index, print_area_set, dispatch_prot):

        workbook.save()
        application.kill()

        excel_ = win32com.client.Dispatch("Excel.Application")
        excel_.Interactive = False
        excel_.ScreenUpdating = False
        excel_.DisplayAlerts = False
        excel_.EnableEvents = False
        workbook1 = excel_.Workbooks.Open(os.path.abspath("MINUTAS.xlsx"))

        worksheet_index_list = [sheet_index]
        folder_path = folderpath.replace("/", "\\")

        path_to_pdf_save = os.path.abspath(
            f"{folder_path}\\{identificator}{transportadora} - {get_branch(orig)} → {get_branch(dest)}.pdf"
        )

        workbook1.Worksheets(worksheet_index_list).Select()
        workbook1.ActiveSheet.PageSetup.BlackAndWhite = False
        workbook1.ActiveSheet.PageSetup.PrintArea = print_area_set
        workbook1.ActiveSheet.PageSetup.TopMargin = 10
        workbook1.ActiveSheet.PageSetup.FitToPagesTall = 1
        workbook1.ActiveSheet.PageSetup.FitToPagesWide = 1
        workbook1.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf_save)

        workbook1.Close(False)
        del workbook1
        excel_.Interactive = True
        excel_.ScreenUpdating = True
        excel_.DisplayAlerts = True
        excel_.EnableEvents = True
        del excel_
        dispatch_prot.set("")

    if driver_fox is None:
        firefox = Service(
            GeckoDriverManager().install(),
            log_path=f"{tempfile.gettempdir()}/geckodriver.log"
        )
    else:
        firefox = driver_fox

    CoInitialize()

    protocolo = int(input_prot)

    if multiple is True:
        identificator = f"{prot_strings} - "
    else:
        identificator = f"{protocolo} - "

    options = Options()
    options.set_preference("pdfjs.migrationVersion", 2)
    options.set_preference("pdfjs.enabledCache.state", False)

    filename_minutas = "MINUTAS.xlsx"

    services = r.request_public('https://transportebiologico.com.br/api/public/service')
    bases = r.request_public('https://transportebiologico.com.br/api/public/branch')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')

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
            services['step'] == 'pendingGeloSecoMaintenance'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO',
            'AGUARDANDO ALOCAÇÃO',
            'EM ROTA DE ENTREGA',
            'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA',
            'DESEMBARCANDO',
            'VALIDAR EMBARQUE',
            'AGENDADO',
            'COLETANDO',
            'EM ROTA DE EMBARQUE',
            'EMBARCANDO SERVIÇO',
            'AGUARDANDO MANUTENÇÃO'],
        default=0)

    services['CROSSDOCKING'] = np.where(services['serviceIDRequested.crossdocking_collector_id'].str.len() > 0, 1, 0)

    services['Identificador'] = np.select(
        condlist=[(services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] != 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] == 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 2)],
        choicelist=[0, 1, 1],
        default=0
    )

    services['transportadora1'] = services['serviceIDRequested.source_branch_id'].map(get_transp)
    services['transportadora2'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_transp)

    # services['Coletador de Origem'] = services['serviceIDRequested.source_collector_id'].map(get_collector)
    # services['Coletador intermediário'] = services['serviceIDRequested.crossdocking_collector_id'].map(get_collector)
    # services['Coletador de destino'] = services['serviceIDRequested.destination_collector_id'].map(get_collector)

    services['COLETADOR DESTINO'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.crossdocking_collector_id'],
                    services['serviceIDRequested.destination_collector_id'],
                    services['serviceIDRequested.destination_collector_id'],
                    '-'],
        default=0)

    services['COLETADOR EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0)],
        choicelist=[services['serviceIDRequested.source_collector_id'],
                    services['serviceIDRequested.crossdocking_collector_id'],
                    services['serviceIDRequested.source_collector_id']],
        default=0
    )

    orig_collector_id = services.loc[
        services['protocol'] == protocolo, 'COLETADOR EMBARQUE'].values.item()

    dest_collector_id = services.loc[
        services['protocol'] == protocolo, 'COLETADOR DESTINO'].values.item()

    # orig_id = services.loc[services['protocol'] == protocolo, 'serviceIDRequested.source_branch_id'].values.item()
    # dest_id = services.loc[
    #     services['protocol'] == protocolo, 'serviceIDRequested.destination_branch_id'].values.item()

    name_orig = collector.loc[collector['id'] == orig_collector_id, 'trading_name'].values.item()
    name_dest = collector.loc[collector['id'] == dest_collector_id, 'trading_name'].values.item()
    # comp_name_orig = collector.loc[collector['id'] == orig_collector_id, 'company_name'].values.item()
    comp_name_dest = collector.loc[collector['id'] == dest_collector_id, 'company_name'].values.item()
    cnpj_orig = collector.loc[collector['id'] == orig_collector_id, 'cnpj'].values.item()
    cnpj_dest = collector.loc[collector['id'] == dest_collector_id, 'cnpj'].values.item()
    city_orig = collector.loc[collector['id'] == orig_collector_id, 'city'].values.item()
    city_dest = collector.loc[collector['id'] == dest_collector_id, 'city'].values.item()
    street_orig = collector.loc[collector['id'] == orig_collector_id, 'street'].values.item()
    street_dest = collector.loc[collector['id'] == dest_collector_id, 'street'].values.item()
    number_orig = collector.loc[collector['id'] == orig_collector_id, 'number'].values.item()
    number_dest = collector.loc[collector['id'] == dest_collector_id, 'number'].values.item()
    cep_orig = collector.loc[collector['id'] == orig_collector_id, 'cep'].values.item()
    cep_dest = collector.loc[collector['id'] == dest_collector_id, 'cep'].values.item()
    # distr_orig = collector.loc[collector['id'] == orig_collector_id, 'neighborhood'].values.item()
    distr_dest = collector.loc[collector['id'] == dest_collector_id, 'neighborhood'].values.item()
    complement_dest = collector.loc[collector['id'] == dest_collector_id, 'complement'].values.item()
    state_register_dest = collector.loc[collector['id'] == dest_collector_id, 'municipal_register'].values.item()

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
            services['step'] == 'boardingService'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE', 'AGENDADO', 'COLETANDO',
            'EM ROTA DE EMBARQUE', 'EMBARCANDO SERVIÇO'],
        default=0)

    service_type = services.loc[services['protocol'] == protocolo, 'serviceIDRequested.service_type'].values.item()

    services['BASE ORIGEM'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.source_branch_id'],
                    services['serviceIDRequested.source_crossdocking_branch_id'],
                    services['serviceIDRequested.source_branch_id'],
                    '-'],
        default=None)

    services['BASE DESTINO'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.destination_crossdocking_branch_id'],
                    services['serviceIDRequested.destination_branch_id'],
                    services['serviceIDRequested.destination_branch_id'],
                    '-'],
        default=None)

    services['Transportadora'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0) & (services['serviceIDRequested.service_type'] == "FRACIONADO"),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['transportadora1'],
                    services['transportadora2'],
                    services['transportadora1'],
                    services['serviceIDRequested.service_type']],
        default=None)

    if service_type == "DEDICADO" and city_orig == city_dest:
        board_client = True
        transportadora = "LATAM CARGO"
        email = services.loc[services['protocol'] == protocolo, 'customerIDService.operational_email'].values.item()
        orig = np.select(condlist=[(city_orig == 'MACEIO') | (city_orig == 'MACEIÓ'),
                                   (city_orig == 'BELEM') | (city_orig == 'BELÉM'),
                                   (city_orig == 'FORTALEZA')],
                         choicelist=['MCZ', 'BEL', 'FOR'],
                         default=0)
        payer = '//*[@id="rcpPayer1"]'
        dest = "GRU"
    else:
        board_client = False
        transportadora = services.loc[services['protocol'] == protocolo, 'Transportadora'].values.item()
        email = "operacional@loglifelogistica.com.br"
        orig = services.loc[services['protocol'] == protocolo, 'BASE ORIGEM'].values.item()
        dest = services.loc[services['protocol'] == protocolo, 'BASE DESTINO'].values.item()
        payer = '//*[@id="sndPayer1"]'

    if transportadora == 'SOL CARGAS':

        sheet = 'SOL'

        app = xw.App(visible=False)
        wb = xw.Book(filename_minutas)
        ws = wb.sheets[sheet]

        ws['H1'].value = bases.loc[bases['id'] == orig, 'hubIDBranch.name'].values.item()
        ws['H2'].value = bases.loc[bases['id'] == dest, 'hubIDBranch.name'].values.item()
        ws['A4'].value = f'REMETENTE: {name_orig}'
        ws['E4'].value = f'DESTINATÁRIO: {name_dest}'
        ws['A5'].value = f'CNPJ: {cnpj_orig}'
        ws['E5'].value = f'CNPJ: {cnpj_dest}'
        ws['A6'].value = f'CIDADE: {city_orig}'
        ws['E6'].value = f'CIDADE: {city_dest}'

        save_to_pdf(wb, app, 5, "A1:I23", prot_entry)

    elif transportadora == 'JEM':

        sheet = 'JEM'

        app = xw.App(visible=False)
        wb = xw.Book(filename_minutas)
        ws = wb.sheets[sheet]
        uf = wb.sheets['UF']
        uf['C1'].value = collector.loc[collector['id'] == orig_collector_id, 'state'].values.item()
        uf['C2'].value = collector.loc[collector['id'] == dest_collector_id, 'state'].values.item()

        ws['I6'].value = bases.loc[bases['id'] == orig, 'hub'].values.item()
        ws['AU6'].value = bases.loc[bases['id'] == dest, 'hub'].values.item()
        ws['I9'].value = name_orig
        ws['K15'].value = name_dest
        ws['I10'].value = cnpj_orig
        ws['I16'].value = cnpj_dest
        ws['G12'].value = city_orig
        ws['G18'].value = city_dest
        ws['I11'].value = street_orig
        ws['I17'].value = street_dest
        ws['BC11'].value = number_orig
        ws['BC17'].value = number_dest
        ws['BM12'].value = cep_orig
        ws['BM18'].value = cep_dest

        save_to_pdf(wb, app, 4, "A1:CA46", prot_entry)

    elif transportadora == 'AZUL CARGO':

        sheet = 'AZUL'

        app = xw.App(visible=False)
        wb = xw.Book(filename_minutas)
        ws = wb.sheets[sheet]

        uf = wb.sheets['UF']
        uf['C3'].value = collector.loc[collector['id'] == orig_collector_id, 'state'].values.item()
        uf['C4'].value = collector.loc[collector['id'] == dest_collector_id, 'state'].values.item()

        ws['F12'].value = f'SIM {get_branch(dest)}'
        ws['A15'].value = "SURE LOGÍSTICA EIRELI"
        ws['M15'].value = name_dest
        ws['A16'].value = 'CNPJ/CPF: 17.062.517/0001-08'
        ws['M16'].value = f'CNPJ/CPF: {cnpj_dest}'
        ws['G20'].value = 'Cidade: BELO HORIZONTE'
        ws['S20'].value = f'Cidade: {city_dest}'
        ws['A18'].value = 'RUA/AV: AV. RAJA GABÁGLIA'
        ws['M18'].value = f'RUA/AV: {street_dest}'
        ws['L18'].value = 'Nº: 4055'
        ws['X18'].value = f'Nº: {number_dest}'
        ws['J21'].value = 'CEP: 30380-103'
        ws['V21'].value = f'CEP: {cep_dest}'
        ws['F19'].value = 'BAIRRO: CIDADE JARDIM'
        ws['R19'].value = f'BAIRRO: {distr_dest}'

        save_to_pdf(wb, app, 1, "A1:X47", prot_entry)

    elif transportadora == 'BUSLOG':

        sheet = 'BUSLOG'

        app = xw.App(visible=False)
        wb = xw.Book(filename_minutas)
        ws = wb.sheets[sheet]

        uf = wb.sheets['UF']
        uf['C5'].value = collector.loc[collector['id'] == orig_collector_id, 'state'].values.item()
        uf['C6'].value = collector.loc[collector['id'] == dest_collector_id, 'state'].values.item()

        ws['C11'].value = f'Remetente: SURE LOGÍSTICA LTDA'
        ws['C18'].value = f'Destinatário: {name_dest}'
        ws['C12'].value = f'CNPJ: 17.062.517/0001-08'
        ws['C19'].value = f'CNPJ: {cnpj_dest}'
        ws['C13'].value = f'Endereço: AVENIDA RAJA GABAGLIA, 4055 - BELO HORIZONTE'
        ws['C20'].value = f'Endereço: {street_dest}, {number_dest} - {city_dest}'
        ws['I14'].value = f'CEP: 30380-103'
        ws['I21'].value = f'CEP: {cep_dest}'
        ws['C14'].value = f'Bairro: SANTA LÚCIA'
        ws['C21'].value = f'Bairro: {distr_dest}'

        save_to_pdf(wb, app, 3, "A1:N44", prot_entry)

    elif transportadora == 'GOLLOG':

        sheet = 'GOLLOG'

        app = xw.App(visible=False)
        wb = xw.Book(filename_minutas)
        ws = wb.sheets[sheet]

        uf = wb.sheets['UF']
        uf['C7'].value = collector.loc[collector['id'] == dest_collector_id, 'state'].values.item()

        ws['Q10'].value = f'{get_branch(dest)}'
        ws['M12'].value = f'{comp_name_dest} - {name_dest}'
        ws['M13'].value = f'CNPJ/CPF: {cnpj_dest}'
        ws['M14'].value = f'ENDEREÇO: {street_dest}'
        ws['R14'].value = f'Nº {number_dest}'
        ws['M15'].value = f'COMPLEMENTO: {complement_dest}'
        ws['O15'].value = f'BAIRRO: {distr_dest}'
        ws['M16'].value = f'CIDADE: {city_dest}'
        ws['Q16'].value = f'CEP: {cep_dest}'

        save_to_pdf(wb, app, 2, "B2:S30", prot_entry)

    elif transportadora == 'LATAM CARGO':

        sender = "17062517000108"

        folder_path1 = folderpath.replace("/", "\\")
        download_path = downloadpath.replace("/", "\\")

        try:
            os.remove(f'{download_path}\\Minuta.pdf')
        except FileNotFoundError:
            pass

        orig_base = f'{get_branch(orig)}-'
        if orig_base == "POA-":
            orig_base = "POA-S"
        elif orig_base == "IOS-":
            orig_base = "IOS-E"
        dest_base = f'{get_branch(dest)}-'
        if dest_base == "POA-":
            dest_base = "POA-S"
        elif dest_base == "IOS-":
            dest_base = "IOS-E"

        if services.loc[
            services['protocol'] == protocolo, 'customerIDService.trading_firstname'].values.item() == 'DLE' and \
                dest in ['SDU', 'GIG']:
            recipient = services.loc[services['protocol'] == protocolo, 'customerIDService.cnpj_cpf'].values.item()
        elif board_client:
            recipient = services.loc[services['protocol'] == protocolo, 'customerIDService.cnpj_cpf'].values.item()
        else:
            recipient = collector.loc[collector['id'] == dest_collector_id, 'cnpj'].values.item()

        if flight_service == 0:
            serv = "VLZ"
        else:
            serv = "STD"

        if material_type == 0:
            tipo = "0786"  # material biológico
        elif material_type == 1:
            tipo = "0001"  # carga geral
        elif material_type == 2:
            tipo = "5600"  # material infectante
        else:
            tipo = "0797"  # medicamentos

        n = webdriver.Firefox(service=firefox, options=options)

        n.set_window_size(1024, 768)

        # if multiple is True:
        #     n.minimize_window()

        # Open mozilla and access site
        while True:
            try:
                # n.get("https://minutaweb.lancargo.com/MinutaWEB-3.0/public/client.jsf")
                n.get('https://www.latamcargo.com/en/eminutaclient')
                break
            except SessionNotCreatedException:
                n.quit()
                continue
            except InvalidSessionIdException:
                n.close()
                n.quit()
                continue

        # Input sender CNPJ (autofill the rest)
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="cpfcnpjsender"]').click()
                n.find_element(By.XPATH, '//*[@id="cpfcnpjsender"]').send_keys(sender + Keys.TAB)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Input recipient CNPJ
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="rcpcpfcnpjsender"]').click()
                n.find_element(By.XPATH, '//*[@id="rcpcpfcnpjsender"]').send_keys(recipient + Keys.TAB)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Input name + IE + CEP (recipient)
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="rcpinputnome"]').click()
                n.find_element(By.XPATH, '//*[@id="rcpinputnome"]').send_keys(Keys.CONTROL + "A")
                n.find_element(By.XPATH, '//*[@id="rcpinputnome"]').send_keys(comp_name_dest)
                n.find_element(By.XPATH, '//*[@id="rcpinputie"]').click()
                n.find_element(By.XPATH, '//*[@id="rcpinputie"]').send_keys(Keys.CONTROL + "A")
                n.find_element(By.XPATH, '//*[@id="rcpinputie"]').send_keys(state_register_dest)
                n.find_element(By.XPATH, '//*[@id="rcpinputcep"]').click()
                n.find_element(By.XPATH, '//*[@id="rcpinputcep"]').send_keys(Keys.CONTROL + "A")
                n.find_element(By.XPATH, '//*[@id="rcpinputcep"]').send_keys(cep_dest)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Input address street + number
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="rcp_inputendereco"]').click()
                n.find_element(By.XPATH, '//*[@id="rcp_inputendereco"]').send_keys(Keys.CONTROL + "A")
                n.find_element(By.XPATH, '//*[@id="rcp_inputendereco"]').send_keys(street_dest)
                n.find_element(By.XPATH, '//*[@id="rcp_inputnumero"]').click()
                n.find_element(By.XPATH, '//*[@id="rcp_inputnumero"]').send_keys(Keys.CONTROL + "A")
                n.find_element(By.XPATH, '//*[@id="rcp_inputnumero"]').send_keys(number_dest)
                n.find_element(By.XPATH, '//*[@id="rcp_inputbairro"]').send_keys(distr_dest)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Click 'Pay with billing record?'
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="flexCheckChecked1"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Define payer
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, payer).click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Input e-mail
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="inputemail"]').click()
                n.find_element(By.XPATH, '//*[@id="inputemail"]').send_keys(Keys.CONTROL + "A")
                n.find_element(By.XPATH, '//*[@id="inputemail"]').send_keys(Keys.BACKSPACE)
                n.find_element(By.XPATH, '//*[@id="inputemail"]').send_keys(email)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.1)
                continue

        # Input origin airport
        while True:
            try:
                n.find_element(By.XPATH, '//*[@id="select2-origin-select-container"]').click()
                n.find_element(By.XPATH, '/html/body/span/span/span[1]/'
                                         'input').send_keys(orig_base + Keys.ENTER)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.1)
                continue

        # Input destination airport
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="select2-dest-select-container"]').click()
                n.find_element(By.XPATH, '/html/body/span/span/span[1]/'
                                         'input').send_keys(dest_base + Keys.ENTER)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Finish first page and advance
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="avanzar"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Input quantity, weight and volume
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH,
                               '/html/body/div[3]/div/div[2]/div/div/div[3]/div/table/tbody/tr/td/div[1]/div[2]/div[2]/'
                               'input').send_keys(Keys.BACKSPACE + Keys.DELETE + str(vols) + Keys.TAB)
                n.find_element(By.XPATH,
                               '/html/body/div[3]/div/div[2]/div/div/div[3]/div/table/tbody/tr/td/div[1]/div[3]/div[2]/'
                               'input').send_keys(str(kg_record) + Keys.TAB)
                n.find_element(By.XPATH,
                               '/html/body/div[3]/div/div[2]/div/div/div[3]/div/table/tbody/tr/td/div[1]/div[4]/div[2]/'
                               'input').send_keys(Keys.BACKSPACE + Keys.DELETE + '10' + Keys.TAB)
                n.find_element(By.XPATH,
                               '/html/body/div[3]/div/div[2]/div/div/div[3]/div/table/tbody/tr/td/div[1]/div[5]/div[2]/'
                               'input').send_keys(Keys.BACKSPACE + Keys.DELETE + '10' + Keys.TAB)
                n.find_element(By.XPATH,
                               '/html/body/div[3]/div/div[2]/div/div/div[3]/div/table/tbody/tr/td/div[1]/div[6]/div[2]/'
                               'input').send_keys(Keys.BACKSPACE + Keys.DELETE + '10')
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Packaging information
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, "//*[@title='SELECCIONE...']").click()
                time.sleep(0.1)
                n.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('PAP' + Keys.ENTER)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Material Type
        while True:
            try:
                n.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td/div/div[2]/'
                                         'div[2]/span/span[1]/span').click()
                time.sleep(0.5)
                n.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(tipo + Keys.ENTER)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Service
        while True:
            try:

                time.sleep(0.5)
                n.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td/div/div[3]/'
                                         'div[2]/span/span[1]/span').click()
                time.sleep(0.5)
                n.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(serv + Keys.ENTER)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Country
        while True:
            try:
                time.sleep(0.5)
                n.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td/div/div[4]/'
                                         'div[2]/span/span[1]/span').click()
                time.sleep(0.5)
                n.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys('B' + Keys.ENTER)
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Finish second page and advance3f5
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="avanzar"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Click 'UNSECURED'
        while True:
            try:
                time.sleep(0.1)
                n.find_element(By.XPATH, '//*[@id="unSecured"]')
                n.find_element(By.XPATH, '//*[@id="unSecured"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Click 'WITHOUT DOCUMENT'
        while True:
            try:
                time.sleep(0.1)
                n.find_element(
                    By.XPATH, '//*[@id="noDocument"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Accept terms and click on captcha input box
        while True:
            try:
                time.sleep(0.1)
                n.find_element(
                    By.XPATH,
                    '//*[@id="flexCheckChecked"]').click()
                n.find_element(By.XPATH, '//*[@id="txtTwoWord"]').click()
                # get the image source.
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # get captcha image and solve it with 2captcha
        while True:
            try:
                time.sleep(0.5)
                ele_captcha = n.find_element(By.XPATH, '//*[@id="captchaImage"]')
                # get the captcha as a base64 string
                img_captcha_base64 = n.execute_async_script("""
                    var ele = arguments[0], callback = arguments[1];
                    ele.addEventListener('load', function fn(){
                      ele.removeEventListener('load', fn, false);
                      var cnv = document.createElement('canvas');
                      cnv.width = this.width; cnv.height = this.height;
                      cnv.getContext('2d').drawImage(this, 0, 0);
                      callback(cnv.toDataURL('image/jpeg').substring(22));
                    }, false);
                    ele.dispatchEvent(new Event('load'));
                    """, ele_captcha)

                # save the captcha to a file
                with open(f"{captcha_image}", 'wb') as f:
                    f.write(base64.b64decode(img_captcha_base64))

                ai_captcha = solve_captcha(captcha_image)
                n.find_element(By.XPATH, '//*[@id="txtTwoWord"]').send_keys(ai_captcha)
                time.sleep(0.5)
                n.find_element(By.XPATH, '//*[@id="quote"]').click()
                time.sleep(0.5)
                n.find_element(By.XPATH, '//*[@id="divError"]')
                continue
            except NoSuchElementException:
                break

        # Quotation result and shutdown
        while True:
            try:
                n.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div[3]/'
                                         '/div/table/tbody/tr/td[6]/div/button').click()
                n.find_element(By.XPATH, '//*[@id="shutdown"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # Press print and save file
        while True:
            try:
                n.find_element(By.XPATH, '//*[@id="minutaSO"]')
                time.sleep(0.5)
                n.find_element(By.XPATH, '//*[@id="printPdf"]').click()
                break
            except ElementClickInterceptedException:
                time.sleep(0.1)
                continue
            except NoSuchElementException:
                time.sleep(0.1)
                continue
            except StaleElementReferenceException:
                time.sleep(0.5)
                continue

        # tam_number = n.find_element(By.TAG_NAME, 'h3')
        # tam_text = tam_number.text
        # tam_text = tam_text.replace("-", "")
        time.sleep(0.4)
        # Move file from download folder to destination, renaming it.
        while True:
            try:
                # os.replace(
                #     f"{download_path}\\MinutaWEB-{tam_text[8:]}.pdf",
                #     f"{folder_path1}\\{identificator}{transportadora} - {get_branch(orig)} → {get_branch(dest)}.pdf"
                # )
                os.replace(
                    f"{download_path}\\Minuta.pdf",
                    f"{folder_path1}\\{identificator}{transportadora} - {get_branch(orig)} → {get_branch(dest)}.pdf"
                )
                break
            except FileNotFoundError:
                time.sleep(0.1)
                continue

        close_firefox(n, prot_entry)
    else:
        print('\nTransportadora não é válida!')

    name_info = f"{identificator}{transportadora} - {get_branch(orig)} → {get_branch(dest)}.pdf"

    return name_info


# noinspection PyTypeChecker
def minutas_all_api(start_date, end_date, folderpath, downloadpath, prot_entry, root):
    def get_transp(id_t):
        try:
            if id_t is not None:
                transp = bases.loc[bases['id'] == id_t, 'shippingIDBranch.company_name'].values.item()
            else:
                transp = ""
        except ValueError:
            transp = ""
        return transp

    def protocol_string(p_list):
        count = 0
        p_string = ''
        for p in p_list:
            if count == 0:
                p_string = str(p)
            # elif count == len(p_list) - 1:
            #     p_string += ' e ' + p
            else:
                p_string += ', ' + str(p)
            count += 1
        return p_string

    firefox_driver = Service(
        GeckoDriverManager().install(),
        log_path=f"{tempfile.gettempdir()}/geckodriver.log")

    # Pegar datas

    di = start_date.get_date()
    df = end_date.get_date()

    di = dt.datetime.strftime(di, '%d/%m/%Y')
    df = dt.datetime.strftime(df, '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    df_dt += dt.timedelta(days=1)  # Correção para poder buscar somente um dia

    # Request

    services = r.request_public('https://transportebiologico.com.br/api/public/service')
    bases = r.request_public('https://transportebiologico.com.br/api/public/branch')

    services = pd.concat(
        [services[services['assigned_pdf'].isnull()],
         services[services['assigned_pdf'] == 'nan'],
         services[services['assigned_pdf'] == 'a9c1bc0bb4127ad9f915c113ef9ffb61-38020 - GOLLOG - CWB → GRU.pdf']],
        ignore_index=True
    )

    # CD + identifier

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
            services['step'] == 'boardingService'],
        choicelist=[
            'AGUARDANDO DISPONIBILIZAÇÃO', 'AGUARDANDO ALOCAÇÃO', 'EM ROTA DE ENTREGA', 'ENTREGANDO',
            'DISPONÍVEL PARA RETIRADA', 'DESEMBARCANDO', 'VALIDAR EMBARQUE', 'AGENDADO', 'COLETANDO',
            'EM ROTA DE EMBARQUE', 'EMBARCANDO SERVIÇO'],
        default=0)

    services = pd.concat([
        services[services['ETAPA'] == 'AGENDADO'],
        services[services['ETAPA'] == 'COLETANDO'],
        services[services['ETAPA'] == 'EM ROTA DE EMBARQUE'],
        services[services['ETAPA'] == 'EMBARCANDO SERVIÇO']
    ], ignore_index=True)

    services['CROSSDOCKING'] = np.where(services['serviceIDRequested.crossdocking_collector_id'].str.len() > 0, 1, 0)

    services['Identificador'] = np.select(
        condlist=[(services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] != 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 1) & (services['ETAPA'] == 'EM ROTA DE EMBARQUE'),
                  (services['serviceIDBoard'].str.len() == 2)],
        choicelist=[0, 1, 1],
        default=0)

    services['transportadora1'] = services['serviceIDRequested.source_branch_id'].map(get_transp)
    services['transportadora2'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_transp)

    services['MODAL'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.crossdocking_modal'],
                    services['serviceIDRequested.modal'],
                    services['serviceIDRequested.modal'],
                    '-'],
        default=0)

    # Filtrar aéreo e embarques do dia

    services = services.loc[services['MODAL'] == 'AÉREO']

    # Definir demais variáveis

    services['COLETADOR DESTINO'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.crossdocking_collector_id'],
                    services['serviceIDRequested.destination_collector_id'],
                    services['serviceIDRequested.destination_collector_id'],
                    '-'],
        default=0)

    services['COLETADOR EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0)],
        choicelist=[services['serviceIDRequested.source_collector_id'],
                    services['serviceIDRequested.crossdocking_collector_id'],
                    services['serviceIDRequested.source_collector_id']],
        default=0)

    services['BASE ORIGEM'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.source_branch_id'],
                    services['serviceIDRequested.source_crossdocking_branch_id'],
                    services['serviceIDRequested.source_branch_id'],
                    '-'],
        default=None)

    services['BASE DESTINO'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['serviceIDRequested.destination_crossdocking_branch_id'],
                    services['serviceIDRequested.destination_branch_id'],
                    services['serviceIDRequested.destination_branch_id'],
                    '-'],
        default=None)

    services['Transportadora'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0) & (services['serviceIDRequested.service_type'] == "FRACIONADO"),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['transportadora1'],
                    services['transportadora2'],
                    services['transportadora1'],
                    services['serviceIDRequested.service_type']],
        default=None)

    services = services.loc[services['Transportadora'] == 'LATAM CARGO']

    # Fix hour and date

    time_offset = dt.timedelta(hours=3)

    services['boardDate'] = pd.to_datetime(services['serviceIDRequested.board_date']) - time_offset
    services['boardHour'] = pd.to_datetime(services['serviceIDRequested.board_hour']) - time_offset
    services['CDboardDate'] = pd.to_datetime(services['serviceIDRequested.crossdocking_board_date']) - time_offset
    services['CDboardHour'] = pd.to_datetime(services['serviceIDRequested.crossdocking_board_hour']) - time_offset
    services['boardDate'] = services['boardDate'].dt.tz_localize(None)
    services['boardHour'] = services['boardHour'].dt.tz_localize(None)
    services['CDboardDate'] = services['CDboardDate'].dt.tz_localize(None)
    services['CDboardHour'] = services['CDboardHour'].dt.tz_localize(None)

    services['DATA LIMITE DE EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['CDboardDate'],
                    services['boardDate'],
                    services['boardDate'],
                    services['boardDate']],
        default=services['boardDate'])

    # Filtrar datas

    services.drop(services[services['DATA LIMITE DE EMBARQUE'] < di_dt].index, inplace=True)
    services.drop(services[services['DATA LIMITE DE EMBARQUE'] > df_dt].index, inplace=True)

    services['HORÁRIO LIMITE DE EMBARQUE'] = np.select(
        condlist=[(services['CROSSDOCKING'] == 1) & (services['Identificador'] == 0),
                  (services['CROSSDOCKING'] == 1) & (services['Identificador'] == 1),
                  (services['CROSSDOCKING'] == 0),
                  (services['serviceIDRequested.service_type'] == "DEDICADO")],
        choicelist=[services['CDboardHour'],
                    services['boardHour'],
                    services['boardHour'],
                    services['boardHour']],
        default=services['boardHour'])

    services['HORÁRIO LIMITE DE EMBARQUE T'] = services['HORÁRIO LIMITE DE EMBARQUE'].dt.strftime(date_format="%H:%M")
    services['DATA LIMITE DE EMBARQUE T'] = services['DATA LIMITE DE EMBARQUE'].dt.strftime(date_format="%d-%m-%Y")

    services['TIPO DE MATERIAL'] = np.select(
        condlist=[services['serviceIDRequested.material_type'] == 'BIOLÓGICO UN3373',
                  services['serviceIDRequested.material_type'] == 'INFECTANTE UN2814/UN2900',
                  services['serviceIDRequested.material_type'] == 'BIOLÓGICO RISCO MÍNIMO/ISENTO',
                  services['serviceIDRequested.material_type'] == 'CARGA GERAL',
                  services['serviceIDRequested.material_type'] == 'CORRELATOS'],
        choicelist=['BIO',
                    'INF',
                    'GER',
                    'GER',
                    'GER'],
        default=None
    )

    services['TIPO DE ENVIO'] = np.where(services['serviceIDRequested.deadline'] < 5, 'PRV', 'CONV')

    # Concatenar → Col. Origem + Col. Destino + Base Origem + Base Destino + Horário de embarque + Data de embarque

    services['uniqueDispatch'] = services['HORÁRIO LIMITE DE EMBARQUE T'].astype('str') + \
                                 services['DATA LIMITE DE EMBARQUE T'].astype('str') + \
                                 services['BASE ORIGEM'].astype('str') + \
                                 services['BASE DESTINO'].astype('str') + \
                                 services['COLETADOR EMBARQUE'].astype('str') + \
                                 services['COLETADOR DESTINO'].astype('str') + \
                                 services['Transportadora'].astype('str') + \
                                 services['TIPO DE MATERIAL'].astype('str') + \
                                 services['TIPO DE ENVIO'].astype('str')

    services.to_excel(f'{temp_folder}/Excel/Services_all_minutas.xlsx', index=False)

    # Lista a partir de coluna e np.unique

    dispatches = pd.DataFrame(columns=[])
    dispatches['uniqueDispatch'] = np.unique(services['uniqueDispatch'])
    dispatches['Protocolo'] = np.empty((len(dispatches), 0)).tolist()

    dispatches.to_excel(f'{temp_folder}/Excel/Dispatches.xlsx', index=False)

    # Verificação inversa (count=0=row e for i in list)

    for uDu, uProt in zip(dispatches['uniqueDispatch'], dispatches['Protocolo']):
        for uDs, prot in zip(services['uniqueDispatch'], services['protocol']):
            if uDs == uDu:
                uProt.append(prot)

    dispatches['p_strings'] = dispatches['Protocolo'].map(protocol_string)
    dispatches['Arquivo PDF'] = "nan"

    print(dispatches['p_strings'])

    current_row = 0
    folder_path1 = folderpath.replace("/", "\\")

    for protocol, p_strings in zip(dispatches['Protocolo'], dispatches['p_strings']):
        print(protocol[0])

        # Definir parâmetros de emissão
        tipo_material = services.loc[services['protocol'] == protocol[0], 'TIPO DE MATERIAL'].values.item()
        tipo_envio = services.loc[services['protocol'] == protocol[0], 'TIPO DE ENVIO'].values.item()
        if tipo_material == "BIO":
            material_type = 0
            flight_service = 0
        elif tipo_material == "GER" and tipo_envio == "PRV":
            material_type = 1
            flight_service = 0
        elif tipo_material == "INF":
            material_type = 2
            flight_service = 0
        else:
            material_type = 1
            flight_service = 1

        # Emissão trocando o nome
        archive_name = minutas_api(str(protocol[0]), True, p_strings, flight_service,
                                   material_type, '1', '1', folderpath, downloadpath,
                                   prot_entry, firefox_driver)

        # Define dataframe report bases on dispatches (protList, archive_name)
        dispatches.at[dispatches.index[current_row], 'Arquivo PDF'] = archive_name
        report = dispatches[['Protocolo', 'Arquivo PDF']].copy()

        # Explode report on protList
        report = report.explode('Protocolo')

        # Export to csv
        report.to_csv(f'{folder_path1}\\UploadPDF.csv', index=False)

        r.post_file("https://transportebiologico.com.br/api/pdf",
                    f'{folder_path1}\\{archive_name}',
                    upload_type="MINUTA",
                    file_format="application/pdf",
                    file_type="pdf_files")

        r.post_file('https://transportebiologico.com.br/api/pdf/associate',
                    f'{folder_path1}\\UploadPDF.csv',
                    upload_type="MINUTA")

        os.remove(f'{folder_path1}\\{archive_name}')

        current_row += 1

    confirmation_pop_up(root)


def cargas_api(filename):
    def get_collector(col_id):
        try:
            col_name = collector.loc[collector['id'] == col_id, 'trading_name'].values.item()
        except ValueError:
            col_name = ""
        return col_name

    def get_transp(id_t):
        try:
            transp = bases.loc[bases['id'] == id_t, 'shippingIDBranch.company_name'].values.item()
        except ValueError:
            transp = ""
        return transp

    def get_branch(id_t):
        try:
            branch = bases.loc[bases['id'] == id_t, 'nickname'].values.item()
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

    def excel_date(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days)

    def excel_time(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.seconds) / 86400

    # request dataframes

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    bases = r.request_public('https://transportebiologico.com.br/api/public/branch')
    collector = r.request_public('https://transportebiologico.com.br/api/public/collector')
    services = r.request_public('https://transportebiologico.com.br/api/public/service')

    # remove busines services

    services = services.loc[services['is_business'] == False]

    # drop unwanted stages

    services.drop(services[services['step'] == 'toCollectService'].index, inplace=True)
    services.drop(services[services['step'] == 'collectingService'].index, inplace=True)
    services.drop(services[services['step'] == 'toBoardService'].index, inplace=True)
    services.drop(services[services['step'] == 'boardingService'].index, inplace=True)

    # rename the stages

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

    # create crossdocking and identificator columns

    services['CROSSDOCKING'] = np.where(services['serviceIDRequested.crossdocking_collector_id'].str.len() > 0, 1, 0)

    services['Identificador'] = np.select(
        condlist=[(services['serviceIDBoard'].str.len() == 1),
                  (services['serviceIDBoard'].str.len() == 2)],
        choicelist=[1, 2],
        default=0
    )

    # get names from ids

    services['Coletador de Origem'] = services['serviceIDRequested.source_collector_id'].map(get_collector)
    services['Coletador intermediário'] = services['serviceIDRequested.crossdocking_collector_id'].map(get_collector)
    services['Coletador de destino'] = services['serviceIDRequested.destination_collector_id'].map(get_collector)

    services['transportadora1'] = services['serviceIDRequested.source_branch_id'].map(get_transp)
    services['transportadora2'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_transp)

    services['Base_origem'] = services['serviceIDRequested.source_branch_id'].map(get_branch)
    services['cd_Base_destino'] = services['serviceIDRequested.destination_crossdocking_branch_id'].map(get_branch)
    services['cd_Base_origem'] = services['serviceIDRequested.source_crossdocking_branch_id'].map(get_branch)
    services['Base_destino'] = services['serviceIDRequested.destination_branch_id'].map(get_branch)

    # get data from both dispatches

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

    # define which data to use in each service

    services['VOLUMES'] = np.where(services['tempo1'] >= services['tempo2'], services['VOLUMES1'], services['VOLUMES2'])
    services['CTE'] = np.where(services['tempo1'] >= services['tempo2'], services['CTE1'], services['CTE2'])
    services['RASTREADOR'] = np.where(
        services['tempo1'] >= services['tempo2'], services['RASTREADOR1'], services['RASTREADOR2'])

    # define the correct data to display, corresponding to the right stage

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
        default='')

    # correct the timeframes, API responise is UTC 0, we want UTC -3

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

    services['LIMITE DE ENTREGA'] = services['DATA DE ENTREGA'].map(excel_date) + \
                                    services['HORÁRIO LIMITE DE ENTREGA'].map(excel_time)

    services['CIDADES ORIGEM'] = services['serviceIDRequested.source_address_id'].map(get_address_city)
    services['CIDADES DESTINO'] = services['serviceIDRequested.destination_address_id'].map(get_address_city)

    # select the columns that go into the final report

    cg = services[[
        'protocol',
        'customerIDService.trading_firstname',
        'DATA DA COLETA',
        'LIMITE DE ENTREGA',
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
        'serviceIDRequested.gelo_seco',
        'ETAPA',
        'serviceIDRequested.service_type',
        'CROSSDOCKING'
    ]].copy()

    # sort the rows accordingly

    cg.sort_values(by="LIMITE DE ENTREGA", axis=0, ascending=True, inplace=True, kind='quicksort',
                   na_position='last')
    cg.sort_values(by="RASTREADOR", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')
    cg.sort_values(by="Transportadora", axis=0, ascending=True, inplace=True, kind='stable', na_position='last')

    # separate depending on the phase

    cg1 = cg.loc[(cg['ETAPA'] != 'EM ROTA DE ENTREGA') & (cg['ETAPA'] != 'ENTREGANDO')]
    cg2 = cg.loc[(cg['ETAPA'] == 'EM ROTA DE ENTREGA') | (cg['ETAPA'] == 'ENTREGANDO')]

    # treat the name of the file to export to

    file_name = filename.replace('/', '\\')

    # CoInitialize to let xlwings run while threading and export to excel

    CoInitialize()

    export_to_excel(cg1, file_name, 'Desembarque', 'B4:T500', autofit=False, change_header=False, start_write="B4",
                    clear_filters=True)
    export_to_excel(cg2, file_name, 'Entrega', 'B4:T500', autofit=False, change_header=False, start_write="B4",
                    clear_filters=True)


def fleury_sheet(date, filename):

    def get_city(add):
        add_lenght = len(add)
        count = 0
        add_total = ''
        for i in range(add_lenght):

            add_city = address.loc[address['id'] == add[i], 'cityIDAddress.name'].values.item()
            count += 1

            if count != add_lenght:
                space = ", "
            else:
                space = ""

            add_total += add_city.title() + space
        return add_total

    def excel_date(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.days)

    def excel_time(date1):
        temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return float(delta.seconds) / 86400

    di = date.get_date()
    df = date.get_date()

    di = dt.datetime.strftime(di, '%d/%m/%Y')
    df = dt.datetime.strftime(df, '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    di_temp = di_dt - dt.timedelta(days=5)
    df_temp = df_dt + dt.timedelta(days=5)

    di = dt.datetime.strftime(di_temp, '%Y-%m-%d')
    df = dt.datetime.strftime(df_temp, '%Y-%m-%d')

    # Requesting data from API

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    services = r.request_private(
        f'https://transportebiologico.com.br/api/consult/'
        f'service?startFilter={di}T00:00:00.000-03:00&endFilter={df}T23:59:00.000-03:00')
    services = services.loc[services['step'] != 'toValidateCancelRequest']
    services.to_excel(f'{temp_folder}/Excel/ServicesAPI.xlsx', index=False)

    services.drop(
        services[
            (services['customerIDService.trading_firstname'] != 'FLEURY') &
            (services['customerIDService.trading_firstname'] != 'PRETTI') &
            (services['customerIDService.trading_firstname'] != 'BIOCLINICO') &
            (services['customerIDService.trading_firstname'] != 'HERMES PARDINI')
        ].index,
        inplace=True)

    services['collectDatetime'] = pd.to_datetime(services['serviceIDRequested.collect_date']) - dt.timedelta(hours=3)
    services['collectDatetime'] = services['collectDatetime'].dt.tz_localize(None)

    df_dt += dt.timedelta(days=1)  # Correção para poder buscar somente um dia

    services.drop(services[services['collectDatetime'] < di_dt].index, inplace=True)
    services.drop(services[services['collectDatetime'] > df_dt].index, inplace=True)

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
            services['step'] == 'pendingGeloSecoMaintenance',
            services['step'] == 'toMaintenanceService',
            services['step'] == 'toValidateFiscalRetention',
            services['step'] == 'finishedService',
            services['step'] == 'cancelledService',
            services['step'] == 'unsuccessService'
        ],
        choicelist=[
            'EMBARCADO',
            'EMBARCADO',
            'EM ROTA DE ENTREGA',
            'EM ROTA DE ENTREGA',
            'EMBARCADO',
            'EMBARCADO',
            'EMBARCADO',
            'AGENDADO',
            'AGENDADO',
            'COLETADO',
            'COLETADO',
            'EMBARCADO',
            'EMBARCADO',
            'EMBARCADO',
            'ENTREGUE',
            'SEM MATERIAL',
            'SEM MATERIAL'],
        default=0)

    services['HORA INÍCIO'] = pd.to_datetime(services['serviceIDRequested.collect_hour_start']) - dt.timedelta(hours=3)
    services['HORA INÍCIO'] = services['HORA INÍCIO'].dt.strftime(date_format="%H:%M")

    services['hourEnd'] = pd.to_datetime(services['serviceIDRequested.collect_hour_end']) - dt.timedelta(hours=3)
    services['HORA FIM'] = services['hourEnd'].dt.strftime(date_format="%H:%M")

    services['dataEntrega'] = pd.to_datetime(services['serviceIDRequested.delivery_date']) - dt.timedelta(hours=3)
    services['dataEntrega'] = services['dataEntrega'].dt.tz_localize(None)

    services['horaEntrega'] = pd.to_datetime(services['serviceIDRequested.delivery_hour']) - dt.timedelta(hours=3)
    services['horaEntrega'] = services['horaEntrega'].dt.tz_localize(None)

    services['HORA ENTREGA'] = services['horaEntrega'].dt.strftime(date_format="%H:%M")

    services['DEADLINE'] = ('D+' +
                            services['serviceIDRequested.deadline'].astype(int).astype(str) +
                            ' ' +
                            services['HORA ENTREGA'])

    services['DATA ENTREGA'] = services['dataEntrega'].map(excel_date)

    services['COLETA'] = services['serviceIDCollect'].str[0]

    services['VOL_COLETA'] = services['COLETA'].map(
        lambda x: x.get('collect_volume', np.nan) if isinstance(x, dict) else np.nan)

    services['EMBARQUE'] = services['serviceIDBoard'].str[0]

    services['VOLUMES'] = services['EMBARQUE'].map(
        lambda x: x.get('board_volume', np.nan) if isinstance(x, dict) else np.nan)

    services['RASTREADOR'] = services['EMBARQUE'].map(
        lambda x: x.get('operational_number', np.nan) if isinstance(x, dict) else np.nan)

    services['RASTREADOR'] = np.select(
        condlist=[services['ETAPA'] == 'AGENDADO',
                  services['ETAPA'] == 'COLETADO',
                  (services['ETAPA'] == 'EMBARCADO') & (services['serviceIDRequested.modal'] == 'AÉREO'),
                  (services['ETAPA'] == 'EMBARCADO') & (services['serviceIDRequested.modal'] == 'RODOVIÁRIO'),
                  (services['ETAPA'] == 'EM ROTA DE ENTREGA') & (services['serviceIDRequested.modal'] == 'AÉREO'),
                  (services['ETAPA'] == 'EM ROTA DE ENTREGA') & (services['serviceIDRequested.modal'] == 'RODOVIÁRIO'),
                  (services['ETAPA'] == 'ENTREGUE') & (services['serviceIDRequested.modal'] == 'AÉREO'),
                  (services['ETAPA'] == 'ENTREGUE') & (services['serviceIDRequested.modal'] == 'RODOVIÁRIO')],
        choicelist=['',
                    '',
                    services['RASTREADOR'],
                    'RODOVIÁRIO',
                    services['RASTREADOR'],
                    'RODOVIÁRIO',
                    services['RASTREADOR'],
                    'RODOVIÁRIO'],
        default="-")

    services['ENTREGA'] = services['serviceIDDelivery'].str[0]

    services['datetimeDelivery'] = services['ENTREGA'].map(
        lambda x: x.get('real_arrival_timestamp', np.nan) if isinstance(x, dict) else np.nan)

    services['datetimeDelivery'] = pd.to_datetime(services['datetimeDelivery']) - dt.timedelta(hours=3)
    services['datetimeDelivery'] = services['datetimeDelivery'].dt.tz_localize(None)

    services['HORA ENTREGUE'] = services['datetimeDelivery'].map(excel_time)
    services['DATA ENTREGUE'] = services['datetimeDelivery'].map(excel_date)

    services['RECEBEDOR'] = services['ENTREGA'].map(
        lambda x: x.get('responsible_name', np.nan) if isinstance(x, dict) else np.nan)

    services['VOLUMES'] = np.select(
        condlist=[services['ETAPA'] == 'AGENDADO',
                  services['ETAPA'] == 'COLETADO',
                  services['ETAPA'] == 'EMBARCADO',
                  (services['ETAPA'] == 'EM ROTA DE ENTREGA') & (
                          services['serviceIDRequested.service_type'] == 'FRACIONADO'),
                  (services['ETAPA'] == 'EM ROTA DE ENTREGA') & (
                              services['serviceIDRequested.service_type'] == 'DEDICADO'),
                  (services['ETAPA'] == 'ENTREGUE') & (
                              services['serviceIDRequested.service_type'] == 'FRACIONADO'),
                  (services['ETAPA'] == 'ENTREGUE') & (
                          services['serviceIDRequested.service_type'] == 'DEDICADO')
                  ],
        choicelist=['',
                    services['VOL_COLETA'],
                    services['VOLUMES'],
                    services['VOLUMES'],
                    services['VOL_COLETA'],
                    services['VOLUMES'],
                    services['VOL_COLETA']],
        default='-')

    services['DATA ENTREGA'] = np.select(
        condlist=[services['ETAPA'] != 'ENTREGUE',
                  services['ETAPA'] == 'ENTREGUE'],
        choicelist=[services['DATA ENTREGA'],
                    services['DATA ENTREGUE']])
    services['HORA ENTREGA'] = np.select(
        condlist=[services['ETAPA'] != 'ENTREGUE',
                  services['ETAPA'] == 'ENTREGUE'],
        choicelist=['',
                    services['HORA ENTREGUE']])
    services['RESPONSÁVEL'] = np.select(
        condlist=[services['ETAPA'] != 'ENTREGUE',
                  services['ETAPA'] == 'ENTREGUE'],
        choicelist=['',
                    services['RECEBEDOR']])

    services['CIDADE'] = services['serviceIDRequested.source_address_id'].map(get_city)
    services['HOUR'] = services['HORA INÍCIO'] + ' \\ ' + services['HORA FIM']
    services['DATA'] = excel_date(di_dt)
    services['SETOR'] = np.select(
        condlist=[(services['serviceIDRequested.budget_id'] == 'da9c8052-a97b-4453-928c-d0d5d03dab54') |
                  (services['serviceIDRequested.budget_id'] == '7a835dda-6548-40d9-9c3f-86cc3b3c9dad') |
                  (services['serviceIDRequested.budget_id'] == 'c1abe48d-3aaf-4045-9776-159fed8e0875'),
                  services['serviceIDRequested.budget_id'] == '46a4a2c1-f1a8-4c39-ab01-d750da9f4e5d'],
        choicelist=['NOVA AMÉRICA A.T.',
                    'SMD'],
        default='PRET'
    )

    day_of_week = di_dt.strftime("%A")
    sheet_day = weekday_translation[day_of_week]

    file_name = filename.replace('/', '\\')

    df = pd.read_excel(file_name, sheet_name=sheet_day)

    id_list = df.iloc[:, 0].dropna().tolist()

    print(id_list)

    temp_df = services[[
        'CIDADE',
        'HOUR',
        'DEADLINE',
        'protocol',
        'RASTREADOR',
        'VOLUMES',
        'DATA',
        'SETOR',
        'ETAPA',
        'DATA ENTREGA',
        'HORA ENTREGA',
        'RESPONSÁVEL',
        'serviceIDRequested.budget_id'
    ]].copy()

    CoInitialize()

    clear_data(file_name, f'{sheet_day};E4:Q50')

    change_cell(file_name, "D2", excel_date(di_dt), sheet_day)

    for service_id in id_list:
        # Create a filtered DataFrame just for rows matching the current `service_id`
        filtered_df = temp_df[temp_df['serviceIDRequested.budget_id'] == service_id].copy()

        if not filtered_df.empty:
            # Find cell location in Excel
            result = df.isin([service_id])
            cell_location = None
            for row_index, col_index in zip(*result.to_numpy().nonzero()):
                col_letter = chr(col_index + 69)  # Adjust column to match Excel format
                cell_location = f"{col_letter}{row_index + 2}"
                break  # Stop after finding the first match

            filtered_df.drop(columns=['serviceIDRequested.budget_id'], inplace=True)

            # Export the filtered DataFrame to Excel
            export_to_excel(
                df=filtered_df, excel_name=file_name, sheet=sheet_day,
                start_write=cell_location, change_header=False
            )

    print('finished')


temp_folder = f'{tempfile.gettempdir()}/Operação LogLife'
os.makedirs(f'{temp_folder}/Excel', exist_ok=True)

r = RequestDataFrame()

weekday_translation = {
    "Monday": "SEGUNDA",
    "Tuesday": "TERÇA",
    "Wednesday": "QUARTA",
    "Thursday": "QUINTA",
    "Friday": "SEXTA",
    "Saturday": "SÁBADO",
    "Sunday": "DOMINGO"
}
os.makedirs(f'{temp_folder}/resolver_captcha', exist_ok=True)
captcha_image = f'{temp_folder}\\resolver_captcha\\captcha.png'
