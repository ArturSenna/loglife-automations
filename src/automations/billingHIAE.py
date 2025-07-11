import datetime as dt
import os
import time
import base64
import numpy as np
import pandas as pd
import win32com.client
from pythoncom import CoInitialize
import tempfile
from functions import *


def hiae_billing(start_date, final_date):

    def get_address_id(trading_name):
        try:
            add = address.loc[address['trading_name'] == trading_name, 'id'].values.item()
        except Exception as e:
            add = e
        return add

    def iterate_value(servideIDcollecss):

        values_list = []

        for item in servideIDcollecss:



    # now = dt.datetime.now()
    # now_date = dt.datetime.strftime(start_date, '%d-%m-%Y')
    # now = dt.datetime.strftime(now, "%H-%M")

    di = dt.datetime.strftime(start_date.get_date(), '%d/%m/%Y')
    df = dt.datetime.strftime(final_date.get_date(), '%d/%m/%Y')

    di_dt = dt.datetime.strptime(di, '%d/%m/%Y')
    df_dt = dt.datetime.strptime(df, '%d/%m/%Y')

    di_temp = di_dt - dt.timedelta(days=5)
    df_temp = df_dt + dt.timedelta(days=5)

    di = dt.datetime.strftime(di_temp, '%d/%m/%Y')
    df = dt.datetime.strftime(df_temp, '%d/%m/%Y')

    di_formatted = dt.datetime.strftime(di_dt, '%Y-%m-%d')
    df_formatted = dt.datetime.strftime(df_dt, '%Y-%m-%d')

    pl = {
        "shipping_id": None,
        "is_driver": False,
        "customer_id": "40251716-e314-478f-98fb-929b24c940bc",
        "collector_id": None,
        "startFilter": f"{di_formatted}T00:00:00.000-03:00",
        "endFilter": f"{df_formatted}T23:59:00.000-03:00",
        "is_customer": False,
        "is_collector": False
    }

    billing = r.request_private(
        link='https://transportebiologico.com.br/api/report/billing',
        request_type='post',
        payload=pl
    )

    address = r.request_public('https://transportebiologico.com.br/api/public/address')
    address = address.loc[
        (address['customerIDAddress.trading_firstname'] == 'HIAE - HOSPITAL ALBERT EINSTEIN') |
        (address['customerIDAddress.trading_firstname'] == 'HOE SALVADOR - HOSPITAL ALBERT EINSTEIN ')
    ]

    sv = r.request_private('https://transportebiologico.com.br/api/consult/service?customer_id='
                           '40251716-e314-478f-98fb-929b24c940bc&'
                           f'startFilter={di_formatted}T00:00:00.000-03:00&'
                           f'endFilter={df_formatted}T23:59:00.000-03:00')

    sv = sv.explode('serviceIDCollect')
    billing['split_address'] = billing['sourceAddress'].str.split("\n")
    billing = billing.explode('split_address')
    billing['split_address'] = billing['split_address'].str[3:]
    billing['address_id'] = billing['split_address'].map(get_address_id)

    sv['address_id'] = sv['serviceIDCollect'].map(
        lambda x: x.get('address_id', np.nan) if isinstance(x, dict) else np.nan)

    sv['address_step'] = sv['serviceIDCollect'].map(
        lambda x: x.get('step', np.nan) if isinstance(x, dict) else np.nan)

    sv.to_excel('Excel/sv.xlsx', index=False)
    billing.to_excel('Excel/Billing.xlsx', index=False)

    # Merge the data based on 'address_id' and 'protocol' columns
    merged_df = billing.merge(sv[['protocol', 'address_id', 'address_step']],
                                 on=['protocol', 'address_id'],
                                 how='left')

    merged_df.to_excel('Excel/Merged.xlsx', index=False)

    address.to_excel("Excel/Address.xlsx", index=False)


r = RequestDataFrame()
