from json import loads
import pandas as pd
import requests


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
        print(response.text)

        return response

    def post_private(self, link, payload):

        response = requests.post(url=link, headers=self.auth, data=payload)

        return response


r = RequestDataFrame()

hubs = r.request_public('https://transportebiologico.com.br/api/public/hub')
transp = r.request_public('https://transportebiologico.com.br/api/public/shipping')
bases = r.request_private('https://transportebiologico.com.br/api/branch')

hubs.rename(columns={"id": "id_hubs", "nome": "Hub de Origem"}, errors="raise", inplace=True)
transp.rename(columns={"id": "id_transp", "nome": "Transportadora"}, errors="raise", inplace=True)
bases.rename(columns={"id": "id_bases", "nickname": "Base"}, errors="raise", inplace=True)

hubs['Hub Destino'] = hubs['Hub de Origem']
bases['Base-final-origem'] = bases[['Base', 'shippingIDBranch.company_name']].agg('-'.join, axis=1)
bases['Base-final-destino'] = bases['Base-final-origem']
bases.drop(bases[bases['situation'] == 'inactive'].index, inplace=True)

malha = pd.read_excel('src/excel/Tabela_Malha.xlsx')

malha.drop('Unnamed: 0', axis=1, inplace=True)
malha.columns = malha.iloc[0]
malha = malha.iloc[1:, :]
malha.reset_index()

malha1 = pd.merge(malha, transp, on='Transportadora', how='inner')
malha1 = pd.merge(malha1, hubs[['Hub de Origem', 'id_hubs']], on='Hub de Origem', how='inner')
malha1.rename(columns={'id_hubs': 'id_hub_origem'}, errors='raise', inplace=True)
malha1 = pd.merge(malha1, hubs[['Hub Destino', 'id_hubs']], on='Hub Destino', how='inner')
malha1.rename(columns={'id_hubs': 'id_hub_destino'}, errors='raise', inplace=True)
malha1['Base-final-origem'] = malha1[['Base Origem', 'Transportadora']].agg('-'.join, axis=1)
malha1['Base-final-destino'] = malha1[['Base Destino', 'Transportadora']].agg('-'.join, axis=1)
malha1 = pd.merge(malha1, bases[['Base-final-origem', 'id_bases']], on='Base-final-origem', how='inner')
malha1.rename(columns={'id_bases': 'id_base_origem'}, inplace=True, errors='raise')
malha1 = pd.merge(malha1, bases[['Base-final-destino', 'id_bases']], on='Base-final-destino', how='inner')
malha1.rename(columns={'id_bases': 'id_base_destino'}, errors='raise', inplace=True)

print(malha1)

malha_final = malha1[['id_transp', 'id_hub_origem', 'id_hub_destino', 'id_base_origem', 'id_base_destino', 'Data Saída',
                      'Data Chegada', 'Horário de Saída', 'Horário de Chegada', 'Voo / Caminhão', 'Modal']]
malha_final = malha_final.rename(columns={'id_transp': 'Transportadora', 'id_hub_origem': 'Hub de Origem',
                                          'id_hub_destino': 'Hub Destino', 'id_base_origem': 'Base Origem',
                                          'id_base_destino': 'Base Destino', 'Voo / Caminhão': 'Voo/Caminhão'})

malha_final.to_csv("Tabela_Malha_ID.csv", index=False)
malha_final.to_excel("Tabela_Malha_ID.xlsx", index=False)

# r = RequestDataFrame()
# results = r.post_file('https://transportebiologico.com.br/api/route-network', 'Tabela_Malha_ID.csv')
