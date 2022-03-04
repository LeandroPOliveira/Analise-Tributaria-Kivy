from kivy.config import Config
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
import os
from datetime import datetime, date
from kivy.utils import get_color_from_hex
import pandas as pd
import win32clipboard
from kivy.core.clipboard import Clipboard

# from kivy.core.window import Window
# Window.size = (1280, 720)
from kivymd.uix.label import MDLabel
from kivymd.uix.textfield import MDTextField, MDTextFieldRect, MDTextFieldRound

Config.set('graphics', 'resizable', '1')
Config.set('graphics', 'width', '1280')
Config.set('graphics', 'height', '720')
Config.write()


class TelaLogin(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)




class AnalisesPendentes(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.data_tables = None
        self.gere = 'gere'

    def add_datatable(self):
        self.lista = []
        self.pasta1 = os.listdir('G:\GECOT\Análise Contábil_Tributária_Licitações\\2022\\1Pendentes\\')
        self.pasta = []
        for item in self.pasta1:
            self.pasta.append(item)

        for i, n in enumerate(self.pasta):
            mod = os.path.getctime('G:\GECOT\Análise Contábil_Tributária_Licitações\\2022\\1Pendentes\\')
            mod = datetime.fromtimestamp(mod)
            data = date.strftime(mod, '%d/%m/%Y')
            # self.lista.append([])
            self.lista.append((n, data))
            # self.lista[i].append(data)

        self.data_tables = MDDataTable(pos_hint={'center_x': 0.5, 'center_y': 0.5},
                                       size_hint=(0.4, 0.75),
                                       check=True, use_pagination=True,
                                       background_color_header=get_color_from_hex("#65275d"),

                                       column_data=[("[color=#ffffff]Análise[/color]", dp(40)),
                                                    ("[color=#ffffff]Data[/color]", dp(40))],
                                       row_data=self.lista, elevation=1)


        self.add_widget(self.data_tables)


        self.data_tables.bind(on_check_press=self.checked)
        self.data_tables.bind(on_row_press=self.row_checked)

        # self.theme_cls.theme_style = 'Light'
        # self.theme_cls.primary_palette = 'BlueGray'
        # Adicionar tabela na tela


        self.lista2 = []

    def checked(self, instance_table, current_row):
        self.lista2.append(current_row[0])
        arquivo = current_row
        # os.startfile('C:\\Users\leandro\Desktop\pendente\\' + arquivo[0])
        # print(self.lista2)

    def row_checked(self, instance_table, current_row):
        if self.data_tables.get_row_checks():
            pass
        else:
            print(type(self.data_tables.get_row_checks()))
            os.startfile('G:\GECOT\Análise Contábil_Tributária_Licitações\\2022\\1Pendentes\\' + current_row.text)


class NovaAnalise(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)


    def teste(self):
        self.pasta_principal = 'G:\\GECOT\Análise Contábil_Tributária_Licitações\\2022\\1Pendentes\\'
        self.lista_mat = [[], [], [], [], [], [], [], []]
        self.entradas_mat = []
        self.data_mat = ['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI', 'PIS', 'COFINS']

        for num, item in enumerate(self.data_mat):
            if num == 1:
                largura = .2
            elif num == 0:
                largura = .1
            elif num == 3:
                largura = .1
            else:
                largura = .05
            self.cabecalho = MDLabel(text=item, size_hint=(largura, .05), halign='center')
            self.ids.grid_teste.add_widget(self.cabecalho)

        for i in range(30):
            for c in range(8):
                if c == 1:
                    largura = .2
                elif c == 0:
                    largura = .1
                elif c == 3:
                    largura = .1
                else:
                    largura = .05
                self.mater = MDTextFieldRect(multiline=False, size_hint=(largura, .05))
                self.entradas_mat.append(self.mater)
                self.lista_mat[c].append(self.mater)

                # self.entradas_mat.append(mater)
                # self.lista_mat[c].append(mater)


                self.ids.grid_teste.add_widget(self.mater)


        for i, n in enumerate(self.entradas_mat):
            if i % 8 == 0:
                self.entradas_mat[i].bind(on_text_validate=self.colar)




    def colar(self, instance):
        for i, l in enumerate(self.entradas_mat):
            if i % 8 == 0:
                if l.text == '':
                    self.posicao1 = int(i / 8) if i > 0 else i
                    print(self.posicao1)
                    break
        cad_mat = pd.read_excel(self.pasta_principal + 'material.xlsx', sheet_name='materiais')
        cad_mat = pd.DataFrame(cad_mat)
        cad_mat['Material'] = cad_mat['Material'].astype(str)
        win32clipboard.OpenClipboard()
        rows = win32clipboard.GetClipboardData()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        rows = rows.split('\n')

        rows.pop() if len(rows) > 1 else rows
        for r, val in enumerate(rows):
            values = val.split('\t')
            print(values)
            if len(values) > 1:
                del values[1:]
            for b, value in enumerate(values):
                for index, row in cad_mat.iterrows():
                    self.lista_mat[b][r + self.posicao1].text = value
                    if value.strip() == row['Material']:
                        campo = cad_mat.loc[index, 'Texto breve material']
                        campo = campo[:32]
                        self.lista_mat[b + 1][r + self.posicao1].text = campo
                        self.lista_mat[b + 3][r + self.posicao1].text = cad_mat.loc[index, 'Ncm']




    def preenche_aliq(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 8 == 0:
                if item.text != '':
                    self.entradas_mat[e + 4].text = '18%'
                    self.entradas_mat[e + 5].text = '15%'
                    self.entradas_mat[e + 6].text = '1,65%'
                    self.entradas_mat[e + 7].text = '7,6%'

    def preenche_iva(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 8 == 0 and e != 0:
                if item.text != '':
                    self.entradas_mat[e + 2].text = self.entradas_mat[2].text


    def preenche_ncm(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 8 == 0 and e != 0:
                if item.text != '':
                    self.entradas_mat[e + 3].text = self.entradas_mat[3].text

    def campos_serv(self):
        self.lista = [[], [], []]
        self.entradas = []
        self.data = [['DESCRIÇÃO', 'CÓDIGO', 'C.C']]

        for i in range(60):
            for c in range(3):
                if c == 1:
                    largura = 30
                else:
                    largura = 15
                self.serv = MDTextFieldRect(multiline=False, size_hint=(largura, .05))
                self.entradas.append(self.serv)
                self.lista[c].append(self.serv)

                self.ids.grid_serv.add_widget(self.serv)

        for i, n in enumerate(self.entradas):
            if i % 3 == 0:
                self.entradas[i].bind(on_text_validate=self.colar_serv)




    def colar_serv(self, instance):

        for i, l in enumerate(self.entradas):
            if l.text == '':
                self.posicao = int(i / 3) if i > 0 else i
                print(self.posicao)
                break
        serv_cad = pd.read_excel(self.pasta_principal + 'material.xlsx', sheet_name='servicos')
        serv_cad = pd.DataFrame(serv_cad)
        serv_cad['Nº de serviço'] = serv_cad['Nº de serviço'].astype(str)
        win32clipboard.OpenClipboard()
        rows = win32clipboard.GetClipboardData()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        rows = rows.split('\n')


        rows.pop() if len(rows) > 1 else rows

        for r, val in enumerate(rows):
            values = val.split('\t')
            if len(values) > 1:
                del values[1:]
            for b, value in enumerate(values):
                for index, row in serv_cad.iterrows():
                    self.lista[b][r + self.posicao].text = value
                    if value == row['Nº de serviço']:
                        self.lista[b + 1][r + self.posicao].text = serv_cad.loc[index, 'Denominação']
                        self.lista[b + 2][r + self.posicao].text = str(int(serv_cad.loc[index, 'Classe avaliaç.']))





class WindowManager(ScreenManager):
    pass


class Example(MDApp):

    def build(self):
        self.items = [{'icon': 'Sim'}, {'icon': 'Não'}]
        return Builder.load_file('analisetribut.kv')


Example().run()