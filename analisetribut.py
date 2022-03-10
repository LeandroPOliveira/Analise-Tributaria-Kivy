import pickle
from time import sleep

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
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from fpdf import FPDF
from kivy.core.clipboard import Clipboard

# from kivy.core.window import Window
# Window.size = (1280, 720)
from kivymd.uix.label import MDLabel
from kivymd.uix.selectioncontrol import MDCheckbox
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

    def tabela_materiais(self):
        self.pasta_principal = 'G:\\GECOT\Análise Contábil_Tributária_Licitações\\2022\\1Pendentes\\'
        self.lista_mat = [[], [], [], [], [], [], [], []]
        self.entradas_mat = []
        self.data_mat = ['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI', 'PIS', 'COFINS']

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
                self.mater = MDTextFieldRect(multiline=False, size_hint=(largura, .05), write_tab=False)
                self.entradas_mat.append(self.mater)
                self.lista_mat[c].append(self.mater)

                # self.entradas_mat.append(mater)
                # self.lista_mat[c].append(mater)

                self.ids.grid_teste.add_widget(self.mater)

        for i, n in enumerate(self.entradas_mat):
            if i % 8 == 0:
                self.entradas_mat[i].bind(on_text_validate=self.colar)
                self.entradas_mat[i + 1].bind(focus=self.colar2)

    def colar2(self, instance, widget):
        cad_mat = pd.read_excel(self.pasta_principal + 'material.xlsx', sheet_name='materiais')
        cad_mat = pd.DataFrame(cad_mat)
        cad_mat['Material'] = cad_mat['Material'].astype(str)

        for i, l in enumerate(self.entradas_mat):
            if i % 8 == 0:
                if l.text != '' and self.entradas_mat[i + 1].text == '':
                    self.posicao1 = int(i / 8) if i > 0 else i
                    print(self.posicao1)

                    for index, row in cad_mat.iterrows():
                        if l.text == row['Material']:
                            campo = cad_mat.loc[index, 'Texto breve material']
                            campo = campo[:32]
                            self.entradas_mat[i + 1].text = campo
                            self.entradas_mat[i + 3].text = cad_mat.loc[index, 'Ncm']
                            break

    def colar(self, instance):
        for i, l in enumerate(self.entradas_mat):
            print(f'{i} e {l.text}')
            if i % 8 == 0:
                if l.text != '' and self.entradas_mat[i + 1].text == '':
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

    def preenche_iva(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 8 == 0 and e != 0:
                if item.text != '':
                    self.entradas_mat[e + 2].text = self.entradas_mat[2].text
                    self.ids.check_iva.active = False

    def preenche_ncm(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 8 == 0 and e != 0:
                if item.text != '':
                    self.entradas_mat[e + 3].text = self.entradas_mat[3].text

    def preenche_aliq(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 8 == 0:
                if item.text != '':
                    self.entradas_mat[e + 4].text = '18%'
                    self.entradas_mat[e + 5].text = '15%'
                    self.entradas_mat[e + 6].text = '1,65%'
                    self.entradas_mat[e + 7].text = '7,6%'

    def limpar(self):
        for lin in self.entradas_mat:
            lin.text = ''

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
                self.serv = MDTextFieldRect(multiline=False, size_hint=(largura, .05), write_tab=False)
                self.entradas.append(self.serv)
                self.lista[c].append(self.serv)

                self.ids.grid_serv.add_widget(self.serv)

        for i, n in enumerate(self.entradas):
            if i % 3 == 0:
                self.entradas[i].bind(on_text_validate=self.colar_serv)
                self.entradas[i + 1].bind(focus=self.colar_serv2)

    def colar_serv2(self, instance, widget):
        serv_cad = pd.read_excel(self.pasta_principal + 'material.xlsx', sheet_name='servicos')
        serv_cad = pd.DataFrame(serv_cad)
        serv_cad['Nº de serviço'] = serv_cad['Nº de serviço'].astype(str)

        for i, l in enumerate(self.entradas):
            if l.text != '' and self.entradas[i + 1].text == '':
                self.posicao = int(i / 3) if i > 0 else i

                for index, row in serv_cad.iterrows():
                    if l.text == row['Nº de serviço']:
                        self.entradas[i + 1].text = serv_cad.loc[index, 'Denominação']
                        self.entradas[i + 2].text = str(int(serv_cad.loc[index, 'Classe avaliaç.']))
                        break

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

    def busca_servico(self):
        # self.path = 'G:\GECOT\Análise Contábil_Tributária_Licitações\\2022\\1Pendentes\\'
        data_serv = pd.read_excel(self.pasta_principal + 'material.xlsx', sheet_name='116', dtype=str)
        data_serv = pd.DataFrame(data_serv)
        self.descricao = []
        for index, row in data_serv.iterrows():
            if self.ids.cod_serv.text == row['servico']:
                self.descricao.append(row['servico'] + ' - ' + data_serv.loc[index, 'descricao'] + '\n')
                self.descricao.append('\n')
                self.descricao.append(data_serv.loc[index, 'obs'] + '\n')
                self.descricao.append('\n')
                self.descricao.append(data_serv.loc[index, 'irrf'] + '\n')
                self.descricao.append(data_serv.loc[index, 'crf'] + '\n')
                self.descricao.append(data_serv.loc[index, 'inss'] + '\n')
                self.descricao.append(data_serv.loc[index, 'iss'] + '\n')
        self.ids.serv.text = ''.join([str(item) for item in self.descricao])

    def clausulas(self):
        self.nomes = ['N/A', 'Minuta', 'Minuta 2', 'Redação', 'Redação 2', '2.3.7.', '2.3.7.1', '2.3.7.2', '2.3.7.3',
                      '2.3.7.4', '6.7.2', '15.1', '3.10.1', '3.9-10-11', 'Anexo 2']

        self.lista_check = []
        self.infos = []

        for i in range(15):

            self.checks = MDCheckbox(size_hint=(.05, .15))
            self.num_claus = MDLabel(text=self.nomes[i], size_hint=(.1, .15))
            self.clausulas = TextInput(size_hint=(.85, .3))

            self.ids.grid_clausulas.add_widget(self.checks)
            self.ids.grid_clausulas.add_widget(self.num_claus)
            self.ids.grid_clausulas.add_widget(self.clausulas)
            self.infos.append(self.clausulas)
            self.lista_check.append(self.checks)

        with open(self.pasta_principal + 'texto.txt', 'r', encoding='latin-1') as read_obj:
            csv_reader = read_obj.readlines()

            for index, row in enumerate(csv_reader):
                self.infos[index].text = row

    def salvar(self):
        # ============================== GUARDAR DADOS ========================================#
        # lista_nova = []
        # lista_entr = []
        # cont = 0
        # for i in self.entradas:
        #     lista_entr.append(i.text)
        #     cont += 1
        #     if cont == 3:
        #         lista2 = lista_entr.copy()
        #         lista_nova.extend([lista2])
        #
        #         lista_entr.clear()
        #         cont = 0
        #
        # lista_nova_mat = []
        # lista_entr_mat = []
        # cont = 0
        # for i in self.entradas_mat:
        #     lista_entr_mat.append(i.text)
        #     cont += 1
        #     if cont == 8:
        #         lista3 = lista_entr_mat.copy()
        #         lista_nova_mat.extend([lista3])
        #
        #         lista_entr_mat.clear()
        #         cont = 0
        # print(lista_nova_mat)
        # dados_salvos = []
        # dados_salvos.extend([datetime.now().strftime('%d/%m/%Y, %H:%M:%S'), self.ids.gere.text, self.ids.proc.text,
        #                      self.ids.req.text, self.ids.orcam.text, self.ids.objcust.text, self.ids.tipo1.text,
        #                      self.ids.tipo2.text,
        #                      self.ids.tipo3.text, self.ids.objeto.text, self.ids.valor.text, self.ids.complem.text,
        #                      lista_nova_mat, self.ids.linha_mat.text, self.ids.serv.text, self.ids.iva.text,
        #                      lista_nova, self.ids.linha_serv.text, self.ids.obs.text, self.ids.obs_serv.text,
        #                      self.ids.obs1.text, self.ids.obs2.text, [i.text for i in self.infos],
        #                      [i.text for i in self.lista_check]])
        #
        # with open(self.pasta_principal + "Base.txt",
        #           "rb+") as fp:  # Pickling
        #     pickle_list = []
        #     nova_lista = []
        #     while True:
        #         try:
        #             pickle_list.append(pickle.load(fp))
        #         except EOFError:
        #             break
        #     for i in pickle_list:
        #         if i[2] != dados_salvos[2]:
        #             nova_lista.append(i)
        # with open(self.pasta_principal + "Base.txt", "wb") as fp:
        #     nova_lista.append(dados_salvos)
        #     for lis in nova_lista:
        #         pickle.dump(lis, fp)

        # ============================== CRIAR PDF ============================================#
        self.pdf = FPDF(orientation='P', unit='mm', format='A4')
        self.pdf.add_page()

        self.pdf_w = 210
        self.pdf_h = 297
        self.pdf.rect(5.0, 5.0, 200.0, 280.0)
        self.pdf.rect(5.0, 5.0, 200.0, 20.0)

        self.pdf.image(self.pasta_principal + 'logo.jpg', x=7.0, y=7.0,
                       h=15.0, w=50.0)
        self.pdf.line(70.0, 5.0, 70.0, 25.0)

        self.pdf.set_font('Arial', 'B', 10)
        self.pdf.set_xy(75.0, 9.0)
        self.pdf.multi_cell(w=125, h=5,
                            txt='Análise Contábil e Tributária para Processos de Licitação e ou Contratação Direta')

        self.pdf.rect(5.0, 30.0, 200.0, 25.0)
        self.pdf.line(5.0, 40.0, 205.0, 40.0)
        self.pdf.line(88.0, 30.0, 88.0, 55.0)
        self.pdf.line(137.0, 30.0, 137.0, 40.0)

        # ===================================== INFORMAÇÕES INICIAIS ==================================#
        self.pdf.set_xy(10.0, 25.0)
        self.pdf.cell(w=40, h=20, txt='Gerência Contratante:')
        self.pdf.set_xy(50.0, 25.0)
        self.pdf.cell(w=40, h=20, txt=self.ids.gere.text)
        self.pdf.set_xy(90.0, 22.5)
        self.pdf.cell(w=40, h=20, txt='N° do Processo GECBS:')
        self.pdf.set_xy(90.0, 26.5)
        self.pdf.cell(w=40, h=20, txt=self.ids.proc.text)
        self.pdf.set_xy(142.0, 22.5)
        self.pdf.cell(w=40, h=20, txt='Requisição de Compras: ')
        self.pdf.set_xy(142.0, 26.5)
        self.pdf.cell(w=40, h=20, txt=self.ids.req.text)
        self.pdf.set_xy(10.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Objeto de Custos:')
        self.pdf.set_xy(43.0, 42.5)
        self.pdf.multi_cell(w=45, h=5, align='L', txt=self.ids.objcust.text)
        self.pdf.set_xy(90.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Consta no Orçamento? ')
        self.pdf.set_xy(148.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Sim')
        self.pdf.set_xy(175.0, 35.0)
        self.pdf.cell(w=40, h=20, txt='Não')
        if self.ids.orcam_sim.active is True:
            self.pdf.set_xy(138.0, 35.0)
            self.pdf.cell(w=40, h=20, txt='X')
        else:
            self.pdf.set_xy(169.0, 35.0)
            self.pdf.cell(w=40, h=20, txt='X')
        self.pdf.line(5.0, 66.0, 205.0, 66.0)
        self.pdf.set_xy(100.0, 51.0)
        self.pdf.cell(w=40, h=20, txt='ANÁLISE')
        self.pdf.line(5.0, 75.0, 205.0, 75.0)
        self.pdf.set_xy(30.0, 60.5)
        self.pdf.cell(w=40, h=20, txt='Serviço')
        self.pdf.set_xy(90.0, 60.5)
        self.pdf.cell(w=40, h=20, txt='Material')
        self.pdf.set_xy(135.0, 60.5)
        self.pdf.cell(w=40, h=20, txt='Serviço com Fornecimento de Material')
        if self.ids.check1.active is True:
            self.pdf.set_xy(25.0, 60.5)
            self.pdf.cell(w=40, h=20, txt='X')
        if self.ids.check2.active is True:
            self.pdf.set_xy(85.0, 60.5)
            self.pdf.cell(w=40, h=20, txt='X')
        if self.ids.check3.active is True:
            self.pdf.set_xy(130.0, 60.5)
            self.pdf.cell(w=40, h=20, txt='X')
        self.pdf.set_xy(10.0, 77.0)
        self.pdf.cell(w=40, h=5, txt='Objeto: ')
        self.pdf.set_xy(30.0, 77.0)
        self.pdf.set_font('')
        self.pdf.multi_cell(w=160, h=5, txt=self.ids.objeto.text)
        self.pdf.set_font('arial', 'B', 10)
        self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
        self.pdf.cell(w=40, h=5, txt='Valor estimado: ')
        self.pdf.set_xy(40.0, self.pdf.get_y())
        self.pdf.set_font('')
        self.pdf.cell(w=40, h=5, txt=self.ids.valor.text)
        self.pdf.set_font('arial', 'B', 10)
        self.pdf.set_xy(15.0, self.pdf.get_y() + 5)
        self.pdf.cell(w=40, h=5, txt=self.ids.complem.text)
        self.pdf.set_xy(15.0, self.pdf.get_y() + 10)

        self.pdf.set_auto_page_break(True, 20.0)
        # self.pdf.set_auto_page_break(True, 20.0)

        # ======================================= MATERIAIS =================================#
        if self.entradas_mat[0].text != '':
            self.pdf.set_y(self.pdf.get_y() + (float(self.ids.linha_mat.text) * 5))
            if int(self.ids.linha_mat.text) > 0:
                self.pdf.rect(5.0, 5.0, 200.0, 280.0)

            q1 = self.pdf.get_y()
            self.pdf.set_font('')
            self.pdf.multi_cell(w=180, h=5, txt='Informações Tributárias: ')
            self.pdf.set_xy(10, self.pdf.get_y() + 5)
            self.pdf.multi_cell(w=180, h=5, txt=' - Materiais serão acobertados por Nota Fiscal eletrônica modelo 55. ')
            self.pdf.set_font('Arial', 'B', 10)

            self.dados_faltantes = []
            self.pdf.set_xy(10, self.pdf.get_y() + 5)
            cont_lista = 0
            cont = 1
            px = 10
            py = self.pdf.get_y()

            self.data_mat = [['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI', 'PIS', 'COFINS']]
            self.data_mat[1:].clear()
            cont2 = 0
            mat_list = []
            for lin in self.entradas_mat:
                if lin.text != '':
                    mat_list.append(lin.text)
                    cont2 += 1
                    if cont2 == 8:
                        lista_nova_mat = mat_list.copy()
                        self.data_mat.append(lista_nova_mat)
                        mat_list.clear()
                        cont2 = 0

            for row in self.data_mat:
                for datum in row:
                    if cont == 1:
                        self.pdf.set_font('') if cont_lista != 0 else self.pdf.set_font('Arial', 'B', 10)
                        self.pdf.set_xy(px, py)
                        self.pdf.multi_cell(w=20, h=5, txt=datum, border=1)
                    elif cont == 2:
                        self.pdf.set_xy(px + 20, py)
                        self.pdf.multi_cell(w=75, h=5, txt=datum, border=1)
                    elif cont == 3:
                        self.pdf.set_xy(px + 95, py)
                        self.pdf.multi_cell(w=10, h=5, txt=datum, border=1)
                    elif cont == 4:
                        self.pdf.set_xy(px + 105, py)
                        self.pdf.multi_cell(w=20, h=5, txt=datum, border=1)
                    else:
                        self.pdf.set_xy(px + 125, py)
                        px += 16
                        self.pdf.multi_cell(w=16, h=5, txt=datum, border=1)
                    cont += 1
                px = 10
                py = self.pdf.get_y()
                cont = 1
                cont_lista += 1
                if py > 270:
                    self.dados_faltantes = self.data_mat[cont_lista:]
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 280.0)
                    break



            else:
                pass

            self.pdf.rect(7.5, q1 - 3, 195, self.pdf.get_y() - q1 + 7.5)
            self.pdf.set_xy(10.0, self.pdf.get_y() + 10)
            self.pdf.set_font('Arial', 'U', 10)
            self.pdf.multi_cell(w=180, h=5, txt=self.ids.obs.text)
            self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
            self.pdf.set_font('')
            self.pdf.rect(5.0, 5.0, 200.0, 280.0)
        # ======================================== SERVIÇOS =============================================#

        if self.ids.iva.text != '':

            self.pdf.set_y(self.pdf.get_y() + (float(self.ids.linha_serv.text) * 5))
            if int(self.ids.linha_serv.text) > 0:
                self.pdf.add_page()
                self.pdf.rect(5.0, 5.0, 200.0, 280.0)

            self.pdf.set_xy(10.0, self.pdf.get_y())
            q2 = self.pdf.get_y()
            self.pdf.multi_cell(w=180, h=5,
                                txt=' - Serviços serão acobertados por Nota Fiscal de serviço eletrônica. ')

            self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
            self.pdf.multi_cell(w=180, h=5, txt='O código de imposto (IVA) utilizado no pedido (SAP) '
                                                'deverá ser o ' + self.ids.iva.text + '.')

            self.data = [['DESCRIÇÃO', 'CÓDIGO', 'C.C']]
            self.data[1:].clear()
            cont2 = 0
            temp_list = []
            for lin in self.entradas:
                if lin.text != '':
                    temp_list.append(lin.text)

                    cont2 += 1
                    if cont2 == 3:
                        ordem = [1, 0, 2]
                        temp_list = [temp_list[i] for i in ordem]
                        lista_nova = temp_list.copy()
                        self.data.append(lista_nova)
                        temp_list.clear()
                        cont2 = 0

            self.dados_faltantes = []
            self.pdf.set_xy(10, self.pdf.get_y() + 5)
            cont_lista = 0
            cont = 3
            px = 10
            py = self.pdf.get_y()
            for row in self.data:
                for datum in row:
                    if cont % 3 == 0:
                        self.pdf.set_font('') if cont_lista != 0 else self.pdf.set_font('Arial', 'B', 10)
                        self.pdf.set_xy(px + 20, py)
                        self.pdf.multi_cell(w=150, h=5, txt=datum, border=1)
                    elif cont % 4 == 0:
                        atual = self.pdf.get_y() - py
                        self.pdf.set_xy(px, py)
                        self.pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
                    else:
                        atual = self.pdf.get_y() - py
                        self.pdf.set_xy(px + 170, py)
                        self.pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
                    cont += 1
                px = 10
                py = self.pdf.get_y()
                cont = 3
                cont_lista += 1
                if py > 270:
                    self.dados_faltantes = self.data[cont_lista:]
                    self.pdf.rect(7.5, q2 - 3, 195.0, self.pdf.get_y() - q2 + 7.5)
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 280.0)
                    break

            if self.dados_faltantes:
                q2 = self.pdf.get_y()
                self.dados_faltantes.insert(0, ['DESCRIÇÃO', 'CÓDIGO', 'C.C'])
                cont = 3
                px = 10
                py = self.pdf.get_y()
                for row in self.dados_faltantes:
                    for datum in row:
                        if cont % 3 == 0:
                            self.pdf.set_xy(px + 20, py)
                            self.pdf.multi_cell(w=150, h=5, txt=datum, border=1)
                        elif cont % 4 == 0:
                            atual = self.pdf.get_y() - py
                            self.pdf.set_xy(px, py)
                            self.pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
                        else:
                            atual = self.pdf.get_y() - py
                            self.pdf.set_xy(px + 170, py)
                            self.pdf.multi_cell(w=20, h=atual, txt=datum, border=1)
                        cont += 1
                    px = 10
                    py = self.pdf.get_y()
                    cont = 3

            else:
                pass

            self.pdf.set_xy(15.0, self.pdf.get_y() + 10)

            self.pdf.multi_cell(w=180, h=5, txt='Informações Tributárias: ')

            # self.pdf.set_font('Arial', 'B', 10)
            self.pdf.set_xy(15.0, self.pdf.get_y() + 5)
            self.pdf.set_font('Arial', 'B', 10)
            self.pdf.multi_cell(w=180, h=5, txt=self.ids.serv.text)
            self.pdf.rect(7.5, q2 - 3, 195.0, self.pdf.get_y() - q2 + 7.5)
            self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
            self.pdf.multi_cell(w=180, h=5, txt=self.ids.obs_serv.text)
            self.pdf.set_xy(10.0, self.pdf.get_y() + 5) if self.ids.obs_serv.text != '' else \
                self.pdf.set_xy(10.0, self.pdf.get_y())
            self.pdf.rect(5.0, 5.0, 200.0, 280.0)

        # ============================== OBSERVAÇÕES ===========================================#

        self.pdf.set_y(self.pdf.get_y() + (float(self.ids.linha_obs.text) * 5))
        if int(self.ids.linha_obs.text) > 0:
            self.pdf.rect(5.0, 5.0, 200.0, 280.0)

        self.pdf.rect(5.0, 5.0, 200.0, 280.0)
        self.pdf.set_font('')
        self.pdf.multi_cell(w=180, h=5, txt=self.ids.obs1.text)
        self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
        self.pdf.multi_cell(w=180, h=5, txt=self.ids.obs2.text)
        self.pdf.rect(5.0, 5.0, 200.0, 280.0)

        # =============================  INFORMAÇÕES CONTRATUAIS ===========================================#

        self.pdf.set_y(self.pdf.get_y() + (float(self.ids.linha_cont.text) * 10))
        self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
        self.pdf.cell(w=40, h=5, txt='Informações Contratuais : ')
        self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
        for i, item in enumerate(self.lista_check):
            if item.active is True:
                self.pdf.set_xy(15.0, self.pdf.get_y() + 5)
                self.pdf.multi_cell(w=180, h=5, align='L', txt=self.infos[i].text)
                if self.pdf.get_y() > 270:
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 285.0)

        if int(self.ids.linha_cont.text) > 0:
            self.pdf.rect(5.0, 5.0, 200.0, 280.0)


        # =============================== ASSINATURA E DATA =====================================#
        self.pdf.rect(5.0, 265.0, 200.0, 20.0)
        self.pdf.set_xy(10.0, 270.0)
        self.pdf.cell(w=40, txt='Responsável pela análise: ')
        self.pdf.line(80.0, 265.0, 80.0, 285.0)
        self.pdf.set_xy(90.0, 270.0)
        self.pdf.cell(w=40, txt='DATA: ' + date.today().strftime('%d/%m/%Y'))
        self.pdf.line(130.0, 265.0, 130.0, 285.0)
        self.pdf.set_xy(135.0, 270.0)
        self.pdf.cell(w=40, txt='Revisado pela Gerência: ')
        self.pdf.image(self.pasta_principal + 'Mari1.png', x=7.0, y=265.0, h=25.0, w=45.0)
        troca = self.ids.proc.text.replace('/', '-')
        self.pdf.output(self.pasta_principal + 'Análise Tributária - ' + troca + '.pdf', 'F')
        os.startfile(self.pasta_principal + 'Análise Tributária - ' + troca + '.pdf')


class WindowManager(ScreenManager):
    pass


class Example(MDApp):

    def build(self):
        self.items = [{'icon': 'Sim'}, {'icon': 'Não'}]
        return Builder.load_file('analisetribut.kv')


Example().run()
