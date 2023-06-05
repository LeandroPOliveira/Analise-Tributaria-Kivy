from datetime import datetime, date
from fpdf import FPDF
import getpass
import glob
from kivy.clock import Clock
from kivy.uix.textinput import TextInput
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
from kivymd.uix.dialog import MDDialog
from kivymd.uix.selectioncontrol import MDCheckbox
from kivymd.uix.textfield import MDTextFieldRect
from kivy.utils import get_color_from_hex
from kivy.core.window import Window
import math
import os
import pandas as pd
import pickle
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
import win32clipboard
import win32com.client as win32


class TelaLogin(Screen):
    pass


class AnalisesPendentes(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dialog_permissao_assinat = None
        self.dialog = None
        self.arquivos_assinatura = []
        self.arquivos_pdf = []
        self.tabela_pendentes = None

        # Criar a marca d'agua com a assinatura
        c = canvas.Canvas('watermark.pdf')
        # Desenhar a imagem na posição x e y.
        c.drawImage('assinatura.png', 440, 30, 100, 60, mask='auto')
        c.save()
        # Buscar o arquivo da marca d'agua criado
        self.watermark = PdfFileReader(open(os.path.join('watermark.pdf'), 'rb'))

        # Abrir arquivo contendo a pasta de trabalho e os emails dos usuarios
        with open('dados.txt', 'r', encoding='UTF-8') as bd:
            self.dados = bd.readlines()
            self.diretorio = self.dados[0].rstrip('\n')
        self.diretorio = os.getcwd()

    def add_datatable(self):  # Adicionar tabela com as análises pendentes
        self.arquivos_assinatura.clear()
        self.arquivos_pdf.clear()
        self.arquivos_diretorio = os.listdir(self.diretorio)
        for item in self.arquivos_diretorio:  # Selecionar arquivos pdf para mostrar na tabela
            if item.endswith('.pdf') is True and item != 'watermark.pdf':
                dt_modificacao = os.path.getctime(os.path.join(self.diretorio, item))
                dt_modificacao = datetime.fromtimestamp(dt_modificacao)
                data = date.strftime(dt_modificacao, '%d/%m/%Y')
                self.arquivos_pdf.append((item, data))
        if len(self.arquivos_pdf) == 1:
            self.arquivos_pdf.append(('', ''))

        self.tabela_pendentes = MDDataTable(pos_hint={'center_x': 0.5, 'center_y': 0.5},
                                            size_hint=(0.4, 0.55),
                                            check=True, use_pagination=True, rows_num=10,
                                            background_color_header=get_color_from_hex("#0d7028"),
                                            column_data=[("[color=#ffffff]Análise[/color]", dp(100)),
                                                         ("[color=#ffffff]Data[/color]", dp(30))],
                                            row_data=self.arquivos_pdf, elevation=1)

        self.add_widget(self.tabela_pendentes)

        self.tabela_pendentes.bind(on_row_press=self.abrir_pdf)
        self.tabela_pendentes.bind(on_check_press=self.marcar_pdf)

    def marcar_pdf(self, instance_row, current_row):  # Marcar arquivos para assinar
        self.arquivos_assinatura.append(current_row[0])

    def abrir_pdf(self, instance_table, current_row):  # Abrir pdf
        if self.tabela_pendentes.get_row_checks():
            pass
        else:
            try:
                os.startfile(os.path.join(self.diretorio, current_row.text))
            except FileNotFoundError:
                self.dialog = MDDialog(text="Clique sobre o texto Análise Tributária...!", radius=[20, 7, 20, 7], )
                self.dialog.open()

    def assinatura(self):  # Assinar arquivos selecionados
        self.salvos = []
        for n, arquivo in enumerate(self.arquivos_assinatura):
            os.chdir(self.diretorio)
            self.output_file = PdfFileWriter()
            with open(arquivo, 'rb') as f:
                input_file = PdfFileReader(f)
                # Número de páginas do documento
                page_count = input_file.getNumPages()
                # Percorrer o arquivo para adicionar a marca d'agua
                for page_number in range(page_count):
                    input_page = input_file.getPage(page_number)
                    if page_number == page_count - 1:
                        input_page.mergePage(self.watermark.getPage(0))
                    self.output_file.addPage(input_page)

                self.dir_acima = self.diretorio.split('\\')
                self.dir_acima.insert(1, '\\')
                self.dir_acima = os.path.join(*self.dir_acima[:-1])
                os.chdir(self.dir_acima)
                file = glob.glob(str(arquivo[21:29]) + '*')
                pasta_analise = ''.join(file)
                try:
                    os.chdir(pasta_analise)
                except OSError:
                    os.chdir(self.dir_acima)

                # Gerar o novo arquivo pdf assinado
                with open('Análise Tributária - ' + str(arquivo[21:]), "wb") as outputStream:
                    self.output_file.write(outputStream)

            os.chdir(self.diretorio)
            try:
                os.remove(arquivo)
            except PermissionError:
                self.dialog_permissao_assinat = MDDialog(text='Erro! Arquivo em uso.', radius=[20, 7, 20, 7], )
                self.dialog_permissao_assinat.open()
            self.salvos.append(n)

        troca = 0
        for i in self.salvos:
            self.arquivos_pdf.pop(i - troca)
            troca += 1

        outlook = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook.CreateItem(0)
        # configurar as informações do e-mail e selecionar o endereço pelo arquivo de texto
        email.To = self.dados[2]
        email.Subject = "E-mail automático Análise Tributária"
        email.HTMLBody = f"""
                    <p>Análise(s) Tributária(s) assinada(s) com sucesso.</p>

                    """
        email.Send()

        self.dialog = MDDialog(text='Análise(s) assinada(s) com sucesso!', radius=[20, 7, 20, 7], )
        self.dialog.open()

        self.add_datatable()  # Atualizar lista de análises


class CarregarAnalise(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.lista_analises = []  # Recebe dados do arquivo de texto com analises salvas
        self.temp_list = []  # Recebe dados brutos do pickle
        self.dados_tabela = None  # Criar Tabela para exibição

    def carregar_dados(self):
        self.temp_list.clear(), self.lista_analises.clear()
        with open(os.path.join(self.manager.get_screen('pendentes').diretorio, 'Base.txt'), "rb") as carga:
            while True:
                try:
                    self.temp_list.append(pickle.load(carga))
                except EOFError:
                    break
        for n, item in enumerate(self.temp_list):
            itens = (item[2], item[0])
            self.lista_analises.append(itens)
            self.lista_analises.sort(key=lambda lista: datetime.strptime(lista[1], '%d/%m/%Y, %H:%M:%S'), reverse=True)
        if len(self.lista_analises) == 1:
            self.lista_analises.append(('', ''))

        self.dados_tabela = MDDataTable(pos_hint={'center_x': 0.5, 'center_y': 0.5},
                                        size_hint=(0.6, 0.8), rows_num=20,
                                        use_pagination=True,
                                        background_color_header=get_color_from_hex("#0d7028"),
                                        check=True,
                                        column_data=[("[color=#ffffff]Análise[/color]", dp(70)),
                                                     ("[color=#ffffff]Data[/color]", dp(70))],
                                        row_data=self.lista_analises, elevation=1)

        self.add_widget(self.dados_tabela)

        self.dados_tabela.bind(on_row_press=self.abrir_dados)

    def abrir_dados(self, instance_table, current_row):  # Pegar informações do txt e enviar para os inputs
        verinfo3 = int(current_row.index / 2)
        self.temp_list.sort(key=lambda lista: datetime.strptime(lista[0], '%d/%m/%Y, %H:%M:%S'), reverse=True)
        self.manager.get_screen("nova").ids.gere.text = self.temp_list[int(verinfo3)][1]
        self.manager.get_screen("nova").ids.proc.text = self.temp_list[int(verinfo3)][2]
        self.manager.get_screen("nova").ids.req.text = self.temp_list[int(verinfo3)][3]
        self.manager.get_screen("nova").ids.orcam_sim.state = 'down' if self.temp_list[int(verinfo3)][
                                                                            4] == 'down' else 'normal'
        self.manager.get_screen("nova").ids.orcam_nao.state = 'normal' if self.temp_list[int(verinfo3)][
                                                                              4] == 'down' else 'down'
        self.manager.get_screen("nova").ids.objcust.text = self.temp_list[int(verinfo3)][5]
        self.manager.get_screen("nova").ids.check1.active = self.temp_list[int(verinfo3)][6]
        self.manager.get_screen("nova").ids.check2.active = self.temp_list[int(verinfo3)][7]
        self.manager.get_screen("nova").ids.check3.active = self.temp_list[int(verinfo3)][8]
        self.manager.get_screen("nova").ids.objeto.text = self.temp_list[int(verinfo3)][9].strip()
        self.manager.get_screen("nova").ids.valor.text = self.temp_list[int(verinfo3)][10]
        self.manager.get_screen("nova").ids.complem.text = self.temp_list[int(verinfo3)][11].strip()
        for r, val in enumerate(self.temp_list[int(verinfo3)][12]):
            for b, value in enumerate(val):
                self.manager.get_screen("nova").lista_mat[b][r].text = value

        self.manager.get_screen("nova").ids.linha_mat.text = self.temp_list[int(verinfo3)][13]
        self.manager.get_screen("nova").ids.serv.text = self.temp_list[int(verinfo3)][14].strip()
        self.manager.get_screen("nova").ids.iva.text = self.temp_list[int(verinfo3)][15]
        for r, val in enumerate(self.temp_list[int(verinfo3)][16]):
            for b, value in enumerate(val):
                self.manager.get_screen("nova").lista_serv[b][r].text = value
        self.manager.get_screen("nova").ids.linha_serv.text = self.temp_list[int(verinfo3)][17]
        self.manager.get_screen("nova").ids.obs.text = self.temp_list[int(verinfo3)][18].strip().strip()
        self.manager.get_screen("nova").ids.obs_serv.text = self.temp_list[int(verinfo3)][19].strip()
        self.manager.get_screen("nova").ids.obs1.text = self.temp_list[int(verinfo3)][20].strip()
        self.manager.get_screen("nova").ids.obs2.text = self.temp_list[int(verinfo3)][21].strip()
        for n, i in enumerate(self.temp_list[int(verinfo3)][22]):
            self.manager.get_screen("nova").infos[n].text = i.strip()
        for n, i in enumerate(self.temp_list[int(verinfo3)][23]):
            if i == 'down':
                self.manager.get_screen("nova").lista_check[n].state = 'down'
        self.manager.get_screen("nova").ids.linha_cont.text = self.temp_list[int(verinfo3)][24]
        self.manager.get_screen("nova").ids.linha_obs.text = self.temp_list[int(verinfo3)][25]
        self.manager.current = 'nova'


class NovaAnalise(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dialog_permissao = None
        self.lista_mat = [[], [], [], [], [], [], [], []]  # Lista para trabalhar com 8 posições de colunas de materiais
        self.entradas_mat = []  # Recebe os inputs criados para os dados de materiais
        self.lista_serv = [[], [], []]  # Lista para trabalhar com as 3 posições de colunas de serviços
        self.entradas = []  # Recebe os inputs criados para os dados de servicos
        self.posicao = []  # Identifica qual linha da tabela os dados serão alocados para materiais e serviços
        self.lista_check = []  # Guardar o status dos botões de check das cláusulas
        self.infos = []  # Carregar informações das cláusulas de contrato
        self.data_mat = []  # Lista para incluir dados dos materiais no pdf
        self.data = []  # Lista para incluir dados de serviços no pdf
        self.dados_cadastro = 'cadastro - exemplo.xlsx'

        Clock.schedule_once(self.cria_tabela_materiais)
        Clock.schedule_once(self.cria_tabela_servicos)
        Clock.schedule_once(self.clausulas)
        Clock.schedule_once(self.informacoes_padrao)

    def informacoes_padrao(self, dt):
        self.ids.obs.text = 'Produto para Consumo Final. \nFabricante: Alíquota de ICMS de 18% conforme RICMS-SP/2000,'\
                            ' Livro I, Título III, Capítulo II, Seção II, Artigo 52, Inciso I \nRevendedor: ' \
                            'Informar o ICMS-ST  recolhido anteriormente\n\nEntrega de materiais em local diverso do ' \
                            'destinatário: o endereço deverá constar na nota fiscal em campo específico do xml ' \
                            '(bloco G) e em dados adicionais. (Regime Especial 28558/2018).'

        self.ids.obs1.text = 'Obs 1: Caso o fornecedor possua alguma especificidade que implique tratamento tributário'\
                             ' diverso do exposto acima, ou seja do regime tributário "SIMPLES NACIONAL" deverá' \
                             '  apresentar documentação hábil que comprove sua condição peculiar, a qual será alvo de' \
                             ' análise prévia pela GECOT.'

        self.ids.obs2.text = 'Obs 2: Essa Análise não é exaustiva, podendo sofrer alterações no decorrer do processo ' \
                             'de contratação em relação ao produto/serviço.'

    def cria_tabela_materiais(self, dt):
        for i in range(61):
            for c in range(6):
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

                self.manager.get_screen("nova").ids.grid_teste.add_widget(self.mater)

        for i, n in enumerate(self.entradas_mat):
            if i % 6 == 0:
                self.entradas_mat[i].bind(on_text_validate=self.busca_dados_mat_clipboard)
                self.entradas_mat[i + 1].bind(focus=self.busca_dados_mat)

    def busca_dados_mat(self, instance, widget):
        cad_mat = pd.read_excel(os.path.join(self.manager.get_screen("pendentes").diretorio, self.dados_cadastro),
                                sheet_name='materiais', converters={'Material': str, 'IPI': str})
        cad_mat = pd.DataFrame(cad_mat)

        for i, l in enumerate(self.entradas_mat):
            if i % 6 == 0:
                if l.text != '' and self.entradas_mat[i + 1].text == '':
                    for index, row in cad_mat.iterrows():
                        if l.text == row['Material']:
                            campo = cad_mat.loc[index, 'Texto breve material']
                            campo = campo[:32]
                            self.entradas_mat[i + 1].text = campo
                            self.entradas_mat[i + 3].text = cad_mat.loc[index, 'Ncm']
                            self.entradas_mat[i + 5].text = cad_mat.loc[index, 'IPI']
                            break

    def busca_dados_mat_clipboard(self, instance):
        for i, l in enumerate(self.entradas_mat):
            if i % 6 == 0:
                if l.text != '' and self.entradas_mat[i + 1].text == '':
                    self.posicao = int(i / 6) if i > 0 else i
                    break

        cad_mat = pd.read_excel(os.path.join(self.manager.get_screen("pendentes").diretorio, self.dados_cadastro),
                                sheet_name='materiais', converters={'Material': str, 'IPI': str})
        cad_mat['Material'] = cad_mat['Material'].astype(str)
        win32clipboard.OpenClipboard()
        try:
            rows = win32clipboard.GetClipboardData()
        except TypeError:
            rows = ''
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        rows = rows.split('\n')

        rows.pop() if len(rows) > 1 else rows
        for r, val in enumerate(rows):
            if rows[0] == '':
                break
            values = val.split('\t')
            if len(values) > 1:
                del values[1:]
            for b, value in enumerate(values):
                for index, row in cad_mat.iterrows():
                    self.lista_mat[b][r + self.posicao].text = value
                    if value.strip() == row['Material']:
                        campo = cad_mat.loc[index, 'Texto breve material']
                        campo = campo[:32]
                        self.lista_mat[b + 1][r + self.posicao].text = campo
                        self.lista_mat[b + 3][r + self.posicao].text = cad_mat.loc[index, 'Ncm']
                        self.lista_mat[b + 5][r + self.posicao].text = cad_mat.loc[index, 'IPI']

    def preenche_iva(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 6 == 0 and e != 0:
                if item.text != '':
                    self.entradas_mat[e + 2].text = self.entradas_mat[2].text
                    self.ids.check_iva.active = False

    def preenche_ncm(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 6 == 0 and e != 0:
                if item.text != '':
                    self.entradas_mat[e + 3].text = self.entradas_mat[3].text

    def preenche_aliq(self):
        for e, item in enumerate(self.entradas_mat):
            if e % 6 == 0:
                if item.text != '':
                    self.entradas_mat[e + 4].text = '18%'
                    # self.entradas_mat[e + 6].text = '1,65%'
                    # self.entradas_mat[e + 7].text = '7,6%'

    def limpa_dados_mat(self):
        for lin in self.entradas_mat:
            lin.text = ''

    def limpa_dados_serv(self):
        for lin in self.entradas:
            lin.text = ''

    def cria_tabela_servicos(self, dt):
        for i in range(90):
            for c in range(3):
                if c == 1:
                    largura = 30
                else:
                    largura = 15
                self.serv = MDTextFieldRect(multiline=False, size_hint=(largura, .05), write_tab=False)
                self.entradas.append(self.serv)
                self.lista_serv[c].append(self.serv)
                self.ids.grid_serv.add_widget(self.serv)

        for i, n in enumerate(self.entradas):
            if i % 3 == 0:
                self.entradas[i].bind(on_text_validate=self.busca_dados_serv_clipboard)
                self.entradas[i + 1].bind(focus=self.busca_dados_serv)

    def busca_dados_serv(self, instance, widget):
        serv_cad = pd.read_excel(os.path.join(self.manager.get_screen("pendentes").diretorio, self.dados_cadastro),
                                 sheet_name='servicos')
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

    def busca_dados_serv_clipboard(self, instance):
        for i, l in enumerate(self.entradas):
            if l.text == '':
                self.posicao = int(i / 3) if i > 0 else i
                break
        serv_cad = pd.read_excel(os.path.join(self.manager.get_screen("pendentes").diretorio, self.dados_cadastro),
                                 sheet_name='servicos')
        serv_cad = pd.DataFrame(serv_cad)
        serv_cad['Nº de serviço'] = serv_cad['Nº de serviço'].astype(str)
        win32clipboard.OpenClipboard()
        try:
            rows = win32clipboard.GetClipboardData()
        except TypeError:
            rows = ''
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
                    self.lista_serv[b][r + self.posicao].text = value.strip()
                    if value.strip() == row['Nº de serviço']:
                        self.lista_serv[b + 1][r + self.posicao].text = serv_cad.loc[index, 'Denominação']
                        self.lista_serv[b + 2][r + self.posicao].text = str(int(serv_cad.loc[index, 'Classe avaliaç.']))

    def busca_dados_lei_116(self):
        if self.ids.cod_serv.text != '':
            data_serv = pd.read_excel(os.path.join(self.manager.get_screen("pendentes").diretorio, self.dados_cadastro),
                                      sheet_name='116', dtype=str)
            data_serv = pd.DataFrame(data_serv)
            descricao = []
            for index, row in data_serv.iterrows():
                if self.ids.cod_serv.text == row['servico']:
                    descricao.append(row['servico'] + ' - ' + data_serv.loc[index, 'descricao'] + '\n')
                    descricao.append('\n')
                    descricao.append(data_serv.loc[index, 'obs'] + '\n')
                    descricao.append('\n')
                    descricao.append(data_serv.loc[index, 'irrf'] + '\n')
                    descricao.append(data_serv.loc[index, 'crf'] + '\n')
                    descricao.append(data_serv.loc[index, 'inss'] + '\n')
                    descricao.append(data_serv.loc[index, 'iss'] + '\n')
            self.ids.serv.text = ''.join([str(item) for item in descricao])

    def clausulas(self, dt):
        with open('texto.txt', 'r', encoding='latin-1') as read_obj:
            texto_clausulas = read_obj.readlines()

            for i in range(15):
                self.checks = MDCheckbox(size_hint=(.05, .15))
                self.clausulas = TextInput(size_hint=(.85, .3), multiline=True)
                self.ids.grid_clausulas.add_widget(self.checks)
                self.ids.grid_clausulas.add_widget(self.clausulas)
                self.infos.append(self.clausulas)
                self.lista_check.append(self.checks)

            for index, row in enumerate(texto_clausulas):
                self.infos[index].text = row

    def salvar(self):
        # ============================== GUARDAR DADOS ========================================#
        lista_nova = []
        lista_entr = []
        cont = 0
        for i in self.entradas:
            lista_entr.append(i.text)
            cont += 1
            if cont == 3:
                lista2 = lista_entr.copy()
                lista_nova.extend([lista2])

                lista_entr.clear()
                cont = 0

        lista_nova_mat = []
        lista_entr_mat = []
        cont = 0
        for i in self.entradas_mat:
            lista_entr_mat.append(i.text)
            cont += 1
            if cont == 6:
                lista3 = lista_entr_mat.copy()
                lista_nova_mat.extend([lista3])

                lista_entr_mat.clear()
                cont = 0

        dados_salvos = []
        dados_salvos.extend([datetime.now().strftime('%d/%m/%Y, %H:%M:%S'), self.ids.gere.text, self.ids.proc.text,
                             self.ids.req.text, self.ids.orcam_sim.state, self.ids.objcust.text, self.ids.check1.active,
                             self.ids.check2.active,
                             self.ids.check3.active, self.ids.objeto.text, self.ids.valor.text, self.ids.complem.text,
                             lista_nova_mat, self.ids.linha_mat.text, self.ids.serv.text, self.ids.iva.text,
                             lista_nova, self.ids.linha_serv.text, self.ids.obs.text, self.ids.obs_serv.text,
                             self.ids.obs1.text, self.ids.obs2.text, [i.text for i in self.infos],
                             [i.state for i in self.lista_check], self.ids.linha_cont.text, self.ids.linha_obs.text,
                             getpass.getuser()])

        with open(os.path.join(self.manager.get_screen('pendentes').diretorio, "Base.txt"),
                  "rb+") as fp:  # Pickling
            pickle_list = []
            nova_lista = []
            while True:
                try:
                    pickle_list.append(pickle.load(fp))
                except EOFError:
                    break
            for i in pickle_list:
                if i[2] != dados_salvos[2]:
                    nova_lista.append(i)
        with open(os.path.join(self.manager.get_screen('pendentes').diretorio, "Base.txt"), "wb") as fp:
            nova_lista.append(dados_salvos)
            for lis in nova_lista:
                pickle.dump(lis, fp)

        # ============================== CRIAR PDF ============================================#
        self.pdf = FPDF(orientation='P', unit='mm', format='A4')
        self.pdf.add_page()

        self.pdf_w = 210
        self.pdf_h = 297
        self.pdf.rect(5.0, 5.0, 200.0, 280.0)
        self.pdf.rect(5.0, 5.0, 200.0, 20.0)

        self.pdf.image('logo.jpg', x=7.0, y=7.0,
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
        self.pdf.cell(w=40, h=20, txt=self.ids.gere.text.encode('latin-1', 'ignore').decode("latin-1"))
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
        if self.ids.complem.text != "":
            self.pdf.set_xy(10.0, self.pdf.get_y() + 10)
            self.pdf.cell(w=40, h=5, txt='Observações: ')
        self.pdf.set_x(40.0)
        self.pdf.multi_cell(w=160, h=5, txt=self.ids.complem.text)
        self.pdf.set_xy(15.0, self.pdf.get_y() + 10)
        self.pdf.set_auto_page_break(True, 20.0)

        # ======================================= MATERIAIS =================================#

        # Verificar a quantidade de linhas necessárias para o quadro completo de materiais
        cont = 0
        for line in self.manager.get_screen("nova").ids.serv.text:
            if "\n" in line:
                cont += 1

        if self.entradas_mat[0].text != '':
            if self.ids.linha_mat.text != "":
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

            self.data_mat = [['CÓDIGO', 'DESCRIÇÃO', 'IVA', 'NCM', 'ICMS', 'IPI']]
            self.data_mat[1:].clear()
            cont2 = 0
            mat_list = []
            for lin in self.entradas_mat:
                if lin.text != '':
                    mat_list.append(lin.text)
                    cont2 += 1
                    if cont2 == 6:
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
                        self.pdf.multi_cell(w=75, h=5, txt=datum[:35], border=1)
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
            if self.ids.linha_serv.text != "":
                self.pdf.set_y(self.pdf.get_y() + (float(self.ids.linha_serv.text) * 5))
                if int(self.ids.linha_serv.text) > 0:
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
            self.pdf.set_xy(15.0, self.pdf.get_y() + 5)
            self.pdf.set_font('Arial', 'B', 10)
            self.pdf.multi_cell(w=180, h=5, txt=self.ids.serv.text.encode('latin-1', 'ignore').decode("latin-1"))
            self.pdf.rect(7.5, q2 - 3, 195.0, self.pdf.get_y() - q2 + 7.5)
            self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
            self.pdf.multi_cell(w=180, h=5, txt=self.ids.obs_serv.text)
            self.pdf.set_xy(10.0, self.pdf.get_y() + 5) if self.ids.obs_serv.text != '' else \
                self.pdf.set_xy(10.0, self.pdf.get_y())
            self.pdf.rect(5.0, 5.0, 200.0, 280.0)

        # ============================== OBSERVAÇÕES ===========================================#
        if self.ids.linha_obs.text != "":
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
        bd_cont = 0
        if self.ids.linha_cont.text != "":
            self.pdf.set_y(self.pdf.get_y() + (float(self.ids.linha_cont.text) * 10))
        rel_check = []
        [rel_check.append(i.active) for i in self.lista_check]

        if True in rel_check:
            self.pdf.set_xy(10.0, self.pdf.get_y() + 7.5)

            self.pdf.cell(w=40, h=5, txt='Informações Contratuais : ')
            bd_cont = self.pdf.get_y()
        self.pdf.rect(5.0, 5.0, 200.0, 280.0)
        self.pdf.set_xy(10.0, self.pdf.get_y() + 5)
        for i, item in enumerate(self.lista_check):
            if item.active is True:
                if self.pdf.get_y() + math.ceil(len(self.infos[i].text.encode('latin-1', 'ignore').decode("latin-1")) /
                                                105) * 5 > 270:
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 280.0)
                    self.pdf.set_xy(15.0, self.pdf.get_y() + 5)
                    self.pdf.multi_cell(w=180, h=5, align='L', txt=self.infos[i].text.encode('latin-1', 'ignore').
                                        decode("latin-1"))
                else:
                    self.pdf.set_xy(15.0, self.pdf.get_y() + 5)
                    self.pdf.multi_cell(w=180, h=5, align='L', txt=self.infos[i].text.encode('latin-1', 'ignore').
                                        decode("latin-1"))
                if self.pdf.get_y() > 270:
                    self.pdf.add_page()
                    self.pdf.rect(5.0, 5.0, 200.0, 280.0)

        if self.ids.linha_cont.text != "":
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
        self.pdf.image('assinatura.png', x=7.0, y=265.0, h=25.0, w=45.0)
        troca = self.ids.proc.text.replace('/', '-')
        nome_arquivo = 'Análise Tributária - ' + troca + '.pdf'
        try:
            self.pdf.output(os.path.join(self.manager.get_screen("pendentes").diretorio, nome_arquivo), 'F')
            os.startfile(os.path.join(self.manager.get_screen("pendentes").diretorio, nome_arquivo))
        except PermissionError:
            self.dialog_permissao = MDDialog(text='Erro! Análise em uso.', radius=[20, 7, 20, 7], )
            self.dialog_permissao.open()

    def enviar_email(self):
        outlook = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook.CreateItem(0)
        # configurar as informações do seu e-mail
        email.To = self.manager.get_screen("pendentes").dados[1]
        email.Subject = "E-mail automático Análise Tributária"
        email.HTMLBody = f"""
        <p>Análise Tributária {self.ids.proc.text} está disponível para assinatura.</p>

        """
        email.Send()
        self.dialog = MDDialog(text='Análise enviada com sucesso!', radius=[20, 7, 20, 7], )
        self.dialog.open()


class WindowManager(ScreenManager):
    pass


class Example(MDApp):
    Window.maximize()
    tamanho_tela = Window.size

    def build(self):
        return Builder.load_file('analisetribut.kv')


Example().run()
