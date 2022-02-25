from kivy.config import Config
from kivy.uix.button import Button
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
import os
from datetime import datetime, date
from kivy.utils import get_color_from_hex
# from kivy.core.window import Window
# Window.size = (1280, 720)

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

    def gerar_analise(self):
        print(self.ids.gere.text)

class WindowManager(ScreenManager):
    pass


class Example(MDApp):

    def build(self):
        self.items = [{'icon': 'Sim'}, {'icon': 'Não'}]
        return Builder.load_file('analisetribut.kv')


Example().run()