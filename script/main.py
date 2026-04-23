
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import random
import openpyxl
import et_xmlfile

from kivy.app import App
from kivy.properties import ObjectProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.core.window import Window
from configparser import ConfigParser
from kivy.garden.matplotlib.backend_kivyagg import FigureCanvasKivyAgg
from kivy.clock import Clock

import matplotlib.pyplot as plt
from kivy.garden.qrcode import QRCodeWidget

from kivy.utils import platform
import os


import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

Window.size = (640, 360)
Window.clearcolor = (.2,.1,0,1)

#Graph Backgrounds
plt.rcParams['axes.facecolor']=(202/255, 158/255, 103/255, 1)
plt.rcParams['figure.facecolor']=(202/255, 158/255, 103/255, 1)
plt.rcParams["figure.autolayout"] = True

# import config
config = ConfigParser()
config.read("config.ini")

# import quotes
quotes_eng = open('quotes_eng.txt').read().splitlines()

# Define different Screens

class WindowManager(ScreenManager):
    pass


class MainWindow(Screen):

    def __init__(self, **kwargs):
        super(MainWindow, self).__init__(**kwargs)
        Window.update_viewport()

    user_select = StringProperty('')

    def user_selected(self, instance):
        self.manager.transition = SlideTransition(direction='left')
        self.manager.current = "user_window"
        self.user_select = instance.text

    def go_to_menu(self,*args):
        self.manager.transition = SlideTransition(direction='up')
        self.manager.current = "menu_window"

    def on_enter(self, *args):
        self.clear_widgets()
        #Create BoxLayout
        main_box = BoxLayout(orientation="vertical")
        #First widget: Header as a new box Layout in horizontal orientation
        main_box.header = BoxLayout(orientation="horizontal", size_hint=(1,.5), padding=(0,5))
        #Widgets of header box. Header text + options button in top right corner
        # main_box.header.text = Label(text="Welcome. Select your profile", size_hint=(.9,1), font_size=40)
        main_box.header.text = Label(text="", size_hint=(.15,1))
        main_box.header.image = Image(source="images/main.png")
        #Another BoxLayout in header box to place options button. This time vertical
        main_box.header.opt = BoxLayout(orientation="vertical", size_hint=(.15,1), padding=(0,10))

        main_box.header.opt.btn = Button(background_normal = 'images/settings.png', background_down = 'images/settings_p.png', size_hint=(None,None), height="40dp", width="40dp", border = (0,0,0,0))
        main_box.header.opt.btn.bind(on_press=self.go_to_menu)
        main_box.header.opt.empty = Label(text="", size_hint=(1,.3))
        main_box.header.opt.add_widget(main_box.header.opt.btn)
        main_box.header.opt.add_widget(main_box.header.opt.empty)

        main_box.header.add_widget(main_box.header.text)
        main_box.header.add_widget(main_box.header.image)
        main_box.header.add_widget(main_box.header.opt)

        #Second widget: User profile buttons, as a GridLayout. Should sit in a separate box of the screen.
        main_box.sub_box = GridLayout(cols=3, padding=(20,20))
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        df = df.reindex(sorted(df.columns), axis=1)
        main_box.sub_box.my_buttons = []
        main_box.sub_box.users = list(df)
        for user_name in main_box.sub_box.users:
            main_box.sub_box.button = Button(text=user_name, font_size="25dp", background_normal="images/button.png",background_down ="images/button_p.png", color=(104/255,60/255,17/255,1), border=(0,0,0,0))
            main_box.sub_box.button.bind(on_press=self.user_selected)
            main_box.sub_box.my_buttons.append(main_box.sub_box.button)
            main_box.sub_box.add_widget(main_box.sub_box.button)

        #Place the widgets to the screen in correct order (important for BoxLayout)
        main_box.add_widget(main_box.header)
        main_box.add_widget(main_box.sub_box)

        #Add the screen
        self.add_widget(main_box)


class UserWindow(Screen):
    def __init__(self, **kwargs):
        super(UserWindow, self).__init__(**kwargs)

    username = StringProperty('')
    date = StringProperty('')
    current_day = StringProperty('')
    last_month = StringProperty('')
    coffee_sum_lm = ObjectProperty()
    counts = ObjectProperty()
    stats_count = ObjectProperty()
    addition = ObjectProperty()
    stats_addition = ObjectProperty()
    substract = ObjectProperty()
    stats_substract = ObjectProperty()
    pay_history_list = []
    quote_input = StringProperty()
    user_warning = StringProperty()

    def timeout(self, *args):
        App.get_running_app().root.current = 'main_window'
        App.get_running_app().root.transition.direction = 'left'

    def reset_timeout(self, *args):
        Clock.unschedule(self.timeout)
        Clock.schedule_once(self.timeout, 30)

    def on_enter(self, *args):
        Clock.schedule_once(self.timeout, 30)
        self.username = self.manager.get_screen("main_window").user_select
        self.date = str(datetime.now().strftime("%B")) + " " + str(datetime.now().strftime("%Y"))
        self.current_day = str(datetime.now().strftime("%d"))
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        self.counts = df.at[self.date, self.username]

        df3 = pd.read_excel("pay_history.xlsx",engine="openpyxl", index_col=0, header=0)
        df3 = df3[df3[self.username].notna()]
        self.pay_history_list = df3[self.username].tolist()
        self.pay_history_list = [str(x) for x in self.pay_history_list]
        self.ids.history_spinner.values = self.pay_history_list
        self.check_enable()

    def add_coffee(self, *args):
        #add coffee count to spreadsheet used for payment
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        self.counts = df.at[self.date, self.username]
        self.addition = self.counts + 1
        df.at[self.date, self.username] = self.addition
        df.to_excel("count.xlsx")
        self.counts = df.at[self.date, self.username]

        #add coffee also to stats spreadsheet used for statistics and rankings
        df2 = pd.read_excel("stats.xlsx",engine="openpyxl", index_col=0, header=0)
        self.stats_count = df2.at[self.date, self.username]
        self.stats_addition = self.stats_count + 1
        df2.at[self.date, self.username] = self.stats_addition
        df2.to_excel("stats.xlsx")
        self.stats_count = df2.at[self.date, self.username]


    def minus_coffee(self, *args):
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        self.counts = df.at[self.date, self.username]
        self.substract = self.counts - 1
        df.at[self.date, self.username] = self.substract
        df.to_excel("count.xlsx")
        self.counts = df.at[self.date, self.username]


        # substract coffee also from stats spreadsheet used for statistics and rankings
        df2 = pd.read_excel("stats.xlsx",engine="openpyxl", index_col=0, header=0)
        self.stats_count = df2.at[self.date, self.username]
        self.stats_substract = self.stats_count - 1
        df2.at[self.date, self.username] = self.stats_substract
        df2.to_excel("stats.xlsx")
        self.stats_count = df2.at[self.date, self.username]

    def get_quote(self):
        random_quote = random.choice(quotes_eng)
        self.quote_input = str(random_quote)

    def check_enable(self, *args):
        self.last_month = format(datetime.now() - relativedelta(months=1), '%B %Y')
        df = pd.read_excel("count.xlsx", engine="openpyxl", index_col=0, header=0)
        self.coffee_sum_lm = df.at[self.last_month, self.username]
        if int(self.coffee_sum_lm) != 0:
            self.ids.pay.disabled = False
        else:
            self.ids.pay.disabled = True
        if int(self.current_day) >= 5 and int(self.coffee_sum_lm) != 0:
            self.user_warning = str("Please pay " + str(self.coffee_sum_lm) + " cups from " + str(self.last_month) + "!")
        else:
            self.user_warning = str("")

    def on_touch_down(self, touch):

        self.reset_timeout()
        return super().on_touch_down(touch)

    def on_touch_move(self, touch):
        self.reset_timeout()
        return super().on_touch_move(touch)

    def on_leave(self, *args):
        Clock.unschedule(self.timeout)
        self.ids.history_spinner.text = "payment history"
        self.quote_input = str("")
        self.user_warning = str("")

class PayWindow(Screen):
    username = StringProperty('')
    today = StringProperty('')
    last_month = StringProperty('')
    coffee_sum = ObjectProperty()
    coffee_sum_lm = ObjectProperty()
    amount_c = ObjectProperty()
    amount_e = ObjectProperty()
    admin_username = StringProperty('')
    coffee_price = StringProperty('')
    paypal_mail = StringProperty('')
    paypal_link = StringProperty('')

    def __init__(self, **kwargs):
        super(PayWindow, self).__init__(**kwargs)

    def on_enter(self, *args):
        self.admin_username = config.get("settings", "admin_username")
        self.coffee_price = config.get("settings", "coffee_price")
        self.paypal_mail = config.get("settings", "paypal_mail")
        self.paypal_link = config.get("settings", "paypal_link")
        self.username = self.manager.get_screen("main_window").user_select
        self.today = str(datetime.now().strftime("%d")) + ". " + str(datetime.now().strftime("%B")) + " " + str(datetime.now().strftime("%Y"))
        self.last_month = format(datetime.now() - relativedelta(months=1), '%B %Y')
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        print(self.last_month)
        self.coffee_sum_lm = df.at[self.last_month, self.username]
        print(self.coffee_sum_lm)
        self.coffee_sum = df[self.username].sum()
        self.amount_c = int(self.coffee_sum_lm) * int(self.coffee_price)
        self.amount_e = int(self.amount_c) / 100

    def pay_history(self, *args):
        self.last_month = format(datetime.now() - relativedelta(months=1), '%B %Y')
        self.username = self.manager.get_screen("main_window").user_select
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        df.loc[self.last_month,self.username] = str("0")
        df.to_excel("count.xlsx")
        df3 = pd.read_excel("pay_history.xlsx",engine="openpyxl", index_col=0, header=0)
        test_add = str(self.today) + ": " + str(self.coffee_sum_lm) + " coffees (" + str(self.amount_e) + "€) paid."
        df3 = df3.append({self.username: test_add}, ignore_index=True)
        df3.to_excel("pay_history.xlsx")

class MenuWindow(Screen):
    today = StringProperty('')
    cleaned_history_list = []
    refilled_history_list = []
    def __init__(self, **kwargs):
        super(MenuWindow, self).__init__(**kwargs)

    def on_enter(self, *args):
        self.today = str(datetime.now().strftime("%d")) + ". " + str(datetime.now().strftime("%B")) + " " + str(datetime.now().strftime("%Y"))
        df4 = pd.read_excel("maintenance_history.xlsx", engine="openpyxl", index_col=0, header=0)
        df_cleaned = df4[df4.iloc[:,1].notna()]
        self.cleaned_history_list = df_cleaned.iloc[:,1].tolist()
        self.cleaned_history_list = [str(x) for x in self.cleaned_history_list]
        self.ids.clean_hist.values = self.cleaned_history_list

        df_refilled = df4[df4.iloc[:, 0].notna()]
        self.refilled_history_list = df_refilled.iloc[:, 0].tolist()
        self.refilled_history_list = [str(x) for x in self.refilled_history_list]
        self.ids.refill_hist.values = self.refilled_history_list

    def create_new_user(self, *args):
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        df2 = pd.read_excel("stats.xlsx",engine="openpyxl", index_col=0, header=0)
        df3 = pd.read_excel("pay_history.xlsx",engine="openpyxl",index_col=0, header=0)
        new_name = self.ids.new_user.text
        if new_name is not str(""):
            df[str(new_name)] = 0
            df.to_excel("count.xlsx")
            df2[str(new_name)] = 0
            df2.to_excel("stats.xlsx")
            df3[str(new_name)] = ""
            df3.to_excel("pay_history.xlsx")
            self.ids.not_found.text = ""
        else:
            self.ids.not_found.text = "Enter new user name!"

    def delete_user(self, *args):
        df = pd.read_excel("count.xlsx",engine="openpyxl", index_col=0, header=0)
        df2 = pd.read_excel("stats.xlsx",engine="openpyxl", index_col=0, header=0)
        df3 = pd.read_excel("pay_history.xlsx",engine="openpyxl",index_col=0, header=0)
        delete_name = self.ids.new_user.text
        if delete_name in list(df) and list(df2):
            df.drop(str(delete_name), axis=1, inplace=True)
            df2.drop(str(delete_name), axis=1, inplace=True)
            df3.drop(str(delete_name), axis=1, inplace=True)
            df.to_excel("count.xlsx")
            df2.to_excel("stats.xlsx")
            df3.to_excel("pay_history.xlsx")
            self.ids.not_found.text = ""
        else:
            self.ids.not_found.text = "User was not found!"

    def change_admin(self, *args):
        pass

    def cleaned(self, *args):
        df4 = pd.read_excel("maintenance_history.xlsx",engine="openpyxl",index_col=0, header=0)
        add_cleaned = str(self.today)
        df4 = df4.append({"CLEANED": add_cleaned}, ignore_index=True)
        df4.to_excel("maintenance_history.xlsx")
    def refilled(self, *args):
        df4 = pd.read_excel("maintenance_history.xlsx",engine="openpyxl",index_col=0, header=0)
        add_refill = str(self.today)
        df4 = df4.append({"REFILLED": add_refill}, ignore_index=True)
        df4.to_excel("maintenance_history.xlsx")
    def export_data(self, *args):
        date = str(datetime.now().strftime("%b-%d-%Y"))
        df = pd.read_excel("count.xlsx", engine="openpyxl", index_col=0, header=0)
        df2 = pd.read_excel("stats.xlsx", engine="openpyxl", index_col=0, header=0)
        df3 = pd.read_excel("pay_history.xlsx", engine="openpyxl", index_col=0, header=0)
        df4 = pd.read_excel("maintenance_history.xlsx",engine="openpyxl",index_col=0, header=0)
        if platform == 'android':
            from android.storage import primary_external_storage_path
            dir = primary_external_storage_path()
            download_dir_path = os.path.join(dir, 'Download')
            with pd.ExcelWriter('output.xlsx') as writer:
                df.to_excel(download_dir_path + "/pre_update" + date + ".xlsx",  sheet_name='count')
                df2.to_excel(download_dir_path + "/pre_update" + date + ".xlsx", sheet_name='stats')
                df3.to_excel(download_dir_path + "/pre_update" + date + ".xlsx", sheet_name='pay_history')
                df4.to_excel(download_dir_path + "/pre_update" + date + ".xlsx", sheet_name='maintenance_history.xlsx')

    def on_leave(self, *args):
        self.ids.new_user.text = ""
        self.ids.not_found.text = ""
        self.ids.clean_hist.text = "history"
        self.ids.refill_hist.text = "history"

class AdminWindow(Screen):
    admin_username = config.get("settings", "admin_username")
    paypal_mail = config.get("settings", "paypal_mail")
    paypal_link = config.get("settings", "paypal_link")
    coffee_price = config.get("settings", "coffee_price")
    admin_password = config.get("settings", "admin_password")


    def __init__(self, **kwargs):
        super(AdminWindow, self).__init__(**kwargs)

    def change_config(self, *args):
        admin_password = config.get("settings", "admin_password")
        password_input = self.ids.admin_pw.text
        admin_username = self.ids.admin_name.text
        paypal_mail = self.ids.admin_mail.text
        paypal_link = self.ids.admin_link.text
        coffee_price = self.ids.price.text
        if password_input == admin_password and admin_username is not str(""):
            config.set("settings", "admin_username", str(admin_username))
            self.ids.wrong_pw.text = "Username changed."
        elif password_input == admin_password and paypal_mail is not str(""):
            config.set("settings", "paypal_mail", str(paypal_mail))
            self.ids.wrong_pw.text = "Mail changed."
        elif password_input == admin_password and paypal_link is not str(""):
            config.set("settings", "paypal_link", str(paypal_link))
            self.ids.wrong_pw.text = "PayPal.me Link changed."
        elif password_input == admin_password and coffee_price is not str(""):
            config.set("settings", "coffee_price", str(coffee_price))
            self.ids.wrong_pw.text = "Price per cup changed."
        elif password_input is not admin_password:
            self.ids.wrong_pw.text = "Password incorrect!"
    # def change_admin_name(self, *args):
    #     admin_password = config.get("settings", "admin_password")
    #     password_input = self.ids.admin_pw.text
    #     admin_username = self.ids.admin_name.text
    #     if password_input == admin_password and admin_username is not str(""):
    #         config.set("settings", "admin_username", str(admin_username))
    #         self.ids.wrong_pw.text = "Changes accepted."
    #     elif password_input is not admin_password:
    #         self.ids.wrong_pw.text = "Password incorrect!"

    # def change_paypal_mail(self, *args):
    #     admin_password = config.get("settings", "admin_password")
    #     password_input = self.ids.admin_pw.text
    #     paypal_mail = self.ids.admin_mail.text
    #     if password_input == admin_password and paypal_mail is not str(""):
    #         config.set("settings", "paypal_mail", str(paypal_mail))
    #         self.ids.wrong_pw.text = "Changes accepted."
    #     elif password_input is not admin_password:
    #         self.ids.wrong_pw.text = "Password incorrect!"
    #
    # def change_paypal_link(self, *args):
    #     admin_password = config.get("settings", "admin_password")
    #     password_input = self.ids.admin_pw.text
    #     paypal_link = self.ids.admin_link.text
    #     if password_input == admin_password and paypal_link is not str(""):
    #         config.set("settings", "paypal_link", str(paypal_link))
    #         self.ids.wrong_pw.text = "Changes accepted."
    #     elif password_input is not admin_password:
    #         self.ids.wrong_pw.text = "Password incorrect!"

    def save_config(self):
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

    def on_leave(self, *args):
        self.ids.wrong_pw.text = ""
        self.ids.admin_pw.text = ""
        self.ids.admin_link.text = ""
        self.ids.admin_name.text = ""
        self.ids.admin_mail.text = ""
        self.ids.price.text = self.ids.price.text

class DataTable(Screen):
    def __init__(self, **kwargs):
        super(DataTable, self).__init__(**kwargs)
    def on_enter(self,table="",*args):
        data = pd.read_excel('stats.xlsx',engine="openpyxl",)
        column_titles = [x for x in data.keys()]
        rows_length = len(data[column_titles[0]])
        self.columns = len(column_titles)

        table_data = []
        for y in column_titles:
            table_data.append({'text':str(y),'size_hint_y':None,'width':200,'bcolor':(.79,.62,.4,.8)}) #append the data

        for z in range(rows_length):
            for y in column_titles:
                table_data.append({'text':str(data[y][z]),'size_hint_y':None,'width':200,'bcolor':(.79,.62,.4,1)}) #append the data

        self.ids.table_floor_layout.cols = self.columns #define value of cols to the value of self.columns
        self.ids.table_floor.data = table_data #add table_data to data value

    def on_leave(self, *args):
        data = pd.read_excel('stats.xlsx', engine="openpyxl")
        if platform == 'android':
            from android.storage import primary_external_storage_path
            dir = primary_external_storage_path()
            download_dir_path = os.path.join(dir, 'Download')
            data.to_excel(download_dir_path + "/stats.xlsx")


class GraphWindow(Screen):
    username = StringProperty('')
    year = StringProperty('')
    def __init__(self, **kwargs):
        super(GraphWindow, self).__init__(**kwargs)
    def on_enter(self, *args):
        self.username = self.manager.get_screen("main_window").user_select
        self.year = str(datetime.now().strftime("%Y"))
        # Define what we want to graph
        df2 = pd.read_excel('stats.xlsx',engine="openpyxl",)
        df_filtered = df2.set_index('Month').filter(regex=self.year +'$', axis=0)
        df_plot = df_filtered.filter([self.username], axis=1)
        df_plot = df_plot.set_axis(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
        df_plot.plot(color=(104/255,60/255,17/255,1))
        plt.tight_layout()
        plt.ylabel("cups of coffee")
        canvas = FigureCanvasKivyAgg(plt.gcf())
        canvas.draw()
        graph_box = self.ids.graph_box
        graph_box.add_widget(canvas)

    def on_leave(self, *args):
        plt.cla()
        graph_box = self.ids.graph_box
        graph_box.clear_widgets()

class RankingWindow(Screen):
    username = StringProperty('')
    year = StringProperty('')
    def __init__(self, **kwargs):
        super(RankingWindow, self).__init__(**kwargs)

    def on_enter(self, *args):
        self.username = self.manager.get_screen("main_window").user_select
        self.year = str(datetime.now().strftime("%Y"))
        # Define what we want to graph
        df2 = pd.read_excel('stats.xlsx', engine="openpyxl")
        df_filtered = df2.set_index('Month').filter(regex=self.year +'$', axis=0)
        print(df_filtered)
        #df_sorted = df_filtered.max().sort_values(ascending=False)
        df_sorted = df_filtered.sum().sort_values(ascending=False)
        print(df_sorted)
        df_sorted.plot.bar(color=(104/255,60/255,17/255,1))
        plt.tight_layout()
        plt.xticks(rotation=45, ha='right')
        plt.ylabel("cups of coffee")
        canvas = FigureCanvasKivyAgg(plt.gcf())
        canvas.draw()
        ranking_box = self.ids.ranking_box
        ranking_box.add_widget(canvas)

    def on_leave(self, *args):
        ranking_box = self.ids.ranking_box
        ranking_box.clear_widgets()

class cecoffee(App):
    def on_start(self):
        Window.update_viewport()


cecoffee().run()