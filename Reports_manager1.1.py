# import shutil
#

# shutil.copyfile(original, target)
#


import os
import pandas as pd
import win32com.client as win32
import time


class Path_manager:
"""
find a path
"""
    #filepath0 = 1
    #path = 2
    #log_path = 3
    # filepath1 = 4

    # tbc .......
	pass


class File_manager:

    @staticmethod
    def param_table():
        return pd.read_excel(Path_manager.filepath0, sheet_name='harmonogram')

    @staticmethod
    def end_user():
        return pd.read_excel(Path_manager.filepath0, sheet_name='adresaci')

    @staticmethod
    def file_to_disti(path_file_name):
        # x = pd.read_csv(path_file_name, encoding='utf-8', delimiter=';')
        return pd.read_csv(path_file_name, encoding='utf-8', delimiter=';')

    # tbc .......


class Parameters:

    monit_user = None
    monit_name = None

    def __init__(self, file_param=None, file_endpoint=None, name_file_id=None):
        self.id = name_file_id
        self.file0 = file_param
        self.file1 = file_endpoint

    def names_param(self):
        return self.file0.columns.values

    # def test(self):
    #     print(self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[0]].values[0]
    #           )

    def func(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[0]].values[0]

    def name_file(self):
        Parameters.monit_user = self.user()
        Parameters.monit_name = self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[1]].values[0]
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[1]].values[0]

    def path_file(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[2]].values[0]

    def on_off(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[3]].values[0]

    def group_id(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[4]].values[0]

    def realization(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[5]].values[0]

    def user(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[6]].values[0]

    def endpoint_id(self):
        return self.file1.loc[self.file1.iloc[:, 0] == self.group_id(), 'id'].values[0].split(";")

    def package(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[9]].values[0]


class Tools:

    def __init__(self, path=None, file_name=None, df_sc=None, end_user=None, df=None):
        self.path = path
        self.file_name = file_name
        self.df_sc = df_sc
        self.end_user = end_user
        self.df = df

    def refresh(self):
        """
        # This one is for refreshing the Excel files without having to do it manually
        """
        # xlapp = win32.DispatchEx('Excel.Application')
        # xlapp.DisplayAlerts = False
        # xlapp.Visible = True
        # xlbook = xlapp.Workbooks.Open(self.path + self.file_name)
        # xlbook.RefreshAll()
        # xlapp.CalculateUntilAsyncQueriesDone()
        # xlbook.Save()
        # xlbook.Close()
        # xlapp.Quit()

        pass

    def split_data(self):

        split_data = dict()

        for user in self.end_user:
            # ds = 1
            # x = self.df.loc[(self.df.loc[:, ['id1']] == user) | (self.df.iloc[:, ['id2']] == user)]
            # split_data.update({user: self.df.loc[(self.df.iloc[:, 0] == user) | (self.df.iloc[:, 1] == user)]})

            df = self.df.loc[(self.df['id1'] == user) | (self.df['id2'] == user)]
            split_data.update({user: df})

        return split_data

    def push_csv(self, dict_data):

        for u in dict_data.keys():
            tmp_path = f"{Path_manager.path}{u}\\dane\\{self.file_name}"

            # !!!!!!!!!!!!!!!!!!!!! brak możliwości zapisu --------> zapis plików csv
            dict_data[u].to_csv(tmp_path, sep=';', index=False, encoding='utf-8', decimal=',')
            xv = 1

        # sep = ',', na_rep = '', float_format = None, columns = None, header = True, index = True,
        # index_label = None, mode = 'w',
        # encoding = None, compression = 'infer', quoting = None, quotechar = '"', line_terminator = None,
        # chunksize = None, date_format = None,
        # doublequote = True, escapechar = None, decimal = '.', errors = 'strict', storage_options = None

        pass


class Mail:
    trigger_info = None

    def __init__(self, recipient='aspr323@ab.pl', subject="Raport", body="W załączniku", attachment_path="", tr_info=0):
        self.recipient = recipient
        self.subject = subject
        self.body = body
        self.attachment_path = attachment_path,
        self.tr_info = tr_info

    # @property
    # def tr(self):
    #     return self.tr_info
    #
    # @tr.setter
    # def tr(self, val):
    #     if val is None:
    #         print(f'Brak ustawionego monitu dla {self.subject}')
    #     else:
    #         trigger_info = val

    # def set(self):
    #     trigger_info = self.tr_info

    def _body(self):
        if self.tr_info == 1:
            str_body = "\n\n\n\n Mail wygenerowany z automatu, w przypadku problemów z plikem, \n" \
                                f"proszę o kontakt {Parameters.monit_user+'@ab.pl'}" #!!!!!!!!!!!!!!!!!!!!!!!!!!!--------->param str
            return str_body
        else:
            str_body = ""
            return str_body

    def sending(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = self.recipient
        # mail.CC = None
        # mail.BCC = "aspr323@ab.pl;aspr323@ab.pl"
        # mail.CC = self.recipient
        mail.Subject = self.subject
        # mail.HTMLBody = '<h3>This is HTML Body</h3>'
        mail.Body = self.body + self._body()

        # "W załączniku lista numerów, których nie ma kartotece klienta i klienci z brakiem kontaktu.\n" \
        #         "Proszę uzupełnić kontakt w MSCRM.\n\nPozdrawiam,\nMB\n\n" \
        #         "(Ten email wygenerował automat, w przypadku problemów z załącznikiem, proszę o kontakt.)"

        mail.Attachments.Add(self.attachment_path[0])
        mail.Send()
        Mail.trigger_info = self.tr_info
        time.sleep(1)
        pass

    # def to_monit(self):
    #     trigger_info = self.tr_info


class Monit(Parameters):
    def __init__(self, monit):
        self._m = monit
        super().__init__()

    def send_monit(self):
        if self._m == 1:
            body = "Raport wysłany do adresata"
            Mail(recipient=Parameters.monit_user +'@ab.pl',#!!!!!!!!!!!!!!!!!!!!!!!!!!!--------->param str
                 subject=f'Monit--> {Parameters.monit_name}', body=body,
                 attachment_path=Path_manager.log_path).sending()
        else:
            pass


class Process(Parameters):

    def __init__(self, id_to_process, df_param, df_end_user):
        self.id_to_process = id_to_process
        self.df_param = df_param
        self.df_end_user = df_end_user
        super().__init__(file_param=df_param,
                         file_endpoint=df_end_user,
                         name_file_id=id_to_process)

    def make_deal_with_output(self, m):

        ## sprawdzenie czy dany plik jest dostępny??

        if m == 0:

            # for ALL!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            # refresh file ---->display
            Tools(path=self.path_file(), file_name=self.name_file()).refresh()

            # send email
            Mail(recipient=";".join([h + '@ab.pl' for h in self.endpoint_id()]),
                 attachment_path=self.path_file() + self.name_file(),
                 subject=self.name_file(),
                 tr_info=1
                 ).sending()

            Monit(Mail.trigger_info).send_monit()

        elif m == 1:

            for han, r in zip(self.endpoint_id(), [h + '@ab.pl' for h in
                                                   self.endpoint_id()]):  # ----------------> var string '@ab.pl'!!!!!!!!!!!!!!!!!!

                ## refresh file ---->display
                Tools(path=self.path_file() + han + '\\', file_name=self.name_file()).refresh()#----> to do

                # send email
                Mail(recipient=r,
                     attachment_path=self.path_file() + han + "\\" + self.name_file(),#----> to do
                     subject=self.name_file(),
                     ).sending()

                print(f'send to {han}')

        else:
            pass

    def make_distribution(self):

        ### open file to distributions
        file_disti = File_manager.file_to_disti(self.path_file() + self.name_file())

        ### split data per id
        data_per_user = Tools(end_user=self.endpoint_id(),
                              df=file_disti).split_data()

        ## push data
        Tools(file_name=self.id_to_process).push_csv(data_per_user)

        print('log------>')

    def run(self):

        if self.on_off() == 1:

            # refresh and send for id -------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            if self.func() == 1:
                # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                self.make_deal_with_output(1)

            # file distribution -------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            elif self.func() == 0:

                self.make_distribution()

            # refresh and send for group -------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            elif self.func() == 2:

                self.make_deal_with_output(0)

        elif self.on_off() == 0:
            print('process OFF')


if __name__ == '__main__':

    f0 = File_manager.param_table()  # df param
    f1 = File_manager.end_user()  # df end user
    f0.sort_values(by=f0.columns.values[0], inplace=True)

    # files_to_process = f0.iloc[:, [1, 9]].drop_duplicates()

    # separation process for 0 and 1!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    # first 0 files!!!!!!!!!!!!!!!!!!!!!! multiprocess
    for g in f0.iloc[:, 9].drop_duplicates():  # -------->alerts for duplicates
        to_process = f0.loc[f0.iloc[:, 9] == g]

        for f in to_process.iloc[:, 1].drop_duplicates():  # -------->alerts for duplicates
            print(f)
            Process(f, to_process, f1).run()

print('End process')

