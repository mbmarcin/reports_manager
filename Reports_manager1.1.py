import pandas as pd
import win32com.client as win32
from logger import log_info as log
from logger import dump_log
from check_file_excel_version import Check_formatka as check_ver


class Path_manager:
	pass



class File_manager:

    @staticmethod
    def param_table():
        try:
            return pd.read_excel(Path_manager.filepath0, sheet_name='harmonogram')
        except FileNotFoundError:
            log(val=99, str_info=Path_manager.filepath0)
            pass
        except PermissionError:
            log(val=108, str_info=Path_manager.filepath0)

    @staticmethod
    def end_user():
        try:
            return pd.read_excel(Path_manager.filepath0, sheet_name='adresaci')
        except FileNotFoundError:
            log(val=99, str_info=Path_manager.filepath0)
            pass
        except PermissionError:
            log(val=108, str_info=Path_manager.filepath0)

    @staticmethod
    def file_to_disti(path_file_name):
        try:
            return pd.read_csv(path_file_name, encoding='utf-8', delimiter=';')
        except FileNotFoundError:
            log(val=99, str_info=path_file_name)
            pass
        except PermissionError:
            log(val=108, str_info=path_file_name)


class Parameters:
    monit_user = None
    monit_name = None
    package_ = None

    def __init__(self, file_param=None, file_endpoint=None, name_file_id=None):
        self.id = name_file_id
        self.file0 = file_param
        self.file1 = file_endpoint

    def names_param(self):
        return self.file0.columns.values

    def func(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[0]].values[0]

    def name_file(self):
        p = self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[1]].values[0]
        Parameters.monit_user = self.user()
        Parameters.monit_name = p
        return p

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
        # var1 = self.group_id()
        # Parameters.end_point_id = var1
        return self.file1.loc[self.file1.iloc[:, 0] == self.group_id(), 'id'].values[0].split(";")

    def package(self):
        return self.file0.loc[self.file0.iloc[:, 1] == self.id, self.names_param()[9]].values[0]


class Tools:

    def __init__(self, path=None, file_name=None, df_sc=None, end_user=None, df=None, list_user=None):
        self.path = path
        self.file_name = file_name
        self.df_sc = df_sc
        self.end_user = end_user
        self.df = df
        self.list_user = list_user

    def refresh(self):
        """
        # This one is for refreshing the Excel files without having to do it manually
        """
        # to do --->>>>>>>>>>>>>>>>>>>>>>>TEST
        _ = check_ver(self.path, self.file_name).check_and_copy()
        log(val=109, str_info=f'{self.file_name} dla {self.path[-7:]}')

        if _ in [0, 5, 11, 12, 16]:  # to do list to attention !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

            try:
                # re RefreshAll to test
                xlapp = win32.DispatchEx('Excel.Application')
                xlapp.DisplayAlerts = False
                xlapp.Visible = False
                xlbook = xlapp.Workbooks.Open(self.path + self.file_name)
                xlbook.RefreshAll()
                xlapp.CalculateUntilAsyncQueriesDone()
                xlbook.Save()
                xlbook.Close()
                xlapp.Quit()
                log(val=_, str_info=self.file_name)
            except:  # to test refresh with ERROR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                log(val=100, str_info=f'{self.file_name}')
                log(val=_, str_info=self.file_name)
        else:
            pass

    def split_data(self):

        split_data = dict()

        try:
            for user in self.end_user:
                df = self.df.loc[(self.df['id1'] == user) | (self.df['id2'] == user)]
                split_data.update({user: df})

            return split_data
        except ValueError:
            log(101)

    def push_csv(self, dict_data):

        for u in dict_data.keys():

            """
            sprawdzenie czy istnieją foldery!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            """

            tmp_path = f"{Path_manager.path}{u}\\dane\\{self.file_name}"
            try:
                dict_data[u].to_csv(tmp_path, sep=';', index=False, encoding='utf-8', decimal=',')
            except PermissionError:
                log(val=104, str_info=f'Nie mogę zrobić kopii {self.file_name} do: {u}')

    def to_point(self):
        return [h[:h.index('@')] for h in self.list_user]


class Mail:
    trigger_info = None

    def __init__(self, recipient=None, subject="Raport", body="W załączniku", attachment_path="", tr_info=0):
        self.recipient = recipient
        self.subject = subject
        self.body = body
        self.attachment_path = attachment_path,
        self.tr_info = tr_info

    def _body(self):
        if self.tr_info == 1:
            str_body = "\n\n\n\nMail wygenerowany z automatu, w przypadku problemów z plikem, \n" \
                       f"proszę o kontakt {Parameters.monit_user + '@ab.pl'}"  # !!!!!!!!!!!!!!!!!!!!!!!!!!!--------->param str
            return str_body
        else:
            str_body = ""
            return str_body

    def sending(self):
        try:
            outlook = win32.Dispatch('outlook.application')
            # outlook.DisplayAlerts = False
            mail = outlook.CreateItem(0)
            mail.To = self.recipient #'aspr323@ab.pl'
            # mail.CC = None
            # mail.BCC = "aspr323@ab.pl;aspr323@ab.pl"
            # mail.CC = self.recipient
            mail.Subject = self.subject
            # mail.HTMLBody = '<h3>This is HTML Body</h3>'
            mail.Body = self.body + self._body()

            mail.Attachments.Add(self.attachment_path[0])
            mail.Send()
            Mail.trigger_info = self.tr_info
            ####time.sleep(1)

            if Mail.trigger_info == 0:
                log(val=110, str_info=self.recipient)
            else:
                log(val=105, str_info=self.recipient)

        except BaseException:
            log(val=102)


class Monit(Parameters):
    def __init__(self, monit, end_id=None):
        self.end_id = end_id
        self._m = monit
        super().__init__()

    def send_monit(self):
        if self._m == 1:
            # to do --------------------------------->
            body = f'Raport wysłany do adresata --> {self.end_id} \n\n Odświeżony raport można sprawdzić ' \
                   f'bezpośrednio w folderze adresata na dysku P->DH->d.Analiz->...'
            Mail(recipient=Parameters.monit_user + '@ab.pl',  # !!!!!!!!!!!!!!!!!!!!!!!!!!!--------->param str
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
        self.end_point_users = Tools(list_user=self.endpoint_id()).to_point()

    def make_deal_with_output(self, m):

        if m == 0:

            # for ALL!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            # refresh file ---->display

            Tools(path=self.path_file(), file_name=self.name_file()).refresh()

            Mail(recipient=self.endpoint_id()[0],
                 attachment_path=self.path_file() + self.name_file(),
                 subject=self.name_file(),
                 tr_info=1
                 ).sending()

            Monit(Mail.trigger_info, self.endpoint_id()[0]).send_monit()

        elif m == 1:

            for han, r in zip(self.end_point_users, self.endpoint_id()):
                ## refresh file ---->display
                Tools(path=self.path_file() + han + '\\', file_name=self.name_file()).refresh()

                # send email
                Mail(recipient=r,
                     attachment_path=self.path_file() + han + "\\" + self.name_file(),
                     subject=self.name_file(),
                     tr_info=1
                     ).sending()

                Monit(Mail.trigger_info, han).send_monit()

        else:
            pass

    def make_distribution(self):

        ### open file to distributions
        file_disti = File_manager.file_to_disti(self.path_file() + self.name_file())

        ### split data per id
        data_per_user = Tools(end_user=self.end_point_users,
                              df=file_disti).split_data()

        """
        sprawdzenie czy folder jest dane jest
        
        """

        ## push data
        Tools(file_name=self.id_to_process).push_csv(data_per_user)
        log(val=103)

    def run(self):

        log(val=107, str_info=self.name_file(), gr=self.package())

        if self.on_off() == 1:

            # refresh and send for id -------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            if self.func() == 1:

                self.make_deal_with_output(1)

            # file distribution -------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            elif self.func() == 0:

                self.make_distribution()

            # refresh and send for group -------------------------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            elif self.func() == 2:

                self.make_deal_with_output(0)

        elif self.on_off() == 0:
            log(val=106, str_info=self.name_file())


# to do --------> multiprocess --> for group reports
def main():
    f0 = File_manager.param_table()  # df param
    f1 = File_manager.end_user()  # df end user
    f0.sort_values(by=f0.columns.values[0], inplace=True)

    # separation process for 0 and 1!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    # first 0 files!!!!!!!!!!!!!!!!!!!!!! multiprocess
    for g in f0.iloc[:, 9].drop_duplicates():  # -------->package_--->alerts for duplicates
        to_process = f0.loc[f0.iloc[:, 9] == g]

        for f in to_process.iloc[:, 1].drop_duplicates():  # -------->alerts for duplicates

            Process(f, to_process, f1).run()

    dump_log()


if __name__ == '__main__':
    main()
    log(val=199)

