from botcity.core import DesktopBot
import pandas as pd


class Bot(DesktopBot):
    def action(self, execution=None):

        tabela = pd.read_excel(str(r'C:\Users\efranca\Desktop\email_sender\gmail_sender\gmail_sender\database_gmail_sender.xlsx'), dtype=str)

        self.browse("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")

        if self.find( "optionsFind", matching=0.97, waiting_time=60000):
                self.not_found("optionsFind")
                self.click()

        for i in tabela.index:

            email = str(tabela.loc[i, 'email'])
            subject = str(tabela.loc[i, 'subject'])
            message = str(tabela.loc[i, 'message'])

            print('====Inicio====')
            print(f'email: {email}')
            print(f'subject: {subject}')
            print(f'message: {message}')
            print('==============')
            
            if self.find( "escreverFind", matching=0.97, waiting_time=60000):
                self.not_found("escreverFind")
                self.click()

            if self.find( "emailPaste", matching=0.97, waiting_time=10000):
                self.not_found("emailPaste")
            self.click_relative(64, 44)            
            self.paste(email)
            self.wait(1000)
            self.tab()
            self.wait(1000)
            self.paste(subject)
            self.tab()
            self.wait(1000)
            self.paste(message)
            self.wait(1000)
            self.tab()
            self.enter()

    def not_found(self, label):
        print(f"Element found: {label}")


if __name__ == '__main__':
    Bot.main()











