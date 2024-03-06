import win32com.client as win32


class EncaminharEmail:
    def EnviarEmail(self):
        outlook = win32.Dispatch("Outlook.Application")

        Lista_Email = [
            "alessandrojuliano828@gmail.com",
            "juj829852@gmail.com",
            "guilhermepaulinoribeiro@gmail.com",
            "defaultretro20@gmail.com",
            "ivandeleao@gmail.com",
            "aocleitonsantos@gmail.com",
        ]

        for email in Lista_Email:
            mensagem = outlook.CreateItem(0)
            mensagem.To = email
            mensagem.Subject = "Testando Automação de Emails com Python"
            mensagem.Body = "Deu certo  pelo menos uma parte ksskskksk"
            mensagem.Send()

        print("Mensagem encaminhada com sucesso")
