from botcity.core import DesktopBot
import pandas as pd
from tqdm import tqdm

class Bot(DesktopBot):
    def action(self, execution=None):

        self.browse("http://web.whatsapp.com")

        tabela = pd.read_excel(r'CAMINHO DO ARQUIVO\Contatos.xlsx')
        print(tabela)

        for linha in tqdm(tabela.index):

            nome = tabela.loc[linha, "Nome"]
            contato = tabela.loc[linha, "Contato"]
            data = tabela.loc[linha, "Data_Vencimento"]
            valor = tabela.loc[linha, "Valor"]
            bitly = tabela.loc[linha, "Bitly"]
            cod_barras = tabela.loc[linha, "Cod_Barras"]
            boleto = tabela.loc[linha, "Boleto"]

            if not self.find( "ADD_NUMERO", matching=0.97, waiting_time=30000):
                self.not_found("ADD_NUMERO")
            self.click()
            self.paste(str(contato))

            if not self.find( "ENTRAR_CHAT", matching=0.97, waiting_time=20000):
                self.not_found("ENTRAR_CHAT")
            self.click()

            if self.find( "OK", matching=0.97, waiting_time=5000):
                self.click()
                print(f"{nome};{contato};NUMERO INVALIDO")
                continue

            self.wait(1000)
            if self.find( "CLICAR_MENSAGEM", matching=0.97, waiting_time=4000):
                self.click()
            self.paste(f"Olá *{nome}*, tudo bem? 🤩"
                        "\nVai aqui somente um *lembrete* tá? 📝"
                       f"\n\nMeu contato com você é sobre seu boleto *RIACHUELO*"
                       f"\n🗓️ *Vencimento:* *{data:%d/%m/%Y}*"
                       f"\n💰 *Valor:* *R${str(valor)}*")
            self.enter()

            self.wait(1000)
            if self.find( "CLICAR_MENSAGEM", matching=0.97, waiting_time=4000):
                self.click()
            self.paste(f"📲 Conferir os dados do boleto enviado: *NOME*, *DATA*, *VALOR* e *BENEFICIÁRIO* que deverá ser *MIDWAY*"
                        "\n\n⬇ Segue abaixo a *código de barras* e *boleto* para pagamento:")
            self.enter()

            self.wait(1000)
            if self.find( "CLICAR_MENSAGEM", matching=0.97, waiting_time=4000):
                self.click()
            self.paste(int(cod_barras))
            self.enter()

            self.wait(1000)
            if not self.find("ANEXAR", matching=0.97, waiting_time=40000):
                self.not_found("ANEXAR")
            self.click()
            if not self.find("DOCUMENTO", matching=0.97, waiting_time=40000):
                self.not_found("DOCUMENTO")
            self.click()
            if not self.find("NOME", matching=0.97, waiting_time=40000):
                self.not_found("NOME")
            self.paste(boleto)
            self.enter()
            self.enter()
            if not self.find("ENVIAR", matching=0.97, waiting_time=40000):
                self.enter()
            self.not_found(f"{nome};{contato};ENVIADO")
            self.enter()

            self.wait(1000)
            if self.find( "CLICAR_MENSAGEM", matching=0.97, waiting_time=4000):
                self.click()
            self.paste(f"Agora escolha *uma opção:*"
                       "\n1️⃣ Confirmar o *PAGAMENTO*"
                       "\n2️⃣ Enviar *COMPROVANTE* de pagamento"
                       "\n3️⃣ Falar com um *ATENDENTE*")
            self.enter()

            self.wait(180000)

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()