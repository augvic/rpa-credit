import xlwings as xw
import win32com.client
import time
import pandas as pd
import os
import keyboard
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as opt
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By as By
from datetime import datetime, timedelta

class RPACrédito:
    
    def __init__(self):
        self.Log = ""
        self.Encerrar = False
        self.ReiniciarLoop = False
        self.DataHoraInício = datetime.now().replace(microsecond = 0)
        self.DataHoraInício = self.DataHoraInício.strftime("%d-%m-%Y_%H-%M")
    
    def PrintarMensagem(self, Mensagem: str = None, CharType: str = None, Qtd: int = None, Side: str = None) -> None:
        DataHoraAtual = datetime.now().replace(microsecond = 0)
        DataHoraAtual = DataHoraAtual.strftime("%d/%m/%Y_%H:%M")
        if Mensagem:
            if CharType:
                if Side == "top":
                    print(f"<{DataHoraAtual}>")
                    print(CharType*Qtd)
                    print(Mensagem)
                    self.Log += f"<{DataHoraAtual}>\n{CharType*Qtd}\n{Mensagem}\n"
                if Side == "bot":
                    print(f"<{DataHoraAtual}>")
                    print(Mensagem)
                    print(CharType*Qtd)
                    self.Log += f"<{DataHoraAtual}>\n{Mensagem}\n{CharType*Qtd}\n"
                if Side == "both":
                    print(f"<{DataHoraAtual}>")
                    print(CharType*Qtd)
                    print(Mensagem)
                    print(CharType*Qtd)
                    self.Log += f"<{DataHoraAtual}>\n{CharType*Qtd}\n{Mensagem}\n{CharType*Qtd}\n"
            else:
                print(f"<{DataHoraAtual}>")
                print(Mensagem)
        else:
            print(CharType*Qtd)
    
    def InstanciarNavegador(self) -> webdriver:
        Options = opt()
        Options.add_argument("--log-level=3")
        Options.add_experimental_option("excludeSwitches", ["enable-logging"])
        Options.add_experimental_option("detach", True) 
        Driver = webdriver.Chrome(options=Options) 
        AbasAbertas = Driver.window_handles
        if len(AbasAbertas) > 1:
            Driver.switch_to.window(AbasAbertas[0])
            Driver.close()
        try:
            Driver.switch_to.window(AbasAbertas[0])
        except:
            Driver.switch_to.window(AbasAbertas[1])
        Driver.get(f"https://www.revendedorpositivo.com.br/admin/")
        microsoft_login_botao = None
        try:
            microsoft_login_botao = Driver.find_element(By.ID, value="login-ms-azure-ad")
            microsoft_login_botao.click()
            time.sleep(3)
            body = Driver.find_element(By.TAG_NAME, value="body").text
            if any(login_string in body for login_string in ["Because you're accessing sensitive info, you need to verify your password.", "Sign in", "Pick an account", "Entrar"]):
                self.PrintarMensagem("Necessário logar conta Microsoft.", "=", 50, "bot")
                while True:
                    body = Driver.find_element(By.TAG_NAME, value="body").text
                    if "DASHBOARD" in body:
                        break
                    else:
                        time.sleep(3)
            if "Approve sign in request" in body:
                time.sleep(3)
                codigo = Driver.find_element(By.ID, value="idRichContext_DisplaySign").text
                self.PrintarMensagem(f"Necessário authenticator Microsoft para continuar: {codigo}.", "=", 50, "bot")
                while True:
                    body = Driver.find_element(By.TAG_NAME, value="body").text
                    if "DASHBOARD" in body:
                        break
                    else:
                        time.sleep(3)
        except:
            Driver.get(f"https://www.revendedorpositivo.com.br/admin/index/")
        return Driver
    
    def InstanciarControle(self) -> dict:
        Controle = {}
        CaminhoScript = os.path.abspath(__file__)
        CaminhoControle = CaminhoScript.split(r"\rpa_credit.py")[0] + r"\control.xlsx"
        Controle["BOOK"] = xw.Book(CaminhoControle)
        Controle["PEDIDOS"] = Controle["BOOK"].sheets["PEDIDOS"]
        Controle["LIMITES"] = Controle["BOOK"].sheets["LIMITES"]
        return Controle
    
    def InstanciarSap(self) -> object:
        try:
            Gui = win32com.client.GetObject("SAPGUI")
            App = Gui.GetScriptingEngine
            Con = App.Children(0)
            for Id in range(0, 4):
                Session = Con.Children(Id)
                if Session.ActiveWindow.Text == "SAP Easy Access":
                    return Session
                else:
                    continue
            else:
                self.PrintarMensagem("Não foi encontrado tela SAP disponível para conexão.", "=", 30, "bot")
                exit()
        except:
            self.PrintarMensagem("Não foi encontrado tela SAP disponível para conexão.", "=", 30, "bot")
            exit()
    
    def AcessarPedido(self, Pedido: int) -> None:
        time.sleep(3)
        self.Driver.get(f"https://www.revendedorpositivo.com.br/admin/orders/edit/id/{Pedido}")
    
    def ColetarDataPedido(self) -> datetime:
        Data = self.Driver.find_element(By.XPATH, value="//label[@for='order_date']/following-sibling::div[@class='col-md-12']").text
        Data = datetime.strptime(Data, "%d/%m/%Y %H:%M:%S")
        return Data
    
    def ColetarCondiçãoPagamento(self) -> str:
        FormaPagamento = self.Driver.find_element(By.XPATH, value="//label[@for='payment_slip_installments_description']/following-sibling::div[@class='col-md-12']").text
        return FormaPagamento
    
    def ColetarFormaPagamentoPedido(self) -> str:
        FormaPagamento = self.Driver.find_element(By.XPATH, value="//label[@for='payment_name']/following-sibling::div[@class='col-md-12']").text
        return FormaPagamento
    
    def ColetarCnpj(self) -> str:
        Cnpj = self.Driver.find_element(By.XPATH, value="//label[@for='client_cnpj']/following-sibling::div[@class='col-md-12']").text
        Cnpj = Cnpj[:8]
        return Cnpj
    
    def ColetarCódigoERP(self, CNPJ: str) -> str:
        self.AbrirTransação("XD03")
        self.Session.findById("wnd[1]").sendVKey(4)
        self.Session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
        self.Session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = CNPJ
        self.Session.findById("wnd[2]/tbar[0]/btn[0]").press()
        StatusBarMsg = self.Session.findById("wnd[0]/sbar").text
        if "Nenhum valor para esta seleção" in StatusBarMsg:
            self.Session.findById("wnd[1]").close()
            self.Session.findById("wnd[1]").close()
            return "-"
        self.Session.findById("wnd[2]").sendVKey(2)
        CódigoERP = self.Session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text
        self.Session.findById("wnd[1]").close()
        return CódigoERP
    
    def ColetarValorPedido(self) -> float:
        ValorPedido = self.Driver.find_element(By.XPATH, value="//label[@for='payment_value']/following-sibling::div[@class='col-md-12']").text 
        ValorPedido = ValorPedido.replace("R$", "").replace(".", "").replace(",", ".")
        ValorPedido = float(ValorPedido)
        return ValorPedido
    
    def ColetarStatusPedido(self) -> str:
        try: 
            StatusPedido = self.Driver.find_element(By.NAME, value="distribution_centers[1][status]")
        except: 
            try:
                StatusPedido = self.Driver.find_element(By.NAME, value="distribution_centers[2][status]")
            except:
                StatusPedido = self.Driver.find_element(By.NAME, value="distribution_centers[3][status]") 
        StatusPedido = Select(StatusPedido) 
        StatusPedido = StatusPedido.first_selected_option.text
        if StatusPedido == "Cancelado pela positivo":
            StatusPedido = "CANCELADO"
        elif StatusPedido in ["Expedido", "Expedido parcial"]:
            StatusPedido = "FATURADO"
        elif StatusPedido == "Recusado pelo crédito":
            StatusPedido = "RECUSADO"
        elif StatusPedido in ["Pedido integrado", "Em separação", "Crédito aprovado", "Faturado"]:
            StatusPedido = "LIBERADO"
        elif StatusPedido == "Pedido recebido":
            StatusPedido = "RECEBIDO"
        return StatusPedido
    
    def ColetarClientePedido(self) -> str:
        Cliente = self.Driver.find_element(By.XPATH, value="//label[@for='client_name_corporate']/following-sibling::div[@class='col-md-12']").text
        Cliente = str(Cliente).split(" (")[0]
        return Cliente
    
    def AbrirTransação(self, Transação: str) -> None:
        self.Session.findById("wnd[0]/tbar[0]/okcd").text = "/N" + Transação
        self.Session.findById("wnd[0]").sendVKey(0)
        StatusBarMsg = None
        StatusBarMsg = self.Session.findById("wnd[0]/sbar").text
        if "Sem autorização" in StatusBarMsg:
            Erro = f"Sem acesso à {Transação}."
            self.PrintarMensagem(Erro, "=", 30, "bot")
            self.EncerrarRPA()
    
    def VerificarSeEstáVencido(self, DataParaVerificar: str) -> str:
        DataVencido = datetime.strptime(DataParaVerificar, "%d/%m/%Y").date()
        DataAtual = datetime.now().date()
        if DataVencido < DataAtual:
            DiasVencidos = 0
            DataVencido = DataVencido + timedelta(days = 1)
            while DataVencido < DataAtual:
                if DataVencido.weekday() < 5:
                    DiasVencidos += 1
                DataVencido = DataVencido + timedelta(days = 1)
            if DiasVencidos >= 2:
                return "Vencido"
            else:
                return "Não vencido"
        else:
            return "Não vencido"
    
    def ColetarDadosFinanceiros(self, RaizCnpj: str) -> dict:
        Dados = {}
        self.AbrirTransação("FD33")
        self.Session.findById("wnd[0]").sendVKey(4)
        self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
        i = 1
        while True:
            self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = ""
            self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = f"{RaizCnpj}000{i}*"
            self.Session.findById("wnd[1]/tbar[0]/btn[0]").press()
            Msg = self.Session.findById("wnd[0]/sbar").text
            if "Nenhum valor para esta seleção" in Msg:
                i += 1
            else:
                break
        self.Session.findById("wnd[1]").sendVKey(2)
        self.Session.findById("wnd[0]/usr/ctxtRF02L-KKBER").text = "1000"
        self.Session.findById("wnd[0]/usr/chkRF02L-D0210").selected = True
        self.Session.findById("wnd[0]").sendVKey(0)
        Limite = self.Session.findById("wnd[0]/usr/txtKNKK-KLIMK").text
        Limite = Limite.replace(".", "").replace(",", ".")
        Limite = float(Limite)
        LimiteStr = f"R$ {Limite:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        Vencimento = self.Session.findById("wnd[0]/usr/ctxtKNKK-NXTRV").text
        if not Vencimento == "":
            Vencimento = datetime.strptime(Vencimento, "%d.%m.%Y").date()
            VencimentoStr = datetime.strftime(Vencimento, "%d/%m/%Y")
            self.PrintarMensagem(f"Limite: {LimiteStr}", "=", 30, "bot")
            self.PrintarMensagem(f"Vencimento do limite: {VencimentoStr}", "=", 30, "bot")
        else:
            Vencimento = "-"
            self.PrintarMensagem(f"Limite: {LimiteStr}", "=", 30, "bot")
            self.PrintarMensagem(f"Vencimento do limite: {Vencimento}", "=", 30, "bot")
        self.AbrirTransação("FBL5N")
        self.Session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.Session.findById("wnd[1]/usr/txtENAME-LOW").text = "72776"
        self.Session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.Session.findById("wnd[0]").sendVKey(4)
        self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
        self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = f"{RaizCnpj}*"
        self.Session.findById("wnd[1]").sendVKey(0)
        Contas = []
        for Linha in range(3, 50):
            try:
                Conta = self.Session.findById(f"wnd[1]/usr/lbl[119,{Linha}]").text
            except:
                continue
            if Conta != "":
                Contas.append(Conta)
            else:
                break
        self.Session.findById("wnd[1]/tbar[0]/btn[0]").press()
        Tabela = []
        Empresas = ["1000", "3500"]
        for Conta in Contas:
            for Empresa in Empresas:
                self.Session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = Conta
                self.Session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = Empresa
                self.Session.findById("wnd[0]/tbar[1]/btn[8]").press()
                Msg = self.Session.findById("wnd[0]/sbar").text
                FormaBusca = "SCROLL"
                if Msg not in ["Nenhuma partida selecionada (ver texto descritivo)", "Nenhuma conta preenche as condições de seleção"]:
                    for Linha in range(10, 100):
                        try:
                            Célula = self.Session.findById(f"wnd[0]/usr/lbl[0,{Linha}]").text
                            if Célula == " Cliente":
                                FormaBusca = "ESTÁTICO"
                                break
                        except:
                            continue
                    if FormaBusca == "ESTÁTICO":
                        for Linha in range(10, 30):
                            TabelaDicionario = {}
                            try:
                                Situacao = self.Session.findById(f"wnd[0]/usr/lbl[6,{Linha}]").IconName
                                if Situacao != "S_LEDR":
                                    break
                                else:
                                    Situacao = "Em aberto"
                                FrmPag = self.Session.findById(f"wnd[0]/usr/lbl[39,{Linha}]").text
                                CondPag = self.Session.findById(f"wnd[0]/usr/lbl[132,{Linha}]").text
                                Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,{Linha}]").text
                                Valor = self.Session.findById(f"wnd[0]/usr/lbl[62,{Linha}]").text
                                Valor = Valor.replace(" ", "")
                                if FrmPag in ["7", "2", "M", "G", "J", "Z", "V", "A", "P", "S", "*"] or CondPag in ["0001", "0002", "Z576", "Z577"]:
                                    if not Valor.endswith("-"):
                                        continue
                                Vencido =  self.Session.findById(f"wnd[0]/usr/lbl[42,{Linha}]").IconName
                                if Vencido == "RESUBM":
                                    Vencido = "No prazo"
                                else:
                                    Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,{Linha}]").text
                                    Texto = self.Session.findById(f"wnd[0]/usr/lbl[81,{Linha}]").text
                                    if Conciliação == "CONCILIACAO":
                                        Vencido = "Conciliação"
                                    elif "DEVOLUÇÃO" in Texto:
                                        Vencido = "Devolução"
                                    elif "EXTRAVIO" in Texto:
                                        Vencido = "Extravio"
                                    else:
                                        try:
                                            DataVencimento = Conciliação
                                            if "." in DataVencimento:
                                                DataVencimento = str(DataVencimento).replace(".", "/")
                                            datetime.strptime(DataVencimento, "%d/%m/%Y")
                                        except:
                                            DataVencimento = self.Session.findById(f"wnd[0]/usr/lbl[28,{Linha}]").text
                                            DataVencimento = str(DataVencimento).replace(".", "/")
                                        Resultado = self.VerificarSeEstáVencido(DataVencimento)
                                        if Resultado == "Vencido":
                                            Vencido = "Vencido"
                                        else:
                                            Vencido = "No prazo"
                                Nf = self.Session.findById(f"wnd[0]/usr/lbl[45,{Linha}]").text
                                if Nf == "":
                                    break
                                if Valor.endswith("-"):
                                    Valor = "-" + Valor[:-1]
                                    Vencido = "Crédito"
                                TabelaDicionario["CONTA"] = Conta
                                TabelaDicionario["SITUAÇÃO"] = Situacao
                                TabelaDicionario["FRM. PAGAMENTO"] = FrmPag
                                TabelaDicionario["CND. PAGAMENTO"] = CondPag
                                TabelaDicionario["VENCIMENTO"] = Vencido
                                TabelaDicionario["NF"] = Nf
                                TabelaDicionario["VALOR"] = Valor
                                Tabela.append(TabelaDicionario)
                            except:
                                break
                    else:
                        for Linha in range(0, 500):
                            TabelaDicionario = {}
                            self.Session.findById("wnd[0]/usr").verticalScrollbar.position = Linha
                            try:
                                Situacao = self.Session.findById(f"wnd[0]/usr/lbl[6,10]").IconName
                                if Situacao != "S_LEDR":
                                    break
                                else:
                                    Situacao = "Em aberto"
                                FrmPag = self.Session.findById(f"wnd[0]/usr/lbl[39,10]").text
                                CondPag = self.Session.findById(f"wnd[0]/usr/lbl[132,10]").text
                                Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,10]").text
                                Valor = self.Session.findById(f"wnd[0]/usr/lbl[62,10]").text
                                Valor = Valor.replace(" ", "")
                                if FrmPag in ["7", "2", "M", "G", "J", "Z", "V", "A", "P", "S", "*"] or CondPag in ["0001", "0002", "Z576", "Z577"]:
                                    if not Valor.endswith("-"):
                                        continue
                                Vencido =  self.Session.findById(f"wnd[0]/usr/lbl[42,10]").IconName
                                if Vencido == "RESUBM":
                                    Vencido = "No prazo"
                                else:
                                    Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,10]").text
                                    Texto = self.Session.findById("wnd[0]/usr/lbl[81,10]").text
                                    if Conciliação == "CONCILIACAO":
                                        Vencido = "Conciliação"
                                    elif "DEVOLUÇÃO" in Texto:
                                        Vencido = "Devolução"
                                    elif "EXTRAVIO" in Texto:
                                        Vencido = "Extravio"
                                    else:
                                        try:
                                            DataVencimento = Conciliação
                                            if "." in DataVencimento:
                                                DataVencimento = str(DataVencimento).replace(".", "/")
                                            datetime.strptime(DataVencimento, "%d/%m/%Y")
                                        except:
                                            DataVencimento = self.Session.findById(f"wnd[0]/usr/lbl[28,10]").text
                                            DataVencimento = str(DataVencimento).replace(".", "/")
                                        Resultado = self.VerificarSeEstáVencido(DataVencimento)
                                        if Resultado == "Vencido":
                                            Vencido = "Vencido"
                                        else:
                                            Vencido = "No prazo"
                                Nf = self.Session.findById(f"wnd[0]/usr/lbl[45,10]").text
                                if Nf == "":
                                    break
                                if Valor.endswith("-"):
                                    Valor = "-" + Valor[:-1]
                                    Vencido = "Crédito"
                                TabelaDicionario["CONTA"] = Conta
                                TabelaDicionario["SITUAÇÃO"] = Situacao
                                TabelaDicionario["FRM. PAGAMENTO"] = FrmPag
                                TabelaDicionario["CND. PAGAMENTO"] = CondPag
                                TabelaDicionario["VENCIMENTO"] = Vencido
                                TabelaDicionario["NF"] = Nf
                                TabelaDicionario["VALOR"] = Valor
                                Tabela.append(TabelaDicionario)
                            except:
                                break
                    self.Session.findById("wnd[0]").sendVKey(3)
                else:
                    continue
        if Tabela:
            Df = pd.DataFrame(Tabela)
            Df["VALOR"] = Df["VALOR"].str.replace(".", "").str.replace(",", ".")
            Df["VALOR"] = Df["VALOR"].astype(float)
            SomaTotal = Df["VALOR"].sum()
            NovaLinha = pd.DataFrame({"CONTA": [""], "SITUAÇÃO": [""], "FRM. PAGAMENTO": [""], "CND. PAGAMENTO": [""], "VENCIMENTO": [""], "NF": ["TOTAL"], "VALOR": [SomaTotal]})
            Df = pd.concat([Df, NovaLinha])
            self.PrintarMensagem(f"Valores em aberto do cliente: {RaizCnpj}\n{Df}", "=", 30, "bot")
            EmAberto = SomaTotal
            TotalLinhas = Df.shape[0]
            Mensagem = "As seguintes notas estão vencidas:\n"
            NfVencida = ""
            for Linha in range(0, TotalLinhas):
                if Df.iloc[Linha]["VENCIMENTO"] == "Vencido":
                    if NfVencida == "":
                        NfVencida += Df.iloc[Linha]["NF"]
                    else:
                        NfVencida += " || " + Df.iloc[Linha]["NF"]
            Mensagem = Mensagem + NfVencida
            if NfVencida == "":
                self.PrintarMensagem("Sem vencidos.", "=", 30, "bot")
            else:
                self.PrintarMensagem(Mensagem, "=", 30, "bot")
        else:
            NfVencida = ""
            self.PrintarMensagem(f"Cliente: {RaizCnpj} não possui nada em aberto.", "=", 30, "bot")
            EmAberto = 0
        Dados["NfVencida"] = NfVencida
        Dados["EmAberto"] = EmAberto
        Dados["Limite"] = Limite
        Dados["Vencimento"] = Vencimento
        return Dados
    
    def ColetarVendedorPedido(self) -> str:
        Cnpj = self.Driver.find_element(By.XPATH, value="//label[@for='client_cnpj']/following-sibling::div[@class='col-md-12']").text 
        self.Driver.get("https://www.revendedorpositivo.com.br/admin/clients")
        Pesquisa = self.Driver.find_element(By.ID, value="keyword") 
        Pesquisa.clear()
        Pesquisa.send_keys(Cnpj)
        Pesquisa.send_keys(Keys.ENTER)
        time.sleep(3)
        try:
            Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_elements(By.XPATH, value=".//td")[10].find_element(By.XPATH, value=".//a") 
            Editar = Editar.get_attribute("href")
            self.Driver.get(str(Editar)) 
            time.sleep(3)
            Carteira = self.Driver.find_element(By.XPATH, value="//section").find_elements(By.XPATH, value=".//ul/li")[10].find_element(By.XPATH, value=".//a")
            Carteira.click()
            Carteira = self.Driver.find_element(By.XPATH, value="(//select[@class='form-control select-multiple side2side-selected-options side2side-select-taller'])[1]")
            Carteira = Select(Carteira)
            Carteira = Carteira.options
            Vendedor = Carteira[0].text
        except:
            try:
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/direct-billing-clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword") 
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Ativo = self.Driver.find_element(By.ID, value="active-1")
                Ativo.click()
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_element(By.XPATH, value="//td[contains(@data-title, 'Ações')]/a").get_attribute("href")
                self.Driver.get(str(Editar))
                Cnpj = self.Driver.find_element(By.ID, value="resale_cnpj").get_attribute("value")
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword")
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_elements(By.XPATH, value=".//td")[10].find_element(By.XPATH, value=".//a") 
                Editar = Editar.get_attribute("href") 
                self.Driver.get(str(Editar)) 
                time.sleep(3)
                Carteira = self.Driver.find_element(By.XPATH, value="//section").find_elements(By.XPATH, value=".//ul/li")[10].find_element(By.XPATH, value=".//a") 
                Carteira.click() 
                Carteira = self.Driver.find_element(By.XPATH, value="(//select[@class='form-control select-multiple side2side-selected-options side2side-select-taller'])[1]") 
                Carteira = Select(Carteira) 
                Carteira = Carteira.options 
                Vendedor = Carteira[0].text 
            except:
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/direct-billing-clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword") 
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Inativo = self.Driver.find_element(By.ID, value="active-0")
                Inativo.click()
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_element(By.XPATH, value="//td[contains(@data-title, 'Ações')]/a").get_attribute("href")
                self.Driver.get(str(Editar))
                Cnpj = self.Driver.find_element(By.ID, value="resale_cnpj").get_attribute("value")
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword")
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_elements(By.XPATH, value=".//td")[10].find_element(By.XPATH, value=".//a") 
                Editar = Editar.get_attribute("href") 
                self.Driver.get(str(Editar)) 
                time.sleep(3)
                Carteira = self.Driver.find_element(By.XPATH, value="//section").find_elements(By.XPATH, value=".//ul/li")[10].find_element(By.XPATH, value=".//a") 
                Carteira.click() 
                Carteira = self.Driver.find_element(By.XPATH, value="(//select[@class='form-control select-multiple side2side-selected-options side2side-select-taller'])[1]") 
                Carteira = Select(Carteira)
                Carteira = Carteira.options
                Vendedor = Carteira[0].text
        return Vendedor
    
    def ColetarDadosPedido(self, Pedido: int) -> dict:
        while True:
            DadosPedido = {}
            self.AcessarPedido(Pedido)
            self.PrintarMensagem(f"Coletando dados do pedido {Pedido}.", "=", 30, "bot")
            ConteúdoPágina = self.Driver.find_element(By.TAG_NAME, value="body").text
            if "Application error: Mysqli statement execute error" in ConteúdoPágina:
                self.PrintarMensagem(f"Pedido {Pedido} não inserido no site ainda.", "=", 30, "bot")
                self.ReiniciarLoop = True
                return DadosPedido
            FormaPagamento = self.ColetarFormaPagamentoPedido()
            if FormaPagamento == "Boleto a Prazo":
                DadosPedido["Pedido"] = Pedido
                DadosPedido["Data"] = self.ColetarDataPedido()
                DadosPedido["CondiçãoPagamento"] = self.ColetarCondiçãoPagamento()
                DadosPedido["Razão"] = self.ColetarClientePedido()
                DadosPedido["CNPJCliente"] = self.ColetarCnpj()
                DadosPedido["CódigoERP"] = self.ColetarCódigoERP(DadosPedido["CNPJCliente"])
                DadosPedido["ValorPedido"] = self.ColetarValorPedido()
                DadosPedido["Status"] = self.ColetarStatusPedido()
                DadosPedido["Vendedor"] = self.ColetarVendedorPedido()
                return DadosPedido
            self.PrintarMensagem(f"Pedido {Pedido} não possui forma de pagamento como crédito interno.", "=", 30, "bot")
            Pedido += 1
    
    def RemoverValorLiberadoDoControle(self, Pedido: int, AdicionarEmAberto: bool) -> None:
        Path = self.Controle["BOOK"].fullname
        Df = pd.read_excel(Path, "LIMITES")
        Colunas = [
                    "PEDIDO 1", "PEDIDO 2", "PEDIDO 3", "PEDIDO 4", "PEDIDO 5", "PEDIDO 6", "PEDIDO 7", "PEDIDO 8", "PEDIDO 9", "PEDIDO 10",
                    "PEDIDO 11", "PEDIDO 12", "PEDIDO 13", "PEDIDO 14", "PEDIDO 15", "PEDIDO 16", "PEDIDO 17", "PEDIDO 18", "PEDIDO 19", "PEDIDO 20"
                    ]
        ColunasPedido = ["F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL", "AN", "AP", "AR"]
        ColunasValor = ["G", "I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA", "AC", "AE", "AG", "AI", "AK", "AM", "AO", "AQ", "AS"]
        for i in range(0, 20):
            Linha = Df.index[Df[Colunas[i]] == Pedido].tolist()
            if Linha:
                ColunaPedido = ColunasPedido[i]
                ColunaValor = ColunasValor[i]
                Linha = int(Linha[0])
                Linha = Linha + 2
                break
        if not Linha:
            self.PrintarMensagem("Pedido não encontrado no controle das liberações.", "=", 30, "bot")
            return
        ValorPedido = float(self.Controle["LIMITES"].range(ColunaValor + str(Linha)).value)
        ValorAberto = float(self.Controle["LIMITES"].range("D" + str(Linha)).value)
        Soma = ValorPedido + ValorAberto
        if AdicionarEmAberto == True:
            self.Controle["LIMITES"].range("D" + str(Linha)).value = Soma
            self.PrintarMensagem(f"Pedido: {Pedido} removido das liberações e somado seu valor com o que cliente tem em aberto.", "=", 30, "bot")
        else:
            self.PrintarMensagem(f"Pedido: {Pedido} removido das liberações, pois não foi faturado.", "=", 30, "bot")
        self.Controle["LIMITES"].range(ColunaPedido + str(Linha)).value = ""
        self.Controle["LIMITES"].range(ColunaValor + str(Linha)).value = ""
    
    def SalvarControle(self) -> None:
        Tentativa = 0
        while Tentativa != 10:
            try:
                self.Controle["BOOK"].save()
                return
            except:
                Tentativa += 1
                time.sleep(2)
    
    def ÚltimaLinhaPreenchida(self, Aba: str, Coluna: str) -> int:
        ÚltimaLinha = self.Controle[Aba].range(Coluna + str("99999")).end("up").row
        return ÚltimaLinha
    
    def ImportarDadosFinanceirosNoControle(self, Cliente: int, Vencimento: datetime = None, Limite: float = None, EmAberto: float = None, Pedido: int = None, ValorPedido: float = None) -> None:
        self.SalvarControle()
        Path = self.Controle["BOOK"].fullname
        Df = pd.read_excel(Path, sheet_name = "LIMITES", dtype = {"CLIENTE": str})
        Linha = Df.index[Df['CLIENTE'] == str(Cliente)].tolist()
        if Linha:
            Linha = int(Linha[0])
            Linha = Linha + 2
        else:
            Linha = self.ÚltimaLinhaPreenchida("LIMITES", "A")
            Linha = Linha + 1
            self.Controle["LIMITES"].range("A" + str(Linha)).value = Cliente
        if Vencimento:
            self.Controle["LIMITES"].range("B" + str(Linha)).value = Vencimento
        if Limite or Limite == 0.0:
            self.Controle["LIMITES"].range("C" + str(Linha)).value = Limite
        if EmAberto or EmAberto == 0:
            self.Controle["LIMITES"].range("D" + str(Linha)).value = EmAberto
        if Pedido:
            ColunasPedidos = ["F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL", "AN", "AP", "AR"]
            ColunasValores = ["G", "I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA", "AC", "AE", "AG", "AI", "AK", "AM", "AO", "AQ", "AS"]
            for i in range(0, 20):
                Célula = self.Controle["LIMITES"].range(ColunasPedidos[i] + str(Linha)).value
                if Célula == Pedido:
                    self.Controle["LIMITES"].range(ColunasPedidos[i] + str(Linha)).value = Pedido
                    self.Controle["LIMITES"].range(ColunasValores[i] + str(Linha)).value = ValorPedido
                    break
                if Célula is None:
                    self.Controle["LIMITES"].range(ColunasPedidos[i] + str(Linha)).value = Pedido
                    self.Controle["LIMITES"].range(ColunasValores[i] + str(Linha)).value = ValorPedido
                    break
    
    def ColetarMargem(self, Cliente: str) -> float:
        self.SalvarControle()
        Path = self.Controle["BOOK"].fullname
        Df = pd.read_excel(Path, sheet_name = "LIMITES", dtype = {"CLIENTE": str})
        Linha = Df.index[Df['CLIENTE'] == str(Cliente)].tolist()
        Linha = int(Linha[0])
        Linha = Linha + 2
        Margem = float(self.Controle["LIMITES"].range("E" + str(Linha)).value)
        return Margem
    
    def AnáliseCréditoPedido(self, Pedido: int, RaizCnpj: int, Valor: float) -> dict:
        RespostaAnálise = {}
        DadosFinanceiros = self.ColetarDadosFinanceiros(RaizCnpj = RaizCnpj)
        Cliente = str(RaizCnpj)
        DataAtual = datetime.now().date()
        Vencimento = DadosFinanceiros["Vencimento"] 
        Limite = DadosFinanceiros["Limite"]
        NfVencida = DadosFinanceiros["NfVencida"]
        EmAberto = DadosFinanceiros["EmAberto"]
        ValorPedido = Valor
        ValorPedidoStr = f"R$ {ValorPedido:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.PrintarMensagem(f"Valor do pedido: {ValorPedidoStr}", "=", 30, "bot")
        if Vencimento == "-":
            VencimentoStr = "-"
        else:
            VencimentoStr = datetime.strftime(Vencimento, "%d/%m/%Y")
        self.ImportarDadosFinanceirosNoControle(Cliente, Vencimento, Limite, EmAberto)
        Margem = self.ColetarMargem(Cliente)
        Margem = round(Margem, 2)
        MargemStr = f"R$ {Margem:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.PrintarMensagem(f"Margem: {MargemStr}", "=", 30, "bot")
        LimiteAtivo = True
        Motivos = ""
        Status = "LIBERADO"
        if Limite == 0.0 or Vencimento == "-":
            Motivos += "\n- Sem limite de crédito ativo."
            Status = "NÃO LIBERADO"
            LimiteAtivo = False
        elif Vencimento < DataAtual:
            Motivos += f"\n- Limite vencido em {VencimentoStr}."
            Status = "NÃO LIBERADO"
            LimiteAtivo = False
        if LimiteAtivo == True:
            if Margem < ValorPedido:
                Motivos += f"\n- Valor do pedido excede a margem disponível. Valor do pedido: {ValorPedidoStr} / Margem livre: {MargemStr}."
                Status = "NÃO LIBERADO"
        if NfVencida != "":
            Motivos += f"\n- Possui vencidos: {NfVencida}."
            Status = "NÃO LIBERADO"
        if Status == "LIBERADO":
            RespostaAnálise["MENSAGEM"] = f"Pedido {Pedido} liberado."
            RespostaAnálise["STATUS"] = "LIBERADO"
            self.ImportarDadosFinanceirosNoControle(Cliente = Cliente, Pedido = Pedido, ValorPedido = ValorPedido)
        else:
            RespostaAnálise["MENSAGEM"] = f"Pedido {Pedido} recusado:{Motivos}"
            RespostaAnálise["STATUS"] = "NÃO LIBERADO"
        self.PrintarMensagem(RespostaAnálise["MENSAGEM"], "=", 30, "bot")
        return RespostaAnálise
    
    def Loop(self) -> None:
        while True:
            try:
                for Linha in range(2, 999999):
                    if self.Encerrar == True:
                        self.EncerrarRPA()
                    Pedido = self.Controle["PEDIDOS"].range("A" + str(Linha)).value
                    PrimeiroPedido = self.Controle["PEDIDOS"].range("B2").value
                    if Pedido is None or PrimeiroPedido is None:
                        if PrimeiroPedido is not None:
                            Pedido = int(self.Controle["PEDIDOS"].range("A" + str(Linha - 1)).value)
                            Pedido = Pedido + 1
                        else:
                            Pedido = int(Pedido)
                        DadosPedido = self.ColetarDadosPedido(Pedido)
                        if self.ReiniciarLoop == True:
                            self.ReiniciarLoop = False
                            break
                        self.PrintarMensagem(f"Inserindo dados do pedido {DadosPedido["Pedido"]} no controle.", "=", 30, "bot")
                        self.Controle["PEDIDOS"].range("A" + str(Linha)).value = DadosPedido["Pedido"]
                        self.Controle["PEDIDOS"].range("B" + str(Linha)).value = DadosPedido["Data"]
                        self.Controle["PEDIDOS"].range("C" + str(Linha)).value = DadosPedido["CondiçãoPagamento"]
                        self.Controle["PEDIDOS"].range("D" + str(Linha)).value = DadosPedido["Vendedor"]
                        self.Controle["PEDIDOS"].range("E" + str(Linha)).value = DadosPedido["Razão"]
                        self.Controle["PEDIDOS"].range("F" + str(Linha)).value = DadosPedido["CNPJCliente"]
                        self.Controle["PEDIDOS"].range("G" + str(Linha)).value = DadosPedido["ValorPedido"]
                        self.Controle["PEDIDOS"].range("H" + str(Linha)).value = DadosPedido["Status"]
                        self.Controle["PEDIDOS"].range("I" + str(Linha)).value = "-"
                        self.Controle["PEDIDOS"].range("J" + str(Linha)).value = "-"
                    Pedido = int(self.Controle["PEDIDOS"].range("A" + str(Linha)).value)
                    RaizCnpj = str(self.Controle["PEDIDOS"].range("F" + str(Linha)).value)
                    ValorPedido = float(self.Controle["PEDIDOS"].range("G" + str(Linha)).value)
                    Status = str(self.Controle["PEDIDOS"].range("H" + str(Linha)).value)
                    if Status == "LIBERADO":
                        self.PrintarMensagem(f"Verificando se houve atualização de status do pedido {Pedido}.", "=", 30, "bot")
                        self.AcessarPedido(Pedido)
                        StatusNovo = self.ColetarStatusPedido()
                        if StatusNovo == "FATURADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado: {Status} => {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo
                            self.RemoverValorLiberadoDoControle(Pedido = Pedido, AdicionarEmAberto = True)
                        elif StatusNovo == "CANCELADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado: {Status} => {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo
                            self.RemoverValorLiberadoDoControle(Pedido = Pedido, AdicionarEmAberto = False)
                        else:
                            self.PrintarMensagem(f"Pedido {Pedido} sem atualização de status.", "=", 30, "bot")
                    if Status == "RECEBIDO":
                        self.PrintarMensagem(f"Verificando se houve atualização de status do pedido {Pedido}.", "=", 30, "bot")
                        self.AcessarPedido(Pedido)
                        StatusNovo = self.ColetarStatusPedido()
                        if StatusNovo == "CANCELADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado para {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo
                        else:
                            self.PrintarMensagem(f"Iniciando análise do pedido {Pedido}. Possui status {Status}.", "=", 30, "bot")
                            RespostaAnálise = self.AnáliseCréditoPedido(Pedido = Pedido, RaizCnpj = RaizCnpj, Valor = ValorPedido)
                            self.Controle["PEDIDOS"].range("I" + str(Linha)).value = RespostaAnálise["MENSAGEM"]
                            self.Controle["PEDIDOS"].range("J" + str(Linha)).value = datetime.now()
                            if RespostaAnálise["STATUS"] == "NÃO LIBERADO":
                                self.AlterarPedidoSite(Pedido = Pedido, AlterarStatus = "Recusado pelo crédito", ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "RECUSADO"
                            else:
                                self.AlterarPedidoSite(Pedido = Pedido, AlterarStatus = "Crédito aprovado", ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "LIBERADO"
                    if Status == "RECUSADO":
                        self.PrintarMensagem(f"Verificando se houve atualização de status do pedido {Pedido}.", "=", 30, "bot")
                        self.AcessarPedido(Pedido)
                        StatusNovo = self.ColetarStatusPedido()
                        if StatusNovo == "CANCELADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado para {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo
                        else:
                            self.PrintarMensagem(f"Iniciando reanálise do pedido {Pedido}. Status continua como {Status}.", "=", 30, "bot")
                            RespostaAnálise = self.AnáliseCréditoPedido(Pedido = Pedido, RaizCnpj = RaizCnpj, Valor = ValorPedido)
                            self.Controle["PEDIDOS"].range("I" + str(Linha)).value = RespostaAnálise["MENSAGEM"]
                            self.Controle["PEDIDOS"].range("J" + str(Linha)).value = datetime.now()
                            if RespostaAnálise["STATUS"] == "NÃO LIBERADO":
                                self.AlterarPedidoSite(Pedido = Pedido, ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "RECUSADO"
                            else:
                                self.AlterarPedidoSite(Pedido = Pedido, AlterarStatus = "Crédito aprovado", ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "LIBERADO"
            except Exception as Erro:
                self.PrintarMensagem(f"Ocorreu o seguinte erro na execução: {Erro}. Aguardando 1m para reinício.", "=", 30, "bot")
                time.sleep(60)
    
    def EncerrarRPA(self) -> None:
        self.Driver.quit()
        while True:
            if self.Session.ActiveWindow.Text == "SAP Easy Access":
                break
            else:
                self.Session.findById("wnd[0]").sendVKey(3)
        self.Session = None
        self.SalvarControle()
        self.PrintarMensagem("Encerrando execução do RPA...", "=", 30, "bot")
        self.ExportarLog()
        exit()
    
    def ExportarLog(self) -> None:
        DataHoraFim = datetime.now().replace(microsecond = 0)
        DataHoraFim = DataHoraFim.strftime("%d-%m-%Y_%H-%M")
        CaminhoScript = os.path.abspath(__file__)
        CaminhoLogs = CaminhoScript.split(r"rpa_credit.py")[0] + r"\logs"
        with open(fr"{CaminhoLogs}\{self.DataHoraInício} & {DataHoraFim}.txt", "w", encoding = "utf-8") as LogFile:
            LogFile.write(self.Log)
    
    def ASCII(self) -> None:
        Ascii1 =  r"""#########################################################"""
        Ascii2 =  r"""#                                                       #"""
        Ascii3 =  r"""#  ____  ____   _       ____       __     _ _ _         #"""
        Ascii4 =  r"""# |  _ \|  _ \ / \     / ___|_ __ /_/  __| (_) |_ ___   #"""
        Ascii5 =  r"""# | |_) | |_) / _ \   | |   | '__/ _ \/ _` | | __/ _ \  #"""
        Ascii6 =  r"""# |  _ <|  __/ ___ \  | |___| | |  __/ (_| | | || (_) | #"""
        Ascii7 =  r"""# |_| \_\_| /_/   \_\  \____|_|  \___|\__,_|_|\__\___/  #"""
        Ascii8 =  r"""#                                                       #"""
        Ascii9 =  r"""#########################################################"""
        Ascii = f"{Ascii1}\n{Ascii2}\n{Ascii3}\n{Ascii4}\n{Ascii5}\n{Ascii6}\n{Ascii7}\n{Ascii8}\n{Ascii9}"
        self.PrintarMensagem(Ascii, "=", 30, "bot")
    
    def IniciarRPA(self) -> None:
        self.ASCII()
        self.Session = self.InstanciarSap()
        self.Driver = self.InstanciarNavegador()
        self.Controle = self.InstanciarControle()
        self.Loop()
    
    def MonitarEncerramento(self) -> None:
        while not self.Encerrar == True:
            time.sleep(0.5)
            if keyboard.is_pressed("CTRL+F12"):
                self.DefinirEncerramento()
    
    def DefinirEncerramento(self) -> None:
        self.Encerrar = True
    
    def AlterarPedidoSite(self, Pedido: int, AlterarStatus: str = None, ObservaçãoInterna: str = None) -> None:
        self.AcessarPedido(Pedido)
        if AlterarStatus is not None:
            for i in range(1, 4):
                try: 
                    StatusPedido = self.Driver.find_element(By.NAME, value = f"distribution_centers[{i}][status]")
                    StatusPedido = Select(StatusPedido)
                    StatusPedido.select_by_visible_text(AlterarStatus)
                except:
                    continue
        if ObservaçãoInterna is not None:
            CampoObservação = self.Driver.find_element(By.ID, value = "comment")
            CampoObservação.clear()
            CampoObservação.send_keys(ObservaçãoInterna)
        botãoSalvar = self.Driver.find_element(By.ID, value="save")
        botãoSalvar.click()

if __name__ == "__main__":
    Rpa = RPACrédito()
    thread = threading.Thread(target = Rpa.MonitarEncerramento)
    thread.daemon = True
    thread.start()
    Rpa.IniciarRPA()
