# -*- coding: utf-8 -*-

import wx
import wx.adv
import os
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import locale
import logos
from io import BytesIO
import sys
from PIL import Image

# --- FUNÇÃO "MAPA" PARA ENCONTRAR ARQUIVOS NO .EXE ---


def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funcionando tanto no script quanto no .exe compilado pelo PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- FUNÇÃO PARA OBTER A LISTA DE CLIENTES EXISTENTES ---


def obter_clientes_existentes(pasta_base='ClientesGerados'):
    if not os.path.isdir(pasta_base):
        return []
    clientes = [nome for nome in os.listdir(
        pasta_base) if os.path.isdir(os.path.join(pasta_base, nome))]
    return clientes

# --- FUNÇÃO PARA GERENCIAR O CONTADOR POR CLIENTE ---


def gerenciar_contador_servico(nome_cliente, pasta_base='ClientesGerados'):
    pasta_cliente = os.path.join(pasta_base, nome_cliente)
    os.makedirs(pasta_cliente, exist_ok=True)
    arquivo_contador = os.path.join(pasta_cliente, 'contador.txt')
    numero_atual = 0
    try:
        with open(arquivo_contador, 'r') as f:
            numero_atual = int(f.read().strip())
    except FileNotFoundError:
        numero_atual = 0
    proximo_numero = numero_atual + 1
    with open(arquivo_contador, 'w') as f:
        f.write(str(proximo_numero))
    return proximo_numero


# --- LÓGICA DE GERAÇÃO DE ARQUIVOS (COM TAMANHO REAL DO LOGO) ---
def gerar_arquivos(dados_gerais, lista_itens):
    numero_servico = gerenciar_contador_servico(dados_gerais["nome_cliente"])
    nome_cliente_formatado = dados_gerais["nome_cliente"].replace(' ', '_')
    responsavel_formatado = dados_gerais["responsavel"].replace(' ', '_')
    data_formatada = dados_gerais["data_emissao"].replace('/', '-')
    nome_base_arquivo = (
        f"Nota_serviço_n_{numero_servico}_{nome_cliente_formatado}_{responsavel_formatado}_{data_formatada}")
    nomes_campos = {
        'Serviço': 'Serviço Adicional', 'Peso (kg)': 'Peso Bruto', 'Peso Peça Pronta (kg)': 'Peso Peça Pronta',
        'Quantidade Pedras': 'Qtd. Pedras', 'Gravação': 'Valor Gravação', 'Ródio': 'Valor Ródio',
        'Máquina Laser': 'Valor Máq. Laser', 'Valor do Produto': 'Valor do Produto', 'Valor de Mão de Obra': 'Valor Mão de Obra'
    }
    pasta_base = "ClientesGerados"
    nome_pasta_cliente = os.path.join(pasta_base, dados_gerais["nome_cliente"])
    os.makedirs(nome_pasta_cliente, exist_ok=True)
    df_geral = pd.DataFrame([dados_gerais])
    df_itens = pd.DataFrame(lista_itens).rename(columns=nomes_campos)
    caminho_excel = os.path.join(
        nome_pasta_cliente, f"{nome_base_arquivo}.xlsx")
    with pd.ExcelWriter(caminho_excel) as writer:
        df_geral.to_excel(writer, sheet_name='Dados Gerais', index=False)
        df_itens.to_excel(writer, sheet_name='Itens', index=False)
    caminho_pdf = os.path.join(nome_pasta_cliente, f"{nome_base_arquivo}.pdf")
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    try:
        pdf.add_font('RobotoMono', '', resource_path('RobotoMono-Regular.ttf'))
        pdf.add_font('RobotoMono', 'B', resource_path('RobotoMono-Bold.ttf'))
        pdf.add_font('RobotoMono', 'I', resource_path('RobotoMono-Italic.ttf'))
    except RuntimeError as e:
        wx.MessageBox(
            f"Não foi possível carregar a fonte Roboto Mono...\n\nErro: {e}", "Erro de Fonte", wx.OK | wx.ICON_ERROR)
        return

    # --- ALTERAÇÃO: Carrega o logo específico para o PDF ---
    logo_stream = BytesIO(logos.LogoClientePDF.GetData())
    # --- FIM DA ALTERAÇÃO ---

    pil_img = Image.open(logo_stream)
    largura_pixel, altura_pixel = pil_img.size
    largura_logo_mm = largura_pixel / pdf.k
    altura_logo_mm = altura_pixel / pdf.k
    pdf.set_draw_color(200, 200, 200)
    y_inicio_cabecalho = 15
    altura_cabecalho = altura_logo_mm + 10
    pdf.rect(10, y_inicio_cabecalho, 190, altura_cabecalho)
    pdf.set_font("RobotoMono", 'B', 16)
    texto_titulo = "Nota de Produto/Serviço"
    largura_titulo = pdf.get_string_width(texto_titulo)
    largura_total_bloco = largura_logo_mm + 5 + largura_titulo
    x_logo = (pdf.w - largura_total_bloco) / 2
    x_titulo = x_logo + largura_logo_mm + 5
    y_centro_cabecalho = y_inicio_cabecalho + (altura_cabecalho / 2)
    logo_stream.seek(0)
    pdf.image(logo_stream, x=x_logo, y=y_centro_cabecalho -
              (altura_logo_mm / 2), w=largura_logo_mm)
    pdf.text(x=x_titulo, y=y_centro_cabecalho + 5, txt=texto_titulo)
    pdf.set_y(y_inicio_cabecalho + altura_cabecalho + 8)
    pdf.set_font("RobotoMono", '', 11)
    pdf.cell(
        0, 7, f"Data de Emissão: {dados_gerais['data_emissao']}", align='L', ln=True)
    pdf.cell(
        0, 7, f"Cliente: {dados_gerais['nome_cliente']}", align='L', ln=True)
    pdf.ln(8)
    pdf.set_font("RobotoMono", 'B', 11)
    pdf.cell(0, 7, "Descrição dos Produtos/Serviços:", align='L', ln=True)
    campos_nao_monetarios = [
        "Peso (kg)", "Peso Peça Pronta (kg)", "Quantidade Pedras"]
    valor_total_final = 0
    for item in lista_itens:
        pdf.ln(4)
        pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 190, pdf.get_y())
        pdf.set_font("RobotoMono", 'B', 10)
        nome_item = item.get('Produto') or item.get('Serviço', 'Item')
        pdf.cell(0, 8, f"Item: {nome_item}", ln=True, border=0)
        pdf.set_font("RobotoMono", '', 9)
        for chave, valor in item.items():
            if chave in ['Produto', 'Subtotal'] or not valor:
                if isinstance(valor, (int, float)) and valor == 0:
                    continue
            if isinstance(valor, (int, float)):
                if chave in campos_nao_monetarios:
                    sufixo = " kg" if "Peso" in chave else ""
                    valor_formatado = f"{valor:.3f}{sufixo}" if "Peso" in chave else str(
                        int(valor))
                else:
                    valor_formatado = f"{valor:.2f}"
            else:
                valor_formatado = str(valor)
            pdf.cell(10)
            pdf.cell(60, 6, f"{nomes_campos.get(chave, chave)}:", border=0)
            pdf.cell(0, 6, valor_formatado, ln=True, border=0)
        pdf.set_font("RobotoMono", 'B', 10)
        subtotal_item_formatado = f"R$ {item['Subtotal']:,.2f}".replace(
            ",", "X").replace(".", ",").replace("X", ".")
        pdf.cell(
            0, 6, f"Subtotal do Item: {subtotal_item_formatado}", align='R', ln=True)
        valor_total_final += item['Subtotal']
    pdf.line(pdf.get_x(), pdf.get_y()+5, pdf.get_x() + 190, pdf.get_y()+5)
    pdf.ln(10)
    pdf.set_font("RobotoMono", 'B', 13)
    total_final_formatado = f"R$ {valor_total_final:,.2f}".replace(
        ",", "X").replace(".", ",").replace("X", ".")
    pdf.cell(
        0, 10, f"Valor Total: {total_final_formatado}", align='R', ln=True)
    pdf.ln(20)
    y_assinaturas = pdf.get_y()
    largura_assinatura = 80
    pdf.line(15, y_assinaturas, 15 + largura_assinatura, y_assinaturas)
    pdf.line(115, y_assinaturas, 115 + largura_assinatura, y_assinaturas)
    pdf.set_font("RobotoMono", '', 10)
    pdf.set_y(y_assinaturas + 2)
    pdf.set_x(15)
    pdf.cell(largura_assinatura, 10,
             f"Cliente - {dados_gerais['nome_cliente']}", align='C')
    pdf.set_x(115)
    pdf.cell(largura_assinatura, 10,
             f"Responsável - {dados_gerais['responsavel']}", align='C')
    pdf.ln()
    pdf.set_y(-35)
    pdf.set_font("RobotoMono", '', 10)
    pdf.cell(
        0, 7, f"{dados_gerais['local']}, {dados_gerais['data_emissao']}", align='C', ln=True)
    y_logo_dev = pdf.get_y()
    largura_logo_dev = 30
    x_logo_dev = (pdf.w - largura_logo_dev) / 2 - 20
    pdf.image(BytesIO(logos.LogoDev.GetData()),
              x=x_logo_dev, y=y_logo_dev, w=largura_logo_dev)
    pdf.set_font("RobotoMono", 'I', 8)
    pdf.text(x=x_logo_dev + largura_logo_dev + 2,
             y=y_logo_dev + 5, txt="www.mantini.app.br")
    pdf.output(caminho_pdf)
    return nome_pasta_cliente

# --- CLASSE DA INTERFACE GRÁFICA ---


class AppFrame(wx.Frame):
    def __init__(self, parent, title):
        super(AppFrame, self).__init__(parent, title=title, size=(1270, 900))

        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)
        self.itens = []
        self.main_sizer.AddSpacer(20)

        # --- ALTERAÇÃO: Carrega o logo específico para o sistema ---
        logo_cliente_img = logos.LogoClienteSistema.GetImage()
        logo_cliente_bmp = wx.Bitmap(logo_cliente_img)
        self.logo_cliente = wx.StaticBitmap(
            self.panel, bitmap=logo_cliente_bmp)
        self.main_sizer.Add(self.logo_cliente, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        # --- FIM DA ALTERAÇÃO ---

        geral_box = wx.StaticBox(self.panel, label="Dados Gerais")
        geral_sizer = wx.StaticBoxSizer(geral_box, wx.VERTICAL)
        form_geral_sizer = wx.GridBagSizer(5, 5)
        lista_clientes = obter_clientes_existentes()
        self.field_cliente = wx.ComboBox(
            self.panel, choices=lista_clientes, style=wx.CB_DROPDOWN | wx.CB_SORT)
        self.field_responsavel = wx.TextCtrl(self.panel)
        self.field_data = wx.adv.DatePickerCtrl(
            self.panel, style=wx.adv.DP_DEFAULT | wx.adv.DP_DROPDOWN)
        self.field_local = wx.TextCtrl(self.panel)
        form_geral_sizer.Add(wx.StaticText(self.panel, label="Nome do Cliente:"), pos=(
            0, 0), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
        form_geral_sizer.Add(self.field_cliente, pos=(0, 1))
        form_geral_sizer.Add(wx.StaticText(self.panel, label="Responsável:"), pos=(
            1, 0), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
        form_geral_sizer.Add(self.field_responsavel, pos=(1, 1))
        form_geral_sizer.Add(wx.StaticText(self.panel, label="Data:"), pos=(
            2, 0), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
        form_geral_sizer.Add(self.field_data, pos=(2, 1))
        form_geral_sizer.Add(wx.StaticText(self.panel, label="Local:"), pos=(
            3, 0), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
        form_geral_sizer.Add(self.field_local, pos=(3, 1))
        geral_sizer.Add(form_geral_sizer, 0, wx.ALL, 5)
        self.main_sizer.Add(geral_sizer, 0, wx.EXPAND | wx.ALL, 5)
        itens_box = wx.StaticBox(self.panel, label="Adicionar Produto/Serviço")
        itens_sizer = wx.StaticBoxSizer(itens_box, wx.VERTICAL)
        form_itens_sizer = wx.GridBagSizer(5, 5)
        self.item_fields = {}
        campos_item = [
            "Produto", "Serviço", "Peso (kg)", "Peso Peça Pronta (kg)", "Quantidade Pedras",
            "Gravação", "Ródio", "Máquina Laser", "Valor do Produto", "Valor de Mão de Obra"
        ]
        for i, label in enumerate(campos_item):
            row, col = divmod(i, 4)
            form_itens_sizer.Add(wx.StaticText(self.panel, label=f"{label}:"), pos=(
                row, col*2), flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
            field = wx.TextCtrl(self.panel, name=label)
            if label not in ["Produto", "Serviço"]:
                field.SetValue("0")
            form_itens_sizer.Add(field, pos=(row, col*2+1), flag=wx.EXPAND)
            self.item_fields[label] = field
        for i in [1, 3, 5, 7]:
            form_itens_sizer.AddGrowableCol(i)
        itens_sizer.Add(form_itens_sizer, 1, wx.EXPAND | wx.ALL, 5)
        btn_add_item = wx.Button(self.panel, label="Adicionar Item à Lista")
        btn_add_item.Bind(wx.EVT_BUTTON, self.on_add_item)
        itens_sizer.Add(btn_add_item, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        self.main_sizer.Add(itens_sizer, 0, wx.EXPAND | wx.ALL, 5)
        lista_box = wx.StaticBox(self.panel, label="Itens do Pedido")
        lista_sizer = wx.StaticBoxSizer(lista_box, wx.VERTICAL)
        self.lista_ctrl = wx.ListCtrl(
            self.panel, style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        self.colunas_lista = campos_item + ["Subtotal"]
        larguras = [150, 150, 80, 140, 100, 80, 80, 100, 120, 120, 100]
        for i, (col, larg) in enumerate(zip(self.colunas_lista, larguras)):
            self.lista_ctrl.InsertColumn(i, col, width=larg)
        lista_sizer.Add(self.lista_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        botoes_lista_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_edit_item = wx.Button(self.panel, label="Editar Item Selecionado")
        btn_edit_item.Bind(wx.EVT_BUTTON, self.on_edit_item)
        btn_remove_item = wx.Button(
            self.panel, label="Remover Item Selecionado")
        btn_remove_item.Bind(wx.EVT_BUTTON, self.on_remove_item)
        botoes_lista_sizer.Add(btn_edit_item, 1, wx.ALL, 5)
        botoes_lista_sizer.Add(btn_remove_item, 1, wx.ALL, 5)
        lista_sizer.Add(botoes_lista_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        self.main_sizer.Add(lista_sizer, 1, wx.EXPAND | wx.ALL, 5)
        self.lbl_total_final = wx.StaticText(
            self.panel, label="Valor Total: R$ 0,00")
        font_total = self.lbl_total_final.GetFont()
        font_total.SetPointSize(14)
        font_total.SetWeight(wx.FONTWEIGHT_BOLD)
        self.lbl_total_final.SetFont(font_total)
        self.main_sizer.Add(self.lbl_total_final, 0,
                            wx.ALIGN_RIGHT | wx.ALL, 10)
        btn_generate = wx.Button(self.panel, label="Gerar Arquivo")
        btn_generate.Bind(wx.EVT_BUTTON, self.on_generate)
        self.main_sizer.Add(btn_generate, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        self.main_sizer.AddStretchSpacer(1)
        line = wx.StaticLine(self.panel)
        self.main_sizer.Add(line, 0, wx.EXPAND | wx.ALL, 10)
        logo_dev_img = logos.LogoDev.GetImage()
        logo_dev_bmp = wx.Bitmap(logo_dev_img)
        self.logo_dev = wx.StaticBitmap(self.panel, bitmap=logo_dev_bmp)
        self.main_sizer.Add(self.logo_dev, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        self.panel.SetSizer(self.main_sizer)
        self.Centre()
        self.Show()

    def reset_all_fields(self):
        self.field_cliente.SetValue("")
        self.field_responsavel.SetValue("")
        self.field_local.SetValue("")
        self.field_data.SetValue(wx.DateTime.Now())
        for label, field in self.item_fields.items():
            field.SetValue("0") if label not in [
                "Produto", "Serviço"] else field.SetValue("")
        self.itens.clear()
        self.update_lista_e_total()

    def on_add_item(self, event):
        # --- INÍCIO DA CORREÇÃO ---
        # Validação dos campos obrigatórios "Produto" e "Valor do Produto"

        # 1. Valida o campo "Produto"
        if not self.item_fields["Produto"].GetValue().strip():
            wx.MessageBox("O campo 'Produto' é obrigatório.",
                          "Erro de Validação", wx.OK | wx.ICON_ERROR)
            return

        # 2. Valida o campo "Valor do Produto"
        valor_produto_str = self.item_fields["Valor do Produto"].GetValue(
        ).strip().replace(",", ".")
        try:
            valor_produto = float(valor_produto_str)
            if valor_produto <= 0:
                wx.MessageBox("O 'Valor do Produto' é obrigatório e deve ser maior que zero.",
                              "Erro de Validação", wx.OK | wx.ICON_ERROR)
                return
        except ValueError:
            wx.MessageBox("O 'Valor do Produto' deve ser um número válido.",
                          "Erro de Validação", wx.OK | wx.ICON_ERROR)
            return
        # --- FIM DA CORREÇÃO ---

        item_data = {}
        subtotal = 0
        campos_de_custo = ["Gravação", "Ródio", "Máquina Laser",
                           "Valor do Produto", "Valor de Mão de Obra"]

        # Loop para ler todos os dados e já converter
        for label, field in self.item_fields.items():
            valor_str = field.GetValue().strip().replace(",", ".")
            if label in ["Produto", "Serviço"]:
                item_data[label] = valor_str
            else:
                try:
                    valor_float = float(valor_str)
                    item_data[label] = valor_float
                except ValueError:
                    wx.MessageBox(
                        f"Valor inválido para '{label}'. Insira um número.", "Erro", wx.OK | wx.ICON_ERROR)
                    return

        # Loop para calcular o subtotal com base nos campos de custo
        for campo in campos_de_custo:
            subtotal += item_data.get(campo, 0)

        item_data["Subtotal"] = subtotal
        self.itens.append(item_data)
        self.update_lista_e_total()

        # Limpa os campos para a próxima entrada
        for label, field in self.item_fields.items():
            field.SetValue("0") if label not in [
                "Produto", "Serviço"] else field.SetValue("")

    def on_edit_item(self, event):
        selected_index = self.lista_ctrl.GetFirstSelected()
        if selected_index == -1:
            wx.MessageBox("Selecione um item para editar.",
                          "Aviso", wx.OK | wx.ICON_INFORMATION)
            return
        item_to_edit = self.itens[selected_index]
        for label, field in self.item_fields.items():
            field.SetValue(str(item_to_edit.get(label, "")))
        self.itens.pop(selected_index)
        self.update_lista_e_total()

    def on_remove_item(self, event):
        selected_index = self.lista_ctrl.GetFirstSelected()
        if selected_index != -1:
            self.itens.pop(selected_index)
            self.update_lista_e_total()
        else:
            wx.MessageBox("Selecione um item para remover.",
                          "Aviso", wx.OK | wx.ICON_INFORMATION)

    def update_lista_e_total(self):
        self.lista_ctrl.DeleteAllItems()
        campos_nao_monetarios = [
            "Peso (kg)", "Peso Peça Pronta (kg)", "Quantidade Pedras"]
        total_final = 0
        for item_dict in self.itens:
            index = self.lista_ctrl.InsertItem(
                self.lista_ctrl.GetItemCount(), item_dict.get("Produto", ""))
            for col_idx, chave in enumerate(self.colunas_lista):
                if col_idx == 0:
                    continue
                valor = item_dict.get(chave, "")
                if isinstance(valor, (int, float)):
                    if chave == "Subtotal":
                        valor_str = f"R$ {valor:,.2f}".replace(
                            ",", "X").replace(".", ",").replace("X", ".")
                    elif chave in campos_nao_monetarios:
                        sufixo = " kg" if "Peso" in chave else ""
                        valor_str = f"{valor:.3f}{sufixo}" if "Peso" in chave else str(
                            int(valor))
                    else:
                        valor_str = f"{valor:.2f}"
                else:
                    valor_str = str(valor)
                self.lista_ctrl.SetItem(index, col_idx, valor_str)
            total_final += item_dict['Subtotal']
        total_final_display = f"R$ {total_final:,.2f}".replace(
            ",", "X").replace(".", ",").replace("X", ".")
        self.lbl_total_final.SetLabel(f"Valor Total: {total_final_display}")
        self.main_sizer.Layout()

    def on_generate(self, event):
        dados_gerais = {
            "nome_cliente": self.field_cliente.GetValue().strip(),
            "responsavel": self.field_responsavel.GetValue().strip(),
            "data_emissao": self.field_data.GetValue().Format("%d/%m/%Y"),
            "local": self.field_local.GetValue().strip()
        }
        if not all([dados_gerais["nome_cliente"], dados_gerais["responsavel"], dados_gerais["local"]]):
            wx.MessageBox(
                "Os campos 'Nome do Cliente', 'Responsável' e 'Local' são obrigatórios.", "Erro", wx.OK | wx.ICON_ERROR)
            return
        if not self.itens:
            wx.MessageBox(
                "É necessário adicionar pelo menos um produto/serviço à lista.", "Erro", wx.OK | wx.ICON_ERROR)
            return
        try:
            pasta_criada = gerar_arquivos(dados_gerais, self.itens)
            wx.MessageBox(
                f"Sucesso! Arquivos foram gerados na pasta:\n{os.path.abspath(pasta_criada)}", "Concluído", wx.OK | wx.ICON_INFORMATION)
            novo_cliente = dados_gerais["nome_cliente"]
            if novo_cliente not in self.field_cliente.GetItems():
                self.field_cliente.Append(novo_cliente)
            self.reset_all_fields()
        except Exception as e:
            wx.MessageBox(
                f"Ocorreu um erro inesperado ao gerar os arquivos:\n{e}", "Erro", wx.OK | wx.ICON_ERROR)


# --- INICIALIZAÇÃO DA APLICAÇÃO ---
if __name__ == '__main__':
    wx.InitAllImageHandlers()
    app = wx.App(False)
    frame = AppFrame(
        None, title="Sistema de Nota de Serviço - Fernando Coelho - Desenvolvido por mantini.app")
    app.MainLoop()
