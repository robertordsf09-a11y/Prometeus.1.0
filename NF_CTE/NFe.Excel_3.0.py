import os
import xml.etree.ElementTree as ET
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from datetime import datetime
from openpyxl.styles import numbers

# Configurações de aparência
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Conversor NFe XML para Excel")
        self.geometry("700x500")

        # Variáveis de caminho
        self.xml_dir = ctk.StringVar()
        self.export_dir = ctk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        # Título
        self.label_title = ctk.CTkLabel(self, text="Extrator de Dados NFe (Modelo 55)",
                                        font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=20)

        # Seleção de pasta XML
        self.frame_xml = ctk.CTkFrame(self)
        self.frame_xml.pack(pady=10, padx=20, fill="x")

        self.btn_xml = ctk.CTkButton(self.frame_xml, text="Selecionar Pasta XML", command=self.select_xml_folder)
        self.btn_xml.pack(side="left", padx=10, pady=10)

        self.entry_xml = ctk.CTkEntry(self.frame_xml, textvariable=self.xml_dir, width=350)
        self.entry_xml.pack(side="left", padx=10, fill="x", expand=True)

        # Seleção de pasta de destino
        self.frame_export = ctk.CTkFrame(self)
        self.frame_export.pack(pady=10, padx=20, fill="x")

        self.btn_export = ctk.CTkButton(self.frame_export, text="Pasta de Destino", command=self.select_export_folder)
        self.btn_export.pack(side="left", padx=10, pady=10)

        self.entry_export = ctk.CTkEntry(self.frame_export, textvariable=self.export_dir, width=350)
        self.entry_export.pack(side="left", padx=10, fill="x", expand=True)

        # Frame para opções de processamento
        self.frame_options = ctk.CTkFrame(self)
        self.frame_options.pack(pady=10, padx=20, fill="x")

        self.process_option = ctk.StringVar(value="individual")

        self.radio_individual = ctk.CTkRadioButton(self.frame_options, text="Um Excel por XML",
                                                   variable=self.process_option, value="individual")
        self.radio_individual.pack(side="left", padx=10, pady=5)

        self.radio_unico = ctk.CTkRadioButton(self.frame_options, text="Um único Excel com todos os XMLs",
                                              variable=self.process_option, value="unico")
        self.radio_unico.pack(side="left", padx=10, pady=5)

        # Barra de Progresso
        self.progress_bar = ctk.CTkProgressBar(self, orientation="horizontal")
        self.progress_bar.pack(pady=20, padx=20, fill="x")
        self.progress_bar.set(0)

        # Botão Play
        self.btn_play = ctk.CTkButton(self, text="▶ Iniciar Processamento", command=self.start_thread,
                                      fg_color="green", hover_color="darkgreen",
                                      font=ctk.CTkFont(size=15, weight="bold"))
        self.btn_play.pack(pady=20)

        self.status_label = ctk.CTkLabel(self, text="Aguardando início...")
        self.status_label.pack()

    def select_xml_folder(self):
        path = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
        if path:
            self.xml_dir.set(path)

    def select_export_folder(self):
        path = filedialog.askdirectory(title="Selecione onde salvar o arquivo Excel")
        if path:
            self.export_dir.set(path)

    def format_value(self, value):
        """Converte strings para float e garante o formato numérico."""
        if value is None or value == "":
            return 0.0
        try:
            # Remove possíveis vírgulas e espaços
            if isinstance(value, str):
                value = value.replace(',', '.').strip()
            return float(value)
        except ValueError:
            return 0.0

    def get_tag_value(self, element, tag_name, default=""):
        """Busca uma tag recursivamente no elemento."""
        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
        found = element.find(f".//ns:{tag_name}", ns)
        if found is not None and found.text:
            return found.text.strip()
        return default
    
    def get_tag_value_from_parent(self, parent_element, tag_name, default=""):
        """Busca uma tag especificamente dentro de um elemento pai."""
        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
        if parent_element is not None:
            found = parent_element.find(f".//ns:{tag_name}", ns)
            if found is not None and found.text:
                return found.text.strip()
        return default

    def format_date_br(self, date_str):
        """Converte data do formato ISO (YYYY-MM-DD) para DD/MM/YYYY."""
        if not date_str:
            return ""
        try:
            # Tenta diferentes formatos de data
            if 'T' in date_str:
                date_str = date_str.split('T')[0]

            if '-' in date_str:
                date_obj = datetime.strptime(date_str, "%Y-%m-%d")
                return date_obj.strftime("%d/%m/%Y")
            else:
                return date_str
        except:
            return date_str

    def extract_nf_data(self, root, ns):
        """Extrai os dados gerais da Nota Fiscal."""
        # Dados da NF
        nNF = self.get_tag_value(root, "nNF")
        serie = self.get_tag_value(root, "serie")
        dhEmi = self.format_date_br(self.get_tag_value(root, "dhEmi"))
        chNFe = self.get_tag_value(root, "chNFe")

        # Dados do Emitente
        emit = root.find(".//ns:emit", ns)
        cnpj_emit = self.get_tag_value(emit, "CNPJ") if emit is not None else ""
        xNome_emit = self.get_tag_value(emit, "xNome") if emit is not None else ""
        ie_emit = self.get_tag_value(emit, "IE") if emit is not None else ""

        # Dados do Destinatário
        dest = root.find(".//ns:dest", ns)
        cnpj_dest = self.get_tag_value(dest, "CNPJ") if dest is not None else ""
        xNome_dest = self.get_tag_value(dest, "xNome") if dest is not None else ""
        xMun_dest = self.get_tag_value(dest, "xMun") if dest is not None else ""

        # Totais da NF - ESPECIFICAMENTE do bloco <total><ICMSTot>
        total_element = root.find(".//ns:total", ns)
        icms_tot_element = None
        if total_element is not None:
            icms_tot_element = total_element.find(".//ns:ICMSTot", ns)
        
        # Extrai os totais APENAS do bloco ICMSTot dentro de total
        if icms_tot_element is not None:
            vProd = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vProd"))
            vBC = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vBC"))
            vICMS = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vICMS"))
            vIPI = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vIPI"))
            vBCST = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vBCST"))
            vST = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vST"))
            vNF = self.format_value(self.get_tag_value_from_parent(icms_tot_element, "vNF"))
        else:
            # Fallback caso não encontre o bloco específico
            vProd = self.format_value(self.get_tag_value(root, "vProd"))
            vBC = self.format_value(self.get_tag_value(root, "vBC"))
            vICMS = self.format_value(self.get_tag_value(root, "vICMS"))
            vIPI = self.format_value(self.get_tag_value(root, "vIPI"))
            vBCST = self.format_value(self.get_tag_value(root, "vBCST"))
            vST = self.format_value(self.get_tag_value(root, "vST"))
            vNF = self.format_value(self.get_tag_value(root, "vNF"))

        return {
            "nNF": nNF,
            "serie": serie,
            "dhEmi": dhEmi,
            "chNFe": chNFe,
            "cnpj_emit": cnpj_emit,
            "xNome_emit": xNome_emit,
            "ie_emit": ie_emit,
            "cnpj_dest": cnpj_dest,
            "xNome_dest": xNome_dest,
            "xMun_dest": xMun_dest,
            "vProd": vProd,
            "vBC": vBC,
            "vICMS": vICMS,
            "vIPI": vIPI,
            "vBCST": vBCST,
            "vST": vST,
            "vNF": vNF
        }

    def extract_duplicatas(self, root, ns):
        """Extrai todas as duplicatas da Nota Fiscal."""
        duplicatas = []

        # Busca todas as tags de duplicata
        for dup in root.findall(".//ns:dup", ns):
            nDup = self.get_tag_value(dup, "nDup")
            dVenc = self.format_date_br(self.get_tag_value(dup, "dVenc"))
            vDup = self.format_value(self.get_tag_value(dup, "vDup"))  # Mantém como float

            duplicatas.append({
                "nDup": nDup,
                "dVenc": dVenc,
                "vDup": vDup
            })

        return duplicatas

    def extract_products(self, root, ns):
        """Extrai todos os produtos da Nota Fiscal (mantendo o código original)."""
        produtos = []

        for det in root.findall(".//ns:det", ns):
            prod = det.find("ns:prod", ns)
            imposto = det.find("ns:imposto", ns)

            if prod is not None:
                item = {
                    "CODIGO_PRODUTO": self.get_tag_value(prod, "cProd"),
                    "EAN": self.get_tag_value(prod, "cEAN"),
                    "DESCRIÇÃO": self.get_tag_value(prod, "xProd"),
                    "NCM": self.get_tag_value(prod, "NCM"),
                    "CEST": self.get_tag_value(prod, "CEST", ""),
                    "CFOP": self.get_tag_value(prod, "CFOP"),
                    "QUANTIDADE": self.format_value(self.get_tag_value(prod, "qCom")),
                    "V.UNIT": self.format_value(self.get_tag_value(prod, "vUnCom")),
                    "V.TOT": self.format_value(self.get_tag_value(prod, "vProd")),

                    # Impostos (todos como float)
                    "B.ICM": self.format_value(self.get_tag_value(imposto, "vBC") if imposto is not None else 0),
                    "V.ICM": self.format_value(self.get_tag_value(imposto, "vICMS") if imposto is not None else 0),
                    "AL.ICM": self.format_value(self.get_tag_value(imposto, "pICMS") if imposto is not None else 0),
                    "MVA": self.format_value(self.get_tag_value(imposto, "pMVAST") if imposto is not None else 0),
                    "B.ST": self.format_value(self.get_tag_value(imposto, "vBCST") if imposto is not None else 0),
                    "ICMSTD": self.format_value(self.get_tag_value(imposto, "pICMSST") if imposto is not None else 0),
                    "ST": self.format_value(self.get_tag_value(imposto, "vICMSST") if imposto is not None else 0),
                    "V.IPI": self.format_value(self.get_tag_value(imposto, "vIPI") if imposto is not None else 0),
                    "ALI.IPI": self.format_value(self.get_tag_value(imposto, "pIPI") if imposto is not None else 0),
                }
                produtos.append(item)

        return produtos

    def apply_brazilian_number_format(self, writer, sheet_name, columns_to_format):
        """Aplica formatação de número brasileiro (vírgula decimal) às colunas especificadas."""
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Formato brasileiro para números: #.##0,00
        brazilian_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1  # Formato: #.##0,00

        for column_name in columns_to_format:
            if column_name in writer.df.columns:
                # Encontra a coluna no Excel (letra da coluna)
                col_idx = writer.df.columns.get_loc(column_name) + 1  # +1 porque Excel começa em 1
                col_letter = openpyxl.utils.get_column_letter(col_idx)

                # Aplica o formato à coluna inteira (da linha 2 até o final)
                for row in range(2, len(writer.df) + 2):
                    cell = worksheet[f"{col_letter}{row}"]
                    cell.number_format = brazilian_format

    def create_excel_with_abas(self, nf_data, duplicatas, produtos, output_path):
        """Cria um arquivo Excel com múltiplas abas e formatação brasileira."""
        import openpyxl
        from openpyxl.utils import get_column_letter

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Aba 1: Dados Gerais da NF
            df_geral = pd.DataFrame([{
                "Número NF": nf_data["nNF"],
                "Série": nf_data["serie"],
                "Data Emissão": nf_data["dhEmi"],
                "Chave NFe": nf_data["chNFe"],
                "CNPJ Emitente": nf_data["cnpj_emit"],
                "Nome Emitente": nf_data["xNome_emit"],
                "IE Emitente": nf_data["ie_emit"],
                "CNPJ Destinatário": nf_data["cnpj_dest"],
                "Nome Destinatário": nf_data["xNome_dest"],
                "Município Destinatário": nf_data["xMun_dest"],
                "Valor Total Produtos": nf_data["vProd"],
                "Base ICMS": nf_data["vBC"],
                "Valor ICMS": nf_data["vICMS"],
                "Valor IPI": nf_data["vIPI"],
                "Base ICMS ST": nf_data["vBCST"],
                "Valor ICMS ST": nf_data["vST"],
                "Valor Total NF": nf_data["vNF"]
            }])
            df_geral.to_excel(writer, sheet_name="Dados Gerais", index=False)

            # Aplica formatação de números na aba Dados Gerais
            geral_numeric_columns = [
                "Valor Total Produtos", "Base ICMS", "Valor ICMS",
                "Valor IPI", "Base ICMS ST", "Valor ICMS ST", "Valor Total NF"
            ]

            workbook = writer.book
            worksheet = writer.sheets["Dados Gerais"]
            brazilian_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

            for col_name in geral_numeric_columns:
                if col_name in df_geral.columns:
                    col_idx = df_geral.columns.get_loc(col_name) + 1
                    col_letter = get_column_letter(col_idx)
                    cell = worksheet[f"{col_letter}2"]
                    cell.number_format = brazilian_format

            # Aba 2: Duplicatas (se houver)
            if duplicatas:
                df_duplicatas = pd.DataFrame(duplicatas)
                df_duplicatas.to_excel(writer, sheet_name="Duplicatas", index=False)

                # Formata a coluna vDup como número
                if 'vDup' in df_duplicatas.columns:
                    worksheet_dup = writer.sheets["Duplicatas"]
                    col_idx = df_duplicatas.columns.get_loc('vDup') + 1
                    col_letter = get_column_letter(col_idx)
                    for row in range(2, len(df_duplicatas) + 2):
                        cell = worksheet_dup[f"{col_letter}{row}"]
                        cell.number_format = brazilian_format
            else:
                df_vazio = pd.DataFrame({"Mensagem": ["Não há duplicatas registradas nesta NF"]})
                df_vazio.to_excel(writer, sheet_name="Duplicatas", index=False)

            # Aba 3: Produtos
            if produtos:
                df_produtos = pd.DataFrame(produtos)
                df_produtos.to_excel(writer, sheet_name="Produtos", index=False)

                # Formata todas as colunas numéricas na aba Produtos
                produtos_numeric_columns = [
                    "QUANTIDADE", "V.UNIT", "V.TOT", "B.ICM", "V.ICM",
                    "AL.ICM", "MVA", "B.ST", "ICMSTD", "ST", "V.IPI", "ALI.IPI"
                ]

                worksheet_prod = writer.sheets["Produtos"]

                for col_name in produtos_numeric_columns:
                    if col_name in df_produtos.columns:
                        col_idx = df_produtos.columns.get_loc(col_name) + 1
                        col_letter = get_column_letter(col_idx)
                        for row in range(2, len(df_produtos) + 2):
                            cell = worksheet_prod[f"{col_letter}{row}"]
                            cell.number_format = brazilian_format
            else:
                df_vazio = pd.DataFrame({"Mensagem": ["Não há produtos registrados nesta NF"]})
                df_vazio.to_excel(writer, sheet_name="Produtos", index=False)

    def process_individual_xml(self, xml_path, export_folder, file_name):
        """Processa um único XML e cria um Excel específico para ele."""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

            # Extrai dados
            nf_data = self.extract_nf_data(root, ns)
            duplicatas = self.extract_duplicatas(root, ns)
            produtos = self.extract_products(root, ns)

            # Gera nome do arquivo: nNF + xNome_emit
            nNF = nf_data.get("nNF", "SEM_NUMERO")
            xNome_emit = nf_data.get("xNome_emit", "SEM_NOME")

            # Remove caracteres inválidos para nome de arquivo
            xNome_emit = "".join(c for c in xNome_emit if c.isalnum() or c in (' ', '-', '_')).strip()

            output_filename = f"{nNF}_{xNome_emit}.xlsx"
            output_path = os.path.join(export_folder, output_filename)

            # Cria Excel com abas
            self.create_excel_with_abas(nf_data, duplicatas, produtos, output_path)

            return True, output_filename
        except Exception as e:
            return False, str(e)

    def process_unico_excel(self, xml_files, xml_folder, export_folder):
        """Processa múltiplos XMLs e cria um único Excel com abas separadas por NF."""
        try:
            output_path = os.path.join(export_folder, "Relatorio_NFe_Completo.xlsx")
            import openpyxl
            from openpyxl.utils import get_column_letter

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                all_geral_data = []
                all_duplicatas_data = []
                all_produtos_data = []

                for index, file in enumerate(xml_files):
                    xml_path = os.path.join(xml_folder, file)
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

                    # Extrai dados
                    nf_data = self.extract_nf_data(root, ns)
                    duplicatas = self.extract_duplicatas(root, ns)
                    produtos = self.extract_products(root, ns)

                    # Adiciona identificador da NF
                    identificador = f"NF {nf_data['nNF']} - {nf_data['xNome_emit']}"

                    # Prepara dados gerais (mantém como float)
                    geral_row = {
                        "Identificador": identificador,
                        "Número NF": nf_data["nNF"],
                        "Série": nf_data["serie"],
                        "Data Emissão": nf_data["dhEmi"],
                        "Chave NFe": nf_data["chNFe"],
                        "CNPJ Emitente": nf_data["cnpj_emit"],
                        "Nome Emitente": nf_data["xNome_emit"],
                        "IE Emitente": nf_data["ie_emit"],
                        "CNPJ Destinatário": nf_data["cnpj_dest"],
                        "Nome Destinatário": nf_data["xNome_dest"],
                        "Município Destinatário": nf_data["xMun_dest"],
                        "Valor Total Produtos": nf_data["vProd"],
                        "Base ICMS": nf_data["vBC"],
                        "Valor ICMS": nf_data["vICMS"],
                        "Valor IPI": nf_data["vIPI"],
                        "Base ICMS ST": nf_data["vBCST"],
                        "Valor ICMS ST": nf_data["vST"],
                        "Valor Total NF": nf_data["vNF"]
                    }
                    all_geral_data.append(geral_row)

                    # Prepara duplicatas
                    for dup in duplicatas:
                        dup_row = {
                            "Identificador": identificador,
                            "Número NF": nf_data["nNF"],
                            "Nome Emitente": nf_data["xNome_emit"],
                            "Número Duplicata": dup["nDup"],
                            "Vencimento": dup["dVenc"],
                            "Valor Duplicata": dup["vDup"]  # Mantém como float
                        }
                        all_duplicatas_data.append(dup_row)

                    # Prepara produtos
                    for prod in produtos:
                        prod_row = {
                            "Identificador": identificador,
                            "Número NF": nf_data["nNF"],
                            "Nome Emitente": nf_data["xNome_emit"],
                            "CODIGO_PRODUTO": prod.get("CODIGO_PRODUTO", ""),
                            "EAN": prod.get("EAN", ""),
                            "DESCRIÇÃO": prod.get("DESCRIÇÃO", ""),
                            "NCM": prod.get("NCM", ""),
                            "CEST": prod.get("CEST", ""),
                            "CFOP": prod.get("CFOP", ""),
                            "QUANTIDADE": prod.get("QUANTIDADE", 0.0),
                            "V.UNIT": prod.get("V.UNIT", 0.0),
                            "V.TOT": prod.get("V.TOT", 0.0),
                            "B.ICM": prod.get("B.ICM", 0.0),
                            "V.ICM": prod.get("V.ICM", 0.0),
                            "AL.ICM": prod.get("AL.ICM", 0.0),
                            "MVA": prod.get("MVA", 0.0),
                            "B.ST": prod.get("B.ST", 0.0),
                            "ICMSTD": prod.get("ICMSTD", 0.0),
                            "ST": prod.get("ST", 0.0),
                            "V.IPI": prod.get("V.IPI", 0.0),
                            "ALI.IPI": prod.get("ALI.IPI", 0.0)
                        }
                        all_produtos_data.append(prod_row)

                    # Atualiza progresso
                    progress = (index + 1) / len(xml_files)
                    self.progress_bar.set(progress)
                    self.status_label.configure(text=f"Processando: {index + 1} de {len(xml_files)} - {file}")

                # Cria as abas no Excel único
                df_geral = pd.DataFrame(all_geral_data)
                df_geral.to_excel(writer, sheet_name="Dados Gerais", index=False)

                # Formata números na aba Dados Gerais
                brazilian_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                geral_numeric_cols = ["Valor Total Produtos", "Base ICMS", "Valor ICMS",
                                      "Valor IPI", "Base ICMS ST", "Valor ICMS ST", "Valor Total NF"]

                worksheet_geral = writer.sheets["Dados Gerais"]
                for col_name in geral_numeric_cols:
                    if col_name in df_geral.columns:
                        col_idx = df_geral.columns.get_loc(col_name) + 1
                        col_letter = get_column_letter(col_idx)
                        for row in range(2, len(df_geral) + 2):
                            cell = worksheet_geral[f"{col_letter}{row}"]
                            cell.number_format = brazilian_format

                # Aba de Duplicatas
                if all_duplicatas_data:
                    df_duplicatas = pd.DataFrame(all_duplicatas_data)
                    df_duplicatas.to_excel(writer, sheet_name="Duplicatas", index=False)

                    worksheet_dup = writer.sheets["Duplicatas"]
                    if "Valor Duplicata" in df_duplicatas.columns:
                        col_idx = df_duplicatas.columns.get_loc("Valor Duplicata") + 1
                        col_letter = get_column_letter(col_idx)
                        for row in range(2, len(df_duplicatas) + 2):
                            cell = worksheet_dup[f"{col_letter}{row}"]
                            cell.number_format = brazilian_format

                # Aba de Produtos
                if all_produtos_data:
                    df_produtos = pd.DataFrame(all_produtos_data)
                    df_produtos.to_excel(writer, sheet_name="Produtos", index=False)

                    produtos_numeric_cols = ["QUANTIDADE", "V.UNIT", "V.TOT", "B.ICM", "V.ICM",
                                             "AL.ICM", "MVA", "B.ST", "ICMSTD", "ST", "V.IPI", "ALI.IPI"]

                    worksheet_prod = writer.sheets["Produtos"]
                    for col_name in produtos_numeric_cols:
                        if col_name in df_produtos.columns:
                            col_idx = df_produtos.columns.get_loc(col_name) + 1
                            col_letter = get_column_letter(col_idx)
                            for row in range(2, len(df_produtos) + 2):
                                cell = worksheet_prod[f"{col_letter}{row}"]
                                cell.number_format = brazilian_format

            return True, "Relatorio_NFe_Completo.xlsx"
        except Exception as e:
            return False, str(e)

    def process_xmls(self):
        """Método principal que processa os XMLs conforme opção selecionada."""
        xml_folder = self.xml_dir.get()
        export_folder = self.export_dir.get()
        process_option = self.process_option.get()

        if not xml_folder or not export_folder:
            messagebox.showwarning("Erro", "Por favor, selecione ambas as pastas.")
            return

        files = [f for f in os.listdir(xml_folder) if f.endswith('.xml')]
        if not files:
            messagebox.showwarning("Erro", "Nenhum arquivo XML encontrado na pasta.")
            return

        total_files = len(files)
        success_count = 0
        error_count = 0
        errors_list = []

        if process_option == "individual":
            # Processa cada XML individualmente
            for index, file in enumerate(files):
                xml_path = os.path.join(xml_folder, file)
                self.status_label.configure(text=f"Processando: {file}")

                success, result = self.process_individual_xml(xml_path, export_folder, file)

                if success:
                    success_count += 1
                else:
                    error_count += 1
                    errors_list.append(f"{file}: {result}")

                # Atualiza progresso
                progress = (index + 1) / total_files
                self.progress_bar.set(progress)

            # Mensagem final
            msg = f"Processamento concluído!\n\nSucesso: {success_count} arquivos\nErros: {error_count}"
            if errors_list:
                msg += f"\n\nErros:\n" + "\n".join(errors_list[:5])
                if len(errors_list) > 5:
                    msg += f"\n... e mais {len(errors_list) - 5} erros"

            messagebox.showinfo("Concluído", msg)

        else:  # process_option == "unico"
            self.status_label.configure(text="Processando arquivos para Excel único...")
            success, result = self.process_unico_excel(files, xml_folder, export_folder)

            if success:
                messagebox.showinfo("Sucesso", f"Arquivo Excel único criado com sucesso!\n\n{result}")
            else:
                messagebox.showerror("Erro", f"Falha ao processar:\n{result}")

        self.status_label.configure(text="Concluído!")
        self.progress_bar.set(0)

    def start_thread(self):
        # Roda o processamento em background para não travar a interface
        thread = threading.Thread(target=self.process_xmls)
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    app = App()
    app.mainloop()

