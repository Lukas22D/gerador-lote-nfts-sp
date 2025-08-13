import sys
import json
import re
from typing import Any, List, Dict
import pandas as pd
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QTabWidget, QTextEdit, QScrollArea,
    QFrame, QGridLayout, QSplitter
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont

# Utility functions
def pad_left(value, width: int) -> str:
    return str(value).rjust(width, '0')

def pad_right(value, width: int) -> str:
    return str(value).ljust(width, ' ')

class ExcelImporter:
    """Classe responsável por importar dados do Excel"""
    
    EXPECTED_COLUMNS = [
        "Tipo do Documento", "Numero do Documento", "Série do Documento", 
        "Data da Prestação", "Tributação do Serviço", "Código do Serviço", 
        "Item/SubItem", "Valor", "Aliquota", "ISS Retido pelo Tomador", 
        "Tipo de Prestador", "CNPJ/CPF do Prestador", "Cidade", "UF", 
        "CEP", "Discriminação dos Serviços"
    ]
    
    def __init__(self):
        self.aliquotas_data = self._load_aliquotas()
    
    def _load_aliquotas(self) -> Dict[str, str]:
        """Carrega dados de alíquotas do arquivo JSON"""
        try:
            with open('aliquotas.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
            m: Dict[str, str] = {}
            for item in data:
                cls = item['classificacao'].replace('.', '')
                cnae = item.get('cnae','')
                if cnae:
                    m[cls] = cnae
            return m
        except Exception as e:
            print(f"Erro ao carregar aliquotas.json: {e}")
            return {}
    
    def import_from_excel(self, file_path: str) -> List[Dict[str, Any]]:
        """Importa dados do arquivo Excel e retorna lista de notas"""
        try:
            df = pd.read_excel(file_path)
            
            # Verifica se todas as colunas esperadas estão presentes
            missing_columns = set(self.EXPECTED_COLUMNS) - set(df.columns)
            if missing_columns:
                raise ValueError(f"Colunas faltando no Excel: {', '.join(missing_columns)}")
            
            notes = []
            for _, row in df.iterrows():
                note_data = self._process_row(row)
                if note_data:  # Pula linhas vazias
                    notes.append(note_data)
            
            return notes
            
        except Exception as e:
            raise Exception(f"Erro ao importar Excel: {str(e)}")
    
    def _process_row(self, row) -> Dict[str, Any]:
        """Processa uma linha do Excel e retorna dados da nota"""
        # Pula linhas vazias
        if pd.isna(row["Numero do Documento"]):
            return None
        
        # Processa subitem para buscar código do serviço automaticamente
        subitem = str(row["Item/SubItem"]).strip()
        cod_servico = str(row["Código do Serviço"]).strip()
        
        # Auto-preenchimento do código do serviço se não estiver preenchido
        if len(subitem) >= 4 and not cod_servico:
            cls = subitem.lstrip('0')
            if cls and cls in self.aliquotas_data:
                cnae = self.aliquotas_data[cls]
                cod_servico = cnae.zfill(5) if len(cnae) < 5 else cnae
        
        return {
            'tipo_doc': self._extract_number_prefix(str(row["Tipo do Documento"])),
            'numero': str(row["Numero do Documento"]).strip(),
            'serie': str(row["Série do Documento"]).strip() if not pd.isna(row["Série do Documento"]) else '',
            'data': self._format_date(row["Data da Prestação"]),
            'tributacao': self._extract_letter_prefix(str(row["Tributação do Serviço"])),
            'cod_servico': cod_servico,
            'subitem': subitem,
            'valor_nota': str(row["Valor"]).replace(',', '.').strip(),
            'aliquota': str(row["Aliquota"]).strip() if not pd.isna(row["Aliquota"]) else '',
            'iss_retido': '1' if str(row["ISS Retido pelo Tomador"]).lower() in ['sim', 'yes', '1', 'true'] else '2',
            'tipo_prestador': self._extract_number_prefix(str(row["Tipo de Prestador"])),
            'cnpj_prest': re.sub(r'\D', '', str(row["CNPJ/CPF do Prestador"])),
            'regime': '0',  # Padrão, pode ser ajustado
            'cidade': str(row["Cidade"]).strip().upper(),
            'uf': str(row["UF"]).strip().upper(),
            'cep': re.sub(r'\D', '', str(row["CEP"])),
            'discriminacao': str(row["Discriminação dos Serviços"]).strip()
        }
    
    def _extract_number_prefix(self, text: str) -> str:
        """Extrai o prefixo numérico de um texto (ex: '01 - Dispensado' -> '01')"""
        match = re.match(r'^(\d+)', text.strip())
        return match.group(1) if match else text.strip()
    
    def _extract_letter_prefix(self, text: str) -> str:
        """Extrai o prefixo de letra de um texto (ex: 'T - Operação Normal' -> 'T')"""
        match = re.match(r'^([A-Z])', text.strip().upper())
        return match.group(1) if match else text.strip().upper()
    
    def _format_date(self, date_value) -> str:
        """Formata data para o formato AAAAMMDD"""
        if pd.isna(date_value):
            return QDate.currentDate().toString('yyyyMMdd')
        
        if isinstance(date_value, str):
            # Tenta diferentes formatos de data
            for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                try:
                    from datetime import datetime
                    dt = datetime.strptime(date_value, fmt)
                    return dt.strftime('%Y%m%d')
                except ValueError:
                    continue
            return QDate.currentDate().toString('yyyyMMdd')
        else:
            # Assume que é um objeto datetime
            return date_value.strftime('%Y%m%d')

class DataValidator:
    """Classe responsável por validar os dados das notas"""
    
    @staticmethod
    def validate_header(ccm: str, notes: List[Dict[str, Any]]) -> List[str]:
        """Valida dados do cabeçalho"""
        errs: List[str] = []
        if not re.fullmatch(r"\d{8}", str(ccm)):
            errs.append("CCM inválido: deve ter 8 dígitos numéricos.")
        if not notes:
            errs.append("É necessário ao menos uma nota.")
        return errs
    
    @staticmethod
    def validate_notes(notes: List[Dict[str, Any]]) -> List[str]:
        """Valida dados das notas"""
        errs: List[str] = []
        today = QDate.currentDate().toString('yyyyMMdd')
        
        for idx, n in enumerate(notes, start=1):
            if not re.fullmatch(r"\d{2}", n['tipo_doc']):
                errs.append(f"Nota {idx}: Tipo de documento inválido.")
            if not re.fullmatch(r"\d{1,12}", n['numero']):
                errs.append(f"Nota {idx}: Número inválido.")
            if n['tipo_doc']=='02' and not n['serie'].strip():
                errs.append(f"Nota {idx}: Série obrigatória para Tipo=02.")
            if not re.fullmatch(r"\d{8}", n['data']):
                errs.append(f"Nota {idx}: Data deve ser AAAAMMDD.")
            if n['data']< '20000101' or n['data']> today:
                errs.append(f"Nota {idx}: Data fora do intervalo.")
            if n['tributacao'] not in ['T','I','J']:
                errs.append(f"Nota {idx}: Tributação inválida.")
            if not re.fullmatch(r"\d{1,5}", n['cod_servico']):
                errs.append(f"Nota {idx}: Código de serviço inválido.")
            if not re.fullmatch(r"\d{1,4}", n['subitem']):
                errs.append(f"Nota {idx}: Subitem inválido.")
            if n['aliquota'].strip():
                if not re.fullmatch(r"\d{1,4}", n['aliquota']):
                    errs.append(f"Nota {idx}: Alíquota inválida.")
                elif int(n['aliquota'])>2500:
                    errs.append(f"Nota {idx}: Alíquota não pode exceder 25%.")
            if not re.fullmatch(r"\d+([\.,]\d{2})?", n['valor_nota']):
                errs.append(f"Nota {idx}: Valor da nota inválido.")
            if n['iss_retido'] not in ['1','2']:
                errs.append(f"Nota {idx}: ISS Retido inválido.")
            if n['tipo_prestador'] not in ['1','2','3']:
                errs.append(f"Nota {idx}: Tipo prestador inválido.")
            if not re.fullmatch(r"\d{14}", n['cnpj_prest']):
                errs.append(f"Nota {idx}: CNPJ do prestador deve ter 14 dígitos.")
            if n['regime'] not in ['0','4','5']:
                errs.append(f"Nota {idx}: Regime inválido.")
            if not n['cidade'].strip():
                errs.append(f"Nota {idx}: Cidade é obrigatória.")
            if not re.fullmatch(r"[A-Z]{2}", n['uf']):
                errs.append(f"Nota {idx}: UF inválida.")
            if not re.fullmatch(r"\d{8}", n['cep']):
                errs.append(f"Nota {idx}: CEP deve ter 8 dígitos numéricos.")
            if len(n['discriminacao'])>500:
                errs.append(f"Nota {idx}: Discriminação excede 500 caracteres.")
        
        return errs

class FileGenerator:
    """Classe responsável por gerar o arquivo de saída"""
    
    @staticmethod
    def generate_file(ccm: str, notes: List[Dict[str, Any]], file_path: str):
        """Gera o arquivo de lote NFTS"""
        content = FileGenerator._build_header(ccm, notes)
        for line in FileGenerator._build_details(notes):
            content += line
        content += FileGenerator._build_footer(notes)
        
        try:
            with open(file_path, 'w', encoding='ISO-8859-1', newline='') as f:
                f.write(content)
        except Exception as e:
            raise Exception(f"Erro ao salvar arquivo: {str(e)}")
    
    @staticmethod
    def _build_header(ccm: str, notes: List[Dict[str, Any]]) -> str:
        """Constrói cabeçalho do arquivo"""
        header = '1' + '001' + pad_left(ccm,8)
        dates = [n['data'] for n in notes]
        header += min(dates) + max(dates)
        return header + '\r\n'
    
    @staticmethod
    def _build_details(notes: List[Dict[str, Any]]) -> List[str]:
        """Constrói linhas de detalhe do arquivo"""
        lines: List[str] = []
        for n in notes:
            l = '4'
            l += pad_left(n['tipo_doc'],2)
            l += pad_right(n['serie'],5)
            l += pad_left(n['numero'],12)
            l += n['data']
            l += 'N'
            l += n['tributacao']
            serv = int(float(n['valor_nota'].replace(',','.'))*100)
            l += pad_left(serv,15)
            l += pad_left(0,15)
            l += pad_left(n['cod_servico'],5)
            l += pad_left(n['subitem'],4)
            aliq = n['aliquota'].strip() or '0'
            l += pad_left(aliq,4)
            l += n['iss_retido']
            l += n['tipo_prestador']
            l += pad_left(n['cnpj_prest'],14)
            l += pad_right('',8)   # inscrição municipal prestador
            l += pad_right('',75)  # razão social prestador

            # bloco opcional (173–430): Cidade(50), UF(2), CEP(8), preencher resto
            opt = ''
            # 18) Tipo de Endereço (173–175, 3)
            opt += pad_right('', 3)
            # 19) Logradouro        (176–225, 50)
            opt += pad_right('', 50)
            # 20) Número            (226–235, 10)
            opt += pad_right('', 10)
            # 21) Complemento       (236–265, 30)
            opt += pad_right('', 30)
            # 22) Bairro            (266–295, 30)
            opt += pad_right('', 30)
            # 23) Cidade            (296–345, 50)
            opt += pad_right(n['cidade'],50)
            opt += pad_right(n['uf'],2)
            opt += pad_left(n['cep'],8)
            opt += pad_right('', 258 - len(opt))
            l += opt

            l += pad_left('1',1)    # tipo NFTS
            l += n['regime']        # regime
            l += pad_right('',8)    # data pagamento em branco
            l += pad_right(n['discriminacao'],500)
            lines.append(l + '\r\n')
        return lines
    
    @staticmethod
    def _build_footer(notes: List[Dict[str, Any]]) -> str:
        """Constrói rodapé do arquivo"""
        count = len(notes)
        total = sum(int(float(n['valor_nota'].replace(',','.'))*100) for n in notes)
        footer = '9'
        footer += pad_left(count,7)
        footer += pad_left(total,15)
        footer += pad_left(0,15)
        return footer + '\r\n'

class InfoTab(QWidget):
    """Aba informativa com campos ENUM"""
    
    def __init__(self):
        super().__init__()
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("Guia de Campos ENUM - NFTS São Paulo")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Scroll area para o conteúdo
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # Informações dos campos ENUM
        enum_info = [
            {
                "title": "TIPO DO DOCUMENTO",
                "description": "Identifica o tipo de documento fiscal",
                "values": [
                    ("01", "Dispensado", "Documento dispensado de emissão"),
                    ("02", "Com emissão", "Documento com emissão obrigatória (requer série)"),
                    ("03", "Sem emissão", "Documento sem emissão")
                ]
            },
            {
                "title": "TRIBUTAÇÃO DO SERVIÇO",
                "description": "Define a situação tributária da prestação de serviço",
                "values": [
                    ("T", "Operação Normal", "Tributação normal com ISS"),
                    ("I", "Imune", "Operação imune de ISS"),
                    ("J", "ISS Suspenso", "ISS suspenso por decisão judicial")
                ]
            },
            {
                "title": "ISS RETIDO PELO TOMADOR",
                "description": "Indica se o ISS foi retido pelo tomador do serviço",
                "values": [
                    ("1", "Sim", "ISS foi retido pelo tomador"),
                    ("2", "Não", "ISS não foi retido pelo tomador")
                ]
            },
            {
                "title": "TIPO DE PRESTADOR",
                "description": "Identifica o tipo de documento do prestador",
                "values": [
                    ("1", "CPF", "Pessoa física"),
                    ("2", "CNPJ", "Pessoa jurídica"),
                    ("3", "Exterior", "Prestador no exterior")
                ]
            },
            {
                "title": "REGIME DE TRIBUTAÇÃO",
                "description": "Define o regime tributário do prestador",
                "values": [
                    ("0", "Normal/SN-DAMSP", "Regime normal ou Simples Nacional com DAMSP"),
                    ("4", "SN-DAS", "Simples Nacional com DAS"),
                    ("5", "MEI", "Microempreendedor Individual")
                ]
            },
            {
                "title": "SERVIÇO TOMADO",
                "description": "Indica se o serviço foi tomado (campo fixo)",
                "values": [
                    ("N", "Não", "Serviço não tomado (padrão para prestadores)"),
                    ("S", "Sim", "Serviço tomado (usado em casos específicos)")
                ]
            }
        ]
        
        for enum_data in enum_info:
            frame = self._create_enum_frame(enum_data)
            scroll_layout.addWidget(frame)
        
        # Seção de observações
        obs_frame = QFrame()
        obs_frame.setFrameStyle(QFrame.Box)
        obs_frame.setStyleSheet("QFrame { background-color: #f0f8ff; }")
        obs_layout = QVBoxLayout(obs_frame)
        
        obs_title = QLabel("OBSERVAÇÕES IMPORTANTES")
        obs_title.setFont(QFont("Arial", 12, QFont.Bold))
        obs_layout.addWidget(obs_title)
        
        observations = [
            "• Para Tipo 02 (Com emissão), a série do documento é obrigatória",
            "• Valores monetários devem ser informados em centavos (ex: R$ 15,00 = 1500)",
            "• Alíquotas devem ser informadas em centésimos (ex: 5% = 500)",
            "• CNPJ deve ter exatamente 14 dígitos numéricos",
            "• CEP deve ter exatamente 8 dígitos numéricos",
            "• Cidades devem ser escritas em MAIÚSCULAS",
            "• Discriminação dos serviços tem limite de 500 caracteres"
        ]
        
        for obs in observations:
            label = QLabel(obs)
            label.setWordWrap(True)
            obs_layout.addWidget(label)
        
        scroll_layout.addWidget(obs_frame)
        scroll_layout.addStretch()
        
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)
        
        self.setLayout(layout)
    
    def _create_enum_frame(self, enum_data):
        """Cria um frame para cada campo ENUM"""
        frame = QFrame()
        frame.setFrameStyle(QFrame.Box)
        frame.setStyleSheet("QFrame { background-color: #f9f9f9; margin: 5px; }")
        
        layout = QVBoxLayout(frame)
        
        # Título do campo
        title = QLabel(enum_data["title"])
        title.setFont(QFont("Arial", 12, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        layout.addWidget(title)
        
        # Descrição
        desc = QLabel(enum_data["description"])
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #7f8c8d; font-style: italic;")
        layout.addWidget(desc)
        
        # Tabela de valores
        grid = QGridLayout()
        grid.addWidget(QLabel("Código"), 0, 0)
        grid.addWidget(QLabel("Valor"), 0, 1)
        grid.addWidget(QLabel("Descrição"), 0, 2)
        
        # Estilo dos cabeçalhos
        for i in range(3):
            grid.itemAtPosition(0, i).widget().setStyleSheet("font-weight: bold; color: #34495e;")
        
        for idx, (code, value, description) in enumerate(enum_data["values"], 1):
            code_label = QLabel(code)
            code_label.setStyleSheet("font-family: monospace; font-weight: bold; color: #e74c3c;")
            
            value_label = QLabel(value)
            value_label.setStyleSheet("font-weight: bold; color: #27ae60;")
            
            desc_label = QLabel(description)
            desc_label.setWordWrap(True)
            
            grid.addWidget(code_label, idx, 0)
            grid.addWidget(value_label, idx, 1)
            grid.addWidget(desc_label, idx, 2)
        
        layout.addLayout(grid)
        return frame

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Lote NFTS - Importação Excel")
        self.resize(1200, 800)
        self.notes: List[Dict[str, Any]] = []
        self.excel_importer = ExcelImporter()
        self.data_validator = DataValidator()
        self.file_generator = FileGenerator()
        self._setup_ui()

    def _setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal com tabs
        main_layout = QVBoxLayout(central_widget)
        
        # Criação das abas
        tab_widget = QTabWidget()
        
        # Aba principal - Importação Excel
        main_tab = QWidget()
        self._setup_main_tab(main_tab)
        tab_widget.addTab(main_tab, "Importação Excel")
        
        # Aba informativa
        info_tab = InfoTab()
        tab_widget.addTab(info_tab, "Guia de Campos")
        
        main_layout.addWidget(tab_widget)

    def _setup_main_tab(self, tab_widget):
        layout = QVBoxLayout(tab_widget)
        
        # Seção do CCM
        ccm_layout = QHBoxLayout()
        ccm_layout.addWidget(QLabel("Contribuinte (CCM):"))
        self.ccm_combo = QComboBox()
        self.ccm_combo.addItem("4.165.071-9 – IM Filial", "41650719")
        self.ccm_combo.addItem("7.661.274-0 – IM Matriz", "76612740")
        ccm_layout.addWidget(self.ccm_combo)
        ccm_layout.addStretch()
        layout.addLayout(ccm_layout)

        # Seção de importação Excel
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Importar Excel:"))
        self.import_btn = QPushButton("Selecionar Arquivo Excel")
        self.import_btn.clicked.connect(self._import_excel)
        excel_layout.addWidget(self.import_btn)
        
        self.excel_file_label = QLabel("Nenhum arquivo selecionado")
        excel_layout.addWidget(self.excel_file_label)
        excel_layout.addStretch()
        layout.addLayout(excel_layout)

        # Tabela de notas
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels([
            "Tipo","Série","Número","Data","Serviço","Subitem","Alíquota","Valor(R$)"
        ])
        layout.addWidget(self.table)

        # Botões de ação
        btns_layout = QHBoxLayout()
        
        clear_btn = QPushButton("Limpar Todas")
        clear_btn.clicked.connect(self._clear_all)
        btns_layout.addWidget(clear_btn)
        
        btns_layout.addStretch()
        
        generate_btn = QPushButton("Gerar Arquivo NFTS")
        generate_btn.clicked.connect(self._generate)
        generate_btn.setStyleSheet("QPushButton { background-color: #27ae60; color: white; font-weight: bold; padding: 8px; }")
        btns_layout.addWidget(generate_btn)
        
        layout.addLayout(btns_layout)

    def _clear_all(self):
        """Limpa todas as notas"""
        if self.notes:
            reply = QMessageBox.question(self, "Confirmar", 
                                       "Deseja realmente limpar todas as notas?",
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.notes.clear()
                self.excel_file_label.setText("Nenhum arquivo selecionado")
                self._refresh()
    
    def _import_excel(self):
        """Importa notas do arquivo Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo Excel", "", 
            "Arquivos Excel (*.xlsx *.xls);;Todos os arquivos (*)"
        )
        
        if not file_path:
            return
        
        try:
            imported_notes = self.excel_importer.import_from_excel(file_path)
            
            if imported_notes:
                # Pergunta se deve substituir ou adicionar às notas existentes
                if self.notes:
                    reply = QMessageBox.question(
                        self, "Importar Excel", 
                        f"Foram encontradas {len(imported_notes)} notas no Excel.\n\n"
                        "Deseja substituir as notas existentes ou adicionar às existentes?",
                        QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                    )
                    
                    if reply == QMessageBox.Cancel:
                        return
                    elif reply == QMessageBox.Yes:  # Substituir
                        self.notes = imported_notes
                    else:  # Adicionar
                        self.notes.extend(imported_notes)
                else:
                    self.notes = imported_notes
                
                self.excel_file_label.setText(f"Arquivo: {Path(file_path).name} ({len(imported_notes)} notas)")
                self._refresh()
                
                QMessageBox.information(
                    self, "Sucesso", 
                    f"Importadas {len(imported_notes)} notas do Excel com sucesso!"
                )
            else:
                QMessageBox.warning(self, "Aviso", "Nenhuma nota válida encontrada no arquivo Excel.")
                
        except Exception as e:
            QMessageBox.critical(self, "Erro na Importação", str(e))

    def _refresh(self):
        """Atualiza a tabela com as notas"""
        self.table.setRowCount(0)
        for n in self.notes:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r,0, QTableWidgetItem(n['tipo_doc']))
            self.table.setItem(r,1, QTableWidgetItem(n['serie']))
            self.table.setItem(r,2, QTableWidgetItem(n['numero']))
            self.table.setItem(r,3, QTableWidgetItem(n['data']))
            self.table.setItem(r,4, QTableWidgetItem(n['cod_servico']))
            self.table.setItem(r,5, QTableWidgetItem(n['subitem']))
            self.table.setItem(r,6, QTableWidgetItem(n['aliquota'] or '0'))
            self.table.setItem(r,7, QTableWidgetItem(n['valor_nota']))

    def _validate_data(self) -> List[str]:
        """Valida todos os dados usando a classe DataValidator"""
        ccm = self.ccm_combo.currentData()
        header_errors = self.data_validator.validate_header(ccm, self.notes)
        notes_errors = self.data_validator.validate_notes(self.notes)
        return header_errors + notes_errors

    def _generate(self):
        """Gera o arquivo de lote NFTS usando a classe FileGenerator"""
        errs = self._validate_data()
        if errs:
            QMessageBox.critical(self, "Erros de Validação", "\n".join(errs))
            return

        fn, _ = QFileDialog.getSaveFileName(
            self, "Salvar Arquivo", "", "Arquivo Texto (*.txt)"
        )
        if not fn:
            return
        if not fn.lower().endswith('.txt'):
            fn += '.txt'

        try:
            ccm = self.ccm_combo.currentData()
            self.file_generator.generate_file(ccm, self.notes, fn)
            QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao salvar", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
