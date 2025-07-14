import sys
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QDateEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QDialog, QFormLayout, QTextEdit, QCheckBox, QFileDialog, QMessageBox
)
from PyQt5.QtCore import QDate
from typing import Any, List, Dict
import re

# Utility functions
def pad_left(value, width):
    return str(value).rjust(width, '0')

def pad_right(value, width):
    return str(value).ljust(width, ' ')

class NoteDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adicionar/Editar Nota")
        self.aliquotas_data = self._load_aliquotas()
        self._setup_ui()

    def _load_aliquotas(self) -> Dict[str, str]:
        """Carrega o arquivo aliquotas.json e cria um mapeamento classificação -> cnae"""
        try:
            with open('aliquotas.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Cria um dicionário mapeando classificação -> cnae
            # Remove o ponto da classificação para facilitar a busca
            aliquotas_map = {}
            for item in data:
                classificacao = item['classificacao'].replace('.', '')
                cnae = item['cnae']
                if cnae:  # Só adiciona se o CNAE não estiver vazio
                    aliquotas_map[classificacao] = cnae
            
            return aliquotas_map
        except Exception as e:
            print(f"Erro ao carregar aliquotas.json: {e}")
            return {}

    def _setup_ui(self):
        layout = QVBoxLayout()
        form = QFormLayout()
        # Tipo do Documento com descrição
        self.tipo_doc = QComboBox()
        self.tipo_doc.addItems([
            "01 - Dispensado",
            "02 - Com emissão",
            "03 - Sem emissão"
        ])
        form.addRow("Tipo do Documento:", self.tipo_doc)

        self.numero = QLineEdit()
        form.addRow("Número do Documento:", self.numero)

        # Série do Documento (obrigatória se Tipo=02)
        self.serie = QLineEdit()
        form.addRow("Série do Documento:", self.serie)

        self.data = QDateEdit()
        self.data.setCalendarPopup(True)
        self.data.setDate(QDate.currentDate())
        form.addRow("Data da Prestação:", self.data)

        # Tributação do Serviço com descrição
        self.tributacao = QComboBox()
        self.tributacao.addItems([
            "T - Operação Normal",
            "I - Imune",
            "J - ISS Suspenso"
        ])
        form.addRow("Tributação do Serviço:", self.tributacao)

        self.cod_servico = QLineEdit()
        form.addRow("Código do Serviço:", self.cod_servico)

        self.subitem = QLineEdit()
        self.subitem.setPlaceholderText("Ex: 0107 para classificação 1.07")
        form.addRow("Item/Subitem:", self.subitem)
        # Conectar o evento de mudança de texto
        self.subitem.textChanged.connect(self._on_subitem_changed)

        self.valor_nota = QLineEdit()
        form.addRow("Valor Total da Nota (R$):", self.valor_nota)

        # Alíquota (opcional, formato: 0500 para 5%)
        self.aliquota = QLineEdit()
        self.aliquota.setPlaceholderText("Ex: 0500 para 5% (opcional)")
        form.addRow("Alíquota (%):", self.aliquota)

        self.iss_retido = QCheckBox()
        form.addRow("ISS Retido pelo Tomador:", self.iss_retido)

        # Tipo de Prestador com descrição
        self.tipo_prestador = QComboBox()
        self.tipo_prestador.addItems([
            "1 - CPF",
            "2 - CNPJ",
            "3 - Exterior"
        ])
        form.addRow("Tipo de Prestador:", self.tipo_prestador)

        self.cnpj_prest = QLineEdit()
        form.addRow("CNPJ do Prestador:", self.cnpj_prest)

        # Regime de Tributação com descrição
        self.regime = QComboBox()
        self.regime.addItems([
            "0 - Normal/SN-DAMSP",
            "4 - SN-DAS",
            "5 - MEI"
        ])
        form.addRow("Regime de Tributação:", self.regime)

        self.discriminacao = QTextEdit()
        form.addRow("Discriminação dos Serviços:", self.discriminacao)

        # Buttons
        btns = QHBoxLayout()
        ok = QPushButton("OK")
        cancel = QPushButton("Cancelar")
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        btns.addWidget(ok)
        btns.addWidget(cancel)

        layout.addLayout(form)
        layout.addLayout(btns)
        self.setLayout(layout)

    def _on_subitem_changed(self, text: str):
        """Chamado quando o texto do campo subitem muda"""
        if not text.strip():
            return
        
        # Remove zeros à esquerda para buscar a classificação
        classificacao = text.lstrip('0')
        if not classificacao:
            return
        
        # Busca o CNAE correspondente
        if classificacao in self.aliquotas_data:
            cnae = self.aliquotas_data[classificacao]
            # Adiciona zeros à esquerda se o CNAE tiver menos de 5 dígitos
            if len(cnae) < 5:
                cnae_padded = cnae.zfill(5)
            else:
                cnae_padded = cnae
            
            # Preenche o campo código do serviço
            self.cod_servico.setText(cnae_padded)

    def get_data(self) -> Dict[str, Any]:
        # Extrai somente o código antes do ' - '
        tipo_doc = self.tipo_doc.currentText().split(' - ')[0]
        tributacao = self.tributacao.currentText().split(' - ')[0]
        tipo_prestador = self.tipo_prestador.currentText().split(' - ')[0]
        regime = self.regime.currentText().split(' - ')[0]
        return {
            'tipo_doc': tipo_doc,
            'numero': self.numero.text(),
            'serie': self.serie.text(),
            'data': self.data.date().toString('yyyyMMdd'),
            'tributacao': tributacao,
            'cod_servico': self.cod_servico.text(),
            'subitem': self.subitem.text(),
            'valor_nota': self.valor_nota.text(),
            'aliquota': self.aliquota.text(),
            'iss_retido': '1' if self.iss_retido.isChecked() else '2',
            'tipo_prestador': tipo_prestador,
            'cnpj_prest': self.cnpj_prest.text(),
            'regime': regime,
            'discriminacao': self.discriminacao.toPlainText().replace('\n', '|')
        }

    def _load_data(self, data: Dict[str, Any]):
        """Carrega dados existentes nos campos"""
        # Tipo do documento
        for i in range(self.tipo_doc.count()):
            if self.tipo_doc.itemText(i).startswith(data['tipo_doc']):
                self.tipo_doc.setCurrentIndex(i)
                break
        
        # Número e série
        self.numero.setText(data['numero'])
        self.serie.setText(data.get('serie', ''))
        
        # Data
        if 'data' in data:
            date = QDate.fromString(data['data'], 'yyyyMMdd')
            if date.isValid():
                self.data.setDate(date)
        
        # Tributação
        for i in range(self.tributacao.count()):
            if self.tributacao.itemText(i).startswith(data['tributacao']):
                self.tributacao.setCurrentIndex(i)
                break
        
        # Código do serviço e subitem
        self.cod_servico.setText(data['cod_servico'])
        self.subitem.setText(data['subitem'])
        
        # Valor da nota
        self.valor_nota.setText(data['valor_nota'])

        # Alíquota
        self.aliquota.setText(data.get('aliquota', ''))
        
        # ISS retido
        self.iss_retido.setChecked(data['iss_retido'] == '1')
        
        # Tipo de prestador
        for i in range(self.tipo_prestador.count()):
            if self.tipo_prestador.itemText(i).startswith(data['tipo_prestador']):
                self.tipo_prestador.setCurrentIndex(i)
                break
        
        # CNPJ do prestador
        self.cnpj_prest.setText(data['cnpj_prest'])
        
        # Regime de tributação
        for i in range(self.regime.count()):
            if self.regime.itemText(i).startswith(data['regime']):
                self.regime.setCurrentIndex(i)
                break
        
        # Discriminação
        self.discriminacao.setPlainText(data['discriminacao'].replace('|', '\n'))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Lote NFTS")
        self.resize(800, 600)
        self.notes: List[Dict[str, Any]] = []
        self._setup_ui()

    def _setup_ui(self):
        container = QWidget()
        layout = QVBoxLayout()
        # CCM selector
        hl = QHBoxLayout()
        hl.addWidget(QLabel("Contribuinte (CCM):"))
        self.ccm_combo = QComboBox()
        self.ccm_combo.addItem("4.165.071-9 – IM Filial", "41650719")
        self.ccm_combo.addItem("7.661.274-0 – IM Matriz", "76612740")
        hl.addWidget(self.ccm_combo)
        layout.addLayout(hl)

        # Notes table
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Tipo", "Série", "Número", "Data", "Serviço", "Alíquota", "Valor(R$)"])
        layout.addWidget(self.table)

        # Buttons
        btns = QHBoxLayout()
        self.add = QPushButton("Adicionar Nota")
        self.edit = QPushButton("Editar Nota Selecionada")
        self.remove = QPushButton("Remover Nota Selecionada")
        self.generate = QPushButton("Gerar Arquivo")
        btns.addWidget(self.add)
        btns.addWidget(self.edit)
        btns.addWidget(self.remove)
        btns.addStretch()
        btns.addWidget(self.generate)
        layout.addLayout(btns)

        container.setLayout(layout)
        self.setCentralWidget(container)

        # Signals
        self.add.clicked.connect(self._add)
        self.edit.clicked.connect(self._edit)
        self.remove.clicked.connect(self._remove)
        self.generate.clicked.connect(self._generate)

    def _add(self):
        dlg = NoteDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            self.notes.append(dlg.get_data())
            self._refresh()

    def _edit(self):
        i = self.table.currentRow()
        if i < 0:
            return
        dlg = NoteDialog(self)
        # Carregar dados existentes
        note = self.notes[i]
        dlg._load_data(note)
        if dlg.exec_() == QDialog.Accepted:
            self.notes[i] = dlg.get_data()
            self._refresh()

    def _remove(self):
        i = self.table.currentRow()
        if i < 0:
            return
        del self.notes[i]
        self._refresh()

    def _refresh(self):
        self.table.setRowCount(0)
        for n in self.notes:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(n['tipo_doc']))
            self.table.setItem(r, 1, QTableWidgetItem(n.get('serie', '')))
            self.table.setItem(r, 2, QTableWidgetItem(n['numero']))
            self.table.setItem(r, 3, QTableWidgetItem(n['data']))
            self.table.setItem(r, 4, QTableWidgetItem(n['cod_servico']))
            self.table.setItem(r, 5, QTableWidgetItem(n.get('aliquota', '') or '0'))
            self.table.setItem(r, 6, QTableWidgetItem(n['valor_nota']))

    def _validate_header(self) -> List[str]:
        errors: List[str] = []
        ccm = self.ccm_combo.currentData()
        if not re.fullmatch(r"\d{8}", str(ccm)):
            errors.append("CCM inválido: deve ter 8 dígitos numéricos.")
        if not self.notes:
            errors.append("Deve haver ao menos uma nota no lote.")
        return errors

    def _validate_notes(self) -> List[str]:
        errors: List[str] = []
        today = QDate.currentDate().toString('yyyyMMdd')
        for idx, n in enumerate(self.notes, start=1):
            if not re.fullmatch(r"\d{2}", n['tipo_doc']):
                errors.append(f"Nota {idx}: Tipo de documento deve ter 2 dígitos.")
            if not re.fullmatch(r"\d{1,12}", n['numero']):
                errors.append(f"Nota {idx}: Número deve ter até 12 dígitos.")
            # Validação da série (obrigatória se tipo_doc = '02')
            if n['tipo_doc'] == '02' and not n.get('serie', '').strip():
                errors.append(f"Nota {idx}: Série é obrigatória quando Tipo = 02.")
            if n.get('serie', '') and len(n['serie']) > 5:
                errors.append(f"Nota {idx}: Série deve ter até 5 caracteres.")
            if not re.fullmatch(r"\d{8}", n['data']):
                errors.append(f"Nota {idx}: Data deve estar no formato AAAAMMDD.")
            if n['data'] < '20000101' or n['data'] > today:
                errors.append(f"Nota {idx}: Data fora do intervalo permitido.")
            if n['tributacao'] not in ['T','I','J']:
                errors.append(f"Nota {idx}: Tributação inválida.")
            if not re.fullmatch(r"\d{1,5}", n['cod_servico']):
                errors.append(f"Nota {idx}: Código do serviço deve ter até 5 dígitos.")
            if not re.fullmatch(r"\d{1,4}", n['subitem']):
                errors.append(f"Nota {idx}: Subitem deve ter até 4 dígitos.")
            # Validação da alíquota (opcional, formato: 0500 para 5%)
            if n.get('aliquota', '').strip():
                if not re.fullmatch(r"\d{1,4}", n['aliquota']):
                    errors.append(f"Nota {idx}: Alíquota deve ter até 4 dígitos numéricos.")
                elif int(n['aliquota']) > 2500:  # Máximo 25%
                    errors.append(f"Nota {idx}: Alíquota não pode ser maior que 25% (2500).")
            if not re.fullmatch(r"\d+([\.,]\d{2})?", n['valor_nota']):
                errors.append(f"Nota {idx}: Valor deve ser numérico com até 2 casas decimais.")
            if n['iss_retido'] not in ['1','2']:
                errors.append(f"Nota {idx}: ISS retido inválido.")
            if n['tipo_prestador'] not in ['1','2','3']:
                errors.append(f"Nota {idx}: Tipo de prestador inválido.")
            if not re.fullmatch(r"\d{14}", n['cnpj_prest']):
                errors.append(f"Nota {idx}: CNPJ do prestador deve ter 14 dígitos.")
            if n['regime'] not in ['0','4','5']:
                errors.append(f"Nota {idx}: Regime de tributação inválido.")
            if len(n['discriminacao']) > 500:
                errors.append(f"Nota {idx}: Discriminação excede 500 caracteres.")
        return errors

    def _build_header(self) -> str:
        ccm = self.ccm_combo.currentData()
        header = '1' + '001'
        header += pad_left(ccm, 8)
        dates = [n['data'] for n in self.notes]
        start, end = min(dates), max(dates)
        header += start + end
        return header + '\r\n'

    def _build_details(self) -> List[str]:
        lines: List[str] = []
        for n in self.notes:
            l = '4'
            l += pad_left(n['tipo_doc'], 2)
            # Série do Documento (posições 4–8)
            l += pad_right(n.get('serie', ''), 5)
            l += pad_left(n['numero'], 12)
            l += n['data']
            l += 'N'
            l += n['tributacao']
            # valor do serviço em centavos
            serv = int(float(n['valor_nota'].replace(',', '.')) * 100)
            l += pad_left(serv, 15)
            # deduções
            l += pad_left(0, 15)
            l += pad_left(n['cod_servico'], 5)
            l += pad_left(n['subitem'], 4)
            # alíquota (usar valor informado ou zerar se não informado)
            aliquota = n.get('aliquota', '').strip()
            l += pad_left(aliquota if aliquota else '0', 4)
            l += n['iss_retido']
            l += n['tipo_prestador']
            l += pad_left(n['cnpj_prest'], 14)
            l += pad_right('', 8)  # inscrição municipal prestador (90–97)
            l += pad_right('', 75)  # nome/razão social prestador (98–172)
            # campos opcionais de prestador: tipo endereço, logradouro, nº, complemento,
            # bairro, cidade, UF, CEP e e-mail (posições 173–430)
            l += pad_right('', 258)
            l += pad_left('1', 1)  # tipo NFTS fixo
            l += n['regime']
            l += pad_left(0, 8)  # data pagamento
            l += pad_right(n['discriminacao'], 500)
            lines.append(l + '\r\n')
        return lines

    def _build_footer(self) -> str:
        count = len(self.notes)
        total = sum(int(float(n['valor_nota'].replace(',', '.')) * 100) for n in self.notes)
        footer = '9'
        footer += pad_left(count, 7)
        footer += pad_left(total, 15)
        footer += pad_left(0, 15)
        return footer + '\r\n'

    def _generate(self):
        # validações antes de gerar
        errors = self._validate_header() + self._validate_notes()
        if errors:
            QMessageBox.critical(self, "Erros de Validação", "\n".join(errors))
            return
        content = self._build_header()
        for line in self._build_details():
            content += line
        content += self._build_footer()
        fn, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo", "", "TXT (*.txt)")
        if fn:
            try:
                with open(fn, 'w', encoding='ISO-8859-1') as f:
                    f.write(content)
                QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso.")
            except Exception as e:
                QMessageBox.critical(self, "Erro ao salvar", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
