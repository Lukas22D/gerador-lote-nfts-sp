import sys
import json
import re
from typing import Any, List, Dict

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QLineEdit, QDateEdit, QPushButton, QTableWidget,
    QTableWidgetItem, QDialog, QFormLayout, QTextEdit, QCheckBox,
    QFileDialog, QMessageBox
)
from PyQt5.QtCore import QDate

# Utility functions
def pad_left(value, width: int) -> str:
    return str(value).rjust(width, '0')

def pad_right(value, width: int) -> str:
    return str(value).ljust(width, ' ')

class NoteDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adicionar/Editar Nota")
        self.aliquotas_data = self._load_aliquotas()
        self._setup_ui()

    def _load_aliquotas(self) -> Dict[str, str]:
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

    def _setup_ui(self):
        layout = QVBoxLayout()
        form = QFormLayout()

        # Campos principais
        self.tipo_doc = QComboBox()
        self.tipo_doc.addItems(["01 - Dispensado", "02 - Com emissão", "03 - Sem emissão"])
        form.addRow("Tipo do Documento:", self.tipo_doc)

        self.numero = QLineEdit()
        form.addRow("Número do Documento:", self.numero)

        self.serie = QLineEdit()
        form.addRow("Série do Documento:", self.serie)

        self.data = QDateEdit()
        self.data.setCalendarPopup(True)
        self.data.setDate(QDate.currentDate())
        form.addRow("Data da Prestação:", self.data)

        self.tributacao = QComboBox()
        self.tributacao.addItems(["T - Operação Normal", "I - Imune", "J - ISS Suspenso"])
        form.addRow("Tributação do Serviço:", self.tributacao)

        self.cod_servico = QLineEdit()
        form.addRow("Código do Serviço:", self.cod_servico)

        self.subitem = QLineEdit()
        self.subitem.setPlaceholderText("Ex: 0107 para classificação 1.07")
        self.subitem.textChanged.connect(self._on_subitem_changed)
        form.addRow("Item/Subitem:", self.subitem)

        self.valor_nota = QLineEdit()
        form.addRow("Valor Total da Nota (R$):", self.valor_nota)

        self.aliquota = QLineEdit()
        self.aliquota.setPlaceholderText("Ex: 0500 para 5% (opcional)")
        form.addRow("Alíquota (%):", self.aliquota)

        self.iss_retido = QCheckBox()
        form.addRow("ISS Retido pelo Tomador:", self.iss_retido)

        self.tipo_prestador = QComboBox()
        self.tipo_prestador.addItems(["1 - CPF", "2 - CNPJ", "3 - Exterior"])
        form.addRow("Tipo de Prestador:", self.tipo_prestador)

        self.cnpj_prest = QLineEdit()
        form.addRow("CNPJ do Prestador:", self.cnpj_prest)

        self.regime = QComboBox()
        self.regime.addItems(["0 - Normal/SN-DAMSP", "4 - SN-DAS", "5 - MEI"])
        form.addRow("Regime de Tributação:", self.regime)

        # NOVOS campos: Cidade, UF e CEP
        self.cidade = QLineEdit()
        form.addRow("Cidade:", self.cidade)

        self.uf = QComboBox()
        self.uf.addItems([
            "", "AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG",
            "PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"
        ])
        form.addRow("UF:", self.uf)

        self.cep = QLineEdit()
        self.cep.setPlaceholderText("Somente números (8 dígitos)")
        form.addRow("CEP:", self.cep)

        # Discriminação
        self.discriminacao = QTextEdit()
        form.addRow("Discriminação dos Serviços:", self.discriminacao)

        # Botões
        btns = QHBoxLayout()
        ok = QPushButton("OK"); ok.clicked.connect(self.accept)
        cancel = QPushButton("Cancelar"); cancel.clicked.connect(self.reject)
        btns.addWidget(ok); btns.addWidget(cancel)

        layout.addLayout(form)
        layout.addLayout(btns)
        self.setLayout(layout)

    def _on_subitem_changed(self, text: str):
        # Aguarda pelo menos 4 dígitos antes de buscar
        if len(text) >= 4:
            # Remove zeros à esquerda para buscar a classificação
            cls = text.lstrip('0')
            if cls and cls in self.aliquotas_data:
                cnae = self.aliquotas_data[cls]
                # Adiciona zeros à esquerda se o CNAE tiver menos de 5 dígitos
                if len(cnae) < 5:
                    cnae_padded = cnae.zfill(5)
                else:
                    cnae_padded = cnae
                self.cod_servico.setText(cnae_padded)

    def get_data(self) -> Dict[str, Any]:
        return {
            'tipo_doc': self.tipo_doc.currentText().split(' - ')[0],
            'numero': self.numero.text(),
            'serie': self.serie.text(),
            'data': self.data.date().toString('yyyyMMdd'),
            'tributacao': self.tributacao.currentText().split(' - ')[0],
            'cod_servico': self.cod_servico.text(),
            'subitem': self.subitem.text(),
            'valor_nota': self.valor_nota.text(),
            'aliquota': self.aliquota.text(),
            'iss_retido': '1' if self.iss_retido.isChecked() else '2',
            'tipo_prestador': self.tipo_prestador.currentText().split(' - ')[0],
            'cnpj_prest': self.cnpj_prest.text(),
            'regime': self.regime.currentText().split(' - ')[0],
            'cidade': self.cidade.text(),
            'uf': self.uf.currentText(),
            'cep': self.cep.text(),
            'discriminacao': self.discriminacao.toPlainText().replace('\n', '|')
        }

    def _load_data(self, data: Dict[str, Any]):
        """Carrega dados existentes nos campos para edição"""
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
        
        # Campos de endereço
        self.cidade.setText(data.get('cidade', ''))
        
        # UF
        uf = data.get('uf', '')
        for i in range(self.uf.count()):
            if self.uf.itemText(i) == uf:
                self.uf.setCurrentIndex(i)
                break
        
        # CEP
        self.cep.setText(data.get('cep', ''))
        
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
        w = QWidget(); layout = QVBoxLayout(w)

        hl = QHBoxLayout()
        hl.addWidget(QLabel("Contribuinte (CCM):"))
        self.ccm_combo = QComboBox()
        self.ccm_combo.addItem("4.165.071-9 – IM Filial", "41650719")
        self.ccm_combo.addItem("7.661.274-0 – IM Matriz",   "76612740")
        hl.addWidget(self.ccm_combo)
        layout.addLayout(hl)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            "Tipo","Série","Número","Data","Serviço","Alíquota","Valor(R$)"
        ])
        layout.addWidget(self.table)

        btns = QHBoxLayout()
        for name, slot in [
            ("Adicionar Nota", self._add),
            ("Editar Nota",   self._edit),
            ("Remover Nota",  self._remove)
        ]:
            b = QPushButton(name); b.clicked.connect(slot); btns.addWidget(b)
        btns.addStretch()
        bgen = QPushButton("Gerar Arquivo"); bgen.clicked.connect(self._generate)
        btns.addWidget(bgen)
        layout.addLayout(btns)

        self.setCentralWidget(w)

    def _add(self):
        dlg = NoteDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            self.notes.append(dlg.get_data())
            self._refresh()

    def _edit(self):
        i = self.table.currentRow()
        if i < 0:
            QMessageBox.warning(self, "Aviso", "Selecione uma nota para editar.")
            return
        
        dlg = NoteDialog(self)
        # Carregar dados existentes
        dlg._load_data(self.notes[i])
        if dlg.exec_() == QDialog.Accepted:
            self.notes[i] = dlg.get_data()
            self._refresh()

    def _remove(self):
        i = self.table.currentRow()
        if i >= 0:
            del self.notes[i]
            self._refresh()

    def _refresh(self):
        self.table.setRowCount(0)
        for n in self.notes:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r,0, QTableWidgetItem(n['tipo_doc']))
            self.table.setItem(r,1, QTableWidgetItem(n['serie']))
            self.table.setItem(r,2, QTableWidgetItem(n['numero']))
            self.table.setItem(r,3, QTableWidgetItem(n['data']))
            self.table.setItem(r,4, QTableWidgetItem(n['cod_servico']))
            self.table.setItem(r,5, QTableWidgetItem(n['aliquota'] or '0'))
            self.table.setItem(r,6, QTableWidgetItem(n['valor_nota']))

    def _validate_header(self) -> List[str]:
        errs: List[str] = []
        ccm = self.ccm_combo.currentData()
        if not re.fullmatch(r"\d{8}", str(ccm)):
            errs.append("CCM inválido: deve ter 8 dígitos numéricos.")
        if not self.notes:
            errs.append("É necessário ao menos uma nota.")
        return errs

    def _validate_notes(self) -> List[str]:
        errs: List[str] = []
        today = QDate.currentDate().toString('yyyyMMdd')
        for idx, n in enumerate(self.notes, start=1):
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

    def _build_header(self) -> str:
        ccm = self.ccm_combo.currentData()
        header = '1' + '001' + pad_left(ccm,8)
        dates = [n['data'] for n in self.notes]
        header += min(dates) + max(dates)
        return header + '\r\n'

    def _build_details(self) -> List[str]:
        lines: List[str] = []
        for n in self.notes:
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

    def _build_footer(self) -> str:
        count = len(self.notes)
        total = sum(int(float(n['valor_nota'].replace(',','.'))*100) for n in self.notes)
        footer = '9'
        footer += pad_left(count,7)
        footer += pad_left(total,15)
        footer += pad_left(0,15)
        return footer + '\r\n'

    def _generate(self):
        errs = self._validate_header() + self._validate_notes()
        if errs:
            QMessageBox.critical(self, "Erros de Validação", "\n".join(errs))
            return

        content = self._build_header()
        for line in self._build_details():
            content += line
        content += self._build_footer()

        fn, _ = QFileDialog.getSaveFileName(
            self, "Salvar Arquivo", "", "Arquivo Texto (*.txt)"
        )
        if not fn:
            return
        if not fn.lower().endswith('.txt'):
            fn += '.txt'

        try:
            with open(fn, 'w', encoding='ISO-8859-1', newline='') as f:
                f.write(content)
            QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao salvar", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
