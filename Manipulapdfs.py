"""
Manipulapdfs.py
Suite de manipulacao de PDFs com interface grafica (Tkinter) e modo CLI.

Interface grafica: janela com abas para cada operacao, sem dependencias
extras (Tkinter acompanha o Python). Os processamentos rodam em thread
para nao travar a janela, com painel de log e barra de status.

Modo linha de comando (mantido da versao anterior):
    python Manipulapdfs.py --cli diagnostico
    python Manipulapdfs.py --cli juntar a.pdf b.pdf -o unificado.pdf

Compatibilidade: pypdf (>= 3) ou PyPDF2 legado, e PyMuPDF nos dois nomes
de import (pymupdf / fitz), com deteccao automatica em tempo de execucao.
"""

from __future__ import annotations

import argparse
import io
import os
import shutil
import sys
import threading
import tkinter as tk
from datetime import date
from importlib.metadata import PackageNotFoundError, version as _versao_metadata
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

# ============================================================
# 0. Deteccao de versoes e imports adaptativos
# ============================================================

def _versao_pacote(*nomes: str) -> str:
    """Retorna a versao instalada do primeiro nome de pacote encontrado."""
    for nome in nomes:
        try:
            return _versao_metadata(nome)
        except PackageNotFoundError:
            continue
    return "nao detectada"

def _importar_pdf():
    """Importa pypdf (preferido) ou PyPDF2 (legado), normalizando a API."""
    try:
        import pypdf as _modulo
        from pypdf import PdfReader, PdfWriter
        try:
            from pypdf.errors import PdfReadError
            origem_erro = "pypdf.errors"
        except ImportError:
            from pypdf import PdfReadError
            origem_erro = "pypdf (raiz)"
        return "pypdf", _modulo, PdfReader, PdfWriter, PdfReadError, origem_erro
    except ImportError:
        try:
            import PyPDF2 as _modulo
        except ImportError:
            raise ImportError(
                "Nenhuma biblioteca de PDF encontrada. "
                "Instale com: pip install pypdf"
            )
        try:
            from PyPDF2 import PdfReader, PdfWriter
        except ImportError:
            from PyPDF2 import (
                PdfFileReader as PdfReader,
                PdfFileWriter as PdfWriter,
            )
        try:
            from PyPDF2 import PdfReadError
            origem_erro = "PyPDF2 (raiz)"
        except ImportError:
            from PyPDF2.errors import PdfReadError
            origem_erro = "PyPDF2.errors"
        return "PyPDF2", _modulo, PdfReader, PdfWriter, PdfReadError, origem_erro

def _importar_pymupdf():
    """Importa PyMuPDF nos dois nomes possiveis: pymupdf ou fitz."""
    try:
        import pymupdf
        return "pymupdf", pymupdf
    except ImportError:
        import fitz as pymupdf
        return "fitz", pymupdf

BIBLIOTECA_PDF, _MODULO_PDF, PdfReader, PdfWriter, PdfReadError, ORIGEM_PDF_ERRO = (
    _importar_pdf()
)
BIBLIOTECA_MUPDF, pymupdf = _importar_pymupdf()

VERSAO_PDF = _versao_pacote("pypdf", "PyPDF2")
VERSAO_MUPDF = _versao_pacote("pymupdf", "PyMuPDF")

if BIBLIOTECA_PDF == "PyPDF2":
    print(
        "Aviso: PyPDF2 foi descontinuado. "
        "Recomendado: pip install --upgrade pypdf",
        file=sys.stderr,
    )

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_OK = True
except ImportError:
    PDF2IMAGE_OK = False

try:
    import pytesseract
    TESSERACT_OK = True
except ImportError:
    TESSERACT_OK = False

import pdfplumber
from reportlab.pdfgen import canvas
import img2pdf

A4_LARGURA = 595
A4_ALTURA = 842

# ============================================================
# 0.1 Helpers de compatibilidade entre versoes da API
# ============================================================

def _total_paginas(reader) -> int:
    if hasattr(reader, "getNumPages"):
        return reader.getNumPages()
    return len(reader.pages)

def _paginas_de(reader) -> list:
    if hasattr(reader, "pages"):
        return list(reader.pages)
    return [reader.getPage(i) for i in range(_total_paginas(reader))]

def _esta_criptografado(reader) -> bool:
    if hasattr(reader, "is_encrypted"):
        return reader.is_encrypted
    return reader.isEncrypted

def _mesclar_pagina(pagina, outra) -> None:
    if hasattr(pagina, "merge_page"):
        pagina.merge_page(outra)
    else:
        pagina.mergePage(outra)

def _adicionar_pagina_branca(writer) -> None:
    if hasattr(writer, "add_blank_page"):
        writer.add_blank_page(width=A4_LARGURA, height=A4_ALTURA)
    else:
        writer.addBlankPage(A4_LARGURA, A4_ALTURA)

def abrir_pdf_seguro(caminho: str, senha: str | None = None) -> PdfReader:
    """Abre um PDF tratando senha e arquivos corrompidos de forma limpa."""
    try:
        reader = PdfReader(caminho)
    except PdfReadError:
        raise ValueError(f"'{caminho}' nao e um PDF valido ou esta corrompido.")
    if _esta_criptografado(reader):
        if not senha:
            raise PermissionError("PDF protegido por senha. Informe a senha.")
        if not reader.decrypt(senha):
            raise PermissionError("Senha incorreta para o PDF.")
    return reader

def _salvar_writer(writer: PdfWriter, saida: str) -> None:
    with open(saida, "wb") as f:
        writer.write(f)

# ============================================================
# 1. Juncao, separacao e extracao
# ============================================================

def juntar_pdfs(lista_arquivos: list[str], saida: str, senha: str | None = None) -> None:
    """Junta todos os PDFs da lista na ordem informada."""
    writer = PdfWriter()
    for caminho in lista_arquivos:
        reader = abrir_pdf_seguro(caminho, senha)
        for pagina in _paginas_de(reader):
            writer.add_page(pagina)
    _salvar_writer(writer, saida)

def extrair_paginas(arquivo: str, paginas: list[int], saida: str, senha: str | None = None) -> None:
    """Extrai paginas especificas para um novo PDF (numeracao 1-based)."""
    reader = abrir_pdf_seguro(arquivo, senha)
    total = _total_paginas(reader)
    writer = PdfWriter()
    for numero in paginas:
        if numero < 1 or numero > total:
            raise ValueError(f"Pagina {numero} fora do intervalo (1 a {total}).")
        writer.add_page(_paginas_de(reader)[numero - 1])
    _salvar_writer(writer, saida)

def dividir_um_por_pagina(arquivo: str, prefixo_saida: str, senha: str | None = None) -> None:
    """Gera um PDF isolado para cada pagina do documento original."""
    reader = abrir_pdf_seguro(arquivo, senha)
    for indice, pagina in enumerate(_paginas_de(reader), start=1):
        writer = PdfWriter()
        writer.add_page(pagina)
        _salvar_writer(writer, f"{prefixo_saida}_pag_{indice:03d}.pdf")

def extrair_texto(pdf_entrada: str, txt_saida: str, senha: str | None = None) -> None:
    """Extrai texto preservando a ordem do layout (pdfplumber)."""
    with pdfplumber.open(pdf_entrada, password=senha) as pdf:
        with open(txt_saida, "w", encoding="utf-8") as f:
            for numero, pagina in enumerate(pdf.pages, start=1):
                texto = pagina.extract_text() or ""
                f.write(f"--- Pagina {numero} ---\n")
                f.write(texto + "\n\n")

def extrair_texto_rapido(pdf_entrada: str, txt_saida: str, senha: str | None = None) -> None:
    """Versao mais rapida com PyMuPDF, para documentos longos."""
    doc = pymupdf.open(pdf_entrada, password=senha)
    with open(txt_saida, "w", encoding="utf-8") as f:
        for pagina in doc:
            f.write(pagina.get_text())
            f.write("\n\n")
    doc.close()

def extrair_tabelas(pdf_entrada: str, senha: str | None = None) -> None:
    """Localiza e imprime as tabelas encontradas em cada pagina (pdfplumber)."""
    with pdfplumber.open(pdf_entrada, password=senha) as pdf:
        for numero, pagina in enumerate(pdf.pages, start=1):
            tabelas = pagina.extract_tables()
            if tabelas:
                print(f"Pagina {numero}: {len(tabelas)} tabela(s)")
                for tabela in tabelas:
                    for linha in tabela:
                        print(linha)

# ============================================================
# 2. Edicao, marcas e anotacoes
# ============================================================

def redigir_termos(pdf_entrada: str, pdf_saida: str, termos: list[str], senha: str | None = None) -> None:
    """Remove permanentemente trechos que contem os termos indicados."""
    doc = pymupdf.open(pdf_entrada, password=senha)
    for pagina in doc:
        for termo in termos:
            for retangulo in pagina.search_for(termo):
                pagina.add_redact_annot(retangulo)
        pagina.apply_redactions()
    doc.save(pdf_saida, garbage=4, deflate=True)
    doc.close()

def substituir_texto(pdf_entrada: str, pdf_saida: str, antigo: str, novo: str, senha: str | None = None) -> None:
    """Substitui um trecho por outro, reconstruindo o conteudo no mesmo lugar."""
    doc = pymupdf.open(pdf_entrada, password=senha)
    for pagina in doc:
        retangulos = pagina.search_for(antigo)
        for ret in retangulos:
            pagina.add_redact_annot(ret, fill=(1, 1, 1))
        pagina.apply_redactions()
        for ret in retangulos:
            pagina.insert_text(
                pymupdf.Point(ret.x0, ret.y1 - 4),
                novo,
                fontsize=11,
                fontname="helv",
            )
    doc.save(pdf_saida, garbage=4, deflate=True)
    doc.close()

def gerar_pdf_marca_dagua(texto: str, pdf_saida: str) -> None:
    """Cria um PDF de marca dagua (texto rotacionado em cinza) com reportlab."""
    c = canvas.Canvas(pdf_saida, pagesize=(A4_LARGURA, A4_ALTURA))
    c.saveState()
    c.translate(A4_LARGURA / 2, A4_ALTURA / 2)
    c.rotate(45)
    c.setFillColorRGB(0.8, 0.8, 0.8)
    c.setFont("Helvetica-Bold", 48)
    c.drawCentredString(0, 0, texto)
    c.restoreState()
    c.save()

def aplicar_marca_dagua(pdf_entrada: str, pdf_marca: str, pdf_saida: str, senha: str | None = None) -> None:
    """Sobrepoe o conteudo do PDF de marca em todas as paginas."""
    reader = abrir_pdf_seguro(pdf_entrada, senha)
    marca = PdfReader(pdf_marca).pages[0]
    writer = PdfWriter()
    for pagina in _paginas_de(reader):
        _mesclar_pagina(pagina, marca)
        writer.add_page(pagina)
    _salvar_writer(writer, pdf_saida)

def adicionar_anotacoes(
    pdf_entrada: str,
    pdf_saida: str,
    termo_destaque: str | None = None,
    comentario: str | None = None,
    carimbo: str | None = None,
    senha: str | None = None,
) -> None:
    """Adiciona destaque, comentario flutuante e carimbo na primeira pagina."""
    doc = pymupdf.open(pdf_entrada, password=senha)
    pagina = doc[0]
    if termo_destaque:
        ocorrencias = pagina.search_for(termo_destaque)
        if ocorrencias:
            pagina.add_highlight_annot(ocorrencias[0])
    if comentario:
        pagina.add_text_annot(pymupdf.Point(72, 720), comentario)
    if carimbo:
        pagina.insert_text(
            pymupdf.Point(72, 100),
            carimbo,
            fontsize=16,
            fontname="helv",
            color=(0, 0.5, 0),
        )
    doc.save(pdf_saida, garbage=4, deflate=True)
    doc.close()

# ============================================================
# 3. Conversoes e OCR
# ============================================================

def pdf_para_imagens(pdf_entrada: str, prefixo_saida: str, dpi: int = 200, senha: str | None = None) -> None:
    """Converte cada pagina em um PNG (PyMuPDF, sem dependencia externa)."""
    doc = pymupdf.open(pdf_entrada, password=senha)
    for indice, pagina in enumerate(doc, start=1):
        pix = pagina.get_pixmap(dpi=dpi)
        pix.save(f"{prefixo_saida}_{indice:03d}.png")
    doc.close()

def imagens_para_pdf(lista_imagens: list[str], pdf_saida: str) -> None:
    """Converte imagens em PDF sem perdas (img2pdf)."""
    with open(pdf_saida, "wb") as f:
        f.write(img2pdf.convert(lista_imagens))

def texto_para_pdf(texto: str, pdf_saida: str) -> None:
    """Cria um PDF simples a partir de texto (reportlab), com quebra de pagina."""
    c = canvas.Canvas(pdf_saida, pagesize=(A4_LARGURA, A4_ALTURA))
    y = 800
    for linha in texto.splitlines():
        c.drawString(72, y, linha)
        y -= 14
        if y < 50:
            c.showPage()
            y = 800
    c.save()

def pdf_escaneado_para_texto(pdf_entrada: str, txt_saida: str, dpi: int = 300, senha: str | None = None) -> None:
    """OCR de PDF sem camada de texto (pdf2image + pytesseract)."""
    if not (PDF2IMAGE_OK and TESSERACT_OK):
        raise RuntimeError(
            "Instale pdf2image e pytesseract: pip install pdf2image pytesseract"
        )
    if not shutil.which("pdftoppm"):
        raise RuntimeError(
            "Binario poppler nao encontrado. "
            "Instale e adicione ao PATH (Linux: sudo apt install poppler-utils)."
        )
    if not shutil.which("tesseract"):
        raise RuntimeError(
            "Binario tesseract nao encontrado. "
            "Instale e adicione ao PATH (Linux: sudo apt install "
            "tesseract-ocr tesseract-ocr-por)."
        )
    imagens = convert_from_path(pdf_entrada, dpi=dpi, password=senha)
    with open(txt_saida, "w", encoding="utf-8") as f:
        for imagem in imagens:
            f.write(pytesseract.image_to_string(imagem, lang="por"))

def diagnostico() -> None:
    """Imprime as versoes detectadas e o estado de cada dependencia."""
    import platform

    linhas = [
        f"Python: {platform.python_version()} ({platform.system()} {platform.release()})",
        f"Biblioteca de PDF estrutural: {BIBLIOTECA_PDF} {VERSAO_PDF}",
        f"  PdfReadError importado de: {ORIGEM_PDF_ERRO}",
        f"PyMuPDF: {BIBLIOTECA_MUPDF} {VERSAO_MUPDF}",
        f"pdfplumber: {_versao_pacote('pdfplumber')}",
        f"reportlab: {_versao_pacote('reportlab')}",
        f"img2pdf: {_versao_pacote('img2pdf')}",
        f"Pillow: {_versao_pacote('pillow', 'Pillow')}",
        f"pdf2image: {_versao_pacote('pdf2image')} "
        f"(poppler: {'OK' if shutil.which('pdftoppm') else 'FALTANDO'})",
        f"pytesseract: {_versao_pacote('pytesseract')} "
        f"(tesseract: {'OK' if shutil.which('tesseract') else 'FALTANDO'})",
    ]
    print("\n".join(linhas))

def exemplo_completo() -> None:
    """Fluxo de teste: junta dois arquivos, extrai texto, carimba e gera imagem."""
    juntar_pdfs(["requerimento.pdf", "documentos.pdf"], "processo_unificado.pdf")
    extrair_texto("processo_unificado.pdf", "processo_texto.txt")
    doc = pymupdf.open("processo_unificado.pdf")
    doc[0].insert_text(
        pymupdf.Point(72, 70),
        f"Documento recebido em {date.today().strftime('%d/%m/%Y')}",
        fontsize=14,
        fontname="helv",
    )
    doc.save("processo_carimbado.pdf", garbage=4, deflate=True)
    doc.close()
    pdf_para_imagens("processo_carimbado.pdf", "capa_processo", dpi=200)
    print("Fluxo de exemplo concluido (processo_unificado.pdf, "
          "processo_texto.txt, processo_carimbado.pdf, capa_processo_001.png).")

# ============================================================
# 4. Interface grafica (Tkinter)
# ============================================================

class PdfToolApp:
    """Janela desktop com abas para cada operacao de PDF."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Manipulador de PDFs - SPTrans")
        self.root.geometry("1000x720")
        self.root.minsize(880, 620)

        self._ocupado = False
        self._senha_salva = tk.StringVar()

        self._montar_cabecalho()
        self._montar_notebook()
        self._montar_status_e_log()

    # ---------- construcao da interface ----------

    def _montar_cabecalho(self) -> None:
        topo = ttk.Frame(self.root, padding=(10, 8))
        topo.pack(fill=tk.X)
        ttk.Label(topo, text="Manipulador de PDFs",
                  font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT)
        ttk.Label(topo, text="  Senha (opcional, para PDFs protegidos):",
                  font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(20, 4))
        ttk.Entry(topo, textvariable=self._senha_salva, width=20,
                  show="*").pack(side=tk.LEFT)

    def _montar_notebook(self) -> None:
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        self._aba_juntar = ttk.Frame(self.notebook, padding=10)
        self._aba_separar = ttk.Frame(self.notebook, padding=10)
        self._aba_extrair = ttk.Frame(self.notebook, padding=10)
        self._aba_redigir = ttk.Frame(self.notebook, padding=10)
        self._aba_substituir = ttk.Frame(self.notebook, padding=10)
        self._aba_marca = ttk.Frame(self.notebook, padding=10)
        self._aba_anotar = ttk.Frame(self.notebook, padding=10)
        self._aba_conversoes = ttk.Frame(self.notebook, padding=10)
        self._aba_ocr = ttk.Frame(self.notebook, padding=10)
        self._aba_diag = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self._aba_juntar, text=" Juntar PDFs ")
        self.notebook.add(self._aba_separar, text=" Separar / Dividir ")
        self.notebook.add(self._aba_extrair, text=" Extrair Texto ")
        self.notebook.add(self._aba_redigir, text=" Redigir ")
        self.notebook.add(self._aba_substituir, text=" Substituir ")
        self.notebook.add(self._aba_marca, text=" Marca d'Agua ")
        self.notebook.add(self._aba_anotar, text=" Anotacoes ")
        self.notebook.add(self._aba_conversoes, text=" PDF <> Imagem ")
        self.notebook.add(self._aba_ocr, text=" OCR ")
        self.notebook.add(self._aba_diag, text=" Diagnostico ")

        self._montar_aba_juntar()
        self._montar_aba_separar()
        self._montar_aba_extrair()
        self._montar_aba_redigir()
        self._montar_aba_substituir()
        self._montar_aba_marca()
        self._montar_aba_anotar()
        self._montar_aba_conversoes()
        self._montar_aba_ocr()
        self._montar_aba_diag()

    def _montar_status_e_log(self) -> None:
        self.status_var = tk.StringVar(value="Pronto.")
        tk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN,
                 anchor=tk.W, padx=8).pack(fill=tk.X, side=tk.BOTTOM)

        painel = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        painel.pack(fill=tk.BOTH, expand=True)
        ttk.Label(painel, text="Log de execucao:",
                  font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        self.log_area = ScrolledText(painel, height=9, state=tk.DISABLED,
                                     font=("Consolas", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True)

    # ---------- helpers de campos ----------

    def _linha_arquivo(self, parent, texto, filetypes, para_salvar=False,
                       largura=58) -> tk.StringVar:
        """Cria linha com label, entry e botao 'Selecionar'. Retorna a var."""
        var = tk.StringVar()
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=texto, width=24).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=largura).pack(side=tk.LEFT,
                                                                padx=(4, 4))
        ttk.Button(frame, text="Selecionar", width=10,
                   command=lambda: self._navegar(var, filetypes,
                                                 para_salvar)).pack(side=tk.LEFT)
        return var

    def _navegar(self, var, filetypes, para_salvar) -> None:
        if para_salvar:
            caminho = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=filetypes)
        else:
            caminho = filedialog.askopenfilename(filetypes=filetypes)
        if caminho:
            var.set(caminho)

    def _botao_executar(self, parent, texto, comando) -> None:
        ttk.Button(parent, text=texto, command=comando).pack(
            anchor=tk.W, pady=(8, 2))

    # ---------- rotinas de execucao em thread ----------

    def _rodar(self, func, args=(), kwargs=None, sucesso_msg="Operacao concluida.") -> None:
        """Executa uma operacao em thread, sem travar a interface."""
        if self._ocupado:
            self._log("Uma operacao ja esta em andamento. Aguarde concluir.")
            return
        kwargs = kwargs or {}
        self._ocupado = True
        self.status_var.set("Processando...")
        threading.Thread(target=self._alvo, args=(func, args, kwargs,
                                                  sucesso_msg),
                         daemon=True).start()

    def _alvo(self, func, args, kwargs, sucesso_msg) -> None:
        buf = io.StringIO()
        stdout_original = sys.stdout
        try:
            sys.stdout = buf
            func(*args, **kwargs)
        except Exception as e:
            self.root.after(0, self._concluir, False, None, e)
        else:
            self.root.after(0, self._concluir, True, sucesso_msg, None)
        finally:
            sys.stdout = stdout_original
            saida_capturada = buf.getvalue()
            if saida_capturada.strip():
                self.root.after(0, self._log, saida_capturada.rstrip())

    def _concluir(self, sucesso, mensagem, erro) -> None:
        self._ocupado = False
        if sucesso:
            self.status_var.set("Concluido.")
            self._log(mensagem or "Operacao concluida.")
        else:
            self.status_var.set("Erro na operacao.")
            self._log(f"Erro: {erro}")

    def _log(self, texto: str) -> None:
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, texto + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def _senha(self) -> str | None:
        return self._senha_salva.get().strip() or None

    # ---------- abas ----------

    def _montar_aba_juntar(self) -> None:
        p = self._aba_juntar
        ttk.Label(p, text="Selecione os PDFs na ordem desejada e clique em "
                          "'Juntar'.").pack(anchor=tk.W)
        self._lista_juntar = tk.Listbox(p, selectmode=tk.EXTENDED, height=7)
        self._lista_juntar.pack(fill=tk.X, pady=6)

        botoes = ttk.Frame(p)
        botoes.pack(fill=tk.X)
        ttk.Button(botoes, text="Adicionar PDFs...",
                   command=self._adicionar_pdfs).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(botoes, text="Remover selecionado",
                   command=self._remover_pdfs).pack(side=tk.LEFT, padx=4)
        ttk.Button(botoes, text="Limpar lista",
                   command=lambda: self._lista_juntar.delete(0, tk.END)
                   ).pack(side=tk.LEFT, padx=4)

        self._var_juntar_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Juntar", self._executar_juntar)

    def _adicionar_pdfs(self) -> None:
        caminhos = filedialog.askopenfilenames(
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        for caminho in caminhos:
            self._lista_juntar.insert(tk.END, caminho)

    def _remover_pdfs(self) -> None:
        for indice in reversed(self._lista_juntar.curselection()):
            self._lista_juntar.delete(indice)

    def _executar_juntar(self) -> None:
        arquivos = list(self._lista_juntar.get(0, tk.END))
        saida = self._var_juntar_saida.get().strip()
        if not arquivos:
            messagebox.showwarning("Atencao", "Adicione ao menos um PDF.")
            return
        if not saida:
            messagebox.showwarning("Atencao", "Informe o caminho de saida.")
            return
        self._rodar(juntar_pdfs,
                    args=(arquivos, saida, self._senha()),
                    sucesso_msg=f"{len(arquivos)} PDF(s) juntados em {saida}.")

    def _montar_aba_separar(self) -> None:
        p = self._aba_separar
        ttk.Label(p, text="Extrair paginas especificas (ex.: 1,3,5-7):",
                  font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        self._var_sep_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        self._var_sep_paginas = tk.StringVar(value="1")
        linha = ttk.Frame(p)
        linha.pack(fill=tk.X, pady=2)
        ttk.Label(linha, text="Paginas:", width=24).pack(side=tk.LEFT)
        ttk.Entry(linha, textvariable=self._var_sep_paginas,
                  width=20).pack(side=tk.LEFT, padx=4)
        self._var_sep_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Extrair paginas", self._executar_extrair_paginas)

        ttk.Separator(p).pack(fill=tk.X, pady=10)

        ttk.Label(p, text="Dividir em um PDF por pagina:", font=("Segoe UI",
                  9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        self._var_div_pasta = tk.StringVar()
        linha2 = ttk.Frame(p)
        linha2.pack(fill=tk.X, pady=2)
        ttk.Label(linha2, text="Pasta de destino:", width=24).pack(side=tk.LEFT)
        ttk.Entry(linha2, textvariable=self._var_div_pasta,
                  width=34).pack(side=tk.LEFT, padx=4)
        ttk.Button(linha2, text="Selecionar", width=10,
                   command=self._escolher_pasta_div).pack(side=tk.LEFT)
        self._var_div_prefixo = tk.StringVar(value="pagina")
        linha3 = ttk.Frame(p)
        linha3.pack(fill=tk.X, pady=2)
        ttk.Label(linha3, text="Prefixo dos arquivos:", width=24).pack(side=tk.LEFT)
        ttk.Entry(linha3, textvariable=self._var_div_prefixo,
                  width=20).pack(side=tk.LEFT, padx=4)
        self._botao_executar(p, "Dividir por pagina", self._executar_dividir)

    def _escolher_pasta_div(self) -> None:
        pasta = filedialog.askdirectory()
        if pasta:
            self._var_div_pasta.set(pasta)

    def _executar_extrair_paginas(self) -> None:
        entrada = self._var_sep_entrada.get().strip()
        saida = self._var_sep_saida.get().strip()
        if not entrada or not saida:
            messagebox.showwarning("Atencao", "Informe entrada e saida.")
            return
        try:
            paginas = [int(x.strip()) for x in
                       self._var_sep_paginas.get().replace(" ", "").split(",")
                       if x.strip()]
        except ValueError:
            messagebox.showwarning("Atencao",
                                   "Paginas invalidas. Use formato 1,3,5.")
            return
        self._rodar(extrair_paginas,
                    args=(entrada, paginas, saida, self._senha()),
                    sucesso_msg=f"Paginas extraidas em {saida}.")

    def _executar_dividir(self) -> None:
        entrada = self._var_sep_entrada.get().strip()
        pasta = self._var_div_pasta.get().strip()
        prefixo = self._var_div_prefixo.get().strip() or "pagina"
        if not entrada or not pasta:
            messagebox.showwarning("Atencao", "Informe a entrada e a pasta.")
            return
        caminho_base = os.path.join(pasta, prefixo)
        self._rodar(dividir_um_por_pagina,
                    args=(entrada, caminho_base, self._senha()),
                    sucesso_msg=f"Paginas geradas em {caminho_base}_pag_*.pdf")

    def _montar_aba_extrair(self) -> None:
        p = self._aba_extrair
        self._var_ext_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        self._var_ext_saida = self._linha_arquivo(
            p, "Arquivo .txt de saida:",
            [("Texto", "*.txt"), ("Todos", "*.*")], para_salvar=True)
        self._var_ext_rapido = tk.BooleanVar(value=False)
        ttk.Checkbutton(p, text="Usar extracao rapida (PyMuPDF, otimo para "
                                "documentos longos)",
                        variable=self._var_ext_rapido).pack(anchor=tk.W, pady=4)
        self._botao_executar(p, "Extrair texto", self._executar_extrair_texto)

    def _executar_extrair_texto(self) -> None:
        entrada = self._var_ext_entrada.get().strip()
        saida = self._var_ext_saida.get().strip()
        if not entrada or not saida:
            messagebox.showwarning("Atencao", "Informe entrada e saida.")
            return
        if self._var_ext_rapido.get():
            self._rodar(extrair_texto_rapido,
                        args=(entrada, saida, self._senha()),
                        sucesso_msg=f"Texto extraido em {saida}.")
        else:
            self._rodar(extrair_texto,
                        args=(entrada, saida, self._senha()),
                        sucesso_msg=f"Texto extraido em {saida}.")

    def _montar_aba_redigir(self) -> None:
        p = self._aba_redigir
        ttk.Label(p, text="Remove permanentemente os trechos que contem os "
                          "termos (use um termo por linha).",
                  font=("Segoe UI", 9)).pack(anchor=tk.W)
        self._var_red_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        self._texto_red_termos = tk.Text(p, height=4, width=60)
        self._texto_red_termos.pack(fill=tk.X, pady=4)
        self._var_red_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Redigir", self._executar_redigir)

    def _executar_redigir(self) -> None:
        entrada = self._var_red_entrada.get().strip()
        saida = self._var_red_saida.get().strip()
        termos = [linha.strip() for linha in
                  self._texto_red_termos.get("1.0", tk.END).splitlines()
                  if linha.strip()]
        if not entrada or not saida or not termos:
            messagebox.showwarning("Atencao",
                                   "Informe entrada, saida e ao menos um termo.")
            return
        self._rodar(redigir_termos,
                    args=(entrada, saida, termos, self._senha()),
                    sucesso_msg=f"Termos redigidos em {saida}.")

    def _montar_aba_substituir(self) -> None:
        p = self._aba_substituir
        self._var_sub_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        linha = ttk.Frame(p)
        linha.pack(fill=tk.X, pady=2)
        ttk.Label(linha, text="Texto original:", width=24).pack(side=tk.LEFT)
        self._var_sub_antigo = tk.StringVar()
        ttk.Entry(linha, textvariable=self._var_sub_antigo,
                  width=40).pack(side=tk.LEFT, padx=4)
        linha2 = ttk.Frame(p)
        linha2.pack(fill=tk.X, pady=2)
        ttk.Label(linha2, text="Texto novo:", width=24).pack(side=tk.LEFT)
        self._var_sub_novo = tk.StringVar()
        ttk.Entry(linha2, textvariable=self._var_sub_novo,
                  width=40).pack(side=tk.LEFT, padx=4)
        self._var_sub_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Substituir", self._executar_substituir)

    def _executar_substituir(self) -> None:
        entrada = self._var_sub_entrada.get().strip()
        saida = self._var_sub_saida.get().strip()
        antigo = self._var_sub_antigo.get().strip()
        novo = self._var_sub_novo.get()
        if not entrada or not saida or not antigo:
            messagebox.showwarning("Atencao",
                                   "Informe entrada, saida e texto original.")
            return
        self._rodar(substituir_texto,
                    args=(entrada, saida, antigo, novo, self._senha()),
                    sucesso_msg=f"Substituicao aplicada em {saida}.")

    def _montar_aba_marca(self) -> None:
        p = self._aba_marca
        self._var_marc_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        linha = ttk.Frame(p)
        linha.pack(fill=tk.X, pady=2)
        ttk.Label(linha, text="Texto da marca:", width=24).pack(side=tk.LEFT)
        self._var_marc_texto = tk.StringVar(value="COPIA NAO CONTROLADA")
        ttk.Entry(linha, textvariable=self._var_marc_texto,
                  width=40).pack(side=tk.LEFT, padx=4)
        self._var_marc_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Aplicar marca dagua",
                             self._executar_marca_dagua)

    def _executar_marca_dagua(self) -> None:
        entrada = self._var_marc_entrada.get().strip()
        saida = self._var_marc_saida.get().strip()
        texto = self._var_marc_texto.get().strip()
        if not entrada or not saida or not texto:
            messagebox.showwarning("Atencao",
                                   "Informe entrada, saida e texto da marca.")
            return
        marca_tmp = "marca_dagua_tmp.pdf"

        def _fluxo():
            try:
                gerar_pdf_marca_dagua(texto, marca_tmp)
                aplicar_marca_dagua(entrada, marca_tmp, saida,
                                    senha=self._senha())
            finally:
                if os.path.exists(marca_tmp):
                    os.remove(marca_tmp)

        self._rodar(_fluxo, sucesso_msg=f"Marca dagua aplicada em {saida}.")

    def _montar_aba_anotar(self) -> None:
        p = self._aba_anotar
        self._var_ano_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        campos = [("Termo a destacar:", "destaque"),
                  ("Comentario flutuante:", "comentario"),
                  ("Carimbo (ex.: APROVADO):", "carimbo")]
        for rotulo, atributo in campos:
            linha = ttk.Frame(p)
            linha.pack(fill=tk.X, pady=2)
            ttk.Label(linha, text=rotulo, width=24).pack(side=tk.LEFT)
            var = tk.StringVar()
            ttk.Entry(linha, textvariable=var, width=40).pack(side=tk.LEFT,
                                                              padx=4)
            setattr(self, f"_var_ano_{atributo}", var)
        self._var_ano_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Adicionar anotacoes",
                             self._executar_anotar)

    def _executar_anotar(self) -> None:
        entrada = self._var_ano_entrada.get().strip()
        saida = self._var_ano_saida.get().strip()
        if not entrada or not saida:
            messagebox.showwarning("Atencao", "Informe entrada e saida.")
            return
        self._rodar(adicionar_anotacoes,
                    args=(entrada, saida,
                          self._var_ano_destaque.get().strip() or None,
                          self._var_ano_comentario.get().strip() or None,
                          self._var_ano_carimbo.get().strip() or None,
                          self._senha()),
                    sucesso_msg=f"Anotacoes salvas em {saida}.")

    def _montar_aba_conversoes(self) -> None:
        p = self._aba_conversoes
        ttk.Label(p, text="PDF para imagens (uma imagem por pagina):",
                  font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        self._var_conv_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        self._var_conv_pasta = tk.StringVar()
        linha = ttk.Frame(p)
        linha.pack(fill=tk.X, pady=2)
        ttk.Label(linha, text="Pasta de destino:", width=24).pack(side=tk.LEFT)
        ttk.Entry(linha, textvariable=self._var_conv_pasta,
                  width=34).pack(side=tk.LEFT, padx=4)
        ttk.Button(linha, text="Selecionar", width=10,
                   command=self._escolher_pasta_conv).pack(side=tk.LEFT)
        self._var_conv_prefixo = tk.StringVar(value="pagina")
        linha2 = ttk.Frame(p)
        linha2.pack(fill=tk.X, pady=2)
        ttk.Label(linha2, text="Nome base:", width=24).pack(side=tk.LEFT)
        ttk.Entry(linha2, textvariable=self._var_conv_prefixo,
                  width=20).pack(side=tk.LEFT, padx=4)
        self._var_conv_dpi = tk.IntVar(value=200)
        ttk.Label(linha2, text="DPI:").pack(side=tk.LEFT, padx=(12, 2))
        ttk.Spinbox(linha2, from_=72, to=600, textvariable=self._var_conv_dpi,
                    width=6).pack(side=tk.LEFT)
        self._botao_executar(p, "Converter PDF em imagens",
                             self._executar_pdf_para_imagem)

        ttk.Separator(p).pack(fill=tk.X, pady=10)

        ttk.Label(p, text="Imagens para PDF (sem perdas):",
                  font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        self._lista_img = tk.Listbox(p, selectmode=tk.EXTENDED, height=4)
        self._lista_img.pack(fill=tk.X, pady=4)
        botoes = ttk.Frame(p)
        botoes.pack(fill=tk.X)
        ttk.Button(botoes, text="Adicionar imagens...",
                   command=self._adicionar_imagens).pack(side=tk.LEFT,
                                                         padx=(0, 4))
        ttk.Button(botoes, text="Remover selecionado",
                   command=lambda: [self._lista_img.delete(i) for i in
                                    reversed(self._lista_img.curselection())]
                   ).pack(side=tk.LEFT)
        self._var_img_saida = self._linha_arquivo(
            p, "PDF de saida:", [("PDF", "*.pdf"), ("Todos", "*.*")],
            para_salvar=True)
        self._botao_executar(p, "Converter imagens em PDF",
                             self._executar_imagem_para_pdf)

    def _escolher_pasta_conv(self) -> None:
        pasta = filedialog.askdirectory()
        if pasta:
            self._var_conv_pasta.set(pasta)

    def _adicionar_imagens(self) -> None:
        caminhos = filedialog.askopenfilenames(
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.tiff *.bmp"),
                       ("Todos", "*.*")])
        for caminho in caminhos:
            self._lista_img.insert(tk.END, caminho)

    def _executar_pdf_para_imagem(self) -> None:
        entrada = self._var_conv_entrada.get().strip()
        pasta = self._var_conv_pasta.get().strip()
        prefixo = self._var_conv_prefixo.get().strip() or "pagina"
        dpi = self._var_conv_dpi.get()
        if not entrada or not pasta:
            messagebox.showwarning("Atencao", "Informe a entrada e a pasta.")
            return
        base = os.path.join(pasta, prefixo)
        self._rodar(pdf_para_imagens,
                    args=(entrada, base, dpi, self._senha()),
                    sucesso_msg=f"Imagens geradas em {base}_001.png, ...")

    def _executar_imagem_para_pdf(self) -> None:
        imagens = list(self._lista_img.get(0, tk.END))
        saida = self._var_img_saida.get().strip()
        if not imagens or not saida:
            messagebox.showwarning("Atencao", "Adicione imagens e informe a saida.")
            return
        self._rodar(imagens_para_pdf,
                    args=(imagens, saida),
                    sucesso_msg=f"{len(imagens)} imagem(ns) convertidas em {saida}.")

    def _montar_aba_ocr(self) -> None:
        p = self._aba_ocr
        ttk.Label(p, text="Reconhecimento de texto em PDFs escaneados "
                          "(requer poppler e Tesseract instalados).",
                  font=("Segoe UI", 9)).pack(anchor=tk.W)
        self._var_ocr_entrada = self._linha_arquivo(
            p, "PDF de entrada:", [("PDF", "*.pdf"), ("Todos", "*.*")])
        self._var_ocr_saida = self._linha_arquivo(
            p, "Arquivo .txt de saida:",
            [("Texto", "*.txt"), ("Todos", "*.*")], para_salvar=True)
        self._var_ocr_dpi = tk.IntVar(value=300)
        linha = ttk.Frame(p)
        linha.pack(fill=tk.X, pady=2)
        ttk.Label(linha, text="DPI:", width=24).pack(side=tk.LEFT)
        ttk.Spinbox(linha, from_=150, to=600, textvariable=self._var_ocr_dpi,
                    width=6).pack(side=tk.LEFT)
        self._botao_executar(p, "Executar OCR", self._executar_ocr)

    def _executar_ocr(self) -> None:
        entrada = self._var_ocr_entrada.get().strip()
        saida = self._var_ocr_saida.get().strip()
        if not entrada or not saida:
            messagebox.showwarning("Atencao", "Informe entrada e saida.")
            return
        self._rodar(pdf_escaneado_para_texto,
                    args=(entrada, saida, self._var_ocr_dpi.get(),
                          self._senha()),
                    sucesso_msg=f"OCR concluido em {saida}.")

    def _montar_aba_diag(self) -> None:
        p = self._aba_diag
        ttk.Label(p, text="Mostra as versoes detectadas e o estado das "
                          "dependencias (resultado no log abaixo).",
                  font=("Segoe UI", 9)).pack(anchor=tk.W)
        self._botao_executar(p, "Executar diagnostico",
                             lambda: self._rodar(diagnostico,
                                                 sucesso_msg="Diagnostico "
                                                             "concluido."))

    def mainloop(self) -> None:
        self.root.mainloop()

# ============================================================
# 5. Interface de linha de comando (modo --cli)
# ============================================================

def _montar_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Suite de manipulacao de PDFs com interface grafica. "
                    "Use sem argumentos para abrir a janela, ou subcomandos "
                    "abaixo para modo linha de comando."
    )
    parser.add_argument("--senha", default=None, help="Senha para PDFs protegidos")
    sub = parser.add_subparsers(dest="comando")

    p = sub.add_parser("diagnostico", help="Mostra versoes e dependencias")
    p = sub.add_parser("juntar", help="Junta varios PDFs em um so")
    p.add_argument("arquivos", nargs="+")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("separar", help="Extrai paginas especificas")
    p.add_argument("arquivo")
    p.add_argument("-p", "--paginas", required=True, help="Ex.: 1,3,5")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("dividir", help="Gera um PDF por pagina")
    p.add_argument("arquivo")
    p.add_argument("-o", "--saida", required=True, help="Prefixo de saida")

    p = sub.add_parser("extrair", help="Extrai texto para um .txt")
    p.add_argument("arquivo")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("tabelas", help="Localiza tabelas e imprime")
    p.add_argument("arquivo")

    p = sub.add_parser("redigir", help="Remove permanentemente termos")
    p.add_argument("arquivo")
    p.add_argument("-t", "--termos", nargs="+", required=True)
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("substituir", help="Substitui um trecho por outro")
    p.add_argument("arquivo")
    p.add_argument("antigo")
    p.add_argument("novo")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("marca-dagua", help="Aplica marca dagua em todas as paginas")
    p.add_argument("arquivo")
    p.add_argument("-t", "--texto", required=True)
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("anotar", help="Adiciona anotacoes na 1a pagina")
    p.add_argument("arquivo")
    p.add_argument("-d", "--destaque")
    p.add_argument("-c", "--comentario")
    p.add_argument("-s", "--carimbo")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("pdf-para-imagem", help="Converte paginas em PNG")
    p.add_argument("arquivo")
    p.add_argument("-o", "--saida", required=True)
    p.add_argument("--dpi", type=int, default=200)

    p = sub.add_parser("imagem-para-pdf", help="Converte imagens em PDF")
    p.add_argument("imagens", nargs="+")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("texto-para-pdf", help="Cria PDF a partir de um .txt")
    p.add_argument("arquivo")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("ocr", help="OCR de PDF escaneado")
    p.add_argument("arquivo")
    p.add_argument("-o", "--saida", required=True)

    p = sub.add_parser("exemplo", help="Executa o fluxo de exemplo")
    return parser

def main_cli() -> None:
    parser = _montar_parser()
    args = parser.parse_args()

    if args.comando is None:
        parser.print_help()
        return

    try:
        if args.comando == "diagnostico":
            diagnostico()
        elif args.comando == "juntar":
            juntar_pdfs(args.arquivos, args.saida, senha=args.senha)
        elif args.comando == "separar":
            paginas = [int(p.strip()) for p in args.paginas.split(",")]
            extrair_paginas(args.arquivo, paginas, args.saida,
                            senha=args.senha)
        elif args.comando == "dividir":
            dividir_um_por_pagina(args.arquivo, args.saida, senha=args.senha)
        elif args.comando == "extrair":
            extrair_texto(args.arquivo, args.saida, senha=args.senha)
        elif args.comando == "tabelas":
            extrair_tabelas(args.arquivo, senha=args.senha)
        elif args.comando == "redigir":
            redigir_termos(args.arquivo, args.saida, args.termos,
                           senha=args.senha)
        elif args.comando == "substituir":
            substituir_texto(args.arquivo, args.saida, args.antigo,
                             args.novo, senha=args.senha)
        elif args.comando == "marca-dagua":
            marca_tmp = "marca_dagua_tmp.pdf"
            gerar_pdf_marca_dagua(args.texto, marca_tmp)
            try:
                aplicar_marca_dagua(args.arquivo, marca_tmp, args.saida,
                                    senha=args.senha)
            finally:
                if os.path.exists(marca_tmp):
                    os.remove(marca_tmp)
        elif args.comando == "anotar":
            adicionar_anotacoes(args.arquivo, args.saida,
                                termo_destaque=args.destaque,
                                comentario=args.comentario,
                                carimbo=args.carimbo,
                                senha=args.senha)
        elif args.comando == "pdf-para-imagem":
            pdf_para_imagens(args.arquivo, args.saida, dpi=args.dpi,
                             senha=args.senha)
        elif args.comando == "imagem-para-pdf":
            imagens_para_pdf(args.imagens, args.saida)
        elif args.comando == "texto-para-pdf":
            with open(args.arquivo, "r", encoding="utf-8") as f:
                texto_para_pdf(f.read(), args.saida)
        elif args.comando == "ocr":
            pdf_escaneado_para_texto(args.arquivo, args.saida,
                                     senha=args.senha)
        elif args.comando == "exemplo":
            exemplo_completo()
        else:
            parser.print_help()
    except (ValueError, PermissionError, RuntimeError, ImportError) as e:
        print(f"Erro: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Erro inesperado: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if "--cli" in sys.argv or "-c" in sys.argv:
        sys.argv = [a for a in sys.argv if a not in ("--cli", "-c")]
        main_cli()
    else:
        PdfToolApp().mainloop()