"""
Gerador automático de Fichas de Matrícula (UNIFESO) - versão com interface gráfica
------------------------------------------------------------------------------------
Interface simples em Tkinter (já vem com o Python, não precisa instalar nada
extra): seleciona a ata (PDF), a ficha-modelo (DOCX), escolhe onde salvar o
PDF final, clica em "Gerar Fichas" e pronto.

Requisito: Microsoft Word OU LibreOffice instalado na máquina.
"""

# ---------------------------------------------------------------------------
# IMPORTANTE: isso precisa vir ANTES de qualquer outro import.
#
# Quando o .exe é empacotado com "pyinstaller --windowed", não existe console
# nenhum por trás da janela — e nesse caso o Windows/Python deixa
# sys.stdout e sys.stderr como None (não existe pra onde escrever).
#
# O docx2pdf usa uma barra de progresso (tqdm) que tenta escrever o
# andamento da conversão em sys.stdout/stderr. Como eles são None, dá o erro
# "'NoneType' object has no attribute 'write'" bem no meio da conversão
# pra PDF. É exatamente o erro que apareceu na tela.
#
# A correção é simples: se stdout/stderr vierem None, a gente troca por um
# "arquivo" que só descarta o que for escrito (os.devnull), então tqdm tem
# pra onde escrever e não quebra mais.
# ---------------------------------------------------------------------------
import sys
import os

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from fichas_core import gerar_fichas_completo, FichasError


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Fichas de Matrícula - UNIFESO")
        self.geometry("680x540")
        self.resizable(False, False)

        self.ata_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.saida_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 16, "pady": 6}

        header = tk.Label(
            self, text="Gerador de Fichas de Matrícula",
            font=("Segoe UI", 14, "bold")
        )
        header.pack(anchor="w", **pad)

        subtitle = tk.Label(
            self,
            text=(
                "Preenche Nome e Matrícula automaticamente em todas as fichas,\n"
                "a partir da ata de alunos exportada do TOTVS."
            ),
            font=("Segoe UI", 9), fg="#444", justify="left"
        )
        subtitle.pack(anchor="w", padx=16, pady=(0, 12))

        # --- Seletores de arquivo ---
        self._campo_arquivo(
            "1. Ata de alunos (PDF do TOTVS):",
            self.ata_path,
            self._escolher_ata,
        )
        self._campo_arquivo(
            "2. Ficha-modelo (.docx):",
            self.template_path,
            self._escolher_template,
        )
        self._campo_arquivo(
            "3. Salvar PDF final como:",
            self.saida_path,
            self._escolher_saida,
            salvar=True,
        )

        # --- Botão Gerar ---
        self.btn_gerar = tk.Button(
            self, text="Gerar Fichas", font=("Segoe UI", 11, "bold"),
            bg="#2e7d32", fg="white", height=2, command=self._on_gerar
        )
        self.btn_gerar.pack(fill="x", padx=16, pady=(16, 8))

        # --- Barra de progresso determinada ---
        progress_frame = tk.Frame(self)
        progress_frame.pack(fill="x", padx=16, pady=(0, 4))
        self.progress = ttk.Progressbar(
            progress_frame, mode="determinate", maximum=100
        )
        self.progress.pack(fill="x", expand=True)
        self.progress_label = tk.Label(
            progress_frame, text="", font=("Segoe UI", 8), fg="#666"
        )
        self.progress_label.pack(anchor="e")

        # --- Log de andamento ---
        log_label = tk.Label(
            self, text="Andamento:", font=("Segoe UI", 9, "bold")
        )
        log_label.pack(anchor="w", padx=16)

        self.log_box = tk.Text(
            self, height=14, font=("Consolas", 9), state="disabled"
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(4, 16))

    def _campo_arquivo(self, rotulo, variavel, comando_escolher,
                       salvar=False):
        frame = tk.Frame(self)
        frame.pack(fill="x", padx=16, pady=4)
        tk.Label(
            frame, text=rotulo, font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        subframe = tk.Frame(frame)
        subframe.pack(fill="x", pady=2)
        entry = tk.Entry(subframe, textvariable=variavel)
        entry.pack(side="left", fill="x", expand=True)
        btn = tk.Button(subframe, text="Procurar...",
                        command=comando_escolher)
        btn.pack(side="left", padx=(8, 0))

    # --- Handlers de seleção de arquivo ---
    def _escolher_ata(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a ata de alunos (PDF)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if caminho:
            self.ata_path.set(caminho)

    def _escolher_template(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a ficha-modelo (.docx)",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if caminho:
            self.template_path.set(caminho)

    def _escolher_saida(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar PDF final como",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if caminho:
            self.saida_path.set(caminho)

    # --- Log na textbox ---
    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    # --- Atualização da barra de progresso ---
    def _atualizar_progresso(self, atual, total):
        """Atualiza barra determinada de 0 a 100%."""
        pct = int((atual / total) * 100) if total > 0 else 0
        self.progress["value"] = pct
        self.progress_label["text"] = f"{atual} de {total} fichas"
        self.update_idletasks()

    # --- Callback recebido do core ---
    def _progress_callback(self, etapa, atual, total):
        """Callback chamado pelo core a cada ficha .docx gerada e ao
        iniciar/finalizar conversão PDF."""
        if etapa == "docx":
            self._atualizar_progresso(atual, total)
        elif etapa == "pdf":
            # Na conversão PDF não temos granularidade por ficha,
            # então mostramos 100% quando terminar
            if atual >= total:
                self._atualizar_progresso(total, total)

    # --- Ação principal ---
    def _on_gerar(self):
        if (not self.ata_path.get() or not self.template_path.get()
                or not self.saida_path.get()):
            messagebox.showwarning(
                "Campos faltando",
                "Selecione a ata, a ficha-modelo e o local "
                "para salvar o PDF final."
            )
            return

        # Desabilita botão, reseta progresso e log
        self.btn_gerar.config(state="disabled", text="Gerando...")
        self.progress["value"] = 0
        self.progress_label["text"] = ""
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        thread = threading.Thread(
            target=self._rodar_geracao, daemon=True
        )
        thread.start()

    def _rodar_geracao(self):
        # ---------------------------------------------------------------
        # CRÍTICO: em builds --windowed, o COM do Windows (usado pelo
        # Word via win32com) precisa ser inicializado explicitamente em
        # QUALQUER thread que não seja a principal. Sem CoInitialize(),
        # o Word não abre o documento corretamente e o objeto interno
        # vira None → "'NoneType' object has no attribute 'write'".
        #
        # docx2pdf já chama CoInitialize internamente, mas fazemos aqui
        # também como redundância defensiva. O CoUninitialize vai no
        # finally pra garantir que o COM seja liberado mesmo em erro.
        # ---------------------------------------------------------------
        import pythoncom
        pythoncom.CoInitialize()
        try:
            total = gerar_fichas_completo(
                self.ata_path.get(),
                self.template_path.get(),
                self.saida_path.get(),
                log=self._log,
                progress_callback=self._progress_callback,
            )
            self.after(0, lambda: self._finalizar_sucesso(total))
        except FichasError as e:
            self.after(0, lambda msg=str(e): self._finalizar_erro(msg))
        except Exception as e:
            self.after(
                0,
                lambda msg=str(e): self._finalizar_erro(
                    f"Erro inesperado: {msg}"
                ),
            )
        finally:
            pythoncom.CoUninitialize()

    def _finalizar_sucesso(self, total):
        self.progress["value"] = 100
        self.progress_label["text"] = f"{total} de {total} fichas"
        self.btn_gerar.config(state="normal", text="Gerar Fichas")
        messagebox.showinfo(
            "Concluído",
            f"{total} fichas geradas com sucesso!\n\n"
            f"Salvo em:\n{self.saida_path.get()}"
        )

    def _finalizar_erro(self, mensagem):
        self.progress["value"] = 0
        self.progress_label["text"] = ""
        self.btn_gerar.config(state="normal", text="Gerar Fichas")
        messagebox.showerror("Erro ao gerar fichas", mensagem)


if __name__ == "__main__":
    app = App()
    app.mainloop()