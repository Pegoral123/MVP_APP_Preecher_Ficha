"""
Gerador automático de Fichas de Matrícula (UNIFESO) - versão linha de comando
------------------------------------------------------------------------------
Lê a ata de alunos (PDF exportado do TOTVS) e a ficha-modelo (.docx),
e gera uma ficha preenchida (Nome + Matrícula) para cada aluno MATRICULADO,
juntando tudo em um único PDF pronto para impressão.

Requisito: Microsoft Word OU LibreOffice instalado na máquina.

Uso:
    gerar_fichas_cli.exe ata.pdf ficha_template.docx saida.pdf
"""
import sys
import os

# Proteção defensiva: se por algum motivo esse .exe for empacotado sem
# console (ex: --windowed) ou rodar num contexto sem stdout, evita o
# erro "'NoneType' object has no attribute 'write'".
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

from fichas_core import gerar_fichas_completo, FichasError


def main():
    if len(sys.argv) != 4:
        print("Uso: gerar_fichas_cli.exe <ata.pdf> <template.docx> "
              "<saida.pdf>")
        sys.exit(1)

    ata_pdf, template_docx, saida_pdf = sys.argv[1:4]

    try:
        gerar_fichas_completo(ata_pdf, template_docx, saida_pdf)
    except FichasError as e:
        print(f"\nERRO: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\nERRO inesperado: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()