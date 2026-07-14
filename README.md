# Gerador de Fichas de Matrícula — UNIFESO

Automatiza o preenchimento das fichas de matrícula a partir da ata de alunos exportada do TOTVS.

**Problema resolvido:**
O preenchimento manual de centenas de fichas de matrícula (Nome + RA) era repetitivo, demorado e sujeito a erros de digitação.
Esta ferramenta elimina completamente esse trabalho manual — você seleciona a ata em PDF e o modelo da ficha em Word, e em segundos ela gera um PDF único com todas as fichas preenchidas e prontas para impressão.

---

## Funcionalidades

- **Extração automática** dos alunos (RA e Nome) do PDF "Alunos por Período" exportado do TOTVS
- **Preenchimento automático** do template `.docx` com os dados de cada aluno (placeholders `{{ nome }}` e `{{ matricula }}`)
- **Conversão para PDF** usando Microsoft Word **ou** LibreOffice como fallback
- **Junção de todas as fichas** em um único arquivo PDF pronto para impressão
- **Interface gráfica** simples (Tkinter) — não exige conhecimento técnico
- **Interface de linha de comando** para quem prefere script/integração

---

## Pré-requisitos

### Software necessário

| Conversor         | Obrigatório? | Observação                                  |
| ----------------- | ------------ | ------------------------------------------- |
| Microsoft Word    | Não¹         | Recomendado, já instalado na maioria das máquinas da instituição |
| LibreOffice       | Não¹         | Gratuito: [libreoffice.org](https://pt-br.libreoffice.org/) |

¹ *Pelo menos um dos dois precisa estar instalado.*

### Python 3.10+ (apenas para rodar a partir do código-fonte)

Se for usar o **executável `.exe`** gerado pelo PyInstaller, não é necessário ter Python instalado.

### Dependências Python

```
docxtpl>=0.16.7
pdfplumber>=0.10.0
pypdf>=4.0.0
docx2pdf>=0.1.8
python-docx>=1.1.0
pywin32>=306
```

Instale com:

```bash
pip install -r requirements.txt
```

---

## Como usar

### Interface Gráfica (recomendado para uso diário)

```bash
python gerar_fichas_gui_refatorado.py
```

1. Clique em **Procurar...** para selecionar a ata de alunos (PDF exportado do TOTVS)
2. Selecione a **ficha-modelo** `.docx` (deve conter `{{ nome }}` e `{{ matricula }}`)
3. Escolha **onde salvar** o PDF final
4. Clique em **Gerar Fichas**

A barra de progresso e o log mostram cada etapa do processo.

### Linha de Comando

```bash
python gerar_fichas_cli_refatorado.py ata.pdf ficha_template.docx saida.pdf
```

Parâmetros (obrigatórios e nesta ordem):
- `ata.pdf` — PDF "Alunos por Período" exportado do TOTVS
- `ficha_template.docx` — Arquivo Word com os placeholders `{{ nome }}` e `{{ matricula }}`
- `saida.pdf` — Caminho onde o PDF final será salvo

---

## Formato esperado dos arquivos de entrada

### Ata de alunos (PDF)

O PDF deve ser o relatório **"Alunos por Período"** exportado do sistema TOTVS.
O programa procura linhas no seguinte formato:

```
<RA> <NOME COMPLETO> MATRICULADO CIÊNCIA DA ...
```

Exemplo:
```
06019004 ABNER CAVALCANTI ATOUGUIA MATRICULADO CIÊNCIA DA COMPUTAÇÃO
```

Apenas alunos com status **MATRICULADO** são incluídos.

### Template da ficha (.docx)

O arquivo Word deve conter **obrigatoriamente** os placeholders:
- `{{ nome }}` — onde o nome completo do aluno será inserido
- `{{ matricula }}` — onde o RA do aluno será inserido

Os placeholders podem estar em qualquer lugar do documento (parágrafos, tabelas, caixas de texto, cabeçalhos, etc.).

---

## Estrutura do Projeto

```
MVP_APP_Preecher_Ficha/
├── fichas_core.py                  # Lógica central (extração, geração, conversão, junção)
├── gerar_fichas_gui_refatorado.py  # Interface gráfica (Tkinter)
├── gerar_fichas_cli_refatorado.py  # Interface de linha de comando
├── GeradorFichas.spec              # Spec do PyInstaller para gerar .exe
├── requirements.txt                # Dependências Python
└── README.md                       # Este arquivo
```

### `fichas_core.py`

Módulo central com toda a lógica de negócio, compartilhado pelas interfaces GUI e CLI.

**Pipeline de execução (`gerar_fichas_completo`)**:

| Etapa | Descrição                                                              |
| ----- | ---------------------------------------------------------------------- |
| 0/5   | Valida arquivos de entrada e permissões                                |
| 1/5   | Extrai alunos do PDF da ata                                            |
| 2/5   | Gera um `.docx` individual para cada aluno                             |
| 3/5   | Verifica se Word e/ou LibreOffice estão disponíveis                    |
| 4/5   | Converte todos os `.docx` para PDF (Word → LibreOffice como fallback)  |
| 5/5   | Junta todos os PDFs individuais em um único arquivo                    |

**Fallback de conversão**: tenta primeiro Microsoft Word (via `docx2pdf`); se falhar ou não estiver instalado, tenta LibreOffice (via `soffice --headless`). Se nenhum estiver disponível, exibe erro claro com instruções.

### `gerar_fichas_gui_refatorado.py`

Interface gráfica construída com Tkinter (biblioteca nativa do Python, sem dependências extras). Executa a geração em uma **thread separada** para não travar a interface.

### `gerar_fichas_cli_refatorado.py`

Versão de terminal. Útil para automação, scripts `.bat` ou integração com outros sistemas.

---

## Gerando o executável (.exe)

```bash
pip install pyinstaller
pyinstaller GeradorFichas.spec
```

O `.exe` será gerado na pasta `dist/`.

Para builds com interface gráfica (sem console), o spec já está configurado com `--windowed`.

---

## Solução de Problemas

### "Nenhum conversor PDF encontrado"
Instale o Microsoft Word ou o LibreOffice. Um dos dois é obrigatório para converter os `.docx` em PDF.

### "Nenhum aluno MATRICULADO encontrado na ata"
- Verifique se o PDF é realmente o relatório "Alunos por Período" do TOTVS
- Confirme que há alunos com status "MATRICULADO" no relatório
- Se o layout do TOTVS mudou, o regex de extração em `fichas_core.py` (`extrair_alunos`) precisa ser atualizado

### "Template inválido: faltam os placeholders"
O arquivo `.docx` precisa conter `{{ nome }}` e `{{ matricula }}` nos locais onde os dados devem ser preenchidos.

### Erro ao usar o .exe: "NoneType object has no attribute 'write'"
Esse erro já foi tratado no código. Ocorria em builds `--windowed` onde `sys.stdout`/`sys.stderr` ficam como `None`. O programa redireciona automaticamente para `os.devnull`.

### Erro COM do Windows na thread
Na versão GUI, o código chama `pythoncom.CoInitialize()` na thread de geração para garantir compatibilidade com o Word via COM.

---

## Licença

Uso interno — UNIFESO.