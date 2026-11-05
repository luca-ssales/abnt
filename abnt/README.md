# Formatador ABNT 📄✨

Aplicação web full-stack que recebe um arquivo **.docx** do Microsoft Word e devolve o documento com a formatação estritamente ajustada às normas da **ABNT** (NBR 14724, NBR 6024, NBR 10520, NBR 6023 e NBR 6028).

---

## 🛡️ Regra de Ouro e Integridade Documental

- **100% de Preservação do Conteúdo**: O texto original nunca é apagado, resumido, truncado ou reescrito. Uma auditoria de integridade compara cada parágrafo, célula de tabela e nota de rodapé antes e depois da gravação; se houver qualquer divergência de caracteres, a operação é imediatamente cancelada.
- **Elementos Ricos Preservados**: Imagens (`<w:drawing>`, `<w:pict>`), tabelas, notas de rodapé, hiperlinks (`<w:hyperlink>`), listas numeradas/com marcadores e comentários são mantidos intactos.
- **Caixa Alta Visual Segura**: Para atender à NBR 6024 (títulos primários em maiúsculas), o sistema utiliza a propriedade OpenXML `<w:caps/>`, garantindo que o Word renderize em maiúsculas sem modificar os caracteres originais da tag `<w:t>`.
- **Privacidade**: O arquivo original enviado é excluído do servidor imediatamente após o processamento. Os arquivos formatados expiram rapidamente e são removidos após o download.

---

## 🏛️ Regras ABNT Aplicadas

1. **Página e Margens (NBR 14724)**:
   - Papel: A4 (21 x 29,7 cm)
   - Margem Superior: 3,0 cm
   - Margem Esquerda: 3,0 cm
   - Margem Inferior: 2,0 cm
   - Margem Direita: 2,0 cm
2. **Texto Corrido (NBR 14724)**:
   - Fonte: Times New Roman ou Arial (padrão: Times New Roman)
   - Tamanho: 12 pt, cor preta
   - Espaçamento entre linhas: 1,5
   - Espaço antes/depois: 0 pt
   - Recuo da primeira linha: 1,25 cm
   - Alinhamento: Justificado
3. **Citações Diretas Longas com mais de 3 linhas (NBR 10520)**:
   - Recuo da margem esquerda: 4,0 cm
   - Fonte tamanho: 10 pt
   - Espaçamento entre linhas: simples (1,0)
   - Sem recuo de primeira linha, justificado
4. **Títulos e Seções (NBR 6024 e NBR 14724)**:
   - **Seção Primária (1)**: Negrito, MAIÚSCULAS, alinhada à esquerda, quebra de página em TCC/Monografia.
   - **Seção Secundária (1.1)**: Negrito, alinhada à esquerda.
   - **Seção Terciária (1.1.1)**: Normal (sem negrito), alinhada à esquerda.
   - **Seção Quaternária/Quinária**: Normal/itálico, alinhada à esquerda.
   - **Títulos sem indicativo numérico** (RESUMO, ABSTRACT, SUMÁRIO, REFERÊNCIAS, etc.): Centralizados, negrito, MAIÚSCULAS.
   - **Espaçamento**: Espaço de 1,5 linha (18 pt) antes e depois.
5. **Paginação (NBR 14724, item 5.3)**:
   - Numeração no canto superior direito da página, em algarismos arábicos, fonte 10 pt.
   - Folhas pré-textuais contadas mas com número oculto; numeração exibida a partir da primeira folha da parte textual (Introdução).
6. **Notas de Rodapé, Tabelas e Legendas**:
   - Notas de rodapé formatadas via XML (`word/footnotes.xml`) em 10 pt, espaçamento simples, justificado.
   - Tabelas padronizadas internamente em 10 pt e espaçamento simples, sem quebrar colunas ou bordas.
   - Legendas e fontes de ilustrações/tabelas em 10 pt, centralizadas, espaçamento simples.
7. **Referências Bibliográficas (NBR 6023)**:
   - Alinhamento à esquerda, espaçamento simples, separadas entre si por 1 espaço simples em branco (12 pt after).

---

## 📁 Estrutura de Pastas

```text
abnt/
├── app/
│   ├── __init__.py           # Inicializador do pacote
│   ├── abnt_rules.py         # Configuração centralizada de todas as regras ABNT
│   ├── detector.py           # Classificação estrutural e heurística de parágrafos
│   ├── formatter.py          # Motor de aplicação das normas e integridade textual
│   ├── xml_helpers.py        # Manipulação direta do OpenXML (paginação, notas, TOC, hiperlinks)
│   ├── report.py             # Montagem do relatório "Antes e Depois" e avisos
│   ├── pdf_converter.py      # Conversão opcional via LibreOffice headless
│   └── main.py               # Servidor FastAPI e rotas de upload/download
├── static/
│   ├── index.html            # Interface web moderna e responsiva
│   ├── styles.css            # Folha de estilos limpa e acessível
│   └── app.js                # Lógica frontend (drag & drop, upload assíncrono, feedback)
├── tests/
│   ├── __init__.py
│   ├── generate_fixtures.py  # Gerador de documentos .docx de teste
│   ├── test_abnt_formatter.py# Testes da regra de ouro (comparação de texto 1:1)
│   └── test_api.py           # Testes dos endpoints FastAPI
├── requirements.txt          # Dependências do projeto
└── README.md                 # Documentação e instruções de execução
```

---

## 🚀 Passo a Passo para Instalação e Execução Local (Para Iniciantes)

### 1. Pré-requisitos
- **Python 3.11 ou superior** instalado no computador (durante a instalação, marque a opção *"Add Python to PATH"*).

### 2. Abrir o terminal na pasta do projeto
Abra o prompt de comando (PowerShell ou CMD no Windows, Terminal no Linux/Mac) e navegue até a pasta:
```bash
cd c:\Users\Studio\Desktop\abnt
```

### 3. (Opcional, mas recomendado) Criar e ativar um ambiente virtual
```bash
python -m venv venv

# No Windows (PowerShell):
.\venv\Scripts\Activate.ps1

# No Linux ou Mac:
source venv/bin/activate
```

### 4. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 5. Executar o servidor web
```bash
uvicorn app.main:app --reload --port 8000
```

### 6. Acessar no navegador
Abra seu navegador de preferência e acesse:
👉 **`http://localhost:8000`**

---

## 🧪 Como Rodar os Testes Automatizados

Para certificar que todas as regras e a integridade de 100% do texto estão funcionando:
```bash
python -m pytest -v
```
Todos os 5 testes devem passar com `PASSED`.

---

## 📑 Conversão Opcional para PDF (LibreOffice)

O sistema verifica automaticamente se o **LibreOffice** está instalado no sistema operacional:
- Se instalado (`soffice` detectado no sistema), o botão extra **"📑 Baixar também em PDF"** é exibido na tela de resultado.
- Se não estiver instalado, a interface oculta esse botão suavemente, mantendo disponível o download completo em **.DOCX**.

Para habilitar a conversão em PDF:
- **Windows**: Instale o [LibreOffice](https://www.libreoffice.org/download/download/) (instalado por padrão em `C:\Program Files\LibreOffice\program\soffice.exe`).
- **Linux (Ubuntu/Debian)**: `sudo apt-get install -y libreoffice`

---

## ⚠️ Limitações Conhecidas e Decisões de Projeto

1. **Parágrafos Ambíguos**: Quando uma linha curta com destaque não possui numeração (ex.: frase em destaque ou citação de 3 linhas com formatação incompleta), a opção mais segura é mantê-la como texto corrido com fonte padronizada, alertando o usuário no relatório em **"Itens para Revisão Manual"**, evitando alterar a estrutura sem certeza.
2. **Sumário Automático**: O campo de Sumário do Word (TOC) pode ser inserido automaticamente via opção do formulário. Como o LibreOffice e Word recalculam páginas dinamicamente, o usuário só precisa dar um clique com o botão direito no Word e escolher *"Atualizar Campo"* para preencher os números de página finais.
3. **Formatos Antigos (.doc)**: Arquivos `.doc` antigos (Word 97-2003 binário) ou protegidos com senha são rejeitados com orientações amigáveis, pois apenas o formato moderno `.docx` (OpenXML) é suportado.

---

## 🔮 Sugestões de Próximos Passos
- **Extração de PDF e OCR**: Suporte para receber artigos escaneados em PDF e convertê-los em `.docx` estruturado.
- **Formulário para Capa e Folha de Rosto**: Criar gerador opcional de Capa, Folha de Rosto e Folha de Aprovação com dados informados pelo usuário.
- **Normalização de Referências (NBR 6023)**: Parser com IA para identificar autor, título em negrito/itálico, edição e link DOI das referências.
- **Deploy em Nuvem**: Container Docker pronto para deploy em plataformas como Railway, Render ou AWS com LibreOffice pré-configurado.

