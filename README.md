**Transpor Planilha → Formato Final (Dígito1 100% preenchido)**

Este repositório contém um aplicativo Streamlit (`app.py`) que transforma planilhas Excel de entrada em uma planilha final padronizada com 13 colunas na ordem exata, garantindo que o campo `Dígito1` seja detectado e preenchido corretamente.

**Principais Funcionalidades**
- **Descrição**: Converte planilhas com colunas de estudantes e orientadores para um formato longo e padronizado.
- **Detecção inteligente**: Reconhece automaticamente rótulos como `Dígito1`, `CPF`, `Nome completo`, `Data de Nascimento`, entre outros.
- **Formatação**: Normaliza CPF e datas; garante as 13 colunas finais na ordem correta.
- **Saída**: Gera e permite download da planilha `PLANILHA_FINAL_DIGITO1_100_CORRETO.xlsx`.

**Colunas finais (ordem exata)**
- `Nome completo`
- `CPF`
- `Data de Nascimento`
- `Tipo`
- `Número`
- `Instituição Bancária`
- `Agência Bancária (sem dígito)`
- `Digito`
- `Número da Conta Corrente Nominal (sem dígito)`
- `Dígito1`
- `Instituição de ensino superior`
- `campus`
- `Nome da(o) Tutura(or)`

**Requisitos**
- Python 3.8+ (ou versão compatível com os pacotes listados).
- Dependências listadas em `requirements.txt` (ex.: `streamlit`, `pandas`, `openpyxl`).

**Instalação (Windows PowerShell)**
1. Ative o ambiente virtual (se já existir):

```powershell
.\env\Scripts\Activate.ps1
```

2. Instale dependências:

```powershell
pip install -r requirements.txt
```

**Execução**
1. Inicie a aplicação Streamlit:

```powershell
streamlit run app.py
```

2. A interface abre no navegador. Use a barra lateral para fazer upload do arquivo `.xlsx` ou marque `Usar arquivo padrão` para usar o arquivo `modelo da planilha.xlsx` (se presente na raiz).

3. Após o upload, clique em `GERAR PLANILHA FINAL (Dígito1 100% PREENCHIDO)` e, quando concluído, baixe a planilha final com o botão de download.

**Notas importantes**
- O app tenta detectar automaticamente colunas nomeadas de formas variadas. Para garantir que `Dígito1` seja encontrado, use exatamente o rótulo `Dígito1` ou `Digito1` quando possível.
- Se a planilha tiver prefixos como `Estudante 10)` ou rótulos com texto adicional, o app aplica heurísticas para extrair o rótulo correto.
- Se der erro ao ler o arquivo `.xlsx`, verifique se a planilha está em formato Excel válido e se o `openpyxl` está instalado.

**Saída gerada**
- Arquivo gerado para download: `PLANILHA_FINAL_DIGITO1_100_CORRETO.xlsx`.

**Depuração / Troubleshooting**
- Mensagem `Aguardando upload do arquivo...`: faça upload do `.xlsx` ou marque `Usar arquivo padrão` e verifique se `modelo da planilha.xlsx` existe.
- Mensagem de erro ao ler Excel: confirme versão do arquivo e dependências (`openpyxl`).
- `Dígito1 NÃO encontrado`: verifique nomes das colunas na planilha de entrada — o app exibe, na seção de verificação, quais colunas foram identificadas contendo `Dígito1`.

**Contribuições e melhorias**
- Ajustes de detecção de rótulos podem ser feitos em `app.py`, especialmente na função `extrair_label` para mapear novos padrões de nome de coluna.

**Contato**
- Caso precise de ajuda, descreva o problema, anexe um exemplo (sem dados sensíveis) e informe a versão do Python e das dependências.

---
Gerado a partir do conteúdo de `app.py` (ferramenta de transformação de planilhas - Streamlit).
