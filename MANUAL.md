# Manual do Validador Bibliográfico

Este script automatiza a validação de referências bibliográficas lendo arquivos do Microsoft Word (.docx), buscando os metadados em APIs acadêmicas oficiais (Crossref, OpenAlex, PubMed, Semantic Scholar) e gerando um relatório em CSV (que pode ser aberto no Excel).

## Pré-requisitos
Antes de usar o script pela primeira vez, você precisa garantir que todas as bibliotecas necessárias estão instaladas.
Abra o **Windows PowerShell** e digite o seguinte comando na pasta do projeto:
```bash
pip install -r requirements.txt
```

## Como usar passo a passo

### Passo 1: Preparar o arquivo Word
Certifique-se de que o seu arquivo `.docx` contém a seção de referências no final, encabeçada por um título claro, como:
- "Referências"
- "Referências Bibliográficas"
- "References"
*(O script buscará por essa palavra exata em uma linha isolada para começar a capturar os dados).*

### Passo 2: Abrir o Terminal (PowerShell)
Navegue até a pasta onde este script está salvo (`d:\04_UEM\Revistas\ACTA_SCIENTIARUM\Script\Validação_Bibliográfica_Python`).

### Passo 3: Executar o Script
Para rodar a validação do seu arquivo, digite o seguinte comando:
```bash
python main.py "Caminho\completo\para\o\seu\artigo.docx"
```
**Exemplo:**
```bash
python main.py "D:\Meus Documentos\Artigo_Biologia.docx"
```

O script vai abrir, ler todas as referências daquele documento e começar a validar uma por uma na internet. Na tela, ele mostrará uma tabela bonita e colorida informando o status de cada referência (Confirmado, Parcialmente confirmado ou Não confirmado).

### Passo 4: Formatação Automática de Tabelas
Por padrão, ao processar o arquivo `.docx`, o script também formatará automaticamente todas as tabelas do documento no padrão **Acta Scientiarum / APA (Estilo de Três Linhas)**:
- Fonte `PT Serif` 8 pt
- Largura 100%
- Bordas horizontais de 0,5 pt no topo, abaixo do cabeçalho e na base
- Sem bordas verticais ou linhas internas entre dados
- Sem cores de fundo/sombreamentos

*(Caso queira desativar a formatação de tabelas durante a validação, utilize a opção `--no-format-tables`)*.

### Passo 5: Salvar com nome personalizado (Opcional)
Por padrão, o script vai criar um arquivo chamado `relatorio_referencias.csv` e salvar o Word verificado como `<nome_original>_verificado.docx`. Se você quiser escolher o nome do arquivo que será salvo, adicione `-o` no final:
```bash
python main.py "D:\Meus Documentos\Artigo_Biologia.docx" -o "D:\Meus Documentos\Relatorio_Biologia.docx"
```

### Passo 6: Abrir o Relatório no Excel
1. Vá até a pasta onde o relatório foi salvo.
2. Abra o Excel.
3. Arraste o arquivo `.csv` para dentro do Excel (ou clique em Arquivo > Abrir).
4. Você terá uma tabela organizada mostrando:
   - Número da referência
   - Texto original (como estava no Word)
   - Status da validação
   - Fonte (qual API encontrou os dados)
   - Link do DOI (clicável para abrir no navegador)
   - Observações (detalhes sobre algum autor que faltou ou erro encontrado).
