# Pipeline ETL e Modelagem Estocástica por Cadeias de Markov (Ativos GPON)

Automação desenvolvida em **Python** para extração, saneamento e modelagem matemática do ciclo de vida operacional de equipamentos de telecomunicações (ONUs e roteadores GPON), unindo Engenharia de Dados e Pesquisa Operacional.

## O que o sistema faz

1. **Ingestão Multiformato (ETL):** Lê planilhas pesadas (`.ods` e `.xlsx`) via *streaming* XML, extrai tabelas de documentos `.pdf` e realiza leitura óptica (OCR) em imagens, unificando o histórico individual por **número de série GPON / MAC**.
2. **Saneamento de Dados:** Filtra insumos passivos (cabos, conectores), descarta balanços contábeis de inventário físico, corrige datas invertidas (`MM/DD/AA` para `DD/MM/AAAA`) e comprime atualizações internas repetidas no mesmo setor.
3. **Classificação em Cadeia de Markov ($S_0$ a $S_9$):** Mapeia cada movimentação em 10 estados operacionais — preservando o histórico entre anos para diferenciar estoque novo ($S_0$) de estoque reaproveitado ($S_5$) e restringindo o fluxo de garantia/RMA ($S_7$) aos modelos elegíveis.
4. **Análise Temporal e Matrizes:** Calcula o tempo de permanência ($\Delta t$) com tratamento de censura à direita no fim do período e gera uma planilha consolidada contendo:
   - **Matriz de Transição Discreta ($P_{ij}$):** probabilidades condicionadas de mudança entre estados.
   - **Matriz Geradora Contínua ($Q$ - CTMC):** taxas de saída e tempo médio de espera incorporando equipamentos retidos.
   - **Indicadores de Permanência:** média, mediana, mínimo e máximo de dias por etapa.

## Tecnologias Utilizadas
- Python 3
- Pandas & OpenPyXL
- Processamento XML em *streaming* (`xml.etree.ElementTree`)
- PDFPlumber & Pytesseract (OCR)

## Como executar
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Coloque as planilhas de movimentação na pasta `entrada_materiais/` e execute:
```bash
python etl_gpon_markov.py
```
