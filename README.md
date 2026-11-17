# TaxKeyAlfa 2.0

**TaxKeyAlfa 2.0** é uma aplicação desktop nativa Windows de alto desempenho para a identificação morfológica interativa de mosquitos neotropicais (Diptera: Culicidae) vetores de arboviroses e malária, fundamentada nas chaves dicotômicas clássicas de **Oswaldo Paulo Forattini (2002) — *Culicidologia Médica, Vol. 2*** e **Consoli & Lourenço-de-Oliveira (1994)**.

## Principais Recursos
- **Motor Orientado a Dados (DAG em JSON)**: 339 passos dicotômicos cobrindo **todas as subfamílias, tribos, 23 gêneros e mais de 210 espécies** de Culicidae neotropicais, validados automaticamente contra nós órfãos, ciclos ou destinos vazios.
- **Interface Gráfica Nativa Windows (`CustomTkinter`)**:
  - Comparação lado a lado (**Opção A vs Opção B**) com ampliação de pranchas morfológicas.
  - **Trilha de Decisão (*Breadcrumb*)** interativa e botão **Voltar (`Backspace`)**.
  - **Atalhos de Teclado para Laboratório**: `[A]` / `[1]` (Opção A), `[B]` / `[2]` (Opção B), `[Backspace]` (Voltar), `[Ctrl+S]` (Salvar).
  - **Glossário Contextual Automático**: Identifica e exibe automaticamente o significado das 47 siglas anatômicas de Forattini (`Mpn`, `Mcs`, `Esc`, `Clp`, `SAP`, `SPe`, `Ta-1..5`, etc.) presentes em cada passo.
  - **Catálogo & Busca Direta**: Permite buscar qualquer espécie/gênero instantaneamente ou saltar direto para a chave de um gênero específico.
- **Persistência Tripla Automatizada**: Salva em **SQLite (`dados_mosquitos.db`)**, sincroniza automaticamente com **Excel (`dados_mosquitos.xlsx`)**, exporta para **CSV (`;`)** e possui conector opcional para **SQL Server (`pyodbc`)**.

## Estrutura do Projeto
```text
TAX_KEY_ALFA01/
├── main.py                          # Ponto de entrada da aplicação (~38 linhas)
├── core/
│   ├── schema.py                    # Modelos de dados (KeyStep, KeyOption, SpecimenRecord)
│   ├── engine.py                    # Motor de navegação dicotômica e validador de grafo
│   └── repository.py                # Persistência em SQLite, Excel (.xlsx), CSV e SQL Server
├── data/
│   ├── glossary.json                # Glossário das 47 siglas anatômicas de Forattini
│   └── keys/                        # Chaves taxonômicas modulares em JSON (339 passos)
│       ├── 01_subfamilias_tribos.json
│       ├── 02_anophelinae.json
│       ├── 03_aedini.json
│       ├── 04_sabethini.json
│       ├── 05_culicini.json
│       └── 06_outras_tribos.json
├── ui/
│   └── app_window.py                # Interface gráfica nativa Windows (CustomTkinter)
└── tests/
    └── test_keys_and_engine.py      # Suíte de testes automatizados do grafo e persistência
```

## Instalação e Execução

```bash
pip install -r requirements.txt
python main.py
```

## Execução dos Testes Automatizados

```bash
python -m unittest discover -s tests -v
```
