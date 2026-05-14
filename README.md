# Conversor de Perguntas AVD

Transforma o xlsx do cliente (N abas) em `Base_Perguntas-2.xlsx` pronto para o `Bases_AVD.ipynb`.

## Instalação

```bash
pip install -r requirements.txt
```

## Rodar

```bash
streamlit run app.py
```

Abre automaticamente em `http://localhost:8501`

## Fluxo

1. **Upload** — arrasta o xlsx do cliente
2. **Abas & Colunas** — seleciona quais abas processar, confirma mapeamento de colunas (detectado automaticamente)
3. **Validação** — revisa erros e usa ferramentas de correção em massa:
   - **Aplicar valor em lote** — filtra por status/aba/bloco e aplica um valor a N linhas de uma vez
   - **Buscar & substituir** — find & replace em qualquer campo, com prévia antes de confirmar
   - **Preencher vazios** — valor fixo ou fill-down (útil para `question_set` repetido a cada linha)
   - **Edição direta na tabela** — clique em qualquer célula para editar individualmente
4. **Exportar** — baixa `Base_Perguntas-2.xlsx` com as abas `Competências e Perguntas` e `Escalas`

## Deploy local (sua máquina)

```bash
# 1. Instale as dependências
pip install -r requirements.txt

# 2. Rode
streamlit run app.py
# Abre em http://localhost:8501
```

## Deploy em nuvem — Streamlit Cloud (gratuito)

1. Crie um repositório no GitHub e suba os três arquivos:
   - `app.py`
   - `requirements.txt`
   - `README.md`

2. Acesse [share.streamlit.io](https://share.streamlit.io) e faça login com GitHub

3. Clique em **New app** → selecione o repositório → aponte o Main file para `app.py`

4. Clique em **Deploy** — em ~1 minuto o app estará online com uma URL pública

> O Streamlit Cloud é gratuito para repositórios públicos.
> Para repositório privado, a conta gratuita ainda permite 1 app privado.

## Estrutura dos arquivos

```
conversor_avd/
├── app.py            # Aplicação principal
├── requirements.txt  # Dependências
└── README.md         # Este arquivo
```
