# Automação de Cobrança e Follow-up de Pedidos

Aplicação desktop em Python para tratar a base de pedidos, enviar e-mails de acompanhamento a fornecedores pelo Microsoft Outlook Classic, controlar tentativas de envio e consultar respostas recebidas.

O projeto foi desenvolvido para uso no Windows e utiliza uma interface gráfica em Tkinter. A automação trabalha com arquivos Excel e com o perfil do Outlook configurado no computador do usuário.

## Funcionalidades

- Localiza a base de pedidos `.xlsx` mais recente cujo nome contenha `FUP`.
- Lê e trata as abas `APP` e `FORNECEDOR` da base.
- Gera planilhas individuais por fornecedor e envia os e-mails pelo Outlook.
- Mantém histórico dos envios, prazos e respostas em uma planilha de acompanhamento.
- Guarda envios que falharam em uma fila para reenvio posterior.
- Consulta respostas no Outlook e salva anexos Excel na pasta de respostas.
- Permite interromper o processamento em andamento pela interface.

## Requisitos

- Windows.
- Python instalado e disponível pelo comando `py`.
- Microsoft Outlook Classic instalado, configurado e autenticado no perfil do usuário.
- Acesso à pasta corporativa que contém a base de pedidos.

> A integração com e-mail usa automação COM do Outlook via `pywin32`. O novo Outlook e o Outlook Web não substituem o Outlook Classic para essa integração.

## Instalação

Abra o PowerShell na pasta do projeto e execute:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Se a política do PowerShell impedir a ativação do ambiente virtual, use o Python do ambiente diretamente:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Entrada de dados

Por padrão, a automação procura a base nesta pasta do usuário do Windows:

```text
C:\Users\<usuario>\Claro SA\APP-Acompanhamento de Pedidos - Documentos
```

O arquivo deve ser `.xlsx` e ter `FUP` no nome. Se houver mais de um arquivo correspondente, é escolhido o mais recentemente modificado. A planilha precisa conter as abas `APP` e `FORNECEDOR`; a aba `APP` deve conter, no mínimo, as colunas `Nº CONTA DO FORNECEDOR`, `STATUS` e `Previsão atual 2`. A rotina também valida e registra avisos sobre outras colunas utilizadas no tratamento.

O caminho da pasta de origem está definido em `utils/obter_base_pedidos.py`. Caso a organização ou o local da base mude, ajuste essa função antes de executar.

## Como executar

Na raiz do projeto, com o ambiente virtual ativado:

```powershell
python app.py
```

Também é possível iniciar sem ativar o ambiente virtual:

```powershell
.\.venv\Scripts\python.exe app.py
```

### Operações da interface

1. **Tratar Excel**: localiza e copia a base de origem para `data/raw/`, processa as abas e cria `data/cleansing/pedidos_tratados.xlsx`.
2. **Iniciar Envios**: agrupa os pedidos por fornecedor, prepara as planilhas de detalhe e envia as mensagens usando a conta configurada no Outlook.
3. **Reenviar Pendências**: tenta novamente os itens não enviados registrados em `log/base_reenvio.xlsx`.
4. **Baixar Retornos**: procura no Outlook as respostas relacionadas aos envios acompanhados e salva anexos Excel em `respostas_fornecedores/`.
5. **Parar Execução**: solicita a interrupção segura da rotina atual.

Os botões para abrir o log, a base de reenvio e a pasta de respostas também ficam disponíveis na interface.

## Segurança dos envios

A interface inicia com o **modo de teste ativado**. Nesse modo, as mensagens são direcionadas à conta de teste identificada no Outlook, em vez de serem enviadas aos endereços dos fornecedores. Confirme a conta exibida e faça um teste antes de qualquer execução real.

Para enviar aos fornecedores, desative o modo de teste deliberadamente na interface. Antes disso, confira a base, os destinatários, as planilhas anexadas e o perfil do Outlook. O envio normal é automático; a opção para abrir cada mensagem para revisão está definida em `app.py`.

> Não publique dados reais de fornecedores, bases Excel, logs, respostas ou informações pessoais. O perfil e a autenticação do Outlook são locais e não são configurados pelo repositório.

## Estrutura do projeto

```text
.
├── app.py                       # Interface gráfica e coordenação das operações
├── arquivo_executar.py          # Adapta a execução e os logs para a interface
├── main.py                      # Fluxo principal de envio e acompanhamento
├── requirements.txt             # Dependências Python
├── assinatura/
│   ├── imagem_claro.png          # Imagem usada na assinatura dos e-mails
│   └── simbolo_pequeno.png       # Símbolo usado na assinatura dos e-mails
├── utils/
│   ├── destaca_coluna_excel.py  # Formatação de colunas em planilhas Excel
│   ├── enviar_email.py          # Criação e envio de mensagens pelo Outlook
│   ├── logs.py                  # Histórico de envio e acompanhamento
│   ├── obter_base_pedidos.py    # Localização da base FUP de origem
│   ├── reenviar_email.py        # Processamento da fila de reenvio
│   ├── retorno_fornecedores.py  # Consulta de respostas e gravação de anexos
│   ├── tratar_excel.py          # Cópia, validação e tratamento das planilhas
│   └── usuario_atual.py         # Identificação do usuário do Windows
├── data/                        # Criada durante o tratamento da base
│   ├── raw/                     # Cópia da base original processada
│   └── cleansing/               # pedidos_tratados.xlsx
├── base_fornecedores/           # Planilhas individuais geradas para fornecedores
├── log/                         # Acompanhamento e fila de reenvio
└── respostas_fornecedores/      # Anexos recebidos nas respostas
```

As pastas `data/`, `base_fornecedores/`, `log/` e `respostas_fornecedores/` são criadas ou preenchidas durante a execução e seus arquivos operacionais são ignorados pelo Git.

## Arquivos gerados

- `data/raw/<base de origem>.xlsx`: cópia da planilha FUP selecionada.
- `data/cleansing/pedidos_tratados.xlsx`: base tratada usada pelo processo de envio.
- `base_fornecedores/`: planilhas de pedidos separadas por fornecedor.
- `log/acompanhamento_envios_fornecedores.xlsx`: histórico e situação dos retornos.
- `log/base_reenvio.xlsx`: fila de mensagens que precisam de nova tentativa.
- `respostas_fornecedores/`: anexos Excel encontrados nas respostas dos fornecedores.

## Publicação no GitHub

Antes de publicar, verifique se o commit não inclui planilhas operacionais, logs, anexos de respostas, arquivos de ambiente virtual ou dados pessoais. O `.gitignore` do projeto já exclui esses itens. Mantenha no repositório apenas o código, as dependências e os recursos necessários da assinatura.

O projeto depende de uma estrutura de pastas corporativa e de uma conta local do Outlook; portanto, publicar o código no GitHub não disponibiliza esses recursos nem os dados de entrada.

..
