# Gerador de Orçamentos CLAMED

Sistema local para preencher os dados variáveis do orçamento, gerar o arquivo Word, tentar gerar PDF automaticamente e manter um histórico estruturado em SQLite.

## Como usar

1. Instale as dependências, se estiver usando outro Python:

```powershell
python -m pip install -r requirements.txt
```

2. Coloque o modelo Word `.docx` dentro da pasta `modelos/`.

Se quiser recriar o modelo a partir do arquivo original em Downloads:

```powershell
python create_template.py
```

3. Inicie o formulário:

```powershell
python app.py
```

4. Abra no navegador:

```text
http://127.0.0.1:8000
```

Se a porta `8000` estiver ocupada, o sistema tentará automaticamente as próximas portas livres e mostrará o endereço correto no terminal.

## Pastas

- `modelos/`: modelos Word usados como base. Se houver mais de um `.docx`, o sistema usa o primeiro em ordem alfabética.
- `data/`: banco SQLite com o histórico.
- `saida/`: arquivos DOCX e PDF gerados.

## Histórico

A tela `Histórico` permite abrir, baixar e excluir orçamentos individualmente. Ao excluir, o sistema remove o registro do banco, os itens vinculados e os arquivos DOCX/PDF gerados para aquele orçamento.

O campo `E-mail` fica vazio por padrão e é salvo como `NULL` quando não for preenchido. O campo `Data` usa seletor de data no formulário e é salvo no formato `AAAA-MM-DD`; no orçamento Word ele é convertido automaticamente para o texto usado no modelo.

CPF e telefone são normalizados antes de salvar e gerar o Word. O CPF precisa ter 11 dígitos e sai como `000.000.000-00`; o telefone celular precisa ter DDD + 9 dígitos e sai como `(00) 00000-0000`.

Quando algum campo estiver inválido, o sistema mostra o erro no próprio formulário, destaca o campo com problema e mantém os dados já digitados.

O campo `Desc. %` reduz o valor unitário antes de calcular o total da linha. Exemplo: valor unitário `100,00`, desconto `10` e quantidade `1` gera `90,00` em `Valor total`.

## Validar o modelo

Depois de editar ou trocar o arquivo `.docx` dentro de `modelos/`, abra:

```text
http://127.0.0.1:8000/validar-modelo
```

Essa página confere se o modelo ainda tem os campos necessários para gerar novos orçamentos.

## Observação sobre PDF

O DOCX é sempre gerado. O PDF será gerado automaticamente quando houver LibreOffice/soffice instalado e acessível no `PATH`. Se não houver, o sistema registra o orçamento normalmente e mostra o PDF como indisponível.
