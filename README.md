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

## Acesso ao app

No primeiro acesso, o sistema abre a tela de criação do usuário administrador em:

```text
http://127.0.0.1:8000/setup
```

Depois disso, o app passa a exigir login e senha em:

```text
http://127.0.0.1:8000/login
```

As senhas não são salvas em texto puro. O banco guarda apenas hash com sal, e o acesso usa cookie de sessão `HttpOnly` com expiração.

## Pastas

- `modelos/`: modelos Word usados como base. Se houver mais de um `.docx`, o sistema usa o primeiro em ordem alfabética.
- `data/`: banco SQLite com o histórico.
- `saida/`: arquivos DOCX e PDF gerados.

## Histórico

A tela `Histórico` permite abrir, baixar e excluir orçamentos individualmente. Ao excluir, o sistema remove o registro do banco, os itens vinculados e os arquivos DOCX/PDF gerados para aquele orçamento.

## Farmacêuticos responsáveis

A tela `Farmacêuticos` permite cadastrar, editar e excluir os nomes que aparecem no seletor `Farm. resp.` do formulário. A exclusão remove o nome apenas da lista de seleção; orçamentos antigos continuam guardando o texto do responsável que foi usado na criação.

No formulário de orçamento, o responsável fica na última seção, junto ao botão `Salvar e gerar arquivos`, e precisa ser escolhido a partir da lista cadastrada.

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

O DOCX é sempre gerado. No Windows, o sistema tenta gerar o PDF usando o Microsoft Word instalado na máquina. Se o Word não estiver disponível, ele tenta usar LibreOffice/soffice quando estiver instalado e acessível no `PATH`.

Na tela de detalhes do orçamento existe a opção `Gerar PDF`, útil para criar o PDF de orçamentos antigos ou tentar novamente quando a conversão automática não tiver sido concluída.

## Cloudflare / cloudflared

Este sistema não deve ser publicado como Cloudflare Pages/Workers com `npx wrangler deploy`, porque ele não é um site estático. Ele é um servidor Python que usa SQLite, gera arquivos DOCX/PDF e pode acionar o Microsoft Word local para PDF.

O erro abaixo indica justamente isso:

```text
Could not detect a directory containing static files (e.g. html, css and js) for the project
```

Para acessar pela internet via Cloudflare, rode o sistema localmente e exponha com Cloudflare Tunnel:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\start_cloudflare_tunnel.ps1
```

Depois use a URL `https://...trycloudflare.com` exibida pelo `cloudflared`.

Se estiver usando o painel da Cloudflare Pages, remova o deploy command `npx wrangler deploy`. Para este projeto, Cloudflare Pages não é o ambiente correto enquanto o sistema depender de Python, SQLite local e geração de PDF pelo Word.

## Deploy no Render

O projeto mantém duas opções de deploy no Render:

- `render.yaml`: opção principal usando Docker + LibreOffice.
- `render-python.yaml`: opção alternativa usando Python nativo, sem LibreOffice garantido.

Use Docker quando quiser tentar gerar PDF no servidor. Use Python nativo apenas como segunda opção para testar o app gerando DOCX/XLSX.

### Opção principal: Docker

O arquivo `render.yaml` deixa o projeto pronto para um Web Service no Render usando Docker.

Configuração esperada:

```text
Runtime/Language: Docker
Dockerfile Path: ./Dockerfile
```

O serviço usa:

```text
Health Check Path: /healthz
DB_PATH: /var/data/orcamentos.db
APP_ADMIN_USER: definido no painel do Render
APP_ADMIN_PASSWORD: definido no painel do Render
ORCAMENTOS_OUTPUT_DIR: /var/data/saida
TEMPERATURE_MAP_OUTPUT_DIR: /var/data/saida/mapas_temperatura
EXPORT_PATH: /var/data/orcamentos_export.csv
HOST: 0.0.0.0
HOME: /tmp
SAL_USE_VCLPLUGIN: svp
```

O `render.yaml` também define um disco persistente em `/var/data` para guardar o banco SQLite e os arquivos gerados. Isso é importante porque o login, os responsáveis, o histórico e os DOCX/XLSX/PDF gerados dependem desses arquivos. Sem disco persistente, esses dados podem ser perdidos em reinícios ou redeploys.

Se o Render abrir `/setup` depois de um redeploy ou depois de uma tentativa de gerar arquivo, isso indica que o serviço iniciou com um banco SQLite vazio ou em outro caminho. Para evitar esse comportamento, configure `APP_ADMIN_USER` e `APP_ADMIN_PASSWORD` como variáveis secretas no painel do Render. Quando o banco estiver vazio, o app cria esse administrador automaticamente; quando já existir usuário, essas variáveis são ignoradas.

Se você estiver criando manualmente como Python, o DOCX/XLSX funciona, mas o PDF não terá LibreOffice instalado por padrão. Para gerar PDF na nuvem, use Docker. Nesse modo Python nativo, só configure `DB_PATH=/var/data/orcamentos.db` se você realmente tiver um disco persistente montado em `/var/data`; caso contrário, o serviço não terá permissão para criar essa pasta.

Se os logs mostrarem algo como:

```text
Banco de dados em: /opt/render/project/src/data/orcamentos.db
LibreOffice/soffice não encontrado
```

o serviço está rodando como Python nativo, não como Docker. Nesse caso, o `Dockerfile` está sendo ignorado. Crie um novo serviço no Render usando `New +` > `Blueprint` e selecione este repositório, ou recrie o Web Service escolhendo `Runtime: Docker` e `Dockerfile Path: ./Dockerfile`.

Após corrigir o serviço, o log esperado deve mostrar:

```text
Banco de dados em: /var/data/orcamentos.db
```

Se você continuar no runtime Python nativo, no campo `Start Command` do painel do Render, digite somente:

```text
python app.py
```

Não digite `Start Command: python app.py`. Se esse texto completo for colocado no campo, o Render tenta executar um comando chamado `Start` e mostra este erro:

```text
bash: line 1: Start: command not found
```

No Render, o PDF pelo Microsoft Word não funciona, porque o ambiente é Linux e não tem Word instalado. Com o `Dockerfile` deste projeto, o LibreOffice Calc/Writer é instalado dentro da imagem e o sistema usa `soffice --headless` para converter DOCX/XLSX em PDF. Cada conversão usa um perfil temporário isolado do LibreOffice, reduzindo falhas intermitentes quando o serviço recebe mais de uma ação próxima.

### Segunda opção: Python nativo

Se quiser manter um deploy simples para teste sem Docker, use o arquivo `render-python.yaml` como referência.

Configuração esperada:

```text
Runtime/Language: Python
Build Command: pip install -r requirements.txt
Start Command: python app.py
Health Check Path: /healthz
```

Nesse modo, não configure `DB_PATH=/var/data/orcamentos.db` a menos que exista um disco persistente montado em `/var/data`. Sem esse disco, o Render pode retornar erro de permissão ao tentar criar `/var/data`.

No Python nativo, o app continua gerando DOCX e XLSX. A conversão para PDF pode falhar porque o LibreOffice/soffice normalmente não vem instalado nesse runtime.

Atenção: o histórico atual usa SQLite em arquivo local. Em hospedagem gratuita, isso serve para teste/demonstração, mas pode ser perdido em redeploy/restart. Para uso real em nuvem, o próximo passo é trocar o histórico para PostgreSQL.
