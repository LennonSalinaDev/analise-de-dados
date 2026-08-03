# Melhorias futuras para ambiente profissional

Este documento registra melhorias recomendadas para evoluir o app depois da primeira versão funcional.

## Prioridade alta

### 1. Banco persistente em produção

Trocar ou consolidar o armazenamento para um banco persistente confiável.

Opções:

- manter SQLite com disco persistente bem configurado no Render;
- migrar para PostgreSQL em produção;
- criar rotina de backup automático.

Motivo: evitar perda de usuários, histórico, arquivos gerados e sessões após restart/redeploy.

### 2. Página de diagnóstico interno

Criar uma página acessível apenas para administrador exibindo:

- `DB_PATH`;
- `SETUP_LOCK_PATH`;
- quantidade de usuários;
- quantidade de sessões;
- pasta dos arquivos gerados;
- status do LibreOffice/soffice;
- ambiente atual: local, Render, Docker;
- espaço disponível em disco;
- últimos erros de PDF.

Motivo: reduzir dependência dos logs do Render para diagnosticar problemas.

### 3. Gerenciamento de usuários

Criar tela administrativa para:

- cadastrar novo usuário;
- trocar senha;
- excluir/desativar usuário;
- listar usuários;
- diferenciar perfil `admin` e `operador`.

Motivo: hoje o app tem setup/login, mas ainda não tem gestão completa de acesso.

### 4. Backup e exportação

Criar opções para exportar:

- banco SQLite completo;
- histórico de orçamentos;
- lista de farmacêuticos;
- mapas gerados;
- arquivos DOCX/XLSX/PDF.

Motivo: facilitar segurança operacional, auditoria e migração futura.

## Prioridade média

### 5. Logs controlados por variável

Transformar logs detalhados em modo opcional:

```text
DEBUG_AUTH=1
DEBUG_PDF=1
DEBUG_XLSX=1
```

Motivo: manter logs limpos em produção, mas permitir diagnóstico profundo quando necessário.

### 6. Melhor estado visual para PDF

Melhorar o botão PDF com estados:

- aguardando;
- convertendo;
- PDF pronto;
- falhou;
- visualizar erro.

Motivo: deixar claro para o usuário quando a conversão está em andamento ou falhou.

### 7. Registro estruturado de erros

Criar tabela no banco para erros importantes:

- data/hora;
- usuário;
- ação;
- arquivo;
- mensagem técnica;
- ambiente;
- status.

Motivo: facilitar auditoria e suporte sem depender apenas do console.

### 8. Proteção contra envio duplicado

Impedir duplo clique em ações como:

- salvar orçamento;
- gerar mapa;
- gerar PDF;
- excluir arquivo.

Motivo: evitar registros duplicados, conversões simultâneas e erros intermitentes.

## Prioridade baixa

### 9. CSRF simples nos formulários

Adicionar token CSRF nos formulários POST.

Motivo: reforçar segurança se o app for usado por mais pessoas ou exposto em domínio público.

### 10. Histórico avançado

Melhorar histórico com:

- filtros por data;
- filtros por cliente;
- filtros por filial;
- filtros por tipo de mapa;
- busca textual;
- paginação.

Motivo: facilitar uso quando o volume de documentos crescer.

### 11. Organização dos arquivos gerados

Separar arquivos por ano/mês/tipo:

```text
saida/
  orcamentos/2026/08/
  mapas_temperatura/2026/08/
```

Motivo: manter a pasta de saída organizada em uso contínuo.

### 12. Testes automatizados

Criar testes para:

- login/setup;
- geração de orçamento;
- geração de mapa;
- preservação de imagem no XLSX;
- exclusão de arquivos;
- validações de CPF, telefone e formulário.

Motivo: reduzir risco de regressão conforme o app crescer.

## Observações para produção

Para ambiente profissional, recomenda-se:

- usar HTTPS;
- manter senhas fortes;
- configurar variáveis sensíveis fora do código;
- usar banco persistente;
- manter backup;
- restringir acesso administrativo;
- validar periodicamente a geração de DOCX/XLSX/PDF.
