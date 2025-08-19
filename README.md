# Sistema de Abastecimento Automático

Este sistema automatiza o processo de criação de cards no WMS para abastecimento das filiais da empresa.

## Funcionalidades

- **Leitura automática de planilhas**: O sistema identifica automaticamente qual planilha usar baseado no dia da semana
- **Validação de dados**: Verifica se a planilha contém todas as colunas necessárias
- **Criação de cards**: Cria cards no WMS com os itens para envio
- **Divisão automática**: Se uma filial tiver mais de 30 itens, cria múltiplos cards
- **Logging completo**: Registra todas as operações em arquivo de log

## Estrutura da Planilha

A planilha deve conter as seguintes colunas:

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| `CD ABASTECIMENTO` | Nome do CD de origem | CD-SP |
| `BASE A ENVIAR` | Nome da filial de destino | São Paulo |
| `ITEM A ENVIAR` | Nome do item | Roda Traseira |
| `QUANTIDADE` | Quantidade para envio | 20 |

## Arquivos do Sistema

- `abastecimentoAutomatico.py`: Script principal que processa a planilha e cria os cards
- `validar_planilha.py`: Script para validar a estrutura da planilha antes do processamento
- `requirements.txt`: Dependências do projeto
- `README.md`: Este arquivo com instruções

## Como Usar

### 1. Preparação

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

2. Configure o caminho das planilhas no arquivo `abastecimentoAutomatico.py`:
```python
CAMINHO_BASE_PLANILHAS = r"C:\caminho\para\suas\planilhas"
```

3. Ajuste as credenciais de acesso ao sistema:
```python
'username': 'seu_usuario@mottu.com.br',
'password': 'sua_senha',
```

### 2. Validação da Planilha

Antes de executar o processo principal, valide a planilha:

```bash
python validar_planilha.py
```

Este comando irá:
- Verificar se a planilha existe
- Validar se todas as colunas obrigatórias estão presentes
- Mostrar um resumo dos dados
- Identificar possíveis problemas

### 3. Execução do Processo

Execute o script principal:

```bash
python abastecimentoAutomatico.py
```

O sistema irá:
1. Ler a planilha `separacaoAmanha.xlsx`
2. Obter token de autenticação
3. Ler e validar a planilha
4. Agrupar itens por CD e filial
5. Criar cards no WMS (máximo 30 itens por card)
6. Registrar todas as operações no log

## Regras de Negócio

### Divisão de Cards
- Máximo de 30 itens por card
- Se uma filial tiver mais de 30 itens, serão criados múltiplos cards
- Exemplo: 45 itens para São Paulo = 1 card com 30 itens + 1 card com 15 itens

### Planilha
- **Nome da planilha**: `separacaoAmanha.xlsx`
- **Localização**: Configurada na variável `CAMINHO_BASE_PLANILHAS`

## Logs

O sistema gera logs detalhados em:
- **Console**: Mostra progresso em tempo real
- **Arquivo**: `abastecimento_automatico.log`

Exemplo de log:
```
2024-01-15 10:30:00 - INFO - Iniciando processamento da planilha de abastecimento
2024-01-15 10:30:01 - INFO - Obtendo token de autenticação...
2024-01-15 10:30:02 - INFO - Token obtido com sucesso
2024-01-15 10:30:03 - INFO - Usando planilha: separacaoAmanha.xlsx
2024-01-15 10:30:04 - INFO - Planilha válida com 150 linhas
2024-01-15 10:30:05 - INFO - Total de 45 itens para processar
2024-01-15 10:30:06 - INFO - Processando: CD 123 → Filial 456
2024-01-15 10:30:07 - INFO - Total: 45 itens em 2 card(s)
2024-01-15 10:30:08 - INFO - [SUCESSO] Card enviado para filial 456
```

## Tratamento de Erros

O sistema inclui tratamento robusto de erros:

- **Arquivo não encontrado**: Verifica se a planilha existe
- **Colunas faltantes**: Valida estrutura da planilha
- **Erro de autenticação**: Trata problemas de login
- **Erro de rede**: Timeout e retry para requisições
- **Dados inválidos**: Validação de dados antes do processamento

## Configurações

### Variáveis Importantes

```python
MAX_ITENS_POR_CARD = 30  # Máximo de itens por card
REQUESTER_ID = '7e48e47a-8c81-4777-a896-afb2d871ebc7'  # ID do solicitante
URL_WMS = 'https://warehouse-inventory.mottu.cloud/Order/file'  # Endpoint do WMS
```

### Segurança

⚠️ **Importante**: As credenciais estão hardcoded no código. Para produção, considere usar variáveis de ambiente:

```python
import os
data = {
    'username': os.getenv('MOTTU_USERNAME'),
    'password': os.getenv('MOTTU_PASSWORD'),
    'client_id': 'mottu-admin',
    'grant_type': 'password'
}
```

## Suporte

Em caso de problemas:

1. Execute `python validar_planilha.py` para verificar a planilha
2. Verifique o arquivo de log `abastecimento_automatico.log`
3. Confirme se as credenciais estão corretas
4. Verifique se o caminho das planilhas está correto 