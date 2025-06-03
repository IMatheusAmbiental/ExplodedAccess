# Access Database Exploder

Este aplicativo permite explodir bancos de dados Access (.mdb) em arquivos separados por tabela, mantendo apenas os registros que atendem a critérios específicos.

## Pré-requisitos

1. **Microsoft Access Database Driver**
   - O driver ODBC para Access deve estar instalado no computador
   - Normalmente já vem instalado em sistemas Windows com Microsoft Office

2. **Para usuários que querem executar o arquivo .py:**
   - Python 3.8 ou superior (recomendamos a versão mais recente)
   - Você pode baixar o Python em: https://www.python.org/downloads/
   - Durante a instalação do Python, **IMPORTANTE**: Marque a opção "Add Python to PATH"

## Instalação

### Opção 1: Executável (.exe)
1. Baixe o arquivo `Access_Database_Exploder.exe`
2. Execute o arquivo normalmente com duplo clique

### Opção 2: Código Fonte (.py)
1. Instale o Python da página oficial: https://www.python.org/downloads/
2. Abra o terminal (cmd ou PowerShell) como administrador
3. Instale as dependências necessárias:
   ```
   pip install -r requirements.txt
   ```
4. Execute o arquivo `Acess_Explode_GUI.py`

## Como Usar

1. Clique em "Adicionar Arquivo" para selecionar até 2 arquivos .mdb
2. Use "Remover Selecionado" se precisar remover algum arquivo da lista
3. Clique em "Selecionar" para escolher o diretório de saída
4. Clique em "Processar" para iniciar o processamento
5. Aguarde o processamento terminar

## Estrutura de Saída

- Para cada arquivo .mdb processado, será criada uma pasta com prefixo "Exploded_"
- Dentro desta pasta, cada tabela será salva em um arquivo .mdb separado
- Apenas os registros que atendem aos critérios serão mantidos

## Solução de Problemas

1. **Erro de Driver ODBC**
   - Verifique se o Microsoft Access Database Driver está instalado
   - Para Windows 32-bit: Use o driver de 32-bit
   - Para Windows 64-bit: Use o driver de 64-bit

2. **Erro ao Executar o Python**
   - Verifique se o Python está corretamente instalado
   - Abra o cmd e digite `python --version` para verificar
   - Se não funcionar, reinstale o Python marcando "Add Python to PATH"

3. **Outros Erros**
   - Verifique se tem permissão de administrador
   - Verifique se os arquivos .mdb não estão em uso
   - Certifique-se de que há espaço suficiente no disco

## Contato

Para reportar problemas ou sugestões, entre em contato com o administrador do sistema. 
