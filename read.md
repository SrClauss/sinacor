# Instalando o Python

Para começar, você precisa instalar o Python em seu sistema. Siga os passos abaixo:

1. Acesse o site oficial do Python em https://www.python.org/downloads/.
2. Faça o download da versão mais recente do Python para o seu sistema operacional.
3. Execute o instalador e siga as instruções para concluir a instalação.

# Criando uma virtualenv

Uma virtualenv é um ambiente virtual isolado onde você pode instalar pacotes e dependências específicos para um projeto. Siga os passos abaixo para criar uma virtualenv:

1. Abra o terminal ou prompt de comando.
2. Navegue até o diretório do seu projeto usando o comando `cd /caminho/do/seu/projeto`.
3. Execute o seguinte comando para criar uma virtualenv:

    ```
    python -m venv nome_da_virtualenv
    ```

    Substitua `nome_da_virtualenv` pelo nome que você deseja dar à sua virtualenv.

4. Ative a virtualenv executando o seguinte comando:

    - No Windows:

      ```
      nome_da_virtualenv\Scripts\activate
      ```

    - No macOS/Linux:

      ```
      source nome_da_virtualenv/bin/activate
      ```

# Instalando as dependências usando o requirements.txt

O arquivo requirements.txt contém uma lista de pacotes e suas versões específicas que são necessários para o seu projeto. Siga os passos abaixo para instalar as dependências:

1. Certifique-se de que sua virtualenv esteja ativada.
2. No terminal ou prompt de comando, execute o seguinte comando:

    ```
    pip install -r requirements.txt
    ```

    Isso irá instalar todas as dependências listadas no arquivo requirements.txt.

Pronto! Agora você tem o Python instalado, uma virtualenv configurada e as dependências do seu projeto instaladas. Você está pronto para começar a desenvolver!
