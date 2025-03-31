# My Python Project

Este projeto é uma aplicação Python que possui uma interface gráfica para gerenciamento de usuários, incluindo funcionalidades de login, cadastro e um painel principal.

## Estrutura do Projeto

```
my-python-project
├── src
│   ├── screens
│   │   ├── login.py         # Tela de login
│   │   ├── cadastro.py      # Tela de cadastro
│   │   └── dashboard.py     # Tela principal após login
│   ├── database
│   │   └── connection.py    # Gerenciamento da conexão com o banco de dados
│   └── main.py              # Ponto de entrada da aplicação
├── requirements.txt          # Dependências do projeto
└── README.md                 # Documentação do projeto
```

## Instalação

1. Clone o repositório:
   ```
   git clone <URL_DO_REPOSITORIO>
   cd my-python-project
   ```

2. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

## Uso

Para iniciar a aplicação, execute o arquivo `main.py`:
```
python src/main.py
```

## Funcionalidades

- **Tela de Login**: Permite que os usuários façam login no sistema.
- **Tela de Cadastro**: Permite que novos usuários se cadastrem.
- **Tela Principal**: Exibe informações relevantes após o login.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## Licença

Este projeto está licenciado sob a MIT License. Veja o arquivo LICENSE para mais detalhes.