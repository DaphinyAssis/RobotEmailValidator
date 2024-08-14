# EmailRobotValidator

### Descrição:
Uma ferramenta automatizada para filtrar e validar e-mails de notificações de robôs, extraindo e processando informações específicas sobre solicitações e envios. Ideal para monitorar e analisar notificações recebidas de contas de e-mail específicas no Outlook.

### Funcionalidades
- Filtragem de e-mails recebidos de um remetente específico (no-reply@lcontrol.com.br).
- Extração do nome do cliente e contagem de solicitações e envios a partir do corpo do e-mail.
- Agrupamento e listagem de e-mails por cliente.
- Suporte para validação de e-mails recebidos de hoje.

### Pré-requisitos
- Python 3.x
- Bibliotecas Python:
  - **pywin32**
  - **re**

### Instalação
Clone este repositório:

``git clone https://github.com/DaphinyAssis/EmailRobotValidator.git``

Navegue até o diretório do projeto:

``cd EmailRobotValidator``

Instale as dependências necessárias:

``pip install pywin32``

### Uso
Certifique-se de que você tem o Outlook configurado e o acesso à conta de e-mail correta.

Execute o script principal:

``python main.py``

O script irá filtrar os e-mails de hoje do remetente no-reply@lcontrol.com.br, extrair as informações relevantes e listar os resultados agrupados por cliente no terminal.

### Contribuição
Se você deseja contribuir para este projeto, por favor, abra um issue ou envie um pull request com suas melhorias.

### Licença
Este projeto está licenciado sob a Licença MIT. Veja o arquivo [LICENSE](./LICENSE) para mais detalhes.
