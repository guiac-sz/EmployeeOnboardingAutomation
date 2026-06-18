# EmployeeOnboardingAutomation
Ferramenta desktop para automatizar parte do processo de onboarding de novos usuários, integrando diretamente com o Outlook. O sistema lê chamados de cadastro recebidos por e-mail, permite revisar os dados antes do envio, monta e envia um e-mail formatado para a equipe de TI/RH e arquiva o chamado automaticamente.

# O que o sistema faz

1. Lê os e-mails da pasta Chamados Novos > Novos Usuários no Outlook, filtrando apenas os de assunto INTERNAL EMPLOYEE REGISTRATION.
2. Extrai automaticamente nome, sobrenome e usuário do corpo de cada e-mail.
3. Exibe os chamados pendentes em um menu para seleção manual.
4. Ao selecionar um usuário, preenche os campos de nome e login na tela para revisão.
5. Ao confirmar o envio, monta um e-mail em HTML com os dados do novo usuário (incluindo a senha padrão) e envia para a equipe responsável, preservando a assinatura padrão do Outlook.
6. Move o chamado original para a pasta Processados, evitando reprocessamento.

# Por que existe

Em ambientes corporativos, o cadastro de novos usuários frequentemente depende de processos manuais e repetitivos. Esta ferramenta automatiza a leitura, formatação e arquivamento desses chamados, mantendo uma etapa de revisão humana antes de qualquer envio.

# Requisitos

- Windows com Microsoft Outlook instalado e configurado (a integração usa COM via pywin32, portanto não funciona em macOS/Linux).
- Python 3.10 ou superior.
- Outlook deve estar aberto e logado na conta correta antes de iniciar o programa.

# Estrutura do projeto
.
├── main.py                     # Interface gráfica e fluxo principal
├── services/
│   └── outlook_reader.py       # Leitura e parsing dos e-mails do Outlook
└── README.md

# Formato esperado do e-mail de chamado

O corpo do e-mail recebido deve conter as seguintes linhas para que o parsing funcione corretamente:

Name: <nome>
LastName: <sobrenome>
AD: <usuário de rede>

Chamados fora desse padrão terão campos em branco e devem ser conferidos manualmente antes do envio.

# Limitações conhecidas

- A lista de chamados pendentes é carregada apenas na abertura do programa; novos chamados que chegarem durante o uso só aparecem após reiniciar a aplicação.
- A senha padrão é fixa no código-fonte (campo somente leitura na interface). Não há reset de senha automático no Active Directory — essa etapa continua manual, pois a gestão do AD está fora do escopo de acesso da equipe.
- O envio depende do Outlook estar aberto e responsivo; não há fila de retentativa em caso de falha de conexão.
- Não há autenticação ou controle de quem operou cada envio — qualquer pessoa com acesso ao computador e ao Outlook pode usar a ferramenta.

# Tecnologias

- Python
- CustomTkinter
- Outlook (via pywin32)
- SQLite (previsto para quando o histórico for implementado)

Imagem 1 — tela inicial, menu "Selecione um usuário" e campos vazios

<img width="1364" height="717" alt="image" src="https://github.com/user-attachments/assets/79b9da57-b2e3-4af5-aa12-dc89ed4f8c98" />


Imagem 2 — tela após selecionar um usuário pendente, com nome, usuário e senha preenchidos

<img width="1003" height="511" alt="image" src="https://github.com/user-attachments/assets/3458e880-d35b-4f3b-9143-2571a9a13608" />


Imagem 3 — mensagem de sucesso após envio e arquivamento

<img width="729" height="385" alt="image" src="https://github.com/user-attachments/assets/696f6a77-3354-4882-b71b-6a6dc33544a5" />



