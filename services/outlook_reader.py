import win32com.client


def extrair_dados(corpo):
    """Faz o parsing do corpo do e-mail e extrai nome, sobrenome e usuário."""
    dados = {}

    for linha in corpo.splitlines():
        linha = linha.strip()
        if "Name:" in linha and "LastName" not in linha:
            dados["nome"] = linha.split("Name:")[1].strip()
        elif "LastName:" in linha:
            dados["sobrenome"] = linha.split("LastName:")[1].strip()
        elif "AD:" in linha:
            dados["usuario"] = linha.split("AD:")[1].strip()

    return dados


def carregar_usuarios():
    """Lê os chamados pendentes na pasta 'Novos Usuários' e retorna a lista de usuários."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    caixa_entrada = outlook.GetDefaultFolder(6)
    chamados_novos = caixa_entrada.Folders["Chamados Novos"]
    novos_usuarios = chamados_novos.Folders["Novos Usuários"]

    emails = novos_usuarios.Items
    emails.Sort("[ReceivedTime]", True)

    usuarios = []

    for email in emails:
        """Colocar o título do e-mail para puxar as informações"""
        if email.Subject.strip() == "Novo Usuário":
            dados = extrair_dados(email.Body)
            usuarios.append({
                "nome": dados.get("nome", ""),
                "sobrenome": dados.get("sobrenome", ""),
                "usuario": dados.get("usuario", ""),
                "email_obj": email
            })

    return usuarios