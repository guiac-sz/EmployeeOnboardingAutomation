import time
from tkinter import messagebox

import customtkinter as ctk
import win32com.client

from services.outlook_reader import carregar_usuarios

# Configurações visuais
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Criação da janela
app = ctk.CTk()
app.title("Employee Onboarding Automation")

# Abre maximizada (ocupa 100% da tela, mantendo a barra de título do Windows)
app.geometry("900x600")  # tamanho de fallback, usado antes de maximizar
app.after(10, lambda: app.state("zoomed"))

# Carrega os chamados pendentes do Outlook
usuarios = carregar_usuarios()
usuario_selecionado = None

print("\nUsuários carregados:")
for usuario in usuarios:
    print(f"{usuario['nome']} {usuario['sobrenome']} ({usuario['usuario']})")

# Monta as opções do menu (texto exibido -> dados do usuário)
PLACEHOLDER_SELECIONE = "Selecione um usuário"
PLACEHOLDER_VAZIO = "Nenhum usuário pendente"

usuarios_dict = {}
opcoes_usuarios = []

for usuario in usuarios:
    texto = f"{usuario['nome']} {usuario['sobrenome']} ({usuario['usuario']})"
    opcoes_usuarios.append(texto)
    usuarios_dict[texto] = usuario

if opcoes_usuarios:
    # Insere o placeholder como primeira opção, para o menu não abrir
    # já com um usuário selecionado sem o operador ter escolhido nada
    opcoes_usuarios.insert(0, PLACEHOLDER_SELECIONE)
else:
    opcoes_usuarios = [PLACEHOLDER_VAZIO]


def selecionar_usuario(escolha):
    # Preenche os campos de nome e usuário com base na seleção do menu
    global usuario_selecionado

    if escolha in (PLACEHOLDER_SELECIONE, PLACEHOLDER_VAZIO):
        usuario_selecionado = None
        nome_entry.delete(0, "end")
        usuario_entry.delete(0, "end")
        return

    usuario = usuarios_dict[escolha]
    usuario_selecionado = usuario

    nome_completo = f"{usuario['nome']} {usuario['sobrenome']}"

    nome_entry.delete(0, "end")
    nome_entry.insert(0, nome_completo)

    usuario_entry.delete(0, "end")
    usuario_entry.insert(0, usuario["usuario"])


def mover_para_processados():
    """Move o e-mail do chamado para a pasta 'Processados' após o envio.

    Retorna True se moveu com sucesso, False se algo deu errado
    (por exemplo, o e-mail já não estar mais na pasta original).
    """
    global usuario_selecionado

    if usuario_selecionado is None:
        return False

    email_obj = usuario_selecionado["email_obj"]

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        caixa_entrada = outlook.GetDefaultFolder(6)
        chamados_novos = caixa_entrada.Folders["Chamados Novos"]
        pasta_processados = chamados_novos.Folders["Processados"]

        print("Movendo:", email_obj.Subject)
        email_obj.Move(pasta_processados)
    except Exception as erro:
        print("Erro ao mover e-mail:", erro)
        messagebox.showwarning(
            "Aviso",
            "O e-mail foi enviado, mas não foi possível mover o chamado "
            "para a pasta 'Processados'. Pode já ter sido movido antes.\n\n"
            f"Detalhe técnico: {erro}"
        )
        return False

    messagebox.showinfo("Sucesso", "Usuário processado com sucesso.")
    return True


def enviar_email():
    """Monta e envia o e-mail de notificação de novo usuário, depois arquiva o chamado."""
    global usuario_selecionado

    if usuario_selecionado is None:
        messagebox.showwarning(
            "Nenhum usuário selecionado",
            "Selecione um usuário pendente antes de enviar o e-mail."
        )
        return

    # Guarda a referência do usuário atual antes de qualquer reset,
    # para poder removê-lo da lista de opções depois do envio
    usuario_atual = usuario_selecionado
    texto_atual = next(
        (texto for texto, dados in usuarios_dict.items() if dados is usuario_atual),
        None
    )

    nome = nome_entry.get()
    usuario = usuario_entry.get()
    senha = senha_entry.get()
    
    """Domínio de e-mail utilizado para gerar o  e-mail corporativo do usuário"""
    email_usuario = usuario + "@empresa.com"

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Display()

    # Aguarda a assinatura padrão do Outlook ser carregada no corpo do e-mail
    time.sleep(2)
    assinatura = mail.HTMLBody

    """E-mail que receberá as credenciais do novo colaborador"""
    mail.To = "rh@empresa.com"
    mail.Subject = f"Novo Usuário - {nome}"

    corpo = f"""
    <p>Bom dia,</p>

    <p>Um novo usuário foi criado em nosso sistema. Abaixo estão os dados de cadastro dele:</p>

    <table border="1" cellspacing="0" cellpadding="5">
        <tr>
            <td><b>Nome</b></td>
            <td>{nome}</td>
        </tr>
        <tr>
            <td><b>Usuário</b></td>
            <td>{usuario}</td>
        </tr>
        <tr>
            <td><b>E-Mail</b></td>
            <td>{email_usuario}</td>
        </tr>
        <tr>
            <td><b>Senha</b></td>
            <td>{senha}</td>
        </tr>
    </table>
    """

    mail.HTMLBody = corpo + assinatura

    print("Enviando e-mail...")
    mail.Send()

    mover_para_processados()

    # Remove o usuário processado da lista de opções, para que ele não
    # possa ser selecionado de novo e reenviado por engano
    if texto_atual is not None:
        usuarios_dict.pop(texto_atual, None)
        if texto_atual in opcoes_usuarios:
            opcoes_usuarios.remove(texto_atual)

    if not usuarios_dict:
        opcoes_usuarios.clear()
        opcoes_usuarios.append(PLACEHOLDER_VAZIO)
        usuarios_menu.configure(values=opcoes_usuarios)
        usuarios_menu.set(PLACEHOLDER_VAZIO)
    else:
        usuarios_menu.configure(values=opcoes_usuarios)
        usuarios_menu.set(PLACEHOLDER_SELECIONE)

    # Reseta a seleção após o envio, para evitar reenvio acidental
    # para o mesmo usuário caso o botão seja clicado de novo
    usuario_selecionado = None
    nome_entry.delete(0, "end")
    usuario_entry.delete(0, "end")


# --- Construção da interface ---

# Frame central que segura todo o conteúdo, mantendo largura fixa mesmo
# com a janela maximizada (evita que os campos fiquem esticados de ponta a ponta)
conteudo = ctk.CTkFrame(app, fg_color="transparent")
conteudo.pack(expand=True)

titulo = ctk.CTkLabel(conteudo, text="Novo Usuário", font=("Arial", 28, "bold"))
titulo.pack(pady=(40, 30))

selecionar_label = ctk.CTkLabel(conteudo, text="Usuários Pendentes")
selecionar_label.pack()

usuarios_menu = ctk.CTkOptionMenu(
    conteudo,
    values=opcoes_usuarios,
    command=selecionar_usuario,
    width=400
)
usuarios_menu.pack(pady=(5, 25))

nome_label = ctk.CTkLabel(conteudo, text="Nome")
nome_label.pack(pady=(15, 5))

nome_entry = ctk.CTkEntry(conteudo, width=400)
nome_entry.pack()

usuario_label = ctk.CTkLabel(conteudo, text="Usuário")
usuario_label.pack(pady=(15, 5))

usuario_entry = ctk.CTkEntry(conteudo, width=400)
usuario_entry.pack()

senha_label = ctk.CTkLabel(conteudo, text="Senha")
senha_label.pack(pady=(15, 5))

senha_entry = ctk.CTkEntry(conteudo, width=400)
senha_entry.pack()
"""Senha padrão enviada ao novo colaborador, altere conforme a política da sua empresa"""
senha_entry.insert(0, "Company@123")
senha_entry.configure(state="readonly")  # senha padrão, não deve ser editada manualmente

email_button = ctk.CTkButton(
    conteudo,
    text="Enviar E-Mail",
    command=enviar_email,
    width=400,
    height=40
)
email_button.pack(pady=(25, 40))

# Mantém a janela aberta
app.mainloop()