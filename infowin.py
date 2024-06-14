import subprocess
import os
import customtkinter
from tkinter import messagebox, simpledialog

# Função para obter o nome do domínio do Windows
def get_domain_name():
    try:
        command = 'powershell "(Get-WmiObject Win32_ComputerSystem).Domain"'
        domain_name = subprocess.check_output(command, shell=True).decode().strip()

        if domain_name:
            messagebox.showinfo("Nome do Domínio", f"O nome do domínio desta máquina é: {domain_name}")
        else:
            messagebox.showwarning("Erro", "Não foi possível obter o nome do domínio.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao executar o comando PowerShell: {e}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")

# Função para alterar o domínio do Windows
def set_domain_name():
    try:
        # Obter o nome do domínio atual
        command = 'powershell "(Get-WmiObject Win32_ComputerSystem).Domain"'
        current_domain = subprocess.check_output(command, shell=True).decode().strip()

        # Solicita o novo nome do domínio
        new_domain = simpledialog.askstring("Alterar Domínio", "Digite o novo nome do domínio:")
        if not new_domain:
            return

        if new_domain.lower() == current_domain.lower():
            messagebox.showinfo("Nenhuma Alteração", f"A máquina já está no domínio {new_domain}.")
            return

        # Solicita o nome de usuário e senha
        username = simpledialog.askstring("Nome de Usuário", "Digite o nome de usuário (com permissões adequadas):")
        password = simpledialog.askstring("Senha", "Digite a senha:", show='*')
        if not username or not password:
            return

        # Comando PowerShell para alterar o domínio
        command = (
            f'powershell "$securePassword = ConvertTo-SecureString \'{password}\' -AsPlainText -Force; '
            f'$credential = New-Object System.Management.Automation.PSCredential(\'{username}\', $securePassword); '
            f'Add-Computer -DomainName {new_domain} -Credential $credential -Force -PassThru"'
        )
        subprocess.check_call(command, shell=True)

        messagebox.showinfo("Sucesso",
                            f"O domínio foi alterado para {new_domain}. Por favor, reinicie o computador para aplicar as mudanças.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao executar o comando PowerShell: {e}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")

# Função para executar o OCS Inventory Agent
def run_ocs_agent():
    try:
        ocs_path = r"C:\Program Files\OCS Inventory Agent\OcsSystray.exe"
        if os.path.exists(ocs_path):
            subprocess.Popen(ocs_path)
        else:
            messagebox.showerror("Erro", "O OCS Inventory Agent não foi encontrado no caminho especificado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao executar o OCS Inventory Agent: {e}")

# Função para instalar o AnyDesk a partir de um local de rede
def install_anydesk():
    try:
        anydesk_path = r"\\IMPRESSORAS\gpo\AnyDesk-CM.exe"
        if os.path.exists(anydesk_path):
            subprocess.Popen(anydesk_path)
        else:
            messagebox.showerror("Erro", "O AnyDesk não foi encontrado no caminho especificado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao instalar o AnyDesk: {e}")

# Função para obter a versão do Windows
def get_windows_version():
    try:
        command = 'powershell "Get-ItemProperty -Path \'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\' | Select-Object -ExpandProperty ProductName"'
        windows_version = subprocess.check_output(command, shell=True).decode().strip()

        if windows_version:
            messagebox.showinfo("Versão do Windows", f"A versão do Windows desta máquina é: {windows_version}")
        else:
            messagebox.showwarning("Erro", "Não foi possível obter a versão do Windows.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao executar o comando PowerShell: {e}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")

# Configuração do CustomTkinter
customtkinter.set_appearance_mode("System")  # Modos: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Temas: blue (default), dark-blue, green

# Criação da janela principal usando CustomTkinter
app = customtkinter.CTk()
app.title("Gerenciar Domínio, OCS e Instalar AnyDesk")

# Configurações do tamanho da janela e cor de fundo
app.geometry("500x500")

# Título da aplicação
label_title = customtkinter.CTkLabel(app, text="Gerenciar Domínio, OCS e Instalar AnyDesk", font=("Helvetica", 18, "bold"))
label_title.pack(pady=20)

# Frame para agrupar os botões
frame = customtkinter.CTkFrame(app)
frame.pack(pady=20, padx=20, fill="both", expand=True)

# Botão para obter o nome do domínio
button_get = customtkinter.CTkButton(frame, text="Obter Nome do Domínio", command=get_domain_name, width=250, height=40)
button_get.pack(pady=10)

# Botão para alterar o nome do domínio
button_set = customtkinter.CTkButton(frame, text="Alterar Nome do Domínio", command=set_domain_name, width=250, height=40)
button_set.pack(pady=10)

# Botão para executar o OCS Inventory Agent
button_ocs = customtkinter.CTkButton(frame, text="Executar OCS Inventory Agent", command=run_ocs_agent, width=250, height=40)
button_ocs.pack(pady=10)

# Botão para instalar o AnyDesk
button_anydesk = customtkinter.CTkButton(frame, text="Instalar AnyDesk", command=install_anydesk, width=250, height=40)
button_anydesk.pack(pady=10)

# Botão para obter a versão do Windows
button_windows_version = customtkinter.CTkButton(frame, text="Verificar Versão do Windows", command=get_windows_version, width=250, height=40)
button_windows_version.pack(pady=10)

# Inicia o loop principal da interface
app.mainloop()
