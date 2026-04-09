import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Variáveis globais para compartilhar entre funções
label_caminho = None
caminho = None

def abrirArquivo():
    global label_caminho
    global caminho

    caminho = filedialog.askopenfilename(
            title="Selecione um arquivo",
            filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
        )
    
    if caminho:
        label_caminho.configure(text=f"Arquivo selecionado:\n{caminho}")

def validarCPF(cpf):
    
    try:
        val1 = cpf[0:9]
        val2 = cpf[0:10]
        soma = 0
        tam = len(val1) + 1
        for i in range(len(val1)):
            soma += (int(val1[i]) * tam)
            tam -= 1
        resto = (soma*10)%11
        if (resto == 10): resto = 0  
        if (resto != int(cpf[9])): return False

        soma = 0
        tam = len(val2) + 1
        for i in range(len(val2)):
            soma += (int(val2[i]) * tam)
            tam -= 1
        resto = (soma*10)%11
        if (resto == 10): resto = 0  
        if (resto != int(cpf[10])): return False
        return True
    except:
        return False

def acertarCPF(cpf_inicial):
    
    cpf_inicial = str(cpf_inicial).strip()
    if len(cpf_inicial) == 10: cpf_inicial = '0' + cpf_inicial
    
    cpf_limpo = ''
    nums = '1234567890'

    for char in cpf_inicial:
        if char.isalpha(): return '(Nao pode haver letras no CPF!)'
        elif char in nums: cpf_limpo += char
    
    if len(cpf_limpo) < 11: return 'Quantidade de digitos menor que o exigido!'
    elif len(cpf_limpo) > 11: return 'Quantidade de digitos maior que o exigido!'
    elif not(validarCPF(cpf_limpo)): return 'CPF formalmente invalido!'

    return f"{cpf_limpo[0:3]}.{cpf_limpo[3:6]}.{cpf_limpo[6:9]}-{cpf_limpo[9:11]}"

def aplicarMod():
    global caminho
    
    if not caminho:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo primeiro!")
        return

    try:
        # Lendo o arquivo garantindo que CPF e NOME sejam strings
        df = pd.read_excel(caminho, dtype={'CPF': str, 'NOME': str})

        for index, linha in df.iterrows():
            nome_atual = str(linha['NOME'])
            novo_nome = nome_atual.title()

            cpf_atual = str(linha['CPF'])
            novo_cpf = acertarCPF(cpf_atual)
                
            df.at[index, 'NOME'] = novo_nome
            df.at[index, 'CPF'] = novo_cpf

        local_destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile="relacao-corrigida.xlsx"
        )
        
        if local_destino:
            df.to_excel(local_destino, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo processado com sucesso!\nSalvo como: {local_destino}")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar: {e}")

def main():
    global label_caminho

    fonte_titulo = ("Segoe UI", 24, "bold")
    fonte_botao = ("Segoe UI", 13, "bold")
    fonte_caminho = ("Consolas", 11, "bold")
    
    janela = ctk.CTk()
    ctk.set_appearance_mode("dark")
    janela.title("Corretor de Nome e CPF")
    janela.geometry("600x400")

    # Frame para organizar melhor
    frame_principal = ctk.CTkFrame(janela, corner_radius=15)
    frame_principal.pack(pady=40, padx=40, fill="both", expand=True)

    # Label Caminho
    label_titulo = ctk.CTkLabel(frame_principal, text="CORRETOR DE NOME E CPF", wraplength=500, font=fonte_titulo)
    label_titulo.pack(pady=(30, 10))

    # Botão Selecionar
    btn_selecionar = ctk.CTkButton(frame_principal, text="Selecionar Arquivo", command=abrirArquivo, fg_color="#242424", hover_color="#333333", font=fonte_botao, corner_radius=10)
    btn_selecionar.pack(pady=10)

    # Label Caminho
    label_caminho = ctk.CTkLabel(frame_principal, text="Nenhum arquivo selecionado", wraplength=500, font=fonte_caminho)
    label_caminho.pack(pady=10)

    # Botão Aplicar
    btn_aplicar = ctk.CTkButton(frame_principal, text="Aplicar Correções", command=aplicarMod, fg_color="#27ae60", hover_color="#219150", font=fonte_botao, corner_radius=10, height=40)
    btn_aplicar.pack(pady=(20, 10), padx=50, fill="x")

    janela.mainloop()

if __name__ == "__main__":
    main()