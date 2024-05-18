import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def selecionar_arquivo(titulo):
    caminho_arquivo = filedialog.askopenfilename(title=titulo, filetypes=[("Excel files", "*.xlsx")])
    return caminho_arquivo

def copiar_todos_os_registros(caminho_plan1, plan2):
    # Ler a planilha
    plan1 = pd.read_excel(caminho_plan1)
    
    # Verificar se a plan1 está vazia
    if plan1.empty:
        # Copiar todos os registros de plan2 para plan1
        plan1 = plan2.copy()
        # Salvar a plan1 com os novos registros
        plan1.to_excel(caminho_plan1, index=False)
        # Informar o usuário que os arquivos foram copiados
        messagebox.showinfo("Sucesso", "A planilha plan1 estava vazia. Os registros foram copiados de plan2 para plan1.")
    else:
        # Verificar a última linha preenchida em ambas as planilhas
        ultima_linha_plan1 = plan1.iloc[-1]
        ultima_linha_plan2 = plan2.iloc[-1]

        # Comparar a última linha preenchida e o número de itens (linhas) nas planilhas
        if not ultima_linha_plan1.equals(ultima_linha_plan2) or len(plan1) != len(plan2):
            # Identificar novos registros em plan2 que não estão em plan1
            novos_registros = plan2[~plan2.isin(plan1.to_dict(orient='list')).all(axis=1)]

            if not novos_registros.empty:
                # Copiar novos registros de plan2 para plan1
                plan1 = pd.concat([plan1, novos_registros], ignore_index=True)
                # Salvar a plan1 com os novos registros
                plan1.to_excel(caminho_plan1, index=False)
                # Informar o usuário que os arquivos foram copiados
                messagebox.showinfo("Sucesso", "Os novos registros foram copiados de plan2 para plan1.")
            else:
                messagebox.showinfo("Informação", "Não há novos registros para copiar de plan2 para plan1.")
        else:
            # Informar o usuário que os itens já foram copiados
            messagebox.showinfo("Informação", "Os itens já foram copiados anteriormente. As últimas linhas e a quantidade de itens são iguais.")

def comparar_e_atualizar(caminho_plan1, caminho_plan2):
    # Ler as planilhas
    plan1 = pd.read_excel(caminho_plan1)
    plan2 = pd.read_excel(caminho_plan2)

    # Verificar se as colunas são iguais
    if list(plan1.columns) != list(plan2.columns):
        messagebox.showerror("Erro", "As colunas das planilhas são diferentes.")
        return

    # Identificar células diferentes entre as planilhas
    celulas_diferentes = (plan1 != plan2).stack()
    indices_diferentes = celulas_diferentes[celulas_diferentes].index

    # Atualizar plan1 com os valores diferentes de plan2
    for idx in indices_diferentes:
        plan1.at[idx] = plan2.at[idx]

    # Salvar a plan1 atualizada
    plan1.to_excel(caminho_plan1, index=False)
    messagebox.showinfo("Sucesso", "As diferenças foram atualizadas na planilha plan1.")

def selecionar_opcao():
    # Criar a janela principal
    root = tk.Tk()
    root.title("Opções")

    # Função para executar a opção selecionada
    def executar_opcao(opcao):
        if opcao == 1:
            caminho_plan1 = selecionar_arquivo("Selecione o arquivo de destino (plan1)")
            caminho_plan2 = selecionar_arquivo("Selecione o arquivo de origem (plan2)")
            if caminho_plan1 and caminho_plan2:
                copiar_todos_os_registros(caminho_plan1, pd.read_excel(caminho_plan2))
        elif opcao == 2:
            caminho_plan1 = selecionar_arquivo("Selecione o arquivo de destino (plan1)")
            caminho_plan2 = selecionar_arquivo("Selecione o arquivo de origem (plan2)")
            if caminho_plan1 and caminho_plan2:
                comparar_e_atualizar(caminho_plan1, caminho_plan2)
        root.destroy()

    # Adicionar botões de opção
    botao_opcao1 = tk.Button(root, text="1 - Copiar todos os registros de plan2 para plan1", command=lambda: executar_opcao(1))
    botao_opcao1.pack(pady=10)
    botao_opcao2 = tk.Button(root, text="2 - Comparar e atualizar apenas as diferenças", command=lambda: executar_opcao(2))
    botao_opcao2.pack(pady=10)

    # Iniciar o loop da janela
    root.mainloop()

def mostrar_instrucoes():
    # Criar a janela principal
    root = tk.Tk()
    root.title("Instruções")

    # Texto de instruções
    instrucoes = """**** INFORMAÇÕES IMPORTANTES ****
    
    1 - Escolha entre as opções Copiar todos os registros entre 2 planilhas ou Comparar e atualizar apenas as diferenças
    2 - Escolha a planilha de Destino.
    3 - Escolha a planilha de Origem.
    4 - A opção atualizar dados ira apenas atualizar as informações que estejam diferente entre a planilha de Origem e destino.

    Desenvolvido por: Samuel Santos V 1.0"""

    # Adicionar o texto à janela
    label_instrucoes = tk.Label(root, text=instrucoes, justify="left", padx=10, pady=10, font=("Arial", 10, "bold"))
    label_instrucoes.pack()

    # Adicionar o botão "INICIAR CÓPIA"
    botao_iniciar = tk.Button(root, text="Iniciar", command=lambda: [root.destroy(), selecionar_opcao()])
    botao_iniciar.pack(pady=20)

    # Iniciar o loop da janela
    root.mainloop()

if __name__ == "__main__":
    mostrar_instrucoes()
