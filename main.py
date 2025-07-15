import csv
import os
import locale
from tkinter import *
from tkinter import messagebox, ttk
from docx import Document
from docx.shared import Cm

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
ARQUIVO_ESTOQUE = "estoque.csv"

def centralizar_janela(janela, largura, altura):
    janela.update_idletasks()
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela // 2) - (largura // 2)
    y = (altura_tela // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

def abrir_arquivo_default(path):
    try:
        os.startfile(path)
    except Exception as e:
        messagebox.showwarning("Aviso", f"Não foi possível abrir o arquivo automaticamente.\nAbra manualmente: {path}")

def carregar_estoque():
    estoque = []
    if not os.path.exists(ARQUIVO_ESTOQUE):
        with open(ARQUIVO_ESTOQUE, "w", newline='', encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Modelo", "Ano", "Preço"])
        return estoque
    with open(ARQUIVO_ESTOQUE, "r", newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row['Modelo']:
                estoque.append({
                    "Modelo": row["Modelo"],
                    "Ano": int(row["Ano"]),
                    "Preço": float(row["Preço"])
                })
    return estoque

def salvar_estoque(estoque):
    with open(ARQUIVO_ESTOQUE, "w", newline='', encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Modelo", "Ano", "Preço"])
        for carro in estoque:
            writer.writerow([carro["Modelo"], carro["Ano"], carro["Preço"]])

def atualizar_treeview(tree, estoque):
    tree.delete(*tree.get_children())
    for idx, carro in enumerate(estoque):
        preco = locale.currency(carro["Preço"], grouping=True)
        tree.insert("", "end", iid=idx, text=idx, values=(carro["Modelo"], carro["Ano"], preco))

def adicionar_carro_gui(tree, estoque):
    def salvar():
        modelo = entry_modelo.get().strip()
        ano = entry_ano.get().strip()
        preco = entry_preco.get().strip().replace(",", ".")
        if not modelo:
            messagebox.showerror("Erro", "Modelo não pode ser vazio.")
            return
        try:
            ano = int(ano)
            if ano < 1900 or ano > 2100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Ano inválido.")
            return
        try:
            preco = float(preco)
            if preco < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Preço inválido.")
            return
        estoque.append({"Modelo": modelo, "Ano": ano, "Preço": preco})
        salvar_estoque(estoque)
        atualizar_treeview(tree, estoque)
        win_add.destroy()
        messagebox.showinfo("Sucesso", f"Carro '{modelo}' adicionado!")
    win_add = Toplevel()
    win_add = Toplevel()
    win_add.title("Adicionar Carro")
    centralizar_janela(win_add, 350, 200)
    win_add.resizable(False, False)
    win_add.grab_set()
    Label(win_add, text="Modelo:", font=('Arial', 13)).grid(row=0, column=0, sticky="e", padx=8, pady=8)
    entry_modelo = Entry(win_add, font=('Arial', 13))
    entry_modelo.grid(row=0, column=1, padx=8, pady=8)
    Label(win_add, text="Ano:", font=('Arial', 13)).grid(row=1, column=0, sticky="e", padx=8, pady=8)
    entry_ano = Entry(win_add, font=('Arial', 13))
    entry_ano.grid(row=1, column=1, padx=8, pady=8)
    Label(win_add, text="Preço:", font=('Arial', 13)).grid(row=2, column=0, sticky="e", padx=8, pady=8)
    entry_preco = Entry(win_add, font=('Arial', 13))
    entry_preco.grid(row=2, column=1, padx=8, pady=8)
    Button(win_add, text="Salvar", font=('Arial', 12), width=13, command=salvar).grid(row=3, column=0, columnspan=2, pady=14)

def editar_carro_gui(tree, estoque):
    selected = tree.focus()
    if not selected:
        messagebox.showerror("Erro", "Selecione um carro para editar.")
        return
    idx = int(selected)
    carro = estoque[idx]
    def salvar():
        modelo = entry_modelo.get().strip()
        ano = entry_ano.get().strip()
        preco = entry_preco.get().strip().replace(",", ".")
        if not modelo:
            messagebox.showerror("Erro", "Modelo não pode ser vazio.")
            return
        try:
            ano = int(ano)
            if ano < 1900 or ano > 2100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Ano inválido.")
            return
        try:
            preco = float(preco)
            if preco < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Preço inválido.")
            return
        carro["Modelo"] = modelo
        carro["Ano"] = ano
        carro["Preço"] = preco
        salvar_estoque(estoque)
        atualizar_treeview(tree, estoque)
        win_edit.destroy()
        messagebox.showinfo("Sucesso", "Carro editado com sucesso!")
    win_edit = Toplevel()
    win_edit.title("Editar Carro")
    centralizar_janela(win_edit, 350, 200)
    win_edit.resizable(False, False)
    win_edit.grab_set()
    Label(win_edit, text="Modelo:", font=('Arial', 13)).grid(row=0, column=0, sticky="e", padx=8, pady=8)
    entry_modelo = Entry(win_edit, font=('Arial', 13))
    entry_modelo.insert(0, carro["Modelo"])
    entry_modelo.grid(row=0, column=1, padx=8, pady=8)
    Label(win_edit, text="Ano:", font=('Arial', 13)).grid(row=1, column=0, sticky="e", padx=8, pady=8)
    entry_ano = Entry(win_edit, font=('Arial', 13))
    entry_ano.insert(0, carro["Ano"])
    entry_ano.grid(row=1, column=1, padx=8, pady=8)
    Label(win_edit, text="Preço:", font=('Arial', 13)).grid(row=2, column=0, sticky="e", padx=8, pady=8)
    entry_preco = Entry(win_edit, font=('Arial', 13))
    entry_preco.insert(0, carro["Preço"])
    entry_preco.grid(row=2, column=1, padx=8, pady=8)
    Button(win_edit, text="Salvar", font=('Arial', 12), width=13, command=salvar).grid(row=3, column=0, columnspan=2, pady=14)

def excluir_carro_gui(tree, estoque):
    selected = tree.focus()
    if not selected:
        messagebox.showerror("Erro", "Selecione um carro para excluir.")
        return
    idx = int(selected)
    carro = estoque[idx]
    confirm = messagebox.askyesno("Confirmar Exclusão", f"Excluir '{carro['Modelo']} {carro['Ano']}'?")
    if confirm:
        estoque.pop(idx)
        salvar_estoque(estoque)
        atualizar_treeview(tree, estoque)
        messagebox.showinfo("Sucesso", "Carro excluído.")

def gerar_arquivo_para_impressao(estoque):
    if not estoque:
        messagebox.showerror("Erro", "Estoque vazio! Nenhum arquivo gerado.")
        return
    doc = Document()
    doc.add_heading("Estoque de Veículos", 0)
    tabela = doc.add_table(rows=1, cols=6)
    tabela.style = "Table Grid"
    hdr_cells = tabela.rows[0].cells
    hdr_cells[0].text = 'Modelo'
    hdr_cells[1].text = 'Ano'
    hdr_cells[2].text = 'Preço (R$)'
    hdr_cells[3].text = 'Fotos'
    hdr_cells[4].text = 'Site'
    hdr_cells[5].text = 'Stories'
    col_widths = [Cm(7.7), Cm(2.5), Cm(3.8), Cm(0.3), Cm(0.3), Cm(0.3)]
    for i, width in enumerate(col_widths):
        hdr_cells[i].width = width
    for carro in estoque:
        row_cells = tabela.add_row().cells
        row_cells[0].text = carro["Modelo"]
        row_cells[1].text = str(carro["Ano"])
        row_cells[2].text = locale.currency(carro["Preço"], grouping=True)
        row_cells[3].text = ""
        row_cells[4].text = ""
        row_cells[5].text = ""
    doc.add_paragraph("\n")
    nome_arquivo = "estoque_para_impressao.docx"
    doc.save(nome_arquivo)
    messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo}' gerado com sucesso!")
    abrir_arquivo_default(nome_arquivo)

def main():
    estoque = carregar_estoque()
    root = Tk()
    root.title("Estoque de Veículos Usados - Loja")
    root.geometry("750x430")
    # ---- ESTILO TREEVIEW ----
    style = ttk.Style(root)
    style.configure("Treeview.Heading", font=('Arial', 15, 'bold'))
    style.configure("Treeview", font=('Arial', 14), rowheight=38, foreground='black')
    style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
    style.map('Treeview', background=[('selected', '#ececec')])

    Label(root, text="Controle de Estoque de Veículos Usados", font=("Arial", 19, "bold")).pack(pady=10)
    frame = Frame(root)
    frame.pack(fill=BOTH, expand=True)
    columns = ("Modelo", "Ano", "Preço")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    tree.heading("Modelo", text="Modelo", anchor="center")
    tree.heading("Ano", text="Ano", anchor="center")
    tree.heading("Preço", text="Preço", anchor="center")
    tree.column("Modelo", width=320, anchor="center")
    tree.column("Ano", width=100, anchor="center")
    tree.column("Preço", width=170, anchor="center")
    tree.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side=RIGHT, fill=Y)
    atualizar_treeview(tree, estoque)
    btn_frame = Frame(root)
    btn_frame.pack(pady=12)
    Button(btn_frame, text="Adicionar Carro", font=('Arial', 13), width=18, command=lambda: adicionar_carro_gui(tree, estoque)).grid(row=0, column=0, padx=8)
    Button(btn_frame, text="Editar Carro", font=('Arial', 13), width=18, command=lambda: editar_carro_gui(tree, estoque)).grid(row=0, column=1, padx=8)
    Button(btn_frame, text="Excluir Carro", font=('Arial', 13), width=18, command=lambda: excluir_carro_gui(tree, estoque)).grid(row=0, column=2, padx=8)
    Button(btn_frame, text="Gerar Arquivo para Impressão", font=('Arial', 13), width=24, command=lambda: gerar_arquivo_para_impressao(estoque)).grid(row=0, column=3, padx=8)
    Button(btn_frame, text="Sair", font=('Arial', 13), width=10, command=root.destroy).grid(row=0, column=4, padx=8)
    root.mainloop()

if __name__ == "__main__":
    main()