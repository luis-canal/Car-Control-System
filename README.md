# 🚗 Sistema de Controle de Estoque de Veículos

Este é um sistema **desktop em Python** para **controle de estoque de veículos**, com interface gráfica simples, moderna e totalmente funcional.  
Permite **adicionar, editar, excluir e visualizar carros** no estoque, além de **gerar arquivos para impressão** e **manter histórico de vendas** automaticamente.

---

## 🧩 Funcionalidades

✅ **Adicionar carros** com modelo, ano e preço  
✅ **Editar informações** de carros existentes  
✅ **Excluir carros**, registrando a venda automaticamente no histórico  
✅ **Gerar documento (.docx)** com o estoque atual para impressão  
✅ **Visualizar histórico de vendas** salvas em CSV  
✅ **Ordenar carros** por ano ou preço com um clique  
✅ **Cálculo automático** no rodapé:
- Total de veículos  
- Média de preço  
- Valor total em estoque  
- Média de ano dos carros  
✅ **Feedback visual e sonoro** em todas as ações (sucesso/erro)  
✅ **Interface centralizada e responsiva**, feita com Tkinter  

---

## 🖥️ Interface

A interface foi construída com `Tkinter` e `ttk`, oferecendo uma navegação fluida e intuitiva:  

- **Tabela central (Treeview)** com listagem dos carros  
- **Botões de ação** organizados em linha inferior  
- **Rodapé dinâmico** com estatísticas do estoque  
- **Mensagens de feedback** coloridas e discretas  

---

## 🗃️ Estrutura de Arquivos

| Arquivo | Descrição |
|----------|------------|
| `estoque.csv` | Base principal com os carros em estoque |
| `historico_vendas.csv` | Registro de veículos vendidos (data, modelo, ano, preço) |
| `estoque_para_impressao.docx` | Documento gerado automaticamente para impressão |
| `main.py` | Código-fonte principal do sistema |

---

## 🧰 Tecnologias Utilizadas

- **Python 3.8+**
- **Tkinter** – interface gráfica  
- **CSV** – armazenamento local de dados  
- **Python-docx** – geração de arquivos Word  
- **Locale** – formatação monetária e regional  
- **Winsound / Bell** – alertas sonoros  
- **Datetime** – registro de data de venda  

---

## ⚙️ Instalação

### 1️⃣ Clonar o repositório
```bash
git clone https://github.com/seuusuario/controle-estoque-veiculos.git
cd controle-estoque-veiculos
