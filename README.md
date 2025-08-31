# Programa_de_gerenciamente_de_ops

<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/6c67d457-81e2-4536-9829-1199a28df6c7" />
# App de OPs de Calçados 👟

<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/c673c916-e976-4646-963b-5375c805c4fc" />


Este é um **aplicativo profissional para gerenciamento de Ordens de Produção (OPs) de calçados**, desenvolvido em **Python com PyQt5** e banco de dados **SQLite**. Ele permite criar, visualizar, editar e exportar OPs de forma prática, rápida e organizada.

## 🔹 Funcionalidades Principais
- **Tela inicial com lista de OPs**: busca por cliente ou número da OP, abertura por duplo clique, exclusão e exportação para CSV.
- **Criação de OPs**: formulário com validações para cliente, número da OP, total de pares e tipo (Masculino ou Feminino).
- **Distribuição automática de talões**: cada giro (1–5) possui talões pré-definidos por tamanho, com controle do total de pares por talão.
- **Visualização e edição de OPs**:
  - Abas separadas por GIRO (1–5)
  - Edição direta de quantidades por tamanho com `QSpinBox`
  - Linha TOTAL LOTE automática por giro e TOTAL GERAL no rodapé
  - Validação para que cada talão some exatamente o número definido de pares
  - Alternância de status dos talões (Pendente / OK)
  - Botões de ação: salvar alterações e exportar CSV
- **Interface moderna e responsiva**:
  - Tema claro com acento roxo
  - Cantos arredondados e tabela com auto-stretch de colunas
  - Botões alinhados com a grade de OPs na tela inicial

<img width="1919" height="1080" alt="image" src="https://github.com/user-attachments/assets/cec7a9ee-5fcc-488b-9cbe-839ba340974a" />


## 🛠 Tecnologias e Dependências
- **Python 3**
- **PyQt5** (interface gráfica)
- **SQLite3** (banco de dados local)
- Apenas dependências da **stdlib + PyQt5**

## 🚀 Como executar
```bash
# Certifique-se de ter Python 3 e PyQt5 instalado
pip install PyQt5

# Execute o app
python app_op_calcados.py

