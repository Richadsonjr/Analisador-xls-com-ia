
# 📊 Analisador de Dados com IA (Gemini API + Tkinter)

Uma aplicação Python com interface gráfica que permite selecionar arquivos Excel, gerar análises automatizadas usando a API Gemini do Google e salvar os resultados em documentos Word (.docx).

---

## ✨ Funcionalidades

- Leitura de arquivos Excel (.xlsx) e conversão para JSON em memória.
- Envio dos dados para a API **Gemini 1.5 Flash** para análise baseada em prompt.
- Salvamento da resposta da IA em um arquivo **.docx**.
- Interface gráfica simples e intuitiva com status em tempo real via **Tkinter**.
- Processamento em thread separada para não travar a interface.

---

## 📦 Tecnologias Utilizadas

- **Python**
- **Pandas**
- **Google Generative AI SDK**
- **python-docx**
- **Tkinter**
- **threading**

---

## 📂 Estrutura Esperada

```
/seu-projeto
├── main.py
├── config.txt        # sua chave da API Gemini
├── prompt.txt        # prompt personalizado para análise dos dados
├── requisitos.txt    # (opcional, se quiser listar dependências)
└── README.md
```

---

## ⚙️ Pré-Requisitos

- Python 3.11+
- Conta e chave de API no **Google AI Studio (Gemini API)**

---

## 📥 Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/Richadsonjr/Analisador-xls-com-ia.git
   cd seu-repositorio
   ```

2. Instale as dependências:
   ```bash
   pip install pandas google-generativeai python-docx
   ```

---

## 🔐 Configuração

1. Crie um arquivo `config.txt` na raiz do projeto e cole sua chave da API Gemini nele.
2. Crie um arquivo `prompt.txt` contendo o prompt personalizado para a análise dos dados.

---

## 🚀 Como Usar

1. Execute o programa:
   ```bash
   python main.py
   ```

2. Na interface:
   - Selecione um arquivo **Excel (.xlsx)**.
   - Defina o local e nome do arquivo de saída **.docx**.
   - Clique em **INICIAR ANÁLISE**.

3. Acompanhe o log do processamento na própria interface.

---

## 📝 Observação

- A aplicação utiliza chamadas síncronas à API Gemini, mas executa o processamento em uma thread separada para manter a interface responsiva.
- Certifique-se de não exceder os limites de uso da sua conta Google AI Studio.

---

## 📌 Melhorias Futuras

- Suporte a múltiplas planilhas.
- Escolha de modelo Gemini diretamente pela interface.
- Ajuste de parâmetros da IA (temperature, max tokens).
- Exportação em PDF além do Word.

---

## 📑 Licença

Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` para mais detalhes.

---


