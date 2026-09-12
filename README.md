# 🦁 LION APP - Organizador de Imposto de Renda em Excel

> Ferramenta completa desenvolvida em Excel com VBA para centralizar, validar e organizar todas as informações necessárias para a declaração do Imposto de Renda Pessoa Física (IRPF).

Projeto desenvolvido como parte do Desafio da DIO - Construindo um perfil de destaque.

![Excel](https://img.shields.io/badge/Excel-VBA-green?style=for-the-badge&logo=microsoft-excel)
![Status](https://img.shields.io/badge/Status-Concluído-yellow?style=for-the-badge)
![DIO](https://img.shields.io/badge/DIO-Desafio-blue?style=for-the-badge)

---

## 📸 Preview do Sistema

| Aba Titular | Aba Informes | Aba Notas |
| :---: | :---: | :---: |
| Cadastro completo da pessoa física | Lançamento de rendimentos bancários | Controle mensal de entradas |
| ![Titular](/images/titular.png) | ![Informes](/images/informes.png) | ![Notas](/images/notas.png) |

> Crie uma pasta `/images` no seu repositório e coloque os 3 prints que você me mandou lá com os nomes acima.

---

## 🎯 Sobre o Desafio

O objetivo proposto pela DIO foi criar uma ferramenta no Excel que funcionasse como um **agregador de dados para o IR**, onde o usuário pudesse controlar suas entradas de maneira eficiente, validada e com uma interface amigável.

Fui além do básico e construí o **LION APP**, um sistema com identidade visual própria, navegação lateral e automações.

## ✨ Funcionalidades Implementadas

### 1. DADOS DO TITULAR
Formulário completo com validação de dados pessoais:
- NOME, CPF, Nascimento, Título de Eleitor
- Cônjuge, Endereço completo (Rua, Rua Abreviada, CEP)
- Contatos (Telefone, Celular, E-mail)
- Campos de controle: Houve alterações, Dependente Cônjuge, Residente Exterior
- Validação automática e máscaras de CPF, CEP e telefone.

### 2. INFORMES DE RENDIMENTOS BANCÁRIOS
Controle centralizado de contas bancárias:
- Lançamento de até 3 bancos com código + nome (Ex: 33 - Banco Santander)
- Campo de Valor Atual por banco
- Campo de Anexo com ícone (para comprovação em PDF - `bancox.pdf`)
- **Cálculo automático de TOTAL:** `R$ 2.200.000,00` no exemplo, somando todos os informes.
- Botões de navegação `ANTERIOR` e `PRÓXIMO`.

### 3. NOTAS BANCÁRIAS / EXTRATOS DE HOLERITES
Gestão de fluxo de entrada mês a mês:
- Tabela de ENTRADAS com colunas: DATA, CATEGORIA e VALOR
- Sistema de filtros (dropdown) para análise
- Layout em tabela listrada para facilitar a leitura
- Base para cálculo de receita anual.

### 4. Recursos Extras (Diferenciais)
- **UI/UX Personalizada:** Menu lateral fixo com logo LION APP, botões TITULAR / INFORMES / NOTAS com efeito de ativo (amarelo).
- **Navegação Intuitiva:** Sistema de abas com VBA - `Sheets().Visible`
- **Links Rápidos:** Botão LinkedIn e assinatura `SYSTEM BY FABIAO`
- **Validação de Dados:** Listas suspensas, formatação condicional e proteção de células.

## 🛠️ Tecnologias Utilizadas

- **Microsoft Excel**
- **VBA (Visual Basic for Applications)** para navegação e automações
- **Formatação Condicional, Validação de Dados e Funções (SOMA, SE, PROCV)**
- **Git & GitHub para documentação**

## 📂 Estrutura do Repositório

/
├── README.md
├── LION_APP_IRPF.xlsm # Arquivo principal do Excel com macros
├── /images
│ ├── titular.png
│ ├── informes.png
│ └── notas.png
├── /docs
│ ├── bancos_apoio.xlsx
│ └── script_de_alinhamentos.txt


## 🚀 Como Usar

1.  Baixe o arquivo `.xlsm`
2.  Habilite as Macros ao abrir o Excel
3.  Comece pela aba **TITULAR**, preencha seus dados
4.  Clique em `PRÓXIMO` para ir para **INFORMES** e lance seus bancos
5.  Finalize em **NOTAS** lançando suas entradas mensais
6.  Use os valores consolidados para preencher sua declaração oficial no programa da Receita Federal.

## 🧠 Aprendizados

Neste desafio pude aplicar na prática:
- Criação de sistemas organizacionais no Excel
- Uso de VBA para criar navegação tipo APP
- Importância da validação de dados para evitar erros na declaração de IR
- Documentação técnica e versionamento com GitHub
- Design de interface pensando no usuário final

## 👨‍💻 Autor

**Fabio Silva Sauro**
Desenvolvido como System by Fabiao

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Conectar-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/seu-link-aqui)

---
*Projeto para fins educacionais - Desafio DIO*
