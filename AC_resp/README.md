# 🤖 Automação de Atualização de Responsáveis - Sistema Acessórias

Este projeto em Python automatiza o processo de atualização de responsáveis pelos departamentos das empresas cadastradas no sistema [Acessórias](https://app.acessorias.com/index.php), utilizando automação web com Selenium.

---

## 📁 Estrutura do Projeto

```
AC_resp/
├── main.py               # Script principal da automação
├── requirements.txt      # Lista de dependências do projeto
├── listas.xlsx           # Planilha com os dados (não incluída no repositório)
```

> ⚠️ O arquivo `listas.xlsx` contém informações sensíveis e **não deve ser enviado para o GitHub**. Adicione-o manualmente à pasta raiz do projeto.

---

## 📋 Pré-requisitos

Antes de executar o projeto, certifique-se de ter:

- **Python 3.8 ou superior** instalado
- **Google Chrome** instalado
- Permissão para fazer login no sistema [Acessórias](https://app.acessorias.com/index.php)

---

## ⚙️ Instalação

1. **Clone este repositório**:

   ```bash
   git clone https://github.com/seuusuario/AC_resp.git
   cd AC_resp
   ```

2. **Instale os pacotes necessários**:

   ```bash
   pip install -r requirements.txt
   ```

3. **Coloque o arquivo `listas.xlsx` na raiz do projeto.**

   Esse arquivo deve conter pelo menos as seguintes colunas (sem aspas):
   - `Codigo`: Código da empresa no sistema
   - `Departamento`: Nome do departamento (exatamente como aparece no sistema)
   - `Novo Responsavel`: Nome do novo responsável

---

## 🚀 Execução

Execute o script com:

```bash
python main.py
```

### O que esperar:
- O navegador Chrome será aberto automaticamente.
- Você deverá **fazer o login manualmente** no sistema Acessórias.
- Após o login, volte para o terminal e pressione **ENTER** para continuar.
- O script irá iterar por cada linha da planilha e atualizar os responsáveis conforme os dados fornecidos.

---

## 🧠 Funcionamento Interno

- Usa **Selenium WebDriver** com o Chrome.
- Identifica empresas pelo código (`Codigo`) fornecido.
- Acessa a tela de responsáveis por departamento.
- Seleciona o novo responsável conforme informado.
- Salva as alterações.
- Repete para todas as empresas da planilha.

Em caso de falha, o script tenta novamente até 2 vezes antes de seguir para a próxima linha.

---

## 🛑 Departamentos Suportados
(Altere conforme a necessidades do seu sistema)
Os nomes dos departamentos devem estar escritos **exatamente como listado abaixo**:

- ADMINISTRAÇÃO DE PESSOAL
- ARQUIVO - EXPEDIENTE
- ROTEIRISTA
- CONTABILIDADE
- COORDENADORES
- FINANCEIRO
- FISCAL
- INFORMÁTICA
- LEGALIZAÇÃO
- ALVARÁS E LICENÇAS
- CERTIDÕES
- CERTIFICAÇÃO DIGITAL
- COMERCIAL
- TAXA MUNICIPAL
- RECEPÇÃO
- MALA DIRETA

---

## ❌ Possíveis Erros

- Departamento não reconhecido: revise a ortografia na planilha.
- Nome do responsável não aparece na lista suspensa: verifique se o usuário existe no sistema.
- Falhas esporádicas na conexão ou carregamento podem causar reexecuções.

---

## 📝 Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

---

## Autor

Desenvolvido por **Leonardo V. Rosseti.**

Contribuições e melhorias são bem-vindas! 😄
