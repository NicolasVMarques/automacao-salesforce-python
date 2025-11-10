# Automação de Cadastro de Leads no Salesforce  
### *Projeto em Python usando PyAutoGUI + Pandas para automação de processos repetitivos*  

## O que aprendi e quais problemas esse projeto resolveu  

Este projeto representa um avanço significativo na minha habilidade de automatizar processos operacionais usando Python. Durante o desenvolvimento, aprofundei meus conhecimentos em:

- **Automação de interface gráfica com PyAutoGUI**  
  - Captura de coordenadas na tela  
  - Controle de fluxo e timing para interagir com elementos dinâmicos  
  - Prevenção de erros com pausas e validações  
- **Manipulação de dados com Pandas**  
  - Leitura e limpeza de planilhas de Excel  
  - Iteração linha a linha em grandes volumes de dados  
  - Garantia de consistência e integridade dos dados enviados  

### 🎯 Problema resolvido  
Antes da automação, o processo de cadastrar leads na plataforma Salesforce era **manual, repetitivo e altamente sujeito a erros humanos**. Cada lead exigia:

- Abrir a tela de novo lead  
- Preencher vários campos manualmente  
- Validar se os dados estavam corretos  
- Repetir tudo para dezenas ou centenas de registros

Com esta automação:

- O tempo de execução diminuiu drasticamente  
- Os erros foram praticamente eliminados  
- O processo se tornou **padronizado, rastreável e muito mais eficiente**  

---

# 📂 Sobre o Projeto

Este repositório contém um script em Python que automatiza o cadastro de leads no Salesforce usando dados de uma planilha Excel. Ele lê cada linha da planilha e, usando PyAutoGUI, preenche automaticamente todos os campos na plataforma.

---

# 🛠 Tecnologias utilizadas

- **Python 3.x**
- **Pandas** — leitura e manipulação da planilha  
- **PyAutoGUI** — automação da interface gráfica  
- **Time / OS** — controle de fluxo, tempo de execução e manipulação de arquivos  

---

# 📈 Funcionamento geral

1. O script carrega uma planilha Excel contendo os dados dos leads.  
2. Para cada linha, extrai informações como:
   - nome  
   - sobrenome  
   - email  
   - telefone  
   - razão social  
   - CNPJ  
3. Abre (ou assume aberta) a tela do Salesforce.  
4. Usa PyAutoGUI para:
   - clicar em cada campo  
   - preencher os valores  
   - confirmar e registrar o lead  
5. Exibe no terminal o contato que foi enviado e confirma individualmente.
