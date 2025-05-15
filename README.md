# Filtro de Dados por UF em Editais Públicos (MG)
🗓️ Projeto desenvolvido em: **Maio de 2022**

Este projeto automatiza o tratamento de dados de editais públicos no Excel, filtrando especificamente os registros pertencentes ao estado de Minas Gerais (UF = MG). A solução foi desenvolvida em **VBA (Visual Basic for Applications)**.

## 🚀 Funcionalidades

- Tratamento e padronização de dados textuais extraídos de editais
- Conversão de texto bruto em colunas organizadas (Nome, CPF, Protocolo, etc.)
- Filtragem automática para registros de MG
- Geração de relatório de contagem total e filtrada
- Ocultação e proteção de planilhas de controle

## 💻 Tecnologias Utilizadas

- Microsoft Excel + VBA

## 📂 Estrutura

- `FiltroMG.bas`: Código principal da macro
- `Exemplo.xlsx` (opcional): Planilha com estrutura esperada
- `README.md`: Documentação do projeto

## 🧾 Como Usar

1. Abra o Excel e pressione `Alt + F11` para acessar o Editor do VBA.
2. Importe o arquivo `FiltroMG.bas` para o seu projeto VBA.
3. Certifique-se de que a planilha contém os dados na célula A1 da aba “Controle”.
4. Execute a macro `Filtro_MG`.

> O script automaticamente irá:
> - Processar os dados
> - Filtrar registros de MG (UF = 6)
> - Copiar os resultados filtrados para a coluna K
> - Gerar uma mensagem com o total processado e total de MG

- ### 📊 Resultados Obtidos

Este projeto foi desenvolvido sob demanda de um escritório de advocacia previdenciária. O objetivo era automatizar a filtragem de CPFs de pessoas do estado de **Minas Gerais (UF = 6)** que tiveram benefícios do INSS cancelados, com a finalidade de **identificar potenciais clientes** e oferecer suporte jurídico para a reativação desses benefícios.

Antes da automação, a equipe levava em média **20 minutos por edital** para identificar manualmente os registros relevantes. Com a macro em VBA, esse tempo foi reduzido para cerca de **1 minuto**, aumentando a produtividade e a precisão da triagem.

Apesar de o projeto estar **finalizado e funcional**, sua continuidade foi interrompida por decisões internas da empresa. Ainda assim, ele permanece como um exemplo prático de como a automação pode ser aplicada para otimizar processos de análise de dados em ambientes reais.

## 📌 Observações

- O código foi usado em ambiente real, com foco em eficiência e redução de erros humanos.
- Este projeto é de uso educacional e não deve ser comercializado sem autorização do autor.

---

**Desenvolvido por Caio Cesar de Albuquerque**  
📫 [caioalbuquerquedev@gmail.com](mailto:caioalbuquerquedev@gmail.com)  
🔗 [LinkedIn](https://www.linkedin.com/in/caio-cesar-for-hire) | [GitHub](https://github.com/Caio-Cesa)
