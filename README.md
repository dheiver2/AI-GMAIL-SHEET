# Gmail to Google Sheets Automation

Este script automatiza a extração de e-mails do Gmail e os salva em uma planilha do Google Sheets. Ele busca e-mails na caixa de entrada, extrai informações como data, remetente, assunto, corpo do e-mail, anexos e etiquetas, e os organiza em colunas dentro da planilha.

## Recursos
- **Extração de e-mails**: Obtém e-mails diretamente da caixa de entrada do Gmail.
- **Filtragem e limpeza**: Remove tags HTML, scripts e outros elementos desnecessários do corpo do e-mail.
- **Armazenamento de anexos**: Salva os anexos no Google Drive e registra o link na planilha.
- **Etiquetas do Gmail**: Recupera e salva as etiquetas atribuídas aos e-mails.
- **Formatação automática**: Ajusta automaticamente colunas e quebra de linha para melhor legibilidade.

## Como Usar

### Passo 1: Configuração
1. Acesse o [Google Apps Script](https://script.google.com/).
2. Crie um novo projeto e copie o código do script.
3. Certifique-se de estar logado na conta do Google associada ao Gmail e ao Google Sheets.

### Passo 2: Execução
1. No editor de código do Apps Script, clique em `Executar`.
2. Conceda as permissões necessárias quando solicitado.
3. O script buscará e-mails na sua caixa de entrada e os registrará na planilha ativa.

### Passo 3: Automatização
Para automatizar a execução:
1. Clique em `Acionadores` (Triggers) no Apps Script.
2. Configure um acionador para rodar o script em intervalos regulares, como diário ou semanal.

## Configuração Personalizada
O script pode ser personalizado na seguinte seção:
```javascript
var config = {
  limiteCaracteresCorpo: 1000,
  nomePlanilha: "Emails",
  colunas: ["Data", "Remetente", "Assunto", "Corpo do E-mail", "Link do Anexo", "Etiquetas"]
};
```
- **`limiteCaracteresCorpo`**: Define o número máximo de caracteres exibidos no corpo do e-mail.
- **`nomePlanilha`**: Nome da aba onde os e-mails serão armazenados.
- **`colunas`**: Personaliza as colunas exibidas na planilha.

## Erros Comuns e Soluções
- **Erro: `mensagem.getLabels is not a function`**
  - Solução: Substituir `mensagem.getLabels()` por `mensagem.getThread().getLabels()`.
- **Erro ao capturar anexos**
  - Verifique se a conta tem permissão para acessar o Google Drive.
- **Nenhum e-mail está sendo salvo**
  - Certifique-se de que existem e-mails na caixa de entrada e que o script tem permissão para acessar o Gmail.

## Licença
Este projeto é distribuído sob a licença MIT. Sinta-se à vontade para modificá-lo e utilizá-lo conforme sua necessidade.

## Autor
Desenvolvido por **Dheiver Santos**. Conecte-se no [LinkedIn](https://www.linkedin.com/in/dheiver-santos/) para mais projetos e discussões sobre automação e IA.

## Contribuições
Contribuições são bem-vindas! Se você encontrar bugs ou tiver sugestões de melhorias, abra uma issue ou envie um pull request no [repositório do GitHub](https://github.com/dheiversantos/gmail-to-sheets).

