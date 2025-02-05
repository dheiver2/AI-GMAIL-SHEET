function salvarEmailsNoSheets() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var config = {
      limiteCaracteresCorpo: 1000,
      nomePlanilha: "Emails",
      colunas: ["Data", "Remetente", "Assunto", "Corpo do E-mail", "Link do Anexo", "Etiquetas"]
    };

    // Criar cabeçalhos se ainda não existirem
    if (planilha.getLastRow() === 0) {
      planilha.appendRow(config.colunas);
    }

    var emails = GmailApp.search('is:inbox'); // Busca todos os e-mails da caixa de entrada

    emails.forEach(function(thread) {
      var mensagens = thread.getMessages();
      
      mensagens.forEach(function(mensagem) {
        var data = mensagem.getDate();
        var remetente = mensagem.getFrom();
        var assunto = mensagem.getSubject();
        var corpoHTML = mensagem.getBody(); // Obtém o corpo do e-mail em HTML
        var corpoTexto = mensagem.getPlainBody(); // Obtém o corpo do e-mail sem formatação

        // Formatar corpo do e-mail usando a função de limpeza avançada
        var corpoFormatado = limparConteudoAvancado(corpoHTML || corpoTexto);

        // Limita o tamanho do corpo do e-mail para garantir o "clip"
        corpoFormatado = aplicarClip(corpoFormatado, config.limiteCaracteresCorpo);

        // Captura os anexos e envia para o Google Drive
        var linkAnexo = capturarAnexos(mensagem);

        // Obtém as etiquetas (labels) do e-mail
        var etiquetas = obterEtiquetas(mensagem);

        // Valida os dados antes de adicionar à planilha
        if (validarDados(data, remetente, assunto, corpoFormatado, linkAnexo, etiquetas)) {
          planilha.appendRow([data, remetente, assunto, corpoFormatado, linkAnexo, etiquetas]);
        } else {
          Logger.log("Dados inválidos encontrados e ignorados.");
        }
      });
    });

    // Ativa a quebra de linha na coluna de conteúdo (coluna 4, "Corpo do E-mail")
    planilha.getRange(2, 4, planilha.getLastRow() - 1, 1).setWrap(true);

    // Ajuste de largura das colunas para permitir overflow controlado
    planilha.setColumnWidths(1, 6, 300); // Ajusta a largura das colunas conforme necessário

    Logger.log("Processo concluído com sucesso!");
  } catch (error) {
    Logger.log("Erro durante a execução do script: " + error.toString());
  }
}

/**
 * Valida os dados antes de adicionar à planilha
 */
function validarDados(data, remetente, assunto, corpoFormatado, linkAnexo, etiquetas) {
  if (!data || !remetente || !assunto || !corpoFormatado || !etiquetas) {
    return false;
  }
  return true;
}

/**
 * Limita o tamanho do corpo do e-mail para um número máximo de caracteres
 */
function aplicarClip(texto, limite) {
  if (texto.length > limite) {
    return texto.substring(0, limite) + "..."; // Adiciona '...' no final do texto cortado
  }
  return texto;
}

/**
 * Captura os anexos de um e-mail e os envia para o Google Drive
 */
function capturarAnexos(mensagem) {
  try {
    var anexos = mensagem.getAttachments();
    var linkAnexo = "";

    anexos.forEach(function(anexo) {
      // Salva o arquivo no Google Drive
      var arquivo = DriveApp.createFile(anexo);
      // Obtenha o link do arquivo
      linkAnexo = arquivo.getUrl(); // Obtém o link do arquivo no Drive
    });

    return linkAnexo || "Sem anexo"; // Caso não haja anexos, retorna "Sem anexo"
  } catch (error) {
    Logger.log("Erro ao capturar anexos: " + error.toString());
    return "Erro ao capturar anexos";
  }
}

/**
 * Formata o corpo do e-mail para melhorar a visualização no Google Sheets
 */
function limparConteudoAvancado(texto) {
  if (!texto) return "Sem conteúdo";

  // 1. Remove caracteres especiais HTML (como &#847; e &zwnj;)
  texto = texto.replace(/&#[0-9]+;/g, ""); // Remove códigos numéricos HTML
  texto = texto.replace(/&[a-zA-Z]+;/g, ""); // Remove entidades HTML como &nbsp;, &zwnj;, etc.

  // 2. Remove todas as declarações @font-face
  texto = texto.replace(/@font-face\s*{[^}]*}/g, "");

  // 3. Remove todas as tags <style> (CSS)
  texto = texto.replace(/<style[^>]*>[\s\S]*?<\/style>/g, "");

  // 4. Remove todas as tags <script> (JS)
  texto = texto.replace(/<script[^>]*>[\s\S]*?<\/script>/g, "");

  // 5. Remove qualquer tag HTML (tag HTML sem exceção)
  texto = texto.replace(/<[^>]+>/g, "");

  // 6. Remove atributos de links (href, target, etc.)
  texto = texto.replace(/<a\s+[^>]*>(.*?)<\/a>/g, "$1");

  // 7. Remove imagens (tags <img>)
  texto = texto.replace(/<img\s+[^>]*>/g, "");

  // 8. Remove espaços extras, quebras de linha e tabulações
  texto = texto.replace(/\s+/g, " "); // Substitui múltiplos espaços por um único
  texto = texto.replace(/^\s+|\s+$/g, ""); // Remove espaços no início e no final
  texto = texto.replace(/\n+/g, "\n"); // Remove múltiplas quebras de linha

  // 9. Remove comentários HTML
  texto = texto.replace(/<!--[\s\S]*?-->/g, "");

  // 10. Remove links e URLs
  texto = texto.replace(/https?:\/\/[^\s]+/g, "[LINK REMOVIDO]"); // Substitui URLs por uma mensagem
  texto = texto.replace(/www\.[^\s]+/g, "[LINK REMOVIDO]"); // Substitui URLs sem 'http'

  // 11. Limpeza de rodapés de e-mails (remover "Enviado de", "Sent from my" etc.)
  texto = texto.replace(/(\n|^)(Sent from my|Enviado de meu|Envoyé depuis mon|Este e-mail foi enviado por)[^\n]*/g, "");

  // 12. Remove o nome e endereço do remetente, se houver
  texto = texto.replace(/Enviado para: [^\n]+/g, "");
  texto = texto.replace(/Cancelar inscrio[^\n]*/g, "");
  texto = texto.replace(/TikTok Pte. Ltd. [^\n]*/g, ""); // Remove informações de empresas e links de rodapé de TikTok

  // 13. Remove qualquer outro conteúdo que não seja texto (como SVG, fontes, etc.)
  texto = texto.replace(/[^\x20-\x7E]/g, ""); // Remove caracteres não ASCII

  // 14. Normaliza quebras de linha para parágrafos
  texto = texto.replace(/\n/g, " "); // Garante que o texto não tenha quebras de linha extras.

  return texto;
}

/**
 * Obtém as etiquetas (labels) de um e-mail
 */
function obterEtiquetas(mensagem) {
  try {
    var thread = mensagem.getThread(); // Obtém o thread da mensagem
    var etiquetas = thread.getLabels(); // Agora obtemos as etiquetas do thread
    
    // Se não houver etiquetas, retorna "Sem etiquetas"
    if (etiquetas.length === 0) {
      return "Sem etiquetas";
    }

    // Caso contrário, retorna os nomes das etiquetas
    var nomesEtiquetas = etiquetas.map(function(label) {
      return label.getName();
    });
    
    return nomesEtiquetas.join(", "); // Retorna as etiquetas separadas por vírgula
  } catch (error) {
    Logger.log("Erro ao obter etiquetas: " + error.toString());
    return "Erro ao obter etiquetas";
  }
}
