// Serviço centralizado para envio de emails
function enviarEmailGenerico(destinatario, assunto, corpoHTML) {
  try {
    MailApp.sendEmail({
      to: destinatario,
      subject: assunto,
      htmlBody: corpoHTML
    });
    
    console.log("✅ Email enviado para:", destinatario);
    return true;
  } catch (error) {
    console.error("❌ Erro ao enviar email:", error);
    return false;
  }
}