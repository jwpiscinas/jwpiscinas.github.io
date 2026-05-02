const fs = require("fs");
const path = require("path");

const outDir = path.join(process.cwd(), "documentos");
const outFile = path.join(outDir, "Proposta_Envio_Automatico_WhatsApp_JW_Piscinas.pdf");

const W = 595.28;
const H = 841.89;
const margin = 56;
const contentW = W - margin * 2;

function esc(text) {
  return String(text || "")
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)");
}

function clean(text) {
  return String(text || "")
    .replace(/\u2013|\u2014/g, "-")
    .replace(/\u201c|\u201d/g, '"')
    .replace(/\u2018|\u2019/g, "'")
    .replace(/\u2022/g, "-");
}

function pdfText(text) {
  return Buffer.from(esc(clean(text)), "latin1").toString("binary");
}

function wrapText(text, maxChars) {
  const words = clean(text).split(/\s+/);
  const lines = [];
  let line = "";
  for (const word of words) {
    if (!word) continue;
    const next = line ? line + " " + word : word;
    if (next.length > maxChars && line) {
      lines.push(line);
      line = word;
    } else {
      line = next;
    }
  }
  if (line) lines.push(line);
  return lines;
}

function rect(x, y, w, h, color) {
  return `q ${color} rg ${x.toFixed(2)} ${y.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f Q\n`;
}

function strokeRect(x, y, w, h, color, lineWidth = 1) {
  return `q ${color} RG ${lineWidth} w ${x.toFixed(2)} ${y.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re S Q\n`;
}

function textLine(text, x, y, size = 11, font = "F1", color = "0.10 0.14 0.22") {
  return `BT ${color} rg /${font} ${size} Tf 1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm (${pdfText(text)}) Tj ET\n`;
}

function addParagraph(page, text, x, y, opts = {}) {
  const size = opts.size || 11;
  const leading = opts.leading || size + 5;
  const chars = opts.chars || Math.floor((opts.width || contentW) / (size * 0.47));
  const font = opts.font || "F1";
  const color = opts.color || "0.20 0.25 0.33";
  const lines = wrapText(text, chars);
  for (const line of lines) {
    page.push(textLine(line, x, y, size, font, color));
    y -= leading;
  }
  return y;
}

function bullet(page, text, x, y, opts = {}) {
  page.push(textLine("-", x, y, opts.size || 11, "F2", opts.color || "0.00 0.47 0.71"));
  return addParagraph(page, text, x + 14, y, {
    size: opts.size || 11,
    leading: opts.leading || 16,
    chars: opts.chars || 86,
    width: opts.width || contentW - 14,
    color: opts.color || "0.20 0.25 0.33"
  }) - 4;
}

function header(page, section) {
  page.push(rect(0, H - 78, W, 78, "0.00 0.47 0.71"));
  page.push(rect(0, H - 84, W, 6, "0.00 0.70 0.85"));
  page.push(textLine("JW Piscinas", margin, H - 38, 18, "F2", "1 1 1"));
  page.push(textLine(section, margin, H - 60, 10, "F1", "0.86 0.95 1"));
}

function footer(page, num) {
  page.push(textLine("Proposta comercial - envio automatico por WhatsApp", margin, 30, 8, "F1", "0.45 0.50 0.58"));
  page.push(textLine(String(num), W - margin, 30, 8, "F1", "0.45 0.50 0.58"));
}

function card(page, x, y, w, h, title, body, accent) {
  page.push(rect(x, y - h, w, h, "1 1 1"));
  page.push(strokeRect(x, y - h, w, h, "0.88 0.91 0.95", 1));
  page.push(rect(x, y - h, 5, h, accent));
  page.push(textLine(title, x + 16, y - 24, 12, "F2", "0.10 0.14 0.22"));
  addParagraph(page, body, x + 16, y - 45, { size: 9.8, chars: Math.floor((w - 32) / 4.8), width: w - 32, leading: 14 });
}

function planCard(page, y, name, price, body, accent, fill) {
  const h = 88;
  page.push(rect(margin, y - h, contentW, h, fill || "1 1 1"));
  page.push(strokeRect(margin, y - h, contentW, h, "0.86 0.89 0.94", 1));
  page.push(rect(margin, y - h, 6, h, accent));
  page.push(textLine(name, margin + 18, y - 24, 13, "F2", "0.10 0.14 0.22"));
  page.push(textLine(price, margin + contentW - 140, y - 24, 12, "F2", accent));
  addParagraph(page, body, margin + 18, y - 48, {
    size: 10,
    leading: 14,
    chars: 66,
    width: contentW - 160,
    color: "0.24 0.29 0.37"
  });
  return y - h - 14;
}

const pages = [];

{
  const p = [];
  p.push(rect(0, 0, W, H, "0.96 0.98 1"));
  p.push(rect(0, H - 260, W, 260, "0.00 0.47 0.71"));
  p.push(rect(0, H - 270, W, 10, "0.00 0.70 0.85"));
  p.push(textLine("JW Piscinas", margin, H - 78, 22, "F2", "1 1 1"));
  p.push(textLine("Proposta de melhoria para o sistema", margin, H - 108, 12, "F1", "0.86 0.95 1"));
  p.push(textLine("Envio automatico de propostas", margin, H - 162, 26, "F2", "1 1 1"));
  p.push(textLine("por WhatsApp", margin, H - 194, 26, "F2", "1 1 1"));
  p.push(textLine("Mais agilidade no atendimento, mais retorno dos clientes e menos trabalho manual.", margin, H - 224, 12, "F1", "0.90 0.97 1"));
  card(p, margin, H - 320, contentW, 118, "Ideia central", "Quando um cliente solicita um orcamento e a equipe define o valor da proposta, o sistema pode enviar automaticamente uma mensagem pelo WhatsApp para o cliente, alem do e-mail e do registro no historico.", "0.00 0.70 0.85");
  let y = H - 480;
  p.push(textLine("O que muda para o cliente", margin, y, 15, "F2"));
  y -= 28;
  y = bullet(p, "Recebe a proposta no canal em que normalmente responde mais rapido.", margin, y);
  y = bullet(p, "Consegue aprovar, tirar duvidas ou pedir ajuste sem precisar procurar e-mail.", margin, y);
  y = bullet(p, "Tem uma experiencia mais moderna, clara e proxima da empresa.", margin, y);
  y -= 12;
  p.push(textLine("O que muda para a empresa", margin, y, 15, "F2"));
  y -= 28;
  y = bullet(p, "Menos mensagens manuais repetidas para enviar propostas.", margin, y);
  y = bullet(p, "Maior chance de resposta rapida e fechamento do orcamento.", margin, y);
  y = bullet(p, "Mais controle, pois o envio e o status ficam registrados no sistema.", margin, y);
  footer(p, 1);
  pages.push(p.join(""));
}

{
  const p = [];
  header(p, "Como funcionaria no dia a dia");
  let y = H - 120;
  p.push(textLine("Fluxo proposto", margin, y, 18, "F2"));
  y -= 34;
  const steps = [
    ["1. Cliente solicita o orcamento", "Ele preenche o pedido no sistema e pode anexar uma foto ou arquivo."],
    ["2. Equipe analisa", "O administrador visualiza os detalhes da solicitacao e define o valor da proposta."],
    ["3. Sistema envia a proposta", "Ao confirmar o valor, o cliente recebe a mensagem automaticamente pelo WhatsApp."],
    ["4. Cliente responde", "O cliente pode aprovar, pedir ajuste ou tirar duvidas diretamente pelo WhatsApp."],
    ["5. Tudo fica acompanhado", "O sistema registra o envio e mantem o historico do atendimento."]
  ];
  for (const s of steps) {
    card(p, margin, y, contentW, 70, s[0], s[1], "0.00 0.47 0.71");
    y -= 86;
  }
  y -= 4;
  p.push(textLine("Exemplo de mensagem", margin, y, 15, "F2"));
  y -= 26;
  p.push(rect(margin, y - 94, contentW, 94, "0.91 0.98 0.96"));
  p.push(strokeRect(margin, y - 94, contentW, 94, "0.74 0.90 0.84", 1));
  y = addParagraph(p, "Ola, Ana. Recebemos sua solicitacao de orcamento para limpeza da piscina. A proposta ficou em R$ 189,90. Para aprovar ou tirar duvidas, responda esta mensagem.", margin + 18, y - 24, { size: 11, chars: 84, width: contentW - 36, leading: 17 });
  y -= 16;
  p.push(textLine("Observacao importante", margin, y, 13, "F2"));
  y -= 22;
  y = addParagraph(p, "Para que o envio seja seguro e profissional, o ideal e usar uma conta oficial de WhatsApp Business conectada a uma plataforma de automacao. Isso evita depender de celular aberto, WhatsApp Web ou solucoes improvisadas.", margin, y, { size: 10.5, chars: 92 });
  footer(p, 2);
  pages.push(p.join(""));
}

{
  const p = [];
  header(p, "Investimento estimado");
  let y = H - 120;
  p.push(textLine("Faixas de implantacao", margin, y, 18, "F2"));
  y -= 28;
  p.push(textLine("Os valores abaixo sao estimativas comerciais para incluir esse recurso no sistema.", margin, y, 10.5, "F1", "0.35 0.40 0.48"));
  y -= 34;
  y = planCard(p, y, "Plano Inicial", "R$ 2.500 a R$ 4.000", "Envio automatico da proposta por WhatsApp, mensagem padrao, registro do envio no sistema e testes iniciais.", "0.00 0.47 0.71", "1 1 1");
  y = planCard(p, y, "Plano Profissional", "R$ 4.000 a R$ 8.000", "Propostas, avisos de compra, atualizacoes de status, historico no sistema e ajustes das mensagens.", "0.06 0.63 0.50", "0.98 0.99 1");
  y = planCard(p, y, "Plano Completo", "R$ 8.000 a R$ 15.000", "Fluxo mais completo com mensagens por etapa, botoes de acao, painel de acompanhamento e relatorios.", "0.43 0.16 0.85", "1 1 1");
  y -= 12;
  p.push(textLine("Custos recorrentes", margin, y, 18, "F2"));
  y -= 30;
  y = bullet(p, "Plataforma de envio/WhatsApp Business: em geral, de R$ 80 a R$ 400 por mes.", margin, y, { chars: 80, size: 10.2, leading: 15 });
  y = bullet(p, "Mensagens enviadas: pode haver um custo pequeno por mensagem, variando conforme volume e plataforma.", margin, y, { chars: 80, size: 10.2, leading: 15 });
  y = bullet(p, "Manutencao e suporte: recomendavel prever de R$ 300 a R$ 900 por mes.", margin, y, { chars: 80, size: 10.2, leading: 15 });
  y = bullet(p, "Novas melhorias sob demanda: faixa comum de R$ 120 a R$ 250 por hora.", margin, y, { chars: 80, size: 10.2, leading: 15 });
  footer(p, 3);
  pages.push(p.join(""));
}

{
  const p = [];
  header(p, "Recomendacao comercial");
  let y = H - 120;
  p.push(rect(margin, y - 130, contentW, 130, "1.00 0.97 0.88"));
  p.push(strokeRect(margin, y - 130, contentW, 130, "0.95 0.82 0.44", 1));
  p.push(textLine("Melhor caminho para iniciar", margin + 18, y - 28, 16, "F2", "0.35 0.26 0.08"));
  addParagraph(p, "Para iniciar, o Plano Profissional costuma ser o melhor equilibrio. Ele ja permite enviar propostas automaticamente, acompanhar status e deixar espaco para evoluir sem refazer o trabalho depois.", margin + 18, y - 58, { size: 11, chars: 82, width: contentW - 36, color: "0.35 0.26 0.08", leading: 17 });
  y -= 170;
  p.push(textLine("Cuidados importantes", margin, y, 18, "F2"));
  y -= 34;
  y = bullet(p, "O cliente deve autorizar o recebimento de comunicacoes pelo WhatsApp no cadastro ou atendimento.", margin, y, { chars: 88 });
  y = bullet(p, "Mensagens enviadas fora de uma conversa recente precisam seguir um modelo aprovado pelo WhatsApp.", margin, y, { chars: 88 });
  y = bullet(p, "Os valores finais dependem da plataforma escolhida, quantidade de mensagens e nivel de automacao desejado.", margin, y, { chars: 88 });
  y = bullet(p, "O recurso pode ser implantado aos poucos, comecando pelas propostas de orcamento.", margin, y, { chars: 88 });
  y -= 18;
  p.push(textLine("Implantacao sugerida", margin, y, 18, "F2"));
  y -= 32;
  card(p, margin, y, contentW, 74, "Etapa 1 - Envio de propostas", "Automatizar o envio da proposta por WhatsApp quando o administrador definir o valor do orcamento.", "0.00 0.47 0.71");
  y -= 88;
  card(p, margin, y, contentW, 74, "Etapa 2 - Avisos de compra e entrega", "Enviar atualizacoes como produto em transito, entregue ou cancelado.", "0.06 0.63 0.50");
  y -= 88;
  card(p, margin, y, contentW, 74, "Etapa 3 - Acompanhamento e indicadores", "Exibir no painel quais mensagens foram enviadas, respondidas e quais propostas ainda estao pendentes.", "0.43 0.16 0.85");
  y -= 108;
  p.push(rect(margin, y - 110, contentW, 110, "0.93 0.97 1"));
  p.push(strokeRect(margin, y - 110, contentW, 110, "0.65 0.82 0.95", 1));
  p.push(textLine("Resumo para decisao", margin + 16, y - 24, 14, "F2"));
  addParagraph(p, "A automacao por WhatsApp tende a reduzir atrasos no atendimento e aumentar a taxa de retorno dos orcamentos. A melhor estrategia e comecar pelas propostas, medir o resultado e depois expandir para compras, entregas e campanhas de relacionamento.", margin + 16, y - 48, { size: 10.5, chars: 84, width: contentW - 32 });
  footer(p, 4);
  pages.push(p.join(""));
}

function makePdf(pageContents) {
  const objects = [];
  objects.push("<< /Type /Catalog /Pages 2 0 R >>");
  const kids = [];
  const fontRegularObj = 3;
  const fontBoldObj = 4;
  objects.push(null);
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");

  for (let i = 0; i < pageContents.length; i++) {
    const pageObjNum = objects.length + 1;
    const contentObjNum = objects.length + 2;
    kids.push(`${pageObjNum} 0 R`);
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${W} ${H}] /Resources << /Font << /F1 ${fontRegularObj} 0 R /F2 ${fontBoldObj} 0 R >> >> /Contents ${contentObjNum} 0 R >>`);
    const stream = pageContents[i];
    const len = Buffer.byteLength(stream, "binary");
    objects.push(`<< /Length ${len} >>\nstream\n${stream}\nendstream`);
  }

  objects[1] = `<< /Type /Pages /Kids [${kids.join(" ")}] /Count ${pageContents.length} >>`;

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let i = 0; i < objects.length; i++) {
    offsets.push(Buffer.byteLength(pdf, "binary"));
    pdf += `${i + 1} 0 obj\n${objects[i]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(pdf, "binary");
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += "0000000000 65535 f \n";
  for (let i = 1; i < offsets.length; i++) {
    pdf += String(offsets[i]).padStart(10, "0") + " 00000 n \n";
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(pdf, "binary");
}

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outFile, makePdf(pages));
console.log(outFile);
