/**
 * relatorio_framework.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Framework para geração de relatórios Word (.docx) via Node.js + docx library.
 *
 * Instalação:
 *   npm init -y && npm install docx
 *
 * Uso:
 *   node relatorio_framework.js   →  gera output.docx
 *
 * Primitivas disponíveis:
 *   h1(text)                       Título de seção (nível 1)
 *   h2(text)                       Subtítulo (nível 2)
 *   h3(text)                       Sub-subtítulo (nível 3)
 *   p(text, opts?)                 Parágrafo simples
 *   bullet(text, bold_prefix?)     Item de lista com marcador
 *   divider()                      Linha separadora horizontal
 *   callout(text, bg?, border?)    Caixa destacada (estilo alerta/info)
 *   hrow(cols, widths)             Linha de cabeçalho de tabela (fundo escuro)
 *   drow(cols, widths, alts?, opts?) Linha de dados de tabela (com cores por célula)
 *   cell(text, width, opts?)       Célula individual (uso avançado)
 * ─────────────────────────────────────────────────────────────────────────────
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageNumber, Header, Footer, LevelFormat, ExternalHyperlink
} = require("docx");
const fs = require("fs");

// ─── Constantes de layout ─────────────────────────────────────────────────────
const CONTENT_W = 9360; // largura útil em DXA (Letter 1" margins)
                        // Para A4 com margem 2cm: use 9638

const border   = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders  = { top: border, bottom: border, left: border, right: border };

// Paleta de cores padrão
const COLORS = {
  primary:    "2E4057",   // azul escuro — títulos, cabeçalhos
  secondary:  "3D5A80",   // azul médio — h3
  text:       "222222",   // texto corpo
  muted:      "888888",   // texto secundário, rodapés
  faint:      "AAAAAA",   // texto muito fraco

  green_bg:   "EAF3DE",   // fundo verde claro (positivo)
  red_bg:     "FCEBEB",   // fundo vermelho claro (negativo)
  yellow_bg:  "FFF3CD",   // fundo amarelo claro (alerta)

  green_bdr:  "27AE60",   // borda verde (callout positivo)
  red_bdr:    "E74C3C",   // borda vermelha (callout negativo)
  yellow_bdr: "FFC107",   // borda amarela (callout neutro/alerta)
};

// ─── Primitivas de célula e tabela ───────────────────────────────────────────

/**
 * Cria uma célula de tabela.
 * @param {string} text   Conteúdo
 * @param {number} w      Largura em DXA
 * @param {object} opts   { bold, bg, align, color }
 */
function cell(text, w, opts = {}) {
  const { bold = false, bg = null, align = AlignmentType.LEFT, color = "000000" } = opts;
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, font: "Arial", size: 20, color })]
    })]
  });
}

/**
 * Linha de cabeçalho — fundo azul escuro, texto branco.
 * @param {string[]} cols    Rótulos das colunas
 * @param {number[]} widths  Larguras em DXA (soma deve ser <= CONTENT_W)
 */
function hrow(cols, widths) {
  return new TableRow({
    tableHeader: true,
    children: cols.map((c, i) => cell(c, widths[i], { bold: true, bg: COLORS.primary, color: "FFFFFF" }))
  });
}

/**
 * Linha de dados com coloração opcional por célula.
 * @param {string[]} cols    Valores
 * @param {number[]} widths  Larguras em DXA
 * @param {string[]} alts    (não usado — reservado para compatibilidade)
 * @param {object[]} opts    Array de opts por célula: [{ bold?, bg?, align?, color? }, ...]
 *
 * Exemplo:
 *   drow(["Colômbia","1,6%","−0,13"], [2500,2000,1800], [], [{}, {}, {bg:"FCEBEB"}])
 */
function drow(cols, widths, alts = [], opts = []) {
  return new TableRow({
    children: cols.map((c, i) => cell(c, widths[i], opts[i] || {}))
  });
}

// ─── Primitivas de parágrafo ─────────────────────────────────────────────────

/** Título nível 1 */
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 32, color: COLORS.primary })]
  });
}

/** Título nível 2 */
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 26, color: COLORS.primary })]
  });
}

/** Título nível 3 */
function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 22, color: COLORS.secondary })]
  });
}

/**
 * Parágrafo de texto corrido.
 * @param {object} opts  { bold?, italic?, color?, size? }
 */
function p(text, opts = {}) {
  const { bold = false, italic = false, color = COLORS.text, size = 20 } = opts;
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, bold, italic, font: "Arial", size, color })]
  });
}

/**
 * Item de lista com marcador.
 * @param {string} text         Texto do item
 * @param {string} bold_prefix  Prefixo em negrito (ex: "Atenção: ")
 */
function bullet(text, bold_prefix = "") {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [
      bold_prefix ? new TextRun({ text: bold_prefix, bold: true, font: "Arial", size: 20 }) : null,
      new TextRun({ text, font: "Arial", size: 20 })
    ].filter(Boolean)
  });
}

/** Linha divisória horizontal */
function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
    children: [new TextRun("")]
  });
}

/**
 * Caixa de destaque (callout) — borda esquerda grossa colorida.
 * @param {string} text         Conteúdo da caixa
 * @param {string} color_bg     Cor de fundo hex (ex: "FFF3CD")
 * @param {string} color_border Cor da borda hex (ex: "FFC107")
 *
 * Presets:
 *   callout(texto)                          → amarelo (alerta)
 *   callout(texto, COLORS.green_bg, COLORS.green_bdr)  → verde (positivo)
 *   callout(texto, COLORS.red_bg,   COLORS.red_bdr)    → vermelho (negativo)
 */
function callout(text, color_bg = COLORS.yellow_bg, color_border = COLORS.yellow_bdr) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 6,  color: color_border },
        bottom: { style: BorderStyle.SINGLE, size: 6,  color: color_border },
        left:   { style: BorderStyle.THICK,  size: 12, color: color_border },
        right:  { style: BorderStyle.SINGLE, size: 6,  color: color_border },
      },
      shading: { fill: color_bg, type: ShadingType.CLEAR },
      width: { size: CONTENT_W, type: WidthType.DXA },
      margins: { top: 120, bottom: 120, left: 200, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20 })] })]
    })]})],
  });
}

// ─── Documento ───────────────────────────────────────────────────────────────

const TITULO       = "TÍTULO DO RELATÓRIO";
const SUBTITULO    = "Subtítulo ou Descrição";
const DATA_STR     = "18 de junho de 2026";
const HEADER_TEXT  = "Relatório  |  " + DATA_STR;

const doc = new Document({
  // Configuração de lista com marcadores
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  // Estilos globais
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: COLORS.primary },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: COLORS.primary },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      // Página Letter (12240 × 15840), margens 1" (1440 DXA)
      // Para A4 use: width: 11906, height: 16838
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    // Cabeçalho
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: HEADER_TEXT, font: "Arial", size: 16, color: COLORS.muted })]
      })] })
    },
    // Rodapé com numeração
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Página ", font: "Arial", size: 16, color: COLORS.muted }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: COLORS.muted }),
          new TextRun({ text: " — Confidencial", font: "Arial", size: 16, color: COLORS.muted }),
        ]
      })] })
    },
    children: [

      // ── CAPA ──────────────────────────────────────────────────────────────
      new Paragraph({ spacing: { before: 2400, after: 80 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: TITULO, bold: true, font: "Arial", size: 52, color: COLORS.primary })] }),
      new Paragraph({ spacing: { before: 0, after: 80 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: SUBTITULO, font: "Arial", size: 32, color: COLORS.secondary })] }),
      new Paragraph({ spacing: { before: 0, after: 1600 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: DATA_STR, font: "Arial", size: 24, color: COLORS.muted, italics: true })] }),
      divider(),
      new Paragraph({ spacing: { before: 400, after: 0 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Gerado automaticamente", font: "Arial", size: 18, color: COLORS.faint, italics: true })] }),
      new Paragraph({ children: [new TextRun({ text: "" })], pageBreakBefore: true }),

      // ── 1. SEÇÃO EXEMPLO ──────────────────────────────────────────────────
      h1("1. Título da Seção"),
      p("Texto de exemplo. Use a função p() para parágrafos comuns."),

      h2("1.1 Subseção"),
      p("Texto com destaque:", { bold: true }),
      p("Texto normal após destaque."),

      h3("1.1.1 Sub-subseção"),
      p("Exemplo de callout neutro (amarelo):"),
      callout("Atenção: este é um callout padrão (alerta)."),

      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      p("Exemplo de callout positivo (verde):"),
      callout("Resultado positivo confirmado.", COLORS.green_bg, COLORS.green_bdr),

      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      p("Exemplo de callout negativo (vermelho):"),
      callout("Atenção: impacto negativo identificado.", COLORS.red_bg, COLORS.red_bdr),

      divider(),
      new Paragraph({ children: [new TextRun({ text: "" })], pageBreakBefore: true }),

      // ── 2. TABELA EXEMPLO ─────────────────────────────────────────────────
      h1("2. Tabelas"),
      p("Exemplo de tabela com cabeçalho e células coloridas:"),

      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [3000, 2000, 2000, 2360],
        rows: [
          hrow(["Item", "Valor A", "Valor B", "Observação"], [3000, 2000, 2000, 2360]),
          drow(["Linha 1", "10,5%", "12,0%", "Positivo"], [3000, 2000, 2000, 2360], [],
            [{}, { bg: COLORS.green_bg }, { bg: COLORS.green_bg }, {}]),
          drow(["Linha 2", "8,0%",  "5,5%",  "Negativo"], [3000, 2000, 2000, 2360], [],
            [{}, { bg: COLORS.red_bg  }, { bg: COLORS.red_bg  }, {}]),
          drow(["Linha 3", "3,5%",  "3,5%",  "Neutro"],   [3000, 2000, 2000, 2360]),
          drow(["TOTAL",   "",      "",       ""],         [3000, 2000, 2000, 2360], [],
            [{ bold: true }, {}, {}, {}]),
        ]
      }),
      new Paragraph({ spacing: { before: 60, after: 0 }, children: [
        new TextRun({ text: "Fonte: exemplo de rodapé de tabela.", font: "Arial", size: 16, italics: true, color: COLORS.muted })
      ]}),

      divider(),
      new Paragraph({ children: [new TextRun({ text: "" })], pageBreakBefore: true }),

      // ── 3. LISTAS E ALERTAS ───────────────────────────────────────────────
      h1("3. Listas e Alertas"),
      bullet("Primeiro item da lista."),
      bullet("Segundo item com prefixo em negrito.", "Atenção: "),
      bullet("Terceiro item normal."),

      new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),
      p("Nota de rodapé de seção:", { italic: true, color: COLORS.muted, size: 18 }),

    ]
  }]
});

// ─── Geração do arquivo ───────────────────────────────────────────────────────
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("output.docx", buffer);
  console.log("OK — output.docx gerado.");
});
