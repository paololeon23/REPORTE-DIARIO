/** Excel SpreadsheetML — Reporte diario de cosecha */
window.APP = window.APP || {};

APP.Excel = (() => {
  const NAVY = "#0B3A66";
  const GREEN = "#1F6B3A";
  const TEAL = "#1A8FA8";
  const PEACH = "#F6D0B0";
  const BROWN = "#C47848";
  const LGREEN = "#C8E6A0";
  const MGREEN = "#9FCF78";
  const LBLUE = "#B7D4F5";
  const GRAY = "#D9D9D9";
  const WHITE = "#FFFFFF";
  const RED = "#C00000";
  const COLS = 15;

  const BORDER = `<Borders>
  <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#8FA3B5"/>
  <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#8FA3B5"/>
  <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#8FA3B5"/>
  <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#8FA3B5"/>
</Borders>`;

  function xmlEsc(s) {
    return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function cell(opts) {
    const style = opts.style || "CBase";
    const merge = opts.merge ? ` ss:MergeAcross="${opts.merge}"` : "";
    const index = opts.index ? ` ss:Index="${opts.index}"` : "";
    if (opts.num === true) {
      const v = Number(opts.value);
      const n = Number.isFinite(v) ? v : 0;
      return `<Cell${index} ss:StyleID="${style}"${merge}><Data ss:Type="Number">${n}</Data></Cell>`;
    }
    if (opts.blank) return `<Cell${index} ss:StyleID="${style}"${merge}/>`;
    return `<Cell${index} ss:StyleID="${style}"${merge}><Data ss:Type="String">${xmlEsc(opts.value ?? "")}</Data></Cell>`;
  }

  function buildModel(opts) {
    const session = opts.session || {};
    const fecha = opts.fecha || "";
    const records = (opts.records || [])
      .filter((r) => r && String(r.lote || "").trim())
      .slice()
      .sort((a, b) => String(a.lote).localeCompare(String(b.lote), "es", { numeric: true }));

    const rows = records.map((rec) => {
      const L = APP.Data.findLote(rec.lote) || {};
      const d = APP.Data.derive(rec);
      return {
        fundo: rec.fundo || L.fundo || "",
        variedad: L.variedad || rec.variedad || "",
        md: L.md || rec.md || "",
        lote: rec.lote,
        turno: L.turno || rec.turno || "",
        area: Number(rec.avance || rec.area || L.area) || 0,
        jarrasConv: d.jarrasConv,
        kgConv: d.kgConvEff,
        jarrasChina: d.jarrasChina,
        kgChina: d.kgChinaEff,
        totalJarras: d.totalJarras,
        totalKg: d.totalKg,
        filled: true,
      };
    });

    const sumArea = rows.reduce((a, r) => a + (Number(r.area) || 0), 0);
    const sumJConv = rows.reduce((a, r) => a + (Number(r.jarrasConv) || 0), 0);
    const sumKgConv = rows.reduce((a, r) => a + (Number(r.kgConv) || 0), 0);
    const sumJChina = rows.reduce((a, r) => a + (Number(r.jarrasChina) || 0), 0);
    const sumKgChina = rows.reduce((a, r) => a + (Number(r.kgChina) || 0), 0);
    const sumTJarras = rows.reduce((a, r) => a + (Number(r.totalJarras) || 0), 0);
    const sumTKg = APP.Data.round2(rows.reduce((a, r) => a + (Number(r.totalKg) || 0), 0));
    const jornales = Number(session.jornales || 0) || 0;
    return {
      fecha,
      semana: APP.Data.isoWeek(fecha),
      year: APP.CONFIG.YEAR,
      supervisor: APP.Data.fullName(session.supervisorDni, session.supervisorNombre),
      scanner: APP.Data.fullName(session.scannerDni, session.scannerNombre),
      grupo: String(session.grupo || records[0]?.grupo || ""),
      etapa: String(session.etapa || records[0]?.etapa || ""),
      turnoCampo: session.turnoCampo === "Tarde" ? "Tarde" : "Mañana",
      jornales,
      rows,
      totals: {
        area: APP.Data.round2(sumArea),
        jarrasConv: sumJConv,
        kgConv: APP.Data.round2(sumKgConv),
        jarrasChina: sumJChina,
        kgChina: APP.Data.round2(sumKgChina),
        totalJarras: sumTJarras,
        totalKg: sumTKg,
        jornales,
        kgHa: sumArea > 0 ? Math.round(sumTKg / sumArea) : 0,
        kgJn: jornales > 0 ? Math.round(sumTKg / jornales) : 0,
      },
      filledCount: rows.length,
    };
  }

  function style(id, extra) {
    return `<Style ss:ID="${id}">${extra}${BORDER}</Style>`;
  }

  function stylesXml() {
    const align = `<Alignment ss:Vertical="Center" ss:Horizontal="Center" ss:WrapText="1"/>`;
    return `<Styles>
  <Style ss:ID="Default">${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="${WHITE}" ss:Pattern="Solid"/>${BORDER}</Style>
  ${style("Title", `${align}<Font ss:FontName="Calibri" ss:Size="18" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${NAVY}" ss:Pattern="Solid"/>`)}
  ${style("MetaLbl", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${GREEN}" ss:Pattern="Solid"/>`)}
  ${style("MetaVal", `${align}<Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${GREEN}" ss:Pattern="Solid"/>`)}
  ${style("Chip", `${align}<Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${TEAL}" ss:Pattern="Solid"/>`)}
  ${style("HNavy", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${NAVY}" ss:Pattern="Solid"/>`)}
  ${style("HPeach", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1"/><Interior ss:Color="${PEACH}" ss:Pattern="Solid"/>`)}
  ${style("HBrown", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${BROWN}" ss:Pattern="Solid"/>`)}
  ${style("HGreen", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1"/><Interior ss:Color="${LGREEN}" ss:Pattern="Solid"/>`)}
  ${style("HGreen2", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${MGREEN}" ss:Pattern="Solid"/>`)}
  ${style("HBlue", `${align}<Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1"/><Interior ss:Color="${LBLUE}" ss:Pattern="Solid"/>`)}
  ${style("CBase", `${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="${WHITE}" ss:Pattern="Solid"/>`)}
  ${style("CLote", `${align}<Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/><Interior ss:Color="${GRAY}" ss:Pattern="Solid"/>`)}
  ${style("CPeach", `${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="#FBE6D5" ss:Pattern="Solid"/>`)}
  ${style("CBrown", `${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="#E8C2A8" ss:Pattern="Solid"/>`)}
  ${style("CGreen", `${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="#E2F3C9" ss:Pattern="Solid"/>`)}
  ${style("CGreen2", `${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="#D4EBB3" ss:Pattern="Solid"/>`)}
  ${style("CBlue", `${align}<Font ss:FontName="Calibri" ss:Size="10"/><Interior ss:Color="#D6E8FB" ss:Pattern="Solid"/>`)}
  ${style("CNavy", `${align}<Font ss:FontName="Calibri" ss:Size="10" ss:Color="${WHITE}"/><Interior ss:Color="${NAVY}" ss:Pattern="Solid"/>`)}
  ${style("TRed", `${align}<Font ss:Bold="1" ss:Color="${RED}"/><Interior ss:Color="${PEACH}" ss:Pattern="Solid"/>`)}
  ${style("TOrg", `${align}<Font ss:Bold="1"/><Interior ss:Color="#E8C2A8" ss:Pattern="Solid"/>`)}
  ${style("TBlue", `${align}<Font ss:Bold="1"/><Interior ss:Color="${LBLUE}" ss:Pattern="Solid"/>`)}
  ${style("TNavy", `${align}<Font ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${NAVY}" ss:Pattern="Solid"/>`)}
</Styles>`;
  }

  function fmtDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ""));
    return m ? `${m[3]}/${m[2]}/${m[1]}` : iso || "";
  }

  function dataRow(r) {
    return `<Row ss:Height="18">${[
      cell({ value: r.fundo, style: "CBase" }),
      cell({ value: r.variedad, style: "CBase" }),
      cell({ value: r.md, style: "CBase" }),
      cell({ value: r.lote, style: "CLote" }),
      cell({ value: r.turno, style: "CBase" }),
      cell({ value: r.area, style: "CBase", num: true }),
      cell({ value: r.jarrasConv, style: "CPeach", num: true }),
      cell({ value: r.kgConv, style: "CBrown", num: true }),
      cell({ value: r.jarrasChina, style: "CGreen", num: true }),
      cell({ value: r.kgChina, style: "CGreen2", num: true }),
      cell({ value: r.totalJarras, style: "CBlue", num: true }),
      cell({ value: r.totalKg, style: "CBlue", num: true }),
      cell({ blank: true, style: "CNavy" }),
      cell({ blank: true, style: "CNavy" }),
      cell({ blank: true, style: "CNavy" }),
    ].join("")}</Row>`;
  }

  function buildXml(model) {
    const t = model.totals;
    const body = model.rows.map(dataRow).join("");
    return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
${stylesXml()}
<Worksheet ss:Name="REPORTE">
<Table ss:ExpandedColumnCount="${COLS}">
<Column ss:Width="72"/><Column ss:Width="78"/><Column ss:Width="40"/><Column ss:Width="50"/><Column ss:Width="50"/><Column ss:Width="50"/>
<Column ss:Width="58"/><Column ss:Width="58"/><Column ss:Width="58"/><Column ss:Width="58"/><Column ss:Width="70"/><Column ss:Width="70"/>
<Column ss:Width="70"/><Column ss:Width="58"/><Column ss:Width="58"/>
<Row ss:Height="28">${cell({ value: `REPORTE DIARIO DE COSECHA - QBERRIES ${model.year}`, style: "Title", merge: 14 })}</Row>
<Row ss:Height="20">
${cell({ value: "SUPERVISOR", style: "MetaLbl" })}
${cell({ value: model.supervisor, style: "MetaVal", merge: 1 })}
${cell({ value: "SCANNER", style: "MetaLbl", index: 4 })}
${cell({ value: model.scanner, style: "MetaVal", merge: 2 })}
${cell({ value: "N° GRUPO", style: "Chip", index: 8 })}
${cell({ value: model.grupo, style: "Chip" })}
${cell({ value: "ETAPA", style: "Chip" })}
${cell({ value: model.etapa, style: "Chip" })}
${cell({ blank: true, style: "Chip" })}
${cell({ blank: true, style: "Chip" })}
${cell({ blank: true, style: "Chip" })}
${cell({ blank: true, style: "Chip" })}
</Row>
<Row ss:Height="18">
${cell({ value: "FECHA", style: "HNavy" })}
${cell({ value: fmtDate(model.fecha), style: "HNavy" })}
${cell({ blank: true, style: "HNavy" })}
${cell({ blank: true, style: "HNavy" })}
${cell({ blank: true, style: "HNavy" })}
${cell({ blank: true, style: "HNavy" })}
${cell({ value: "CONVENCIONAL", style: "HPeach", merge: 1 })}
${cell({ value: "CHINA", style: "HGreen", index: 9, merge: 1 })}
${cell({ value: "TOTAL", style: "HBlue", index: 11, merge: 1 })}
${cell({ blank: true, style: "HNavy", index: 13 })}
${cell({ blank: true, style: "HNavy" })}
${cell({ blank: true, style: "HNavy" })}
</Row>
<Row ss:Height="20">
${cell({ value: "FUNDO", style: "HNavy" })}
${cell({ value: "VARIEDAD", style: "HNavy" })}
${cell({ value: "MD", style: "HNavy" })}
${cell({ value: "LOTE", style: "HNavy" })}
${cell({ value: "TURNO", style: "HNavy" })}
${cell({ value: "ÁREA", style: "HNavy" })}
${cell({ value: "JARRAS", style: "HPeach" })}
${cell({ value: "KG", style: "HBrown" })}
${cell({ value: "JARRAS", style: "HGreen" })}
${cell({ value: "KG", style: "HGreen2" })}
${cell({ value: "T. JARRAS", style: "HBlue" })}
${cell({ value: "TOTAL KG", style: "HBlue" })}
${cell({ value: "JORNALES", style: "HNavy" })}
${cell({ value: "KG/HA", style: "HNavy" })}
${cell({ value: "KG/JN", style: "HNavy" })}
</Row>
${body}
<Row ss:Height="20">
${cell({ value: "TOTAL", style: "TNavy" })}
${cell({ blank: true, style: "TNavy" })}
${cell({ blank: true, style: "TNavy" })}
${cell({ blank: true, style: "TNavy" })}
${cell({ blank: true, style: "TNavy" })}
${cell({ value: t.area, style: "TNavy", num: true })}
${cell({ value: t.jarrasConv, style: "TRed", num: true })}
${cell({ value: t.kgConv, style: "TOrg", num: true })}
${cell({ value: t.jarrasChina, style: "CGreen", num: true })}
${cell({ value: t.kgChina, style: "CGreen2", num: true })}
${cell({ value: t.totalJarras, style: "TBlue", num: true })}
${cell({ value: t.totalKg, style: "TBlue", num: true })}
${cell({ value: t.jornales, style: "TNavy", num: true })}
${cell({ value: t.kgHa, style: "TNavy", num: true })}
${cell({ value: t.kgJn, style: "TNavy", num: true })}
</Row>
</Table>
<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
  <FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane>
</WorksheetOptions>
</Worksheet>
</Workbook>`;
  }

  function buildFile(opts) {
    const model = buildModel(opts);
    const xml = buildXml(model);
    const name = `reporte-cosecha-${model.fecha}.xls`;
    const type = "application/vnd.ms-excel";
    const blob = new Blob([xml], { type });
    const file = typeof File === "function" ? new File([blob], name, { type }) : blob;
    return { model, blob, file, name, xml };
  }

  function download(opts) {
    const pack = buildFile(opts);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(pack.blob);
    a.download = pack.name;
    a.click();
    URL.revokeObjectURL(a.href);
    return pack.model;
  }

  function ntxt(v) {
    const n = Number(v);
    if (!Number.isFinite(n)) return "0";
    return Number.isInteger(n) ? String(n) : String(APP.Data.round2(n));
  }

  function previewUrl(model) {
    const widths = [52, 56, 28, 34, 34, 34, 40, 40, 40, 40, 48, 48, 46, 38, 38];
    const W = widths.reduce((a, b) => a + b, 0);
    const dataH = 26;
    const heights = [40, 28, 26, 28].concat(model.rows.map(() => dataH)).concat([28]);
    const H = heights.reduce((a, b) => a + b, 0);
    const dpr = 2;
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(W * dpr);
    canvas.height = Math.ceil(H * dpr);
    const ctx = canvas.getContext("2d");
    if (!ctx) return "";
    ctx.scale(dpr, dpr);
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    function xs(i) {
      let x = 0;
      for (let k = 0; k < i; k++) x += widths[k];
      return x;
    }
    function spanW(i, span) {
      let w = 0;
      for (let k = 0; k < span; k++) w += widths[i + k] || 0;
      return w;
    }
    function box(c, y, span, h, bg) {
      const x = xs(c);
      const w = spanW(c, span);
      ctx.fillStyle = bg;
      ctx.fillRect(x, y, w, h);
      ctx.strokeStyle = "#8FA3B5";
      ctx.lineWidth = 0.7;
      ctx.strokeRect(x + 0.35, y + 0.35, w - 0.7, h - 0.7);
      return { x, y, w, h };
    }
    function label(c, y, span, h, bg, fg, value, size, bold) {
      const r = box(c, y, span, h, bg);
      ctx.fillStyle = fg;
      ctx.font = `${bold ? "800" : "600"} ${size}px "Plus Jakarta Sans", Calibri, Arial, sans-serif`;
      ctx.fillText(String(value ?? ""), r.x + r.w / 2, r.y + r.h / 2, r.w - 5);
    }

    let y = 0;
    label(0, y, 15, heights[0], NAVY, WHITE, `REPORTE DIARIO DE COSECHA - QBERRIES ${model.year}`, 13, true);
    y += heights[0];
    label(0, y, 1, heights[1], GREEN, WHITE, "SUPERVISOR", 8, true);
    label(1, y, 2, heights[1], GREEN, WHITE, model.supervisor || "—", 10, true);
    label(3, y, 1, heights[1], GREEN, WHITE, "SCANNER", 8, true);
    label(4, y, 3, heights[1], GREEN, WHITE, model.scanner || "—", 10, true);
    label(7, y, 1, heights[1], TEAL, WHITE, "N° GRUPO", 8, true);
    label(8, y, 1, heights[1], TEAL, WHITE, model.grupo || "—", 11, true);
    label(9, y, 1, heights[1], TEAL, WHITE, "ETAPA", 8, true);
    label(10, y, 1, heights[1], TEAL, WHITE, model.etapa || "—", 11, true);
    label(11, y, 4, heights[1], TEAL, WHITE, "", 10, true);
    y += heights[1];
    label(0, y, 1, heights[2], NAVY, WHITE, "FECHA", 8, true);
    label(1, y, 1, heights[2], NAVY, WHITE, fmtDate(model.fecha), 10, true);
    label(2, y, 4, heights[2], NAVY, WHITE, "", 10, true);
    label(6, y, 2, heights[2], PEACH, "#1A1A1A", "CONVENCIONAL", 8, true);
    label(8, y, 2, heights[2], LGREEN, "#1A1A1A", "CHINA", 8, true);
    label(10, y, 2, heights[2], LBLUE, "#1A1A1A", "TOTAL", 8, true);
    label(12, y, 3, heights[2], NAVY, WHITE, "", 8, true);
    y += heights[2];
    const headerRow = [
      { t: "FUNDO", bg: NAVY, fg: WHITE },
      { t: "VARIEDAD", bg: NAVY, fg: WHITE },
      { t: "MD", bg: NAVY, fg: WHITE },
      { t: "LOTE", bg: NAVY, fg: WHITE },
      { t: "TURNO", bg: NAVY, fg: WHITE },
      { t: "ÁREA", bg: NAVY, fg: WHITE },
      { t: "JARRAS", bg: PEACH, fg: "#1A1A1A" },
      { t: "KG", bg: BROWN, fg: WHITE },
      { t: "JARRAS", bg: LGREEN, fg: "#1A1A1A" },
      { t: "KG", bg: MGREEN, fg: WHITE },
      { t: "T. JARRAS", bg: LBLUE, fg: "#1A1A1A" },
      { t: "TOTAL KG", bg: LBLUE, fg: "#1A1A1A" },
      { t: "JORNALES", bg: NAVY, fg: WHITE },
      { t: "KG/HA", bg: NAVY, fg: WHITE },
      { t: "KG/JN", bg: NAVY, fg: WHITE },
    ];
    headerRow.forEach((h, i) => label(i, y, 1, heights[3], h.bg, h.fg, h.t, 8, true));
    y += heights[3];

    model.rows.forEach((r) => {
      const vals = [
        [r.fundo, WHITE, "#111"],
        [r.variedad, WHITE, "#111"],
        [r.md, WHITE, "#111"],
        [r.lote, GRAY, "#111"],
        [r.turno, WHITE, "#111"],
        [ntxt(r.area), WHITE, "#111"],
        [ntxt(r.jarrasConv), "#FBE6D5", "#111"],
        [ntxt(r.kgConv), "#E8C2A8", "#111"],
        [ntxt(r.jarrasChina), "#E2F3C9", "#111"],
        [ntxt(r.kgChina), "#D4EBB3", "#111"],
        [ntxt(r.totalJarras), "#D6E8FB", "#111"],
        [ntxt(r.totalKg), "#D6E8FB", "#111"],
        ["", NAVY, WHITE],
        ["", NAVY, WHITE],
        ["", NAVY, WHITE],
      ];
      vals.forEach((v, i) => label(i, y, 1, dataH, v[1], v[2], v[0], i === 3 ? 11 : 9, i === 3));
      y += dataH;
    });

    const t = model.totals;
    const tot = [
      ["TOTAL", NAVY, WHITE],
      ["", NAVY, WHITE],
      ["", NAVY, WHITE],
      ["", NAVY, WHITE],
      ["", NAVY, WHITE],
      [ntxt(t.area), NAVY, WHITE],
      [ntxt(t.jarrasConv), PEACH, RED],
      [ntxt(t.kgConv), "#E8C2A8", "#111"],
      [ntxt(t.jarrasChina), "#E2F3C9", "#111"],
      [ntxt(t.kgChina), "#D4EBB3", "#111"],
      [ntxt(t.totalJarras), LBLUE, "#111"],
      [ntxt(t.totalKg), LBLUE, "#111"],
      [ntxt(t.jornales), NAVY, WHITE],
      [ntxt(t.kgHa), NAVY, WHITE],
      [ntxt(t.kgJn), NAVY, WHITE],
    ];
    tot.forEach((v, i) => label(i, y, 1, 22, v[1], v[2], v[0], 9, true));
    return canvas.toDataURL("image/png");
  }

  return { buildModel, buildFile, buildXml, download, previewUrl };
})();
