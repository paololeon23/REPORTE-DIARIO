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

  function xmlEsc(s) {
    return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function numCell(n, style) {
    const v = Number(n);
    if (!Number.isFinite(v)) return `<Cell ss:StyleID="${style}"><Data ss:Type="Number">0</Data></Cell>`;
    return `<Cell ss:StyleID="${style}"><Data ss:Type="Number">${v}</Data></Cell>`;
  }
  function textCell(s, style, merge) {
    const m = merge ? ` ss:MergeAcross="${merge}"` : "";
    return `<Cell ss:StyleID="${style}"${m}><Data ss:Type="String">${xmlEsc(s)}</Data></Cell>`;
  }
  function emptyCell(style) {
    return `<Cell ss:StyleID="${style}"/>`;
  }

  function buildModel(opts) {
    const session = opts.session || {};
    const fecha = opts.fecha || "";
    const records = opts.records || [];
    const byLote = {};
    records.forEach((r) => {
      byLote[String(r.lote)] = r;
    });
    const grupo = String(session.grupo || records[0]?.grupo || "2");
    const assigned = APP.Data.lotes.filter((l) => String(l.grupo) === grupo);
    const source = assigned.slice();
    while (source.length < 10) source.push({ fundo: assigned[0]?.fundo || "LICAPA", variedad: assigned[0]?.variedad || "SEKOYA", md: assigned[0]?.md || "10" });

    const rows = source.map((lot) => {
      const rec = lot.lote ? byLote[String(lot.lote)] : null;
      const filled = !!rec;
      const d = rec ? APP.Data.derive(rec) : null;
      return {
        fundo: rec?.fundo || lot.fundo || "",
        variedad: rec?.variedad || lot.variedad || "",
        md: rec?.md || lot.md || "",
        lote: rec?.lote || lot.lote || "",
        turno: rec?.turno || (filled ? lot.turno : "") || "",
        area: filled ? Number(rec.avance || rec.area) || Number(lot.area) || 0 : lot.lote ? Number(lot.area) || 0 : "",
        jarrasConv: filled ? d.jarrasConv : "",
        kgConv: filled ? d.kgConv : "",
        jarrasChina: filled ? d.jarrasChina : "",
        kgChina: filled ? d.kgChina : "",
        totalJarras: filled ? d.totalJarras : "",
        totalKg: filled ? d.totalKg : "",
        filled,
      };
    });

    const filledRows = rows.filter((r) => r.filled);
    const sumArea = filledRows.reduce((a, r) => a + (Number(r.area) || 0), 0);
    const sumJConv = filledRows.reduce((a, r) => a + (Number(r.jarrasConv) || 0), 0);
    const sumKgConv = filledRows.reduce((a, r) => a + (Number(r.kgConv) || 0), 0);
    const sumJChina = filledRows.reduce((a, r) => a + (Number(r.jarrasChina) || 0), 0);
    const sumKgChina = filledRows.reduce((a, r) => a + (Number(r.kgChina) || 0), 0);
    const sumTJarras = filledRows.reduce((a, r) => a + (Number(r.totalJarras) || 0), 0);
    const sumTKg = APP.Data.round2(filledRows.reduce((a, r) => a + (Number(r.totalKg) || 0), 0));
    const jornales = Number(session.jornales || 0) || 0;
    return {
      fecha,
      semana: APP.Data.isoWeek(fecha),
      year: APP.CONFIG.YEAR,
      supervisor: APP.Data.displayName(session.supervisorNombre),
      scanner: APP.Data.displayName(session.scannerNombre),
      grupo,
      etapa: String(session.etapa || "1"),
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
      filledCount: filledRows.length,
    };
  }

  function stylesXml() {
    return `<Styles>
  <Style ss:ID="Default"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="10"/></Style>
  <Style ss:ID="Title"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="18" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${NAVY}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="MetaLbl"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${GREEN}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="MetaVal"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${GREEN}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="Chip"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${TEAL}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="HNavy"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${NAVY}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="HPeach"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1"/><Interior ss:Color="${PEACH}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="HBrown"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${BROWN}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="HGreen"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1"/><Interior ss:Color="${LGREEN}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="HGreen2"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1" ss:Color="${WHITE}"/><Interior ss:Color="${MGREEN}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="HBlue"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="9" ss:Bold="1"/><Interior ss:Color="${LBLUE}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="CBase"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="10"/></Style>
  <Style ss:ID="CLote"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/><Interior ss:Color="${GRAY}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="CPeach"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Interior ss:Color="#FBE6D5" ss:Pattern="Solid"/></Style>
  <Style ss:ID="CBrown"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Interior ss:Color="#E8C2A8" ss:Pattern="Solid"/></Style>
  <Style ss:ID="CGreen"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Interior ss:Color="#E2F3C9" ss:Pattern="Solid"/></Style>
  <Style ss:ID="CBlue"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Interior ss:Color="#D6E8FB" ss:Pattern="Solid"/></Style>
  <Style ss:ID="TRed"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:Bold="1" ss:Color="${RED}"/><Interior ss:Color="${PEACH}" ss:Pattern="Solid"/></Style>
  <Style ss:ID="TOrg"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:Bold="1"/><Interior ss:Color="#E8C2A8" ss:Pattern="Solid"/></Style>
  <Style ss:ID="TBlue"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:Bold="1"/><Interior ss:Color="${LBLUE}" ss:Pattern="Solid"/></Style>
</Styles>`;
  }

  function fmtDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ""));
    return m ? `${m[3]}/${m[2]}/${m[1]}` : iso || "";
  }

  function buildXml(model) {
    const t = model.totals;
    const body = model.rows
      .map((r) => {
        const cells = [
          textCell(r.fundo, "CBase"),
          textCell(r.variedad, "CBase"),
          textCell(r.md, "CBase"),
          r.lote ? textCell(r.lote, "CLote") : emptyCell("CLote"),
          textCell(r.turno, "CBase"),
          r.area === "" ? emptyCell("CBase") : numCell(r.area, "CBase"),
          r.jarrasConv === "" ? emptyCell("CPeach") : numCell(r.jarrasConv, "CPeach"),
          r.kgConv === "" ? emptyCell("CBrown") : numCell(r.kgConv, "CBrown"),
          r.jarrasChina === "" ? emptyCell("CGreen") : numCell(r.jarrasChina, "CGreen"),
          r.kgChina === "" ? emptyCell("CGreen") : numCell(r.kgChina, "CGreen"),
          r.totalJarras === "" ? emptyCell("CBlue") : numCell(r.totalJarras, "CBlue"),
          r.totalKg === "" ? emptyCell("CBlue") : numCell(r.totalKg, "CBlue"),
        ];
        return `<Row>${cells.join("")}</Row>`;
      })
      .join("");
    return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
${stylesXml()}
<Worksheet ss:Name="REPORTE">
<Table>
<Column ss:Width="70"/><Column ss:Width="70"/><Column ss:Width="40"/><Column ss:Width="50"/><Column ss:Width="50"/><Column ss:Width="50"/>
<Column ss:Width="55"/><Column ss:Width="55"/><Column ss:Width="55"/><Column ss:Width="55"/><Column ss:Width="70"/><Column ss:Width="70"/>
<Column ss:Width="70"/><Column ss:Width="60"/><Column ss:Width="60"/>
<Row ss:Height="28"><Cell ss:MergeAcross="14" ss:StyleID="Title"><Data ss:Type="String">REPORTE DIARIO DE COSECHA - QBERRIES ${model.year}</Data></Cell></Row>
<Row>
${textCell("SUPERVISOR", "MetaLbl")}${textCell(model.supervisor, "MetaVal")}${emptyCell("MetaVal")}
${textCell("SCANNER", "MetaLbl")}${textCell(model.scanner, "MetaVal")}${emptyCell("MetaVal")}${emptyCell("MetaVal")}
${textCell("N° GRUPO", "Chip")}${textCell(model.grupo, "Chip")}
${textCell("ETAPA", "Chip")}${textCell(model.etapa, "Chip")}
</Row>
<Row>
${textCell("FECHA", "HNavy")}${textCell(fmtDate(model.fecha), "HNavy")}
${textCell("CONVENCIONAL", "HPeach")}${emptyCell("HPeach")}
${textCell("CHINA", "HGreen")}${emptyCell("HGreen")}
${textCell("TOTAL", "HBlue")}${emptyCell("HBlue")}${emptyCell("HBlue")}${emptyCell("HBlue")}
</Row>
<Row>
${textCell("FUNDO", "HNavy")}${textCell("VARIEDAD", "HNavy")}${textCell("MD", "HNavy")}${textCell("LOTE", "HNavy")}${textCell("TURNO", "HNavy")}${textCell("ÁREA", "HNavy")}
${textCell("JARRAS", "HPeach")}${textCell("KG", "HBrown")}
${textCell("JARRAS", "HGreen")}${textCell("KG", "HGreen2")}
${textCell("T. JARRAS", "HBlue")}${textCell("TOTAL KG", "HBlue")}
${textCell("JORNALES", "HNavy")}${textCell("KG/HA", "HNavy")}${textCell("KG/JN", "HNavy")}
</Row>
${body}
<Row>
${textCell("TOTAL", "CBase")}${emptyCell("CBase")}${emptyCell("CBase")}${emptyCell("CLote")}${emptyCell("CBase")}
${numCell(t.area, "CBase")}
${numCell(t.jarrasConv, "TRed")}${numCell(t.kgConv, "TOrg")}
${numCell(t.jarrasChina, "CGreen")}${numCell(t.kgChina, "CGreen")}
${numCell(t.totalJarras, "TBlue")}${numCell(t.totalKg, "TBlue")}
${numCell(t.jornales, "CBase")}${numCell(t.kgHa, "CBase")}${numCell(t.kgJn, "CBase")}
</Row>
</Table>
</Worksheet>
</Workbook>`;
  }

  function download(opts) {
    const model = buildModel(opts);
    const xml = buildXml(model);
    const blob = new Blob([xml], { type: "application/vnd.ms-excel" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `reporte-cosecha-${model.fecha}.xls`;
    a.click();
    URL.revokeObjectURL(a.href);
    return model;
  }

  return { buildModel, download };
})();
