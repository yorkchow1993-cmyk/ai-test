import ExcelJS from 'exceljs';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { content, filename } = req.body;

  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');

    // 解析Markdown表格
    const lines = content.split('\n').filter(l => l.trim());
    const tableLines = lines.filter(l => l.includes('|'));
    let rows = [];

    tableLines.forEach(line => {
      if (line.replace(/[\s|\-:]/g, '') === '') return;
      const cells = line.split('|').map(c => c.trim()).filter(c => c !== '');
      if (cells.length > 0) rows.push(cells);
    });

    if (rows.length === 0) {
      lines.forEach(line => {
        const t = line.trim().replace(/^#+\s*/, '').replace(/\*\*/g, '').replace(/^[-*]\s*/, '');
        if (t) rows.push([t]);
      });
    }

    rows.forEach((row, rowIndex) => {
      const excelRow = sheet.addRow(row);
      if (rowIndex === 0) {
        excelRow.eachCell(cell => {
          cell.font = { bold: true, name: 'Microsoft YaHei', size: 11 };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdbeafe' } };
          cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });
      } else {
        excelRow.eachCell(cell => {
          cell.font = { name: 'Microsoft YaHei', size: 11 };
          cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
          cell.alignment = { vertical: 'middle' };
        });
      }
      excelRow.height = 22;
    });

    sheet.columns.forEach(col => { col.width = 20; });

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename || 'spreadsheet.xlsx')}"`);
    return res.send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).json({ error: '生成失败: ' + err.message });
  }
}
