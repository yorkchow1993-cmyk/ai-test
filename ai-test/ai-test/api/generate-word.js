import { Document, Paragraph, TextRun, HeadingLevel, Packer, AlignmentType } from 'docx';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { content, filename } = req.body;

  try {
    const lines = content.split('\n');
    const children = [];

    lines.forEach(line => {
      const t = line.trim();
      if (!t) { children.push(new Paragraph({})); return; }
      if (t.startsWith('# ')) {
        children.push(new Paragraph({ text: t.slice(2), heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }));
      } else if (t.startsWith('## ')) {
        children.push(new Paragraph({ text: t.slice(3), heading: HeadingLevel.HEADING_2 }));
      } else if (t.startsWith('### ')) {
        children.push(new Paragraph({ text: t.slice(4), heading: HeadingLevel.HEADING_3 }));
      } else if (t.startsWith('- ') || t.startsWith('* ')) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.slice(2), size: 24, font: 'Microsoft YaHei' })],
          bullet: { level: 0 }
        }));
      } else {
        const parts = t.split(/\*\*(.+?)\*\*/);
        const runs = parts.map((p, i) =>
          i % 2 === 1
            ? new TextRun({ text: p, bold: true, size: 24, font: 'Microsoft YaHei' })
            : new TextRun({ text: p, size: 24, font: 'Microsoft YaHei' })
        );
        children.push(new Paragraph({ children: runs }));
      }
    });

    const doc = new Document({
      styles: {
        paragraphStyles: [{
          id: 'Normal', name: 'Normal',
          run: { font: 'Microsoft YaHei', size: 24 },
          paragraph: { spacing: { line: 360 } }
        }]
      },
      sections: [{ properties: { page: { margin: { top: 1440, right: 1800, bottom: 1440, left: 1800 } } }, children }]
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename || 'document.docx')}"`);
    return res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: '生成失败: ' + err.message });
  }
}
