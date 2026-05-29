import PptxGenJS from 'pptxgenjs';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { type, content, filename } = req.body;

  try {
    if (type === 'pptx') {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';
      pptx.theme = { headFontFace: 'Microsoft YaHei', bodyFontFace: 'Microsoft YaHei' };

      const lines = content.split('\n');
      let currentSlide = null;
      let bulletPoints = [];

      const saveSlide = () => {
        if (currentSlide && bulletPoints.length > 0) {
          currentSlide.addText(bulletPoints.map(b => ({ text: b, options: { bullet: true, fontSize: 18, color: '333333', breakLine: true } })), {
            x: 0.5, y: 1.5, w: 9, h: 4.5, fontFace: 'Microsoft YaHei', valign: 'top'
          });
        }
      };

      lines.forEach(line => {
        const t = line.trim();
        if (!t) return;
        if (t.startsWith('## ') || t.startsWith('# ')) {
          saveSlide();
          bulletPoints = [];
          currentSlide = pptx.addSlide();
          const title = t.replace(/^#+\s*/, '').replace(/第\d+页[：:]\s*/, '');
          currentSlide.addText(title, {
            x: 0.5, y: 0.3, w: 9, h: 1,
            fontSize: 28, bold: true, color: '1e40af',
            fontFace: 'Microsoft YaHei',
            align: 'left'
          });
          currentSlide.addShape('rect', { x: 0.5, y: 1.2, w: 9, h: 0.04, fill: { color: '3b82f6' } });
        } else if (t.startsWith('- ') || t.startsWith('* ')) {
          bulletPoints.push(t.slice(2));
        } else if (t.match(/^\d+\./)) {
          bulletPoints.push(t.replace(/^\d+\.\s*/, ''));
        } else if (!t.startsWith('#')) {
          bulletPoints.push(t);
        }
      });
      saveSlide();

      if (!currentSlide) {
        const slide = pptx.addSlide();
        slide.addText(content.substring(0, 100), { x: 0.5, y: 2, w: 9, h: 2, fontSize: 20, fontFace: 'Microsoft YaHei' });
      }

      const buffer = await pptx.stream();
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename || 'presentation.pptx')}"`);
      return res.send(Buffer.from(buffer));
    }

    res.status(400).json({ error: '不支持的文件类型' });
  } catch (err) {
    res.status(500).json({ error: '生成失败: ' + err.message });
  }
}
