export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { messages, systemPrompt } = req.body;
    const apiKey = process.env.CLAUDE_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'API Key 未配置' });

    const response = await fetch('https://api.openai-proxy.org/anthropic/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 4000,
        system: systemPrompt || '你是一个专业高效的AI助理，用中文回复，回答简洁清晰，可直接使用。',
        messages: messages
      })
    });

    const data = await response.json();
    if (!response.ok) return res.status(response.status).json({ error: data.error?.message || '请求失败' });
    res.status(200).json({ content: data.content[0].text });
  } catch (err) {
    res.status(500).json({ error: '服务器错误: ' + err.message });
  }
}
