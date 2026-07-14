export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const { texts, from = 'en', to = 'te' } = req.body;
    const translated = [];
    for (const text of (texts||[])) {
      if (!text?.trim()) { translated.push(text); continue; }
      const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${from}&tl=${to}&dt=t&q=${encodeURIComponent(text)}`;
      const r = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });
      const d = await r.json();
      translated.push(d[0]?.map(i=>i[0]).join('') || text);
    }
    return res.status(200).json({ translated });
  } catch(e) {
    return res.status(500).json({ error: e.message });
  }
}
