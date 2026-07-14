export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { texts } = req.body; // array of strings to translate
    
    const translated = await Promise.all(texts.map(async (text) => {
      if (!text || !text.trim()) return text;
      const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=te&dt=t&q=${encodeURIComponent(text)}`;
      const response = await fetch(url);
      const data = await response.json();
      return data[0]?.map(item => item[0]).join('') || text;
    }));

    return res.status(200).json({ translated });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
