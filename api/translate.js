export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const { texts, from = 'en', to = 'te', mode = 'transliterate' } = req.body;
    const translated = [];
    const errors = [];

    for (const text of (texts || [])) {
      if (!text?.trim()) { translated.push(text); continue; }
      try {
        // Use transliteration endpoint for names (not translation)
        const url = mode === 'transliterate'
          ? `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${from}&tl=${to}&dt=t&dt=rm&q=${encodeURIComponent(text)}`
          : `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${from}&tl=${to}&dt=t&q=${encodeURIComponent(text)}`;

        const r = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });

        if (!r.ok) {
          errors.push(`"${text}": Google Translate returned HTTP ${r.status}`);
          translated.push(text);
          continue;
        }

        let d;
        try {
          d = await r.json();
        } catch {
          const bodyText = await r.text().catch(() => "");
          errors.push(`"${text}": Google Translate returned non-JSON response (${bodyText.slice(0, 80)})`);
          translated.push(text);
          continue;
        }

        if (!Array.isArray(d) || !Array.isArray(d[0])) {
          errors.push(`"${text}": unexpected response shape from Google Translate`);
          translated.push(text);
          continue;
        }

        // For transliteration of proper nouns, the Telugu script output is in d[0]
        const result = d[0].map(i => i[0]).join('');
        translated.push(result || text);
      } catch (innerErr) {
        errors.push(`"${text}": ${innerErr.message}`);
        translated.push(text);
      }
    }

    return res.status(200).json({ translated, errors: errors.length ? errors : undefined });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
