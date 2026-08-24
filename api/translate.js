const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function translateOne(text, from, to, mode) {
  const url = mode === 'transliterate'
    ? `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${from}&tl=${to}&dt=t&dt=rm&q=${encodeURIComponent(text)}`
    : `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${from}&tl=${to}&dt=t&q=${encodeURIComponent(text)}`;

  const MAX_ATTEMPTS = 3;
  const BACKOFF_MS = [500, 1500, 3000];

  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    const r = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0' } });

    if (r.status === 429) {
      // Rate limited — wait and retry rather than giving up immediately
      if (attempt < MAX_ATTEMPTS - 1) { await sleep(BACKOFF_MS[attempt]); continue; }
      return { ok: false, error: `Google Translate rate-limited this request (HTTP 429) after ${MAX_ATTEMPTS} attempts` };
    }

    if (!r.ok) {
      return { ok: false, error: `Google Translate returned HTTP ${r.status}` };
    }

    let d;
    try {
      d = await r.json();
    } catch {
      const bodyText = await r.text().catch(() => "");
      return { ok: false, error: `Google Translate returned non-JSON response (${bodyText.slice(0, 80)})` };
    }

    if (!Array.isArray(d) || !Array.isArray(d[0])) {
      return { ok: false, error: `unexpected response shape from Google Translate` };
    }

    // For transliteration of proper nouns, the Telugu script output is in d[0]
    const result = d[0].map(i => i[0]).join('');
    return { ok: true, result: result || text };
  }

  return { ok: false, error: "Google Translate rate-limited this request (HTTP 429)" };
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const { texts, from = 'en', to = 'te', mode = 'transliterate' } = req.body;
    const translated = [];
    const errors = [];

    // Small stagger between requests (not run in parallel) — sending them all at once
    // is exactly what tends to trip Google's undocumented endpoint's rate limiter.
    const STAGGER_MS = 120;
    let first = true;

    for (const text of (texts || [])) {
      if (!text?.trim()) { translated.push(text); continue; }
      if (!first) await sleep(STAGGER_MS);
      first = false;

      try {
        const outcome = await translateOne(text, from, to, mode);
        if (outcome.ok) {
          translated.push(outcome.result);
        } else {
          errors.push(`"${text}": ${outcome.error}`);
          translated.push(text);
        }
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
