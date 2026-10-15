// Vercel serverless proxy for Google Apps Script.
// Configure APPS_SCRIPT_URL in Vercel Project Settings > Environment Variables.
module.exports = async function handler(req, res) {
  const endpoint = process.env.APPS_SCRIPT_URL;
  if (!endpoint || !/^https:\/\/script\.google\.com\/macros\/s\//.test(endpoint)) {
    return res.status(500).json({ ok: false, message: "APPS_SCRIPT_URL is not configured in Vercel." });
  }

  try {
    if (req.method === "GET") {
      const upstream = await fetch(endpoint, { method: "GET", redirect: "follow" });
      const text = await upstream.text();
      let data;
      try { data = JSON.parse(text); }
      catch (_) { return res.status(502).json({ ok: false, message: "Apps Script did not return JSON. Check its deployment URL and access settings." }); }
      return res.status(upstream.ok ? 200 : 502).json(data);
    }

    if (req.method !== "POST") {
      res.setHeader("Allow", "GET, POST");
      return res.status(405).json({ ok: false, message: "Method not allowed." });
    }

    let payload = req.body && req.body.payload;
    if (typeof req.body === "string") {
      try {
        payload = new URLSearchParams(req.body).get("payload") || req.body;
      } catch (_) {}
    }
    if (payload && typeof payload !== "string") payload = JSON.stringify(payload);
    if (!payload) return res.status(400).json({ ok: false, message: "Missing registration payload." });

    const body = new URLSearchParams({ payload });
    const upstream = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body,
      redirect: "follow"
    });
    const text = await upstream.text();
    let data;
    try { data = JSON.parse(text); }
    catch (_) { return res.status(502).json({ ok: false, message: "Apps Script did not return JSON. Check its deployment URL and access settings." }); }
    return res.status(upstream.ok ? 200 : 502).json(data);
  } catch (err) {
    console.error("Waitlist proxy error:", err);
    return res.status(502).json({ ok: false, message: "Could not reach Google Apps Script. Verify APPS_SCRIPT_URL and the web-app deployment." });
  }
};
