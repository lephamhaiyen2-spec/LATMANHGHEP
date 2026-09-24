const TTL = 86400;

function cors(extra = {}) {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store",
    ...extra
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...cors() }
  });
}

function randomCode(len = 8) {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
  const bytes = crypto.getRandomValues(new Uint8Array(len));
  return Array.from(bytes, b => chars[b % chars.length]).join("");
}

async function sha256(text) {
  const data = new TextEncoder().encode(text);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(hash), b => b.toString(16).padStart(2, "0")).join("");
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") return new Response(null, { headers: cors() });

    const url = new URL(request.url);

    try {
      // Create a short assignment link.
      if (url.pathname === "/api/share" && request.method === "POST") {
        const body = await request.json();

        if (!body || !body.img || !Array.isArray(body.qs) || ![3,4,5,6].includes(Number(body.grid))) {
          return json({ error: "INVALID_PAYLOAD" }, 400);
        }

        if (body.qs.length !== Number(body.grid) * Number(body.grid)) {
          return json({ error: "INVALID_QUESTION_COUNT" }, 400);
        }

        // The teacher token is only used to view the teacher's 24-hour history.
        const teacherToken = randomCode(20);
        const code = randomCode(8);
        const now = Date.now();
        const expiresAt = now + TTL * 1000;

        const assignment = {
          v: 4,
          title: body.title || "Lật mảnh ghép",
          grid: Number(body.grid),
          time: Number(body.time) || 300,
          answerMode: body.answerMode === "show" ? "show" : "hide",
          order: body.order === "shuffle" ? "shuffle" : "sequential",
          qs: body.qs,
          img: body.img,
          code,
          createdAt: now,
          expiresAt
        };

        await env.SHARES.put("s:" + code, JSON.stringify(assignment), { expirationTtl: TTL });
        await env.SHARES.put(
          "a:" + teacherToken + ":" + now + ":" + code,
          JSON.stringify({
            code,
            title: assignment.title,
            grid: assignment.grid,
            total: assignment.qs.length,
            createdAt: now,
            expiresAt
          }),
          { expirationTtl: TTL }
        );

        return json({
          code,
          teacherToken,
          expiresAt,
          studentPath: "?s=" + encodeURIComponent(code),
          teacherPath: "?teacher=" + encodeURIComponent(teacherToken)
        });
      }

      // Load an assignment by its short code.
      if (url.pathname === "/api/share" && request.method === "GET") {
        const code = (url.searchParams.get("code") || "").trim();
        if (!/^[A-Za-z0-9]{8}$/.test(code)) return json({ error: "INVALID_CODE" }, 400);

        const value = await env.SHARES.get("s:" + code, { type: "json" });
        if (!value) return json({ error: "NOT_FOUND" }, 404);

        return json(value);
      }

      // Student submits one completed attempt.
      if (url.pathname === "/api/result" && request.method === "POST") {
        const body = await request.json();
        const code = String(body.code || "");
        const name = String(body.name || "Học sinh").trim().slice(0, 80);

        if (!/^[A-Za-z0-9]{8}$/.test(code)) return json({ error: "INVALID_CODE" }, 400);

        const assignment = await env.SHARES.get("s:" + code, { type: "json" });
        if (!assignment) return json({ error: "EXPIRED" }, 404);

        const resultId = randomCode(12);
        const record = {
          id: resultId,
          code,
          name: name || "Học sinh",
          correct: Math.max(0, Number(body.correct) || 0),
          total: assignment.qs.length,
          opened: Math.max(0, Number(body.opened) || 0),
          score: Math.max(0, Number(body.score) || 0),
          timeLeft: Math.max(0, Number(body.timeLeft) || 0),
          completedAt: Date.now()
        };

        await env.SHARES.put(
          "r:" + code + ":" + resultId,
          JSON.stringify(record),
          { expirationTtl: TTL }
        );

        return json({ ok: true, id: resultId, expiresAt: assignment.expiresAt });
      }

      // Teacher history for the last 24 hours.
      if (url.pathname === "/api/history" && request.method === "GET") {
        const token = (url.searchParams.get("token") || "").trim();
        if (!/^[A-Za-z0-9]{20}$/.test(token)) return json({ error: "INVALID_TOKEN" }, 400);

        const assignments = [];
        let cursor;
        do {
          const page = await env.SHARES.list({ prefix: "a:" + token + ":", cursor, limit: 1000 });
          assignments.push(...page.keys.map(k => ({ key: k.name })));
          cursor = page.list_complete ? undefined : page.cursor;
        } while (cursor);

        const details = [];
        for (const item of assignments) {
          const a = await env.SHARES.get(item.key, { type: "json" });
          if (a) details.push(a);
        }

        details.sort((a,b) => b.createdAt - a.createdAt);

        const out = [];
        for (const a of details) {
          const results = [];
          let cursor;
          do {
            const page = await env.SHARES.list({ prefix: "r:" + a.code + ":", cursor, limit: 1000 });
            for (const k of page.keys) {
              const r = await env.SHARES.get(k.name, { type: "json" });
              if (r) results.push(r);
            }
            cursor = page.list_complete ? undefined : page.cursor;
          } while (cursor);

          results.sort((x,y) => y.completedAt - x.completedAt);
          out.push({ ...a, results });
        }

        return json({ generatedAt: Date.now(), assignments: out });
      }

      // Small health endpoint for deployment checks.
      if (url.pathname === "/api/health" && request.method === "GET") {
        return json({ ok: true, service: "LATMANHGHEP", ttlHours: 24 });
      }

      return json({ error: "NOT_FOUND" }, 404);
    } catch (error) {
      return json({ error: "SERVER_ERROR", message: String(error?.message || error) }, 500);
    }
  }
};
