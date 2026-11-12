const PROJECTS = {
  'talk.varis.kr': 'talk-1on1',
  'talkwith.varis.kr': 'talk-group',
  'team.varis.kr': 'team-group',
};

const endpoint = () => process.env.VARIS_SYNC_ENDPOINT;
const token = () => process.env.VARIS_SYNC_TOKEN;

function getBody(req) {
  if (!req.body) return {};
  if (typeof req.body === 'string') {
    try { return JSON.parse(req.body); } catch (error) { return {}; }
  }
  return req.body;
}

function validate(source, projectKey) {
  return Boolean(source && PROJECTS[source] && PROJECTS[source] === projectKey);
}

module.exports = async function handler(req, res) {
  const upstreamEndpoint = endpoint();
  const upstreamToken = token();
  if (!upstreamEndpoint || !upstreamToken) {
    return res.status(500).json({ ok: false, error: 'sync_proxy_not_configured' });
  }

  const body = getBody(req);
  const source = String(req.method === 'GET' ? req.query.source || '' : body.source || '');
  const projectKey = String(req.method === 'GET' ? req.query.projectKey || PROJECTS[source] || '' : body.projectKey || PROJECTS[source] || '');
  if (!validate(source, projectKey)) {
    return res.status(400).json({ ok: false, error: 'invalid_project' });
  }

  try {
    const target = new URL(upstreamEndpoint);
    target.searchParams.set('token', upstreamToken);
    if (req.method === 'GET') {
      target.searchParams.set('action', String(req.query.action || 'history'));
      target.searchParams.set('source', source);
      target.searchParams.set('projectKey', projectKey);
      if (req.query.t) target.searchParams.set('t', String(req.query.t));
    }

    const upstream = await fetch(target, {
      method: req.method === 'GET' ? 'GET' : 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: req.method === 'GET' ? undefined : JSON.stringify({ ...body, token: upstreamToken }),
    });
    const text = await upstream.text();
    let payload;
    try { payload = JSON.parse(text); } catch (error) { payload = { ok: false, error: 'invalid_upstream_response' }; }
    return res.status(upstream.status).json(payload);
  } catch (error) {
    return res.status(502).json({ ok: false, error: 'sync_upstream_unreachable' });
  }
};
