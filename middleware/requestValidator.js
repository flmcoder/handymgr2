const BLOCKED_USER_AGENT = /go-http-client|python-requests|curl|wget|nikto|sqlmap/i;

export default function requestValidator(req, res, next) {
  const userAgent = String(req.get?.('user-agent') ?? req.headers?.['user-agent'] ?? '');

  if (BLOCKED_USER_AGENT.test(userAgent)) {
    res.status(403).end();
    return;
  }

  if (req.method === 'POST' && req.path === '/') {
    res.status(405).end();
    return;
  }

  next();
}