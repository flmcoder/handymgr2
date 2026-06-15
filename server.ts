// ============================================================================
// server.ts — Express Backend Entry Point
// Unifies existing Deno handlers into Express routes
// ============================================================================

import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { db, testConnection, closeConnection } from './db';

// Load environment variables
dotenv.config();

// ── Import Existing Handlers ────────────────────────────────────────────────
import { handleUnits, handleUnitLookup } from './afproxy/handlers/units';
import {
  handleTurns,
  handleUnitTurns,
  handleTurnsIncremental,
  handleUnitTurnsHistory,
  handleClosedTurns,
  handleTurnRecords,
  handleTurnRecordStage,
} from './afproxy/handlers/turns';
import { handleReassignmentQueue } from './afproxy/handlers/queue';
import {
  handleDeviceOtpRequest,
  handleDeviceOtpVerify,
  handleDeviceSetup,
  handleVerifyRole,
  getTrustedDeviceSession,
  handleTrustedDeviceList,
  handleTrustedDeviceRevoke,
  handlePmProxyUserUpsert,
  handlePmProxyUserDelete,
} from './afproxy/handlers/deviceAuth';

// ── Express App Setup ───────────────────────────────────────────────────────
const app = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ──────────────────────────────────────────────────────────────
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// CORS Configuration (strict frontend whitelist)
const allowedOrigins = [
  'https://handymgr.app',
  'https://www.handymgr.app',
  'http://localhost:5173', // Vite dev server
  'http://127.0.0.1:5173',
];

app.use(cors({
  origin: (origin, callback) => {
    // Allow requests with no origin (mobile apps, Postman, etc.)
    if (!origin || allowedOrigins.includes(origin)) {
      callback(null, true);
    } else {
      console.warn('[cors] Blocked origin:', origin);
      callback(new Error('CORS policy violation'));
    }
  },
  credentials: true,
  methods: ['GET', 'POST', 'PATCH', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
}));

// ── Deno Handler Adapter ────────────────────────────────────────────────────
// Bridges Express req/res with Deno-style handler functions
type DenoHandler = (params: Record<string, string>, req?: any) => Promise<any>;

interface WrapOptions {
  allowedMethods?: string[];
  requireAuth?: boolean;
}

function wrapDenoHandler(handler: DenoHandler, options: WrapOptions = {}) {
  const { allowedMethods = ['GET', 'POST'], requireAuth = false } = options;

  return async (req: Request, res: Response) => {
    try {
      // Method validation
      if (!allowedMethods.includes(req.method)) {
        return res.status(405).json({
          ok: false,
          error: `Method ${req.method} not allowed`,
          allowed: allowedMethods,
        });
      }

      // Auth check (if required)
      if (requireAuth) {
        const authHeader = req.headers.authorization;
        if (!authHeader || !authHeader.startsWith('Bearer ')) {
          return res.status(401).json({
            ok: false,
            error: 'Missing or invalid Authorization header',
          });
        }
      }

      // Build params from query string (Deno handler convention)
      const params: Record<string, string> = {};
      for (const [key, value] of Object.entries(req.query)) {
        params[key] = String(value || '');
      }

      // Create lightweight Request-like object for handlers that need it
      const reqLike = {
        method: req.method,
        headers: req.headers,
        json: async () => req.body,
      };

      // Call the Deno handler
      const result = await handler(params, req.method !== 'GET' ? reqLike : undefined);

      // Normalize response
      const status = result.status || (result.ok === false ? 400 : 200);
      return res.status(status).json(result);

    } catch (err: any) {
      console.error('[handler error]', req.path, err.message);
      console.error('[stack trace]', err.stack);

      // Log TCP/tunnel-specific errors clearly
      if (err.code === 'ECONNREFUSED' || err.code === 'ETIMEDOUT') {
        console.error('[TCP TUNNEL ERROR] PgBouncer connection failed:', {
          code: err.code,
          address: err.address,
          port: err.port,
        });
      }

      return res.status(500).json({
        ok: false,
        error: err.message || 'Internal server error',
        code: err.code,
      });
    }
  };
}

// ── Health & System Routes ──────────────────────────────────────────────────
app.get('/health', async (_req: Request, res: Response) => {
  const dbHealthy = await testConnection();
  res.status(dbHealthy ? 200 : 503).json({
    ok: dbHealthy,
    service: 'HandyManager Express Backend',
    version: '1.0.0',
    database: dbHealthy ? 'connected' : 'unavailable',
    timestamp: new Date().toISOString(),
  });
});

app.get('/', (_req: Request, res: Response) => {
  res.json({
    ok: true,
    service: 'HandyManager Express Backend',
    version: '1.0.0',
    endpoints: [
      'GET /health',
      'POST /api/units',
      'GET /api/unit_lookup',
      'POST /api/turns',
      'POST /api/unit_turns',
      'POST /api/turns_incremental',
      'POST /api/unit_turns_history',
      'POST /api/closed_turns',
      'POST /api/turn_records',
      'POST /api/turn_record_stage',
      'POST /api/reassignment_queue',
      'POST /api/device_otp_request',
      'POST /api/device_otp_verify',
      'POST /api/device_setup',
      'POST /api/verify_role',
      'GET /api/session_info',
      'GET /api/trusted_devices',
      'DELETE /api/trusted_devices/:token',
      'POST /api/pm_proxy_user',
      'DELETE /api/pm_proxy_user/:id',
    ],
  });
});

// ── API Routes: Units ───────────────────────────────────────────────────────
app.post('/api/units', wrapDenoHandler(handleUnits));
app.get('/api/unit_lookup', wrapDenoHandler(handleUnitLookup));

// ── API Routes: Turns ───────────────────────────────────────────────────────
app.post('/api/turns', wrapDenoHandler(handleTurns));
app.post('/api/unit_turns', wrapDenoHandler(handleUnitTurns));
app.post('/api/turns_incremental', wrapDenoHandler(handleTurnsIncremental));
app.post('/api/unit_turns_history', wrapDenoHandler(handleUnitTurnsHistory));
app.post('/api/closed_turns', wrapDenoHandler(handleClosedTurns, { allowedMethods: ['GET', 'POST'] }));
app.post('/api/turn_records', wrapDenoHandler(handleTurnRecords, { allowedMethods: ['GET', 'POST'] }));
app.post('/api/turn_record_stage', wrapDenoHandler(handleTurnRecordStage, { allowedMethods: ['POST'] }));

// ── API Routes: Reassignment Queue ──────────────────────────────────────────
app.post('/api/reassignment_queue', wrapDenoHandler(handleReassignmentQueue));

// ── API Routes: Device Auth / PM Login ──────────────────────────────────────
app.post('/api/device_otp_request', wrapDenoHandler(handleDeviceOtpRequest, { allowedMethods: ['POST'] }));
app.post('/api/device_otp_verify', wrapDenoHandler(handleDeviceOtpVerify, { allowedMethods: ['POST'] }));
app.post('/api/device_setup', wrapDenoHandler(handleDeviceSetup, { allowedMethods: ['POST'] }));
app.post('/api/verify_role', wrapDenoHandler(handleVerifyRole, { allowedMethods: ['POST'] }));

// Session info endpoint (token validation)
app.get('/api/session_info', async (req: Request, res: Response) => {
  try {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ ok: false, authenticated: false, error: 'Missing token' });
    }

    const token = authHeader.replace('Bearer ', '').trim();
    const session = await getTrustedDeviceSession(token);

    if (!session) {
      return res.status(401).json({ ok: false, authenticated: false, error: 'Invalid or expired token' });
    }

    return res.json({
      ok: true,
      authenticated: true,
      session: {
        role: session.role || 'full',
        user_name: session.user_name || '',
        login_email: session.login_email || '',
        property_group_uuid: session.property_group_uuid || '',
      },
    });
  } catch (err: any) {
    console.error('[session_info error]', err.message);
    return res.status(500).json({ ok: false, error: err.message });
  }
});

// Trusted devices management
app.get('/api/trusted_devices', wrapDenoHandler(handleTrustedDeviceList));
app.delete('/api/trusted_devices/:token', async (req: Request, res: Response) => {
  try {
    const result = await handleTrustedDeviceRevoke({ token: req.params.token });
    const status = result.status || (result.ok === false ? 400 : 200);
    res.status(status).json(result);
  } catch (err: any) {
    console.error('[trusted_devices revoke error]', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// PM proxy user management
app.post('/api/pm_proxy_user', async (req: Request, res: Response) => {
  try {
    const result = await handlePmProxyUserUpsert({}, req as any);
    const status = result.status || (result.ok === false ? 400 : 200);
    res.status(status).json(result);
  } catch (err: any) {
    console.error('[pm_proxy_user upsert error]', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.delete('/api/pm_proxy_user/:id', async (req: Request, res: Response) => {
  try {
    const result = await handlePmProxyUserDelete({ id: req.params.id });
    const status = result.status || (result.ok === false ? 400 : 200);
    res.status(status).json(result);
  } catch (err: any) {
    console.error('[pm_proxy_user delete error]', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── 404 Handler ─────────────────────────────────────────────────────────────
app.use((_req: Request, res: Response) => {
  res.status(404).json({
    ok: false,
    error: 'Endpoint not found',
    hint: 'See GET / for available endpoints',
  });
});

// ── Error Handler ───────────────────────────────────────────────────────────
app.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  console.error('[express error]', err.message);
  console.error('[stack trace]', err.stack);
  res.status(500).json({
    ok: false,
    error: err.message || 'Internal server error',
  });
});

// ── Server Startup ──────────────────────────────────────────────────────────
const server = app.listen(PORT, () => {
  console.log(`\n🚀 HandyManager Express Backend`);
  console.log(`   Port: ${PORT}`);
  console.log(`   Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`   Health: http://localhost:${PORT}/health`);
  console.log(`\n📡 CORS Whitelist:`);
  allowedOrigins.forEach(origin => console.log(`   - ${origin}`));
  console.log('');
});

// ── Graceful Shutdown ───────────────────────────────────────────────────────
process.on('SIGTERM', async () => {
  console.log('\n[shutdown] SIGTERM received, closing connections...');
  server.close(async () => {
    await closeConnection();
    console.log('[shutdown] Server closed gracefully');
    process.exit(0);
  });
});

process.on('SIGINT', async () => {
  console.log('\n[shutdown] SIGINT received, closing connections...');
  server.close(async () => {
    await closeConnection();
    console.log('[shutdown] Server closed gracefully');
    process.exit(0);
  });
});

export default app;
