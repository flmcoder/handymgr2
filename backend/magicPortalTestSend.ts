import type { Request, Response } from 'express';
import type { MagicPortalInput } from './magicPortal.ts';

type TestPortalSession = {
  magicLink: string;
  shortCode: string;
  expiresAt: string;
};

type TestSendDependencies = {
  requireFullAdminSession: (req: Request, res: Response) => Promise<unknown | null>;
  createSession: (input: MagicPortalInput) => Promise<TestPortalSession>;
  sendSms: (phone: string, message: string) => Promise<{ messageId: string }>;
  buildMessage: (workOrderNumber: string, link: string) => string;
};

export function createMagicPortalTestSendHandler(deps: TestSendDependencies) {
  return async (req: Request, res: Response): Promise<void> => {
    const session = await deps.requireFullAdminSession(req, res);
    if (!session) return;

    const phone = String(req.body?.phone || '').trim();
    if (!/^\+\d{10,15}$/.test(phone)) {
      res.status(400).json({ ok: false, error: 'A test recipient phone in E.164 format is required' });
      return;
    }

    const workOrderNumber = `TEST-${Date.now()}`;
    const portalSession = await deps.createSession({
      woId: `TEST-${workOrderNumber}`,
      woNumber: workOrderNumber,
      techId: 'TEST-TECH',
      techName: 'Dispatch Test',
      techPhone: phone,
      tenantName: 'John Doe (Test Resident)',
      tenantPhone: '+15551232314',
      propertyAddress: 'TEST PROPERTY — 123 Example Street',
      sessionMode: 'test',
    });

    try {
      const sms = await deps.sendSms(phone, deps.buildMessage(workOrderNumber, portalSession.magicLink));
      res.status(200).json({
        ok: true,
        test: true,
        magic_link: portalSession.magicLink,
        short_code: portalSession.shortCode,
        expires_at: portalSession.expiresAt,
        rc_message_id: sms.messageId,
      });
    } catch (error) {
      res.status(502).json({
        ok: false,
        test: true,
        error: String((error as Error)?.message || 'Test SMS handoff failed'),
      });
    }
  };
}