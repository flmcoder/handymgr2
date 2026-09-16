import { handleApiResponse, jsonResp } from '../utils.ts';

export async function handleTechRoster(params: Record<string, unknown>, req: Request): Promise<any> {
  try {
    const action = String(params.action || '').trim();
    if (!action) return jsonResp({ ok: false, error: 'missing action' }, 400);

    // Map action to backend endpoint
    let url: string;
    if (action === 'tech_roster') {
      url = '/api?action=tech_roster';
    } else if (action === 'dispatch_sync_assignees') {
      url = '/api?action=dispatch_sync_assignees';
    } else {
      return jsonResp({ ok: false, error: `unknown action: ${action}` }, 400);
    }

    const result = await fetch(url, {
      method: action === 'tech_roster' && params.method === 'POST' ? 'POST' : 'GET',
      headers: { 'Content-Type': 'application/json' },
      body: action !== 'tech_roster' || params.method !== 'POST' ? JSON.stringify(params) : undefined,
    });
    const data = await result.json();
    return handleApiResponse(data, 'tech roster updated');
  } catch (error) {
    console.error('[techRoster] handler failed', error);
    return jsonResp({ ok: false, error: String((error as any)?.message || error || 'tech roster handler failed') }, 500);
  }
}

export async function handleDispatchSyncAssignees(params: Record<string, unknown>, req: Request): Promise<any> {
  try {
    // This is handled by the same handler via the action parameter
    return handleTechRoster(params, req);
  } catch (error) {
    console.error('[dispatchSyncAssignees] handler failed', error);
    return jsonResp({ ok: false, error: String((error as any)?.message || error || 'dispatch assignees handler failed') }, 500);
  }
}