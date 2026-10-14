import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import { createAuthModule } from './auth.js';

function adminPanelHarness(role, ready) {
  const calls = [];
  const deps = {
    isSessionReady: () => ready,
    getAccessRole: () => role,
    loadPmUsers: async () => { calls.push('users'); },
    loadOtpSettings: async () => { calls.push('settings'); },
  };
  return { calls, deps };
}

for (const role of ['full', 'manager']) {
  test(`refreshes both admin panels for an authenticated ${role}`, async () => {
    const { calls, deps } = adminPanelHarness(role, true);
    await createAuthModule(deps).refreshAdminPanels();
    assert.deepEqual(calls, ['users', 'settings']);
  });
}

test('waits for authentication before loading, then refreshes after sign-in', async () => {
  const { calls, deps } = adminPanelHarness('full', false);
  const auth = createAuthModule(deps);
  await auth.refreshAdminPanels();
  assert.deepEqual(calls, []);
  deps.isSessionReady = () => true;
  await auth.refreshAdminPanels();
  assert.deepEqual(calls, ['users', 'settings']);
});

for (const role of ['pm_readonly', 'vendors', '', undefined, 'admin']) {
  test(`does not request admin data for unsupported role ${role}`, async () => {
    const { calls, deps } = adminPanelHarness(role, true);
    await createAuthModule(deps).refreshAdminPanels();
    assert.deepEqual(calls, []);
  });
}

test('does not suppress a failed panel refresh', async () => {
  const { calls, deps } = adminPanelHarness('full', true);
  deps.loadPmUsers = async () => { throw new Error('Invalid session'); };
  await assert.rejects(createAuthModule(deps).refreshAdminPanels(), /Invalid session/);
  assert.deepEqual(calls, ['settings']);
});

const appSource = readFileSync(new URL('./app.js', import.meta.url), 'utf8');
const shellFunction = appSource.match(/function showUnlockedAppShell\(\) \{[\s\S]*?\n\}/)[0];

test('unlocking the app refreshes admin panels only after the vault is hidden', () => {
  const vault = { style: { display: '' } };
  let unlocked = false;
  let refreshes = 0;
  runInNewContext(shellFunction + '\nshowUnlockedAppShell();', {
    $: (selector) => selector === '#vaultScreen'
      ? vault
      : { classList: { add: () => { unlocked = true; } } },
    _authModule: {
      refreshAdminPanels: () => {
        assert.equal(vault.style.display, 'none');
        assert.equal(unlocked, true);
        refreshes++;
      },
    },
  });
  assert.equal(refreshes, 1);
});

test('opening Database Admin retries the authenticated panels', () => {
  const navigationFunction = appSource.match(/function setActiveMainSection\(tabName\) \{[\s\S]*?\n\}/)[0];
  const calls = [];
  runInNewContext(navigationFunction + "\nsetActiveMainSection('dashboard');\nsetActiveMainSection('dbadmin');", {
    _tabsModule: { setActiveMainSection: (tab) => { calls.push(tab); } },
    _authModule: { refreshAdminPanels: () => { calls.push('refresh'); } },
  });
  assert.deepEqual(calls, ['dashboard', 'dbadmin', 'refresh']);
});

test('password, resumed, and device sessions all use the shared unlock refresh', () => {
  const resumeSource = appSource.match(/async function resumeFromPendingSession\(\) \{[\s\S]*?\n\}/)[0];
  const deviceSource = appSource.match(/async function unlockWithDeviceToken\([^)]*\) \{[\s\S]*?\n\}/)[0];
  const passwordStart = appSource.indexOf('// Verify password against proxy env vars');
  const passwordSource = appSource.slice(passwordStart, appSource.indexOf("if ($('#btnSendOtp'))", passwordStart));
  for (const source of [resumeSource, deviceSource, passwordSource]) {
    assert.match(source, /showUnlockedAppShell\(\);/);
  }
});

test('admin initialization has no premature account or OTP settings load', () => {
  const adminSource = appSource.slice(appSource.indexOf('(function initDbAdmin()'), appSource.indexOf('(function initWebhookLiveDrawer()'));
  assert.doesNotMatch(adminSource, /autoLoadPmUsers/);
  assert.doesNotMatch(adminSource, /^\s{2}loadPmUsers\(\);/m);
  assert.doesNotMatch(adminSource, /^\s{4}loadOtpSettings\(\);/m);
  assert.match(adminSource, /_authModule\.refreshAdminPanels\(\);/);
});
