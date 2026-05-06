import { execFile } from 'node:child_process';
import { writeFileSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

export async function installCertToCurrentUserRoot(certPem: string): Promise<void> {
  if (process.platform !== 'win32') {
    throw new Error('Cert install is only supported on Windows in v1.');
  }
  if (!/-----BEGIN CERTIFICATE-----/.test(certPem)) {
    throw new Error('Expected PEM-encoded certificate');
  }
  const certPath = join(tmpdir(), `autooffice-${Date.now()}.cer`);
  writeFileSync(certPath, certPem, 'utf8');
  try {
    // certutil -addstore targets LocalMachine\Root and requires admin elevation,
    // which the installer always provides. It handles PEM format natively and
    // exits non-zero on any failure — more reliable than Import-Certificate.
    await execFileAsync('certutil', ['-addstore', 'Root', certPath]);
  } finally {
    try { unlinkSync(certPath); } catch { /* noop */ }
  }
}

export async function uninstallCertByFingerprint(fingerprintHex: string): Promise<void> {
  if (process.platform !== 'win32') return;
  const fp = fingerprintHex.replace(/[^A-F0-9]/gi, '').toUpperCase();
  await execFileAsync('certutil', ['-delstore', 'Root', fp]);
}
