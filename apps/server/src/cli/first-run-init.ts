// Invoked by the installer once at install time:
//   "{app}\autoOffice-server.exe" --first-run-init
// Idempotent: safe to run on upgrades — only creates config when absent,
// but always ensures the cert is present in the trust store.
import { mkdirSync, writeFileSync, existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { generateCert } from '../tls/generate';
import { installCertToCurrentUserRoot } from '../tls/install-store';
import { makeFreshConfig, saveConfig, loadConfig, configPath } from '../lifecycle/config';
import { resolveDataDir, dbPath } from '../env';
import { openDb } from '../db';

export async function firstRunInit(): Promise<void> {
  const dataDir = resolveDataDir();
  const cfgPath = configPath(dataDir);
  let cfg = loadConfig(dataDir);

  if (cfg) {
    console.log(`[autoOffice] config exists at ${cfgPath} — ensuring cert is installed …`);
    await ensureCertInstalled(dataDir, cfg.certPath);
    return;
  }

  console.log('[autoOffice] generating bearer token + cert …');
  cfg = makeFreshConfig({ port: 47318 });
  const bundle = generateCert({ commonName: `AutoOffice (${cfg.installId})`, validityYears: 10 });

  const certDir = join(dataDir, 'config');
  mkdirSync(certDir, { recursive: true });
  writeFileSync(join(certDir, 'cert.pem'), bundle.cert, 'utf8');
  writeFileSync(join(certDir, 'key.pem'), bundle.key, 'utf8');
  cfg.certFingerprint = bundle.fingerprint;

  saveConfig(dataDir, cfg);

  console.log('[autoOffice] installing cert to LocalMachine\\Root …');
  await ensureCertInstalled(dataDir, cfg.certPath);

  console.log('[autoOffice] initializing database …');
  openDb({ url: dbPath() }).close();

  console.log(`[autoOffice] init complete. Data dir: ${dataDir}`);
}

async function ensureCertInstalled(dataDir: string, certPath: string): Promise<void> {
  const fullCertPath = join(dataDir, certPath);
  if (!existsSync(fullCertPath)) {
    console.error(`[autoOffice] cert file not found at ${fullCertPath}, skipping install.`);
    return;
  }
  const certPem = readFileSync(fullCertPath, 'utf8');
  try {
    await installCertToCurrentUserRoot(certPem);
    console.log('[autoOffice] cert installed to LocalMachine\\Root.');
  } catch (err) {
    console.error('[autoOffice] cert install failed; the user may need to install it manually.');
    console.error((err as Error).message);
  }
}
