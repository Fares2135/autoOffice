import SysTrayModule from 'systray';
import { spawn } from 'node:child_process';
import { rotateToken } from '../lifecycle/config';
import iconPng from './icon.png' with { type: 'binary' };

// Handle CJS default export interop
const SysTray = (SysTrayModule as unknown as { default?: typeof SysTrayModule }).default ?? SysTrayModule;

const ICON_B64 = Buffer.from(iconPng).toString('base64');

export async function startTray(opts: { port: number; dataDir: string }) {
  const tray = new SysTray({
    menu: {
      icon: ICON_B64,
      title: 'AutoOffice',
      tooltip: `AutoOffice on https://localhost:${opts.port}`,
      items: [
        { title: 'Open guide', tooltip: '', checked: false, enabled: true },
        { title: 'Restart service', tooltip: '', checked: false, enabled: true },
        { title: 'Rotate token', tooltip: 'Invalidate the current bearer and write a new one', checked: false, enabled: true },
        { title: 'Quit', tooltip: '', checked: false, enabled: true },
      ],
    },
    debug: false,
    copyDir: false,
  });

  tray.onClick(async (action) => {
    switch (action.seq_id) {
      case 0: // Open guide
        spawn('rundll32', ['url.dll,FileProtocolHandler', 'https://sivan22.github.io/autoOffice/guide/'], { detached: true, stdio: 'ignore' }).unref();
        break;
      case 1: // Restart
        process.exit(0);
        break;
      case 2: // Rotate token
        rotateToken(opts.dataDir);
        process.exit(0);
        break;
      case 3: // Quit
        tray.kill();
        process.exit(0);
    }
  });
}
