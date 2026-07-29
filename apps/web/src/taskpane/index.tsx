import React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webDarkTheme, webLightTheme, Text } from '@fluentui/react-components';
import { App } from './App.tsx';
import { detectHost, UnsupportedHostError, type HostContext } from './host/context.ts';
import { LanguageProvider, useDirection, useTranslation } from './i18n/index.ts';
import { ThemeProvider, useThemeMode } from './theme/context.tsx';

function Shell({ children }: { children: React.ReactNode }) {
  const dir = useDirection();
  const { resolved } = useThemeMode();
  return (
    <FluentProvider theme={resolved === 'dark' ? webDarkTheme : webLightTheme} dir={dir}>
      {children}
    </FluentProvider>
  );
}

function FatalShell({ message }: { message: string }) {
  const { t } = useTranslation();
  return (
    <Shell>
      <div style={{ padding: '24px' }}>
        <Text size={400} weight="semibold">{t('fatal.cannotStart')}</Text>
        <p>{message}</p>
      </div>
    </Shell>
  );
}

const rootElement = document.getElementById('root')!;
const root = createRoot(rootElement);

function renderApp(host: HostContext) {
  root.render(
    <ThemeProvider>
      <LanguageProvider>
        <Shell>
          <App host={host} />
        </Shell>
      </LanguageProvider>
    </ThemeProvider>,
  );
}

function renderFatal(message: string) {
  root.render(
    <ThemeProvider>
      <LanguageProvider>
        <FatalShell message={message} />
      </LanguageProvider>
    </ThemeProvider>,
  );
}

function start() {
  try {
    renderApp(detectHost());
  } catch (e) {
    if (e instanceof UnsupportedHostError) {
      renderFatal(e.message);
    } else {
      renderFatal(e instanceof Error ? e.message : String(e));
    }
  }
}

if (typeof Office !== 'undefined') {
  // Office.onReady never fires when the SPA is loaded outside an Office host
  // (regular browser, Playwright). Race it against a timeout so the dev /
  // E2E flow still bootstraps quickly into the Word fallback.
  let started = false;
  const safeStart = () => {
    if (started) return;
    started = true;
    start();
  };
  Office.onReady(() => safeStart());
  setTimeout(safeStart, 1500);
} else {
  start();
}
