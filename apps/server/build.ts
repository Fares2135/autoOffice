// Compiles the server + web app into dist/.
// Run with: bun build.ts
import { $ } from 'bun';

const out = './dist/autoOffice-server.exe';
await $`mkdir -p dist`;

console.log('Building web app …');
await $`bun run build`.cwd('../web');
console.log('OK → apps/web/dist');

console.log('Building autoOffice-server.exe …');
await $`bun build ./src/index.ts --compile --target=bun-windows-x64 --outfile=${out} --minify --define 'process.env.NODE_ENV="production"' --external ai-sdk-provider-gemini-cli`;
console.log(`OK → ${out}`);
