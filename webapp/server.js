import fs from 'node:fs';
import path from 'node:path';
import { createAdvisor } from './src/advisor.js';
import { createApp } from './src/app.js';
import { config } from './src/config.js';

function die(message) {
  console.error(`\nCannot start: ${message}\n`);
  process.exit(1);
}

if (!config.appPassword && !config.insecure) {
  die('set GRID_APP_PASSWORD (the one password of the app). For local development only you can set GRID_APP_INSECURE=1 instead: no password, and the server listens on 127.0.0.1 only.');
}
if (config.appPassword && config.appPassword.length < 8) die('GRID_APP_PASSWORD must have at least 8 characters.');
if (!fs.existsSync(path.join(config.gridTools, 'liberty_grid_to_config.py'))) die(`the converter scripts were not found in ${config.gridTools} (set GRID_TOOLS_DIR).`);
if (path.isAbsolute(config.python) && !fs.existsSync(config.python)) die(`the Python interpreter ${config.python} does not exist (set PYTHON).`);

const advisor = process.env.ADVISOR === '1' && process.env.ANTHROPIC_API_KEY
  ? createAdvisor({ apiKey: process.env.ANTHROPIC_API_KEY, model: process.env.ADVISOR_MODEL || undefined, baseUrl: process.env.ANTHROPIC_BASE_URL || undefined })
  : null;
if (process.env.ADVISOR === '1' && !advisor) console.warn('ADVISOR=1 but ANTHROPIC_API_KEY is not set: the advisor is off.');
if (advisor) console.log('Advisor on: cards memory cannot answer may carry a suggested answer (never applied by itself).');
const { app, agent, masters } = createApp({ config, advisor });
const host = process.env.HOST || (config.insecure ? '127.0.0.1' : '0.0.0.0');
const server = app.listen(config.port, host, () => {
  console.log(`Grid agent on http://${host === '0.0.0.0' ? 'localhost' : host}:${config.port}  (converters: ${config.gridTools}, data: ${config.dataDir})`);
  if (config.insecure) console.log('GRID_APP_INSECURE=1: no password is asked. Use this on your own computer only.');
  masters.ensure().then((m) => {
    if (!m.loaded()) console.warn('Warning: the master lists (RTO / make-model) could not be read from the master folder; names cannot be checked or mapped until they can.');
  });
});

server.requestTimeout = 15 * 60 * 1000;          // a 150 MB upload over a slow line takes longer than node's default 5 minutes

async function stop() {
  console.log('\nStopping: waiting for the running jobs to finish their current run...');
  server.close();
  const timer = setTimeout(() => process.exit(0), 60 * 1000);
  timer.unref();
  await agent.idle();
  process.exit(0);
}
process.on('SIGINT', stop);
process.on('SIGTERM', stop);
