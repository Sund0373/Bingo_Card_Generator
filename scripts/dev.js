#!/usr/bin/env node
/**
 * One-command bootstrapper: finds Python, creates .venv, installs dependencies
 * (only when requirements.txt changes), then starts the web UI.
 *
 *   npm run dev                  setup if needed, then start the UI
 *   npm run dev -- --port 8000   extra args are forwarded to app.py
 *   npm run setup                setup only, do not start
 *   npm run start                start without re-checking dependencies
 *
 * Uses Node built-ins only, so there is nothing to npm install.
 */

'use strict';

const { spawn, spawnSync } = require('child_process');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const VENV = path.join(ROOT, '.venv');
const IS_WIN = process.platform === 'win32';
const VENV_PY = IS_WIN
  ? path.join(VENV, 'Scripts', 'python.exe')
  : path.join(VENV, 'bin', 'python');
const REQUIREMENTS = path.join(ROOT, 'requirements.txt');
const STAMP = path.join(VENV, '.deps-stamp');
const MIN_PYTHON = [3, 10];

const argv = process.argv.slice(2);
const setupOnly = argv.includes('--setup-only');
const skipSetup = argv.includes('--skip-setup');
const appArgs = argv.filter((a) => a !== '--setup-only' && a !== '--skip-setup');

const say = (msg) => console.log(`\x1b[36m>\x1b[0m ${msg}`);
const warn = (msg) => console.error(`\x1b[33m!\x1b[0m ${msg}`);

function die(message, hints = []) {
  console.error(`\n\x1b[31mx\x1b[0m ${message}`);
  hints.forEach((h) => console.error(`  ${h}`));
  console.error('');
  process.exit(1);
}

/** Returns {cmd, args} for a Python new enough to run this project. */
function findPython() {
  const candidates = IS_WIN
    ? [['py', ['-3']], ['python', []], ['python3', []]]
    : [['python3', []], ['python', []]];

  const tooOld = [];
  for (const [cmd, base] of candidates) {
    const probe = spawnSync(
      cmd,
      [...base, '-c', 'import sys; print("%d.%d" % sys.version_info[:2])'],
      { encoding: 'utf8' }
    );
    if (probe.status !== 0 || !probe.stdout) continue;

    const [major, minor] = probe.stdout.trim().split('.').map(Number);
    if (major > MIN_PYTHON[0] || (major === MIN_PYTHON[0] && minor >= MIN_PYTHON[1])) {
      return { cmd, args: base, version: `${major}.${minor}` };
    }
    tooOld.push(`${cmd} (found ${major}.${minor})`);
  }

  if (tooOld.length) {
    die(`Python ${MIN_PYTHON.join('.')}+ is required. Too old: ${tooOld.join(', ')}`, [
      'Install a newer Python from https://www.python.org/downloads/',
    ]);
  }
  die('Python was not found on PATH.', [
    'Install Python 3.10+ from https://www.python.org/downloads/',
    IS_WIN ? 'During setup, tick "Add python.exe to PATH".' : '',
  ].filter(Boolean));
}

function run(cmd, args, label) {
  const result = spawnSync(cmd, args, { stdio: 'inherit', cwd: ROOT });
  if (result.error && result.error.code === 'ENOENT') {
    die(`Could not run ${cmd}. Is it installed and on PATH?`);
  }
  if (result.status !== 0) {
    die(`${label} failed.`, [
      'Behind a corporate proxy? Try:',
      '  pip install --proxy http://YOUR_PROXY:PORT -r requirements.txt',
    ]);
  }
}

/** Dependencies are reinstalled only when requirements.txt actually changes. */
function requirementsHash() {
  return crypto.createHash('sha256').update(fs.readFileSync(REQUIREMENTS)).digest('hex');
}

function depsAreCurrent() {
  try {
    return fs.readFileSync(STAMP, 'utf8').trim() === requirementsHash();
  } catch {
    return false;
  }
}

function setup() {
  if (!fs.existsSync(VENV_PY)) {
    const py = findPython();
    say(`Python ${py.version} found. Creating virtual environment in .venv ...`);
    run(py.cmd, [...py.args, '-m', 'venv', VENV], 'Creating the virtual environment');
    // setuptools ships pinned inside venv and is often an old, CVE-flagged build
    run(
      VENV_PY,
      ['-m', 'pip', 'install', '--upgrade', '--quiet', 'pip', 'setuptools'],
      'Upgrading pip and setuptools'
    );
  }

  if (depsAreCurrent()) {
    say('Dependencies are up to date.');
    return;
  }

  say('Installing dependencies ...');
  run(VENV_PY, ['-m', 'pip', 'install', '-r', REQUIREMENTS], 'Installing dependencies');
  fs.writeFileSync(STAMP, requirementsHash());
  say('Dependencies installed.');
}

function start() {
  if (!fs.existsSync(VENV_PY)) {
    die('No virtual environment found.', ['Run: npm run setup']);
  }
  say('Starting the Bingo Card Generator ...\n');

  const child = spawn(VENV_PY, [path.join(ROOT, 'app.py'), ...appArgs], {
    stdio: 'inherit',
    cwd: ROOT,
  });

  // Let Python handle Ctrl+C so it can shut the server down cleanly.
  const forward = (signal) => () => child.kill(signal);
  process.on('SIGINT', forward('SIGINT'));
  process.on('SIGTERM', forward('SIGTERM'));

  child.on('exit', (code, signal) => {
    if (signal) process.exit(0);
    process.exit(code ?? 0);
  });
}

if (!fs.existsSync(REQUIREMENTS)) {
  die(`requirements.txt not found in ${ROOT}`, ['Are you in the project folder?']);
}

if (!skipSetup) setup();
if (setupOnly) {
  say('Setup complete. Run: npm run dev');
} else {
  start();
}
