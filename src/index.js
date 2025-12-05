// Copyright 2025 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

import fs from 'node:fs';
import os from 'node:os';
import * as crypto from 'node:crypto';
import path from 'node:path';

import * as core from '@actions/core';
import * as exec from '@actions/exec';
import * as io from '@actions/io';
import * as tc from '@actions/tool-cache';

const POWERPOINT_PACKAGE_NAME = 'Microsoft_PowerPoint_16.102.25101829_Updater.pkg';
const INSTALLER_URL =
  `https://officecdn.microsoft.com/pr/C1297A47-86C4-4C1F-97FA-950631F94777/MacAutoupdate/${POWERPOINT_PACKAGE_NAME}`;

async function main() {
  try {
    await run();
  } catch (error) {
    core.setFailed(error.message);
  }
}

async function run() {
  const platform = os.platform();
  const isMacRunner = platform === 'darwin';

  if (!isMacRunner) {
    throw new Error(
      `The setup-powerpoint action supports macOS runner only. Detected platform: '${platform}'.`
    );
  }

  core.info('Setting up Microsoft PowerPoint for Mac');

  const installerPath = await downloadInstaller();
  const powerpointAppPath = await installPowerPoint(installerPath);
  await reportInstalledVersion(powerpointAppPath);
  await configurePowerPointPolicies();
  await enableUiAutomation();
}

async function downloadInstaller() {
  core.startGroup('Download Installer');
  try {
    core.info(`Downloading Microsoft PowerPoint installer package...`);
    const tempDirectory = process.env['RUNNER_TEMP'] || os.tmpdir();
    const targetPath = path.join(tempDirectory, crypto.randomUUID(), POWERPOINT_PACKAGE_NAME);
    
    const downloadedPath = await tc.downloadTool(INSTALLER_URL, targetPath);
    core.info('Installer package downloaded successfully.');
    
    return downloadedPath;
  } finally {
    core.endGroup();
  }
}

async function installPowerPoint(installerPath) {
  core.startGroup('Install Microsoft PowerPoint');
  try {
    core.info('Installing Microsoft PowerPoint application...');
    const exitCode = await exec.exec('sudo', ['installer', '-pkg', installerPath, '-target', '/Applications']);
    if (exitCode !== 0) {
      throw new Error(`Microsoft PowerPoint installation failed with code ${exitCode}.`);
    }

    return '/Applications/Microsoft PowerPoint.app';
  } finally {
    core.endGroup();
  }
}

async function enableUiAutomation() {
  core.startGroup('Enable user interface automation');
  try {
    const terminalAppPath = '/System/Applications/Utilities/Terminal.app/';
    const tccDatabasePath = '/Library/Application Support/com.apple.TCC/TCC.db';
    const csreqHex = await generateCsreqHex(terminalAppPath);

    const serviceName = 'kTCCServiceAccessibility';
    const bundleId = 'com.apple.Terminal';
    const clientType = 0;
    const authValue = 2;
    const authReason = 4;
    const authVersion = 1;

    const insertStatement = `INSERT or REPLACE INTO access (service, client, client_type, auth_value, auth_reason, auth_version, csreq) `+
      `VALUES('${serviceName}','${bundleId}',${clientType},${authValue},${authReason},${authVersion},X'${csreqHex}');`;

    const { exitCode, stderr, stdout } = await exec.getExecOutput('sudo', [
      'sqlite3',
      tccDatabasePath,
      insertStatement,
    ], { ignoreReturnCode: true, failOnStdErr: false });

    if (exitCode !== 0) {
      const message = (stderr || stdout || '(unknown SQLite error)').trim();
      throw new Error(`Failed to apply changes to TCC.db database. Exit code ${exitCode}. ${message}`);
    } else {
      core.info('Granted permission to Terminal.app to automate user interface.');
    }
  } finally {
    core.endGroup();
  }
}

async function generateCsreqHex(appPath) {
  const tempDir = path.join(os.tmpdir(), crypto.randomUUID());
  const csreqPath = path.join(tempDir, 'csreq.bin');

  await io.mkdirP(tempDir);

  try {
    const { exitCode, stdout: codesignOutput } = await exec.getExecOutput('codesign', [
      '-d',
      '-r-',
      appPath,
    ]);

    if (exitCode !== 0) {
      throw new Error(`Failed to get code signature designation for '${appPath}'.`);
    }

    const designationMatch = codesignOutput.match(/designated\s*=>\s*(.+)/);
    const appDesignation = designationMatch[1].trim();

    await exec.exec('csreq', ['-r-', '-b', csreqPath], {
      input: Buffer.from(appDesignation, 'utf8'),
    });

    const csreqBinary = fs.readFileSync(csreqPath);
    return csreqBinary.toString('hex');
  } finally {
    await io.rmRF(tempDir);
  }
}

async function reportInstalledVersion( powerpointAppPath) {
  const powerpointPlistPath = `${powerpointAppPath}/Contents/Info.plist`;
  const version = await readPlistValue(powerpointPlistPath, 'CFBundleShortVersionString');
  const buildRaw = await readPlistValue(powerpointPlistPath, 'CFBundleVersion');
  const build = buildRaw ? buildRaw.split('.').pop() : '(unknown)';

  core.notice(`Microsoft PowerPoint version ${version} (${build})`);
}


async function configurePowerPointPolicies() {
  core.startGroup('Configure Microsoft PowerPoint policies');
  try {
    const policiesDir = path.join(import.meta.dirname, '..', 'policies');
    const policyScripts = [
      'policy_ms_autoupdate.sh',
      'policy_ms_office.sh',
      'policy_ms_powerpoint.sh',
    ];

    for (const script of policyScripts) {
      const scriptPath = path.join(policiesDir, script);
      if (!fs.existsSync(scriptPath)) {
        core.info(`Skipping policy script '${script}' as it does not exist.`);
        continue;
      }
      
      core.info(`Applying policy script '${script}'...`);
      core.debug(`Executing script at path: '${scriptPath}'`);
      await exec.exec('bash', [scriptPath], { cwd: policiesDir });
    }
  } finally {
    core.endGroup();
  }
}


async function readPlistValue(plistPath, key) {
  try {
    const { exitCode, stdout } = await exec.getExecOutput('defaults', ['read', plistPath, key]);
    if (exitCode !== 0) {
      core.error(`Failed to read '${plistPath}' key '${key}'. Exit code: ${exitCode}`);
    }

    return stdout.trim();
  } catch (error) {
    core.error(`Failed to read '${plistPath}' key '${key}'. Error: ${error.message}`);
    return '(unknown)';
  }
}

main();
