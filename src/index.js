// Copyright 2025 Cisco Systems, Inc.
// Licensed under MIT-style license (see LICENSE.txt file).

import os from 'node:os';
import * as crypto from 'node:crypto';
import path from 'node:path';

import * as core from '@actions/core';
import * as exec from '@actions/exec';
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

async function reportInstalledVersion( powerpointAppPath) {
  const powerpointPlistPath = `${powerpointAppPath}/Contents/Info.plist`;
  const version = await readPlistValue(powerpointPlistPath, 'CFBundleShortVersionString');
  const buildRaw = await readPlistValue(powerpointPlistPath, 'CFBundleVersion');
  const build = buildRaw ? buildRaw.split('.').pop() : '(unknown)';

  core.notice(`Microsoft PowerPoint version ${version} (${build})`);
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
