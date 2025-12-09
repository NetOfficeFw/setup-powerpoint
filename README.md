# setup-powerpoint

> Action to setup Microsoft PowerPoint for Mac in GitHub Actions runner.

## Requirements

This action requires macOS operating system.

## Usage

```yaml
name: PowerPoint Automation

on:
  - push:

jobs:
  powerpoint-tests:
    runs-on: macos-latest
    steps:
      - uses: actions/checkout@v6

      - name: Install Microsoft PowerPoint for Mac
        uses: NetOfficeFw/setup-powerpoint@v1
        # Optional: install a different Microsoft PowerPoint release
        # with:
        #   package: Microsoft_PowerPoint_16.103.25113013_Updater.pkg

      - name: Run automation
        run: |
          # PowerPoint is installed in /Applications
          ls "/Applications/Microsoft PowerPoint.app"
```

### Action Outputs

The `setup-powerpoint` action will output these variables:

- `path`: Installation path of Microsoft PowerPoint (e.g., `/Applications/Microsoft PowerPoint.app`).
- `version`: Installed Microsoft PowerPoint version (`CFBundleShortVersionString`).
- `build`: Build number (last component of `CFBundleVersion`).
- `package`: Installer package file name used during installation.
- `installer-url`: Full URL used to download the installer.


## License

Source code is licensed under [MIT License](LICENSE.txt).

© 2025 Cisco Systems, Inc. All rights reserved.
