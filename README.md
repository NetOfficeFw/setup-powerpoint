# setup-powerpoint

> Action to setup Microsoft PowerPoint for Mac in GitHub Actions runner.

## Requirements

Ths action requires macOS operating system.

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

      - name: Run automation
        run: |
          # PowerPoint is installed in /Applications
          ls "/Applications/Microsoft PowerPoint.app"
```


## License

Source code is licensed under [MIT License](LICENSE.txt).

© 2025 Cisco Systems, Inc. All rights reserved.
