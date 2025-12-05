#!/usr/bin/env bash
#
# Copyright 2025 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).
#

set -e

POWERPOINT_PREFERENCES="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Preferences"

# Ensure the containerized preferences directory exists
mkdir -p "$POWERPOINT_PREFERENCES"

# First Run Experience settings
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" LoginKeychainEmpty -bool true
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" OUIShouldEstablishWhatsNewBaseline -bool true
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" kSetupUIKeyNotSupportedLicenseFound -bool true
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" kSubUIAppCompletedFirstRunSetup1507 -bool true
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" FirstRunExperienceCompletedO16 -bool true

# Telemetry and diagnostic settings
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" SendAllTelemetryEnabled -bool false
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" DisableTelemetry -bool true

# Privacy and activation settings
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" OfficeActivationWizardHasRun -bool true
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" AcceptedPrivacyPolicyVersion -int 2
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" DisablePrivacySettings -bool true

# Disable additional first-run prompts
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" ShowWhatsNewOnLaunch -bool false
defaults write "$POWERPOINT_PREFERENCES/com.microsoft.Powerpoint.plist" ShowTemplateGallery -bool false
