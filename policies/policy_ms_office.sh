#!/usr/bin/env bash
#
# Copyright 2025 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).
#

set -e

# Diagnostic and telemetry settings
defaults write com.microsoft.office DiagnosticDataTypePreference -string ZeroDiagnosticData
defaults write com.microsoft.office OptionalConnectedExperiencesPreference -bool false
defaults write com.microsoft.office SendAllTelemetryEnabled -bool false
defaults write com.microsoft.office DisableTelemetry -bool true

# First Run Experience (FRE) settings
defaults write com.microsoft.office ConnectedOfficeExperiencesPreference -bool true
defaults write com.microsoft.office HasUserSeenEnterpriseFREDialog -bool true
defaults write com.microsoft.office HasUserSeenFREDialog -bool true
defaults write com.microsoft.office FirstRunExperienceCompletedO16 -bool true
defaults write com.microsoft.office OfficeActivationWizardHasRun -bool true

# Sign-in and connected experiences
defaults write com.microsoft.office OfficeAutoSignIn -bool false
defaults write com.microsoft.office OfficeExperiencesAnalyzingContentPreference -bool true
defaults write com.microsoft.office OfficeExperiencesDownloadingContentPreference -bool true

# Privacy and EULA settings
defaults write com.microsoft.office AcceptedPrivacyPolicyVersion -int 2
defaults write com.microsoft.office UserInfo.AcceptedPrivacyPolicyVersion -int 2
defaults write com.microsoft.office DisablePrivacySettings -bool true
defaults write com.microsoft.office EULAAccepted -bool true
