#!/usr/bin/env bash
#
# Copyright 2025 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).
#

set -e

defaults write com.microsoft.autoupdate2 AcknowledgedDataCollectionPolicy -string RequiredDataOnly
defaults write com.microsoft.autoupdate2 SendAllTelemetryEnabled -bool false
