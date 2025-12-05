#!/usr/bin/env bash
#
# Copyright 2025 Cisco Systems, Inc.
# Licensed under MIT-style license (see LICENSE.txt file).
#

set -e

echo "Waiting for PowerPoint privacy modals..."

close_microsoft_powerpoint_app() {
    if pgrep -x "Microsoft PowerPoint" > /dev/null; then
        echo "Closing Microsoft PowerPoint..."
        osascript >/dev/null 2>&1 <<'EOF'
tell application "Microsoft PowerPoint" to quit
EOF
    fi
}

# Wait up to 20 seconds for the privacy modals to appear and dismiss them
max_attempts=20
attempt=0
first_screen_dismissed=false
second_screen_dismissed=false
third_screen_dismissed=false

while [ $attempt -lt $max_attempts ]; do
    # Check if PowerPoint is running
    if pgrep -x "Microsoft PowerPoint" > /dev/null; then
        echo "PowerPoint is running (attempt $attempt/$max_attempts)..."
        
        # Handle first screen if not yet dismissed
        if [[ "$first_screen_dismissed" == "false" ]]; then
            echo "Attempting to dismiss first privacy screen..."
            
            # Press Return key to skip through screen
            result=$(osascript 2>/dev/null <<EOF
try
    tell application "System Events"
        tell process "Microsoft PowerPoint"
            if exists window 1 then
                keystroke return
                return "pressed"
            end if
        end tell
    end tell
    return "not_found"
on error
    return "error"
end try
EOF
)
            
            if [[ "$result" == "pressed" ]]; then
                echo "✓ Successfully completed first screen (pressed Return)"
                first_screen_dismissed=true
                sleep 3
                continue
            fi
        fi
        
        # Second screen ("Getting better together")
        if [[ "$first_screen_dismissed" == "true" && "$second_screen_dismissed" == "false" ]]; then
            echo "Looking for second privacy screen..."
            
            # Click "No" radio button, then press Return twice to accept
            result=$(osascript 2>/dev/null <<EOF
try
    tell application "System Events"
        tell process "Microsoft PowerPoint"
            if exists window 1 then
                -- Get all UI elements and find the "No" radio button
                set allElements to entire contents of window 1
                set foundRadioButton to false
                repeat with elem in allElements
                    try
                        if class of elem is radio button then
                            set btnName to name of elem as string
                            if btnName contains "No" or btnName contains "don't" then
                                click elem
                                delay 0.5
                                set foundRadioButton to true
                                exit repeat
                            end if
                        end if
                    end try
                end repeat
                
                -- Now press Return twice to activate Accept button
                keystroke return
                delay 0.5
                keystroke return
                
                if foundRadioButton then
                    return "radio_clicked_and_accepted"
                else
                    return "only_returns_pressed"
                end if
            end if
        end tell
    end tell
    return "not_found"
on error errMsg
    return "error: " & errMsg
end try
EOF
)
            
            if [[ "$result" == "radio_clicked_and_accepted" || "$result" == "only_returns_pressed" ]]; then
                echo "✓ Successfully completed second screen (clicked radio + double Return)"
                second_screen_dismissed=true
                sleep 3
                # Don't exit yet, continue to check for third screen
            fi
        fi
        
        # Third screen ("Powering your experiences")
        if [[ "$first_screen_dismissed" == "true" && "$second_screen_dismissed" == "true" && "$third_screen_dismissed" == "false" ]]; then
            echo "Looking for third privacy screen..."
            
            # Press Return key to skip through screen
            result=$(osascript 2>/dev/null <<EOF
try
    tell application "System Events"
        tell process "Microsoft PowerPoint"
            if exists window 1 then
                keystroke return
                return "pressed"
            end if
        end tell
    end tell
    return "not_found"
on error
    return "error"
end try
EOF
)
            
            if [[ "$result" == "pressed" ]]; then
                echo "✓ Successfully completed third screen (pressed Return)"
                third_screen_dismissed=true
                sleep 3
                close_microsoft_powerpoint_app
                exit 0
            fi
        fi
    else
        echo "Waiting for PowerPoint to start..."
    fi
    
    sleep 1
    attempt=$((attempt + 1))
done

if [[ "$first_screen_dismissed" == "true" && "$second_screen_dismissed" == "true" && "$third_screen_dismissed" == "true" ]]; then
    echo "✅ All three privacy screens successfully dismissed"
    close_microsoft_powerpoint_app
    exit 0
elif [[ "$first_screen_dismissed" == "true" && "$second_screen_dismissed" == "true" ]]; then
    echo "⚠ First two screens dismissed, but third screen may still be showing"
    close_microsoft_powerpoint_app
    exit 1
elif [[ "$first_screen_dismissed" == "true" ]]; then
    echo "⚠ First screen dismissed, but later screens may still be showing"
    close_microsoft_powerpoint_app
    exit 1
else
    echo "⚠ Privacy modal handling completed (timeout reached)"
    echo "Note: Modals may have been dismissed or may not have appeared"
    close_microsoft_powerpoint_app
    exit 0
fi
