#!/bin/zsh

curl -s -L "https://github.com/objective-see/KnockKnock/releases/download/v2.3.0/KnockKnock_2.3.0.zip" -o /tmp/KnockKnock.zip
cd /tmp && unzip -q KnockKnock.zip
xattr -r -d com.apple.quarantine KnockKnock.app/Contents/MacOS/KnockKnock
rm /tmp/KnockKnock.zip
KnockKnock.app/Contents/MacOS/KnockKnock -whosthere -pretty
