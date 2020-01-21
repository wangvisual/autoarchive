#!/bin/sh

zip="zip -r"
AllFiles="content locale defaults chrome.manifest manifest.json icon.png bootstrap.js"
version="$(grep '"version"' manifest.json | sed -e 's/.*:.*"\(.*\)".*/\1/g')"

fileName="../beta/awsomeAutoArchive-${version}-tb.xpi"

rm -f ${fileName}
${zip} ${fileName} ${AllFiles}
