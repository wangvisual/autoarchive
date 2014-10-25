#!/bin/sh

zip="zip -r"
AllFiles="content locale defaults chrome.manifest icon.png install.rdf bootstrap.js"
version="$(grep "<em:version>" install.rdf | sed -e 's/.*<em:version>\(.*\)<\/em:version>.*/\1/g')"

fileName="../beta/awsomeAutoArchive-${version}-tb.xpi"

rm -f ${fileName}
${zip} ${fileName} ${AllFiles}
