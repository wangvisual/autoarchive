set PATH=c:\Program Files (x86)\7-Zip;c:\Program Files\7-Zip;d:\Program Files (x86)\7-Zip;d:\Program Files\7-Zip
set zip=7z.exe a -tzip -mx1 -r
set AllFiles=content locale defaults chrome.manifest icon.png install.rdf bootstrap.js
del awsomeAutoArchive-*-tb.xpi
%zip% awsomeAutoArchive-0.5-tb.xpi %AllFiles% -xr!.svn
