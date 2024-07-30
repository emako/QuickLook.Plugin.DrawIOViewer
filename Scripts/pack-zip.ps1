Remove-Item ..\QuickLook.Plugin.DrawIOViewer.qlplugin -ErrorAction SilentlyContinue

$files = Get-ChildItem -Path ..\QuickLook.Plugin\QuickLook.Plugin.DrawIOViewer\bin\Release\ -Exclude *.pdb,*.xml
Compress-Archive $files ..\QuickLook.Plugin.DrawIOViewer.zip
Move-Item ..\QuickLook.Plugin.DrawIOViewer.zip ..\QuickLook.Plugin.DrawIOViewer.qlplugin