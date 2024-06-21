$exclude = @("venv", "bot-sicalc-id.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "bot-sicalc-id.zip" -Force