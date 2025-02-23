$exclude = @("venv", "Bot_coleta_LME.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "Bot_coleta_LME.zip" -Force