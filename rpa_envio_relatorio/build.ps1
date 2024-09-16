$exclude = @("venv", "desafio02.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "desafio02.zip" -Force