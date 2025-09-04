$exclude = @("venv", "botPython.zip", "#material", "json", "downloads", "teste")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "botPython.zip" -Force