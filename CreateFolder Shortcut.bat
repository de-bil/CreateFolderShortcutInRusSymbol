@echo off
REM Поддержка кириллицы в консоли
chcp 65001 >nul

echo Проверка и создание ярлыка через PowerShell...
echo.

set "FULL_PATH=*сюда вставить uml путь до конкретной папки например *\\192.168.126.253\Мое сетевое хранилище\Новая папка* Ярлык новой папки будет создан с название указанным ниже "

REM Выполняем проверку и создание ярлыка через PowerShell
powershell -ExecutionPolicy Bypass -Command "$fullPath='%FULL_PATH%'; $desktopPath=[Environment]::GetFolderPath('Desktop'); $shortcutFullPath=Join-Path $desktopPath '*Имя ярлыка на кириллице в том числе*.lnk'; if(Test-Path $shortcutFullPath) { $shortcut=(New-Object -ComObject WScript.Shell).CreateShortcut($shortcutFullPath); if(($shortcut.TargetPath.Split('\\')[-1].Trim().ToLower() -eq 'explorer.exe') -and ($shortcut.Arguments.Trim() -eq $fullPath.Trim())) { Write-Host 'Ярлык уже настроен верно, замена не требуется.'; Exit } else { Write-Host 'Обнаружено несоответствие целей ярлыков. Удаляю существующий ярлык и создаю новый...'; Remove-Item $shortcutFullPath -Force } } else { Write-Host 'Ярлык не найден, создаём новый...' }; $WshShell=New-Object -ComObject WScript.Shell; $newShortcut=$WshShell.CreateShortcut($shortcutFullPath); $newShortcut.TargetPath='explorer.exe'; $newShortcut.Arguments=$fullPath; $newShortcut.WorkingDirectory='C:\\'; $newShortcut.Description='Сетевая папка'; $newShortcut.IconLocation='shell32.dll,4'; $newShortcut.Save(); Write-Host 'Ярлык успешно создан или обновлён!'"


echo.

