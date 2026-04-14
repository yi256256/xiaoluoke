# xiaoluoke

Windows WinForms auto clicker for XiaoLuoke-style front-end automation.

## Download

Download a ready-to-run package from one of these files:

- [release/AutoClicker-windows.zip](release/AutoClicker-windows.zip)
- [release/AutoClicker.exe](release/AutoClicker.exe)

If GitHub previews the file page, use the download button to save it locally.

## Run

1. Download `release/AutoClicker.exe`.
2. Double-click to start it on Windows.
3. If the target app or game is running as administrator, run `AutoClicker.exe` as administrator too.
4. For full-screen apps, set a start delay in the tool, click `Start`, then switch back to the target app before the countdown ends.

## Features

- Mouse click actions and keyboard key actions
- Configurable interval per action
- Start delay for full-screen apps
- Auto minimize on start
- Global hotkeys for capture, start/pause, and stop
- Config persisted to `%AppData%\\AutoClicker\\config.json`

## Notes

- This build is front-end input automation. The target window should be active when actions run.
- Some exclusive full-screen games may still reject simulated input depending on how they handle keyboard and mouse events.
- The app is built for Windows and requires .NET Framework 4.x runtime available on the system.
