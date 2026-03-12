# AwakeKeeper

AwakeKeeper is a small Windows tray app whose job is simple: keep your computer awake when you need it to stay awake, and get out of the way when you do not.

It is for situations like:
- a dashboard you need visible all afternoon
- a presentation laptop that should not decide to nap mid-meeting
- a machine that interprets "quietly reading something" as "abandoned by civilization"

This is not enterprise software. It is a practical little tray utility with just enough polish to be useful and just enough personality to admit what it is doing.

## What It Does

AwakeKeeper watches your system idle time and, once you have been idle for long enough, uses one of a few different methods to convince Windows that now is not the time for sleep.

It lives in the system tray and shows:
- whether the keeper is on or off
- your current idle time
- the configured idle threshold
- the check interval
- the active keep-awake method
- the selected profile
- whether the Windows sleep override is active
- the last time an action ran
- the last recorded error

The app stores its configuration in your roaming app data folder, so it remembers your settings between runs.

## How It Works

AwakeKeeper loops in the background every few seconds and checks how long it has been since your last mouse or keyboard input.

If:
- the app is enabled, and
- your idle time is greater than or equal to the configured threshold

then it runs the currently selected keep-awake method.

If you become active again and the current method is `preventsleep`, it clears the Windows sleep override so the system can behave normally.

## Keep-Awake Methods

AwakeKeeper supports three methods. They do not all behave the same way.

### 1. `preventsleep`

This is the preferred method.

It calls the Windows API `SetThreadExecutionState(...)` to tell Windows the system and display should remain awake.

Pros:
- does not physically move the mouse
- does not press keys
- usually the cleanest option

Cons:
- depends on Windows honoring the execution state request

Best for:
- desk setups
- dashboards
- passive keep-awake behavior

### 2. `scrolllock`

This uses `WScript.Shell.SendKeys` to tap Scroll Lock on and off.

Pros:
- useful as an alternate compatibility mode

Cons:
- can visibly affect the Scroll Lock indicator on some systems
- may fail in some environments
- falls back to mouse jiggle if the required shell integration is unavailable

Best for:
- situations where `preventsleep` is not doing the job

### 3. `mousejiggle`

This physically moves the cursor by one pixel and puts it back.

Pros:
- simple
- effective in many cases

Cons:
- it actually moves the pointer
- can be annoying if you are right on the edge of being considered idle

Best for:
- fallback behavior
- presentation machines
- environments where the other methods are blocked or ignored

## Profiles

The app includes two built-in profiles:

### Dashboard Mode

Designed for passive viewing situations.

Settings:
- idle threshold: `30` seconds
- check interval: `5` seconds
- method: `preventsleep`

### Presentation Mode

Designed for more aggressive "do not fall asleep in public" behavior.

Settings:
- idle threshold: `60` seconds
- check interval: `10` seconds
- method: `mousejiggle`

If you manually switch methods after selecting a profile, the profile becomes effectively custom behavior even if the tray still started from one of the presets.

## Tray Menu Guide

Everything happens from the system tray icon.

Menu items:

- `Start` / `Stop`
  Turns the keeper on or off.

- `Run Now`
  Immediately runs the current keep-awake method. This is only available while the keeper is on.

- `Dashboard Mode`
  Applies the dashboard-friendly preset.

- `Presentation Mode`
  Applies the presentation-friendly preset.

- `Use PreventSleep`
  Switches the app to the Windows execution-state method.

- `Use ScrollLock`
  Switches the app to the Scroll Lock pulse method.

- `Use MouseJiggle`
  Switches the app to cursor movement mode.

- `Exit`
  Shuts the app down and clears the active sleep override if needed.

## Configuration

AwakeKeeper stores config here:

```text
%APPDATA%\AwakeKeeper\config.json
```

It also writes a log here:

```text
%APPDATA%\AwakeKeeper\awake_keeper.log
```

If the config file becomes invalid, the app will attempt to move it aside as a backup and fall back to defaults instead of dying dramatically.

### Config Fields

Example:

```json
{
  "idle_threshold_seconds": 30,
  "check_interval_seconds": 5,
  "method": "preventsleep",
  "profile": "Dashboard Mode",
  "start_enabled": true
}
```

Field meanings:

- `idle_threshold_seconds`
  How long the system must be idle before AwakeKeeper acts.

- `check_interval_seconds`
  How often the worker loop checks idle state.

- `method`
  One of `preventsleep`, `scrolllock`, or `mousejiggle`.

- `profile`
  A label for the current preset or custom behavior.

- `start_enabled`
  Whether the keeper should start in the enabled state when the app launches.

## Installation / Running

### Option 1: Run from Python

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Run the app:

```powershell
python awake_keeper.py
```

### Option 2: Build the executable

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Build:

```powershell
pyinstaller awake_keeper.spec
```

The packaged executable ends up in:

```text
dist\awake_keeper.exe
```

## Dependencies

The app uses:
- `pystray` for the tray icon and menu
- `Pillow` for drawing the tray icon
- `pywin32` for Windows interaction
- `PyInstaller` for packaging

## Notes and Limitations

- This is Windows-only.
- `mousejiggle` is intentionally a little rude.
- `scrolllock` is slightly more polite, but also a little weird.
- `preventsleep` is the method you probably want most of the time.
- If something goes wrong, check the log file before assuming your laptop has developed opinions.

## Troubleshooting

### The app does not seem to do anything

Check:
- that it is actually `Start`ed
- that the idle threshold has been reached
- that the selected method is appropriate for your system
- `%APPDATA%\AwakeKeeper\awake_keeper.log` for errors

### `scrolllock` does not work

That can happen. The app will try to fall back to `mousejiggle` if shell automation is unavailable or the key pulse fails.

### My editor says `win32api` cannot be resolved from source

That is usually an editor analysis complaint, not a real runtime problem. `pywin32` ships compiled extension modules, so some language servers complain even when the import is valid.

### The cursor moved and I did not enjoy that

You were using `mousejiggle`.

Use `preventsleep` instead.

## Why This Exists

Because Windows sometimes interprets "I am quietly monitoring something important" as "please turn the screen off and undo my productivity."

AwakeKeeper respectfully disagrees.
