# Macro Helper Tool 2

A lightweight macro reminder tool that uses text-to-speech to call out timed alerts for recurring tasks like building workers and other repetitive tasks that you need some help to drill into your memory routines.

It's tailored for Starcraft 2, but could be adapted to be more generic with a little bit of effort.

## Features

- **Race-specific presets** — Built-in timers for Zerg, Protoss, and Terran macro tasks
- **Custom actions (includes Paste function)** — Create your own reminders with flexible timing expressions (initial delays, repeating intervals, finite sequences)
- **Build profiles** — Save and switch between multiple build-specific action sets per race
- **Text-to-speech** — Spoken alerts with configurable voice, rate, volume, and randomization
- **Compact overlay** — Frameless, always-on-top draggable window designed to sit alongside your game
- **System tray** — Minimize to tray to keep it out of the way
- **Per-race theming** — Customizable accent and font colors for each race

## Requirements

- Python 3.10+
- Windows (uses Windows TTS voices via `pyttsx3`)

## Setup

```bash
pip install -r requirements.txt
python main.py
```

## Building a Standalone Executable

```bash
compile.bat
```

This creates `dist/sc2_alarm_app.exe` using PyInstaller.

## Configuration

Settings are stored in `config.json` next to the executable/script and are saved automatically.

## License

MIT
