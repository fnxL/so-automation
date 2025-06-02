# SO automation tool

This is a RPA bot that can run locally and automate complex predefined end-to-end workflows for SO creation in SAP for different customers.

> Disclaimer: No data ever leaves your local computer adhering to privacy concerns.

## Configuration

`sotool.json` file must be present at one of the following:
- Same directory of the program
- In the user's home directory `C:\Users\<YourUserName>` in Windows
- In the user's `Downloads` directory
- In the user's `Desktop` directory

Please take a look at the `sotool_example.json` file for an example.

You can also specify the configuration path using the environment variable `SOTOOL_CONFIG`.

```sh
export SOTOOL_CONFIG=path/to/sotool.json
```

## Notes

Before running the tool:

- Ensure proper config in the `sotool.json` and placed it at one of the above locations.
- Make sure you have already logged onto SAP with your credentials.
- Only Outlook (Classic) is supported at the moment, so have it opened and logged in with your credentials.
- Keep only one window of Outlook open at a time.
- Close all open excel workbooks to ensure smooth operation.

## Requirements

- Python 3.13 or higher
- uv package manager

## Installation

1. Clone the repository using below command or download the zip file
```sh
git clone https://github.com/fnx-io/so-automation.git
```
2. Run the application
```sh
cd so-automation && uv sync --native-tls
uv run sotool
```
### Build instructions

You can also build the application using PyInstaller to obtain a standalone executable.

```sh
uv sync --native-tls
uv run pyinstaller --onefile --windowed --name sotool_app main.py
```
Find the executable in the `dist` directory.
```sh
cd dist && ./sotool_app.exe
```
