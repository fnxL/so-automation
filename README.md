# SO automation tool

This is a RPA bot that can run locally and automate complex predefined end-to-end workflows for SO creation in SAP for different customers.

> Disclaimer: No data ever leaves your local computer adhering to privacy concerns.

## Configuration

`sotool.json` file must be present either in the same directory of the program or in the user's home directory `C:\Users\<YourUserName>` in Windows.

Please take a look at the `sotool_example.json` file for an example.

You can also specify the configuration path using the environment variable `SOTOOL_CONFIG`.
`export SOTOOL_CONFIG=path/to/sotool.json`

## Notes

Before running the tool:

- Make sure you have set the required fields correctly in the `sotool.json` and placed it at correct directory.
- Make sure you have already logged into the SAP with your credentials.
- Only Outlook (Classic) is supported at the moment, so have it opened and logged in with your credentials.
- Keep only one window of Outlook open at a time.
