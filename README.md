# SO automation tool

This is a RPA bot that can run locally and automate complex predefined end-to-end workflows for SO creation in SAP for different customers.

> Disclaimer: No data ever leaves your local computer adhering to privacy concerns.

## Configuration

There must be a `sotool.json` file present either in the same directory of the program or in the user's home directory (C:\Users\<YourUserName>).

Please take a look at the `sotool_example.json` file for an example.

You can also specify the configuration path using environment variable `SOTOOL_CONFIG`.
`export SOTOOL_CONFIG=path/to/sotool.json`
