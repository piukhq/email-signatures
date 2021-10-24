# email-signatures

Simple JXA Script for adding email-signatures into Microsoft Outlook for Mac after asking the user some basic questions about themselves

## Usage

- Open `Script Editor.app`
- Change the `AppleScript` dropdown to `JavaScript`
- Copy the contents of `gen.js` into `Script Editor.app` and make any required changes to the `template`
    - Ensure you keep the `#name#`, `#title#`, and `#phone#` sections as these are used later during string substitution.
- Save the script as an Application and ship it to your users
    - Optional: Replace the icon for the app with the instructions [here](https://support.apple.com/en-gb/guide/mac-help/mchlp2313/mac)
