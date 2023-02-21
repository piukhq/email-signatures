# email-signatures

Simple JXA Script for adding email-signatures into Microsoft Outlook for Mac after asking the user some basic questions about themselves

## Usage

- Open `Script Editor.app`
- Change the `AppleScript` dropdown to `JavaScript`
- Copy the contents of `gen.js` into `Script Editor.app` and make any required changes to the `template`
    - Ensure you keep the `#name#`, `#title#`, and `#phone#` sections as these are used later during string substitution.
- Save the Application into the `./build` directory and with **NO CODE SIGNING** as `Bink Signature Generator.app`
- From the `./build` directory, execute `gon sign.json`
- This should eventually return a `Bink Signature Generator.zip` file
- Extract this zip file, rename the resulting directory with a `.app` suffix.
- ???
- Profit!
