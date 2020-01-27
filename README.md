# EGW Library PowerPoint sidebar addin

This plugin searches the egw.hu database for quotes.

# Developement

- Run `npm install`
- For running the plugin on a live dev server, run `npm start` in the project root.
- For debugging the sidebar in the desktop PowerPoint use IEChooser.exe or F12Chooser.exe `c:\windows\SysWOW64\F12\IEChooser.exe`
- For building, run `npm run build`

# Deploying to the AppSource Office Store

1. after running `npm run build` upload the contents of the dist folder to the web server that'll host your add-in. You can use any type of web server or web hosting service to host your add-in.
2. Open the add-in's manifest file, located in the root directory of the project (manifest.xml). Replace all occurrences of https://localhost:3000 with the URL of the web application that you deployed to a web server in the previous step.
3. Choose the method you'd like to use to [deploy your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish), and follow the instructions to publish the manifest file.

[source](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish-add-in-vs-code)
