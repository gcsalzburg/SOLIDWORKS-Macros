# SOLIDWORKS-Macros
Collection of macros to support SolidWorks development

General notes:

- .swp files are for direct use as a Macro in SolidWorks
- .bas files are an export - to enable source code to be read and checked using GIT

To use a macro, in SolidWorks customise a menu and create a new button / shortcut for a macro, linking to the SWP file.


## Toggle Background Colour

Toggle between gradient and plain background colour.
No need to define colours in the macro, just set them as per normal in System Options > Colors.
Useful if you want to work with a gradient colour as a background (easier on eyes) but need a white background for screenshots.

Usage example:
- Set plain background to white in Tools > Options > System Options > Colors
- Set gradient background to desired colour (e.g. gentle blue fade)
- Assign macro to toggle button on menu or keyboard shortcut
- Use button to toggle from gradient background to white before taking a screenshot for presentation sharing, etc...