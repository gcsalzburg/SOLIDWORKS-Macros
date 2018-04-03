Attribute VB_Name = "Toggle_Background_Colour1"
' ******************************************************************************
' Solidworks Macro to toggle between gradient background and plain background
' Author: George Cave
' Date: April 2018
' www.designedbycave.co.uk
'
' Usage example:
'  - Set plain background to white in Tools > Options > System Options > Colors
'  - Set gradient background to desired colour (e.g. gentle blue fade)
'  - Assign macro to toggle button on menu or keyboard shortcut
'  - Use button to toggle from gradient background to white before taking a screenshot for presentation sharing, etc...
'
' Reference documents:
'  - http://help.solidworks.com/2016/english/api/swconst/so_colors.htm
'  - http://help.solidworks.com/2016/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks~SetUserPreferenceIntegerValue.html
'  - http://help.solidworks.com/2016/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swColorsBackgroundAppearance_e.html

' ******************************************************************************

Option Explicit
Dim swApp As SldWorks.SldWorks
Dim Model As SldWorks.ModelDoc2

Dim BgState As swColorsBackgroundAppearance_e

Sub Main()

Set swApp = Application.SldWorks
Set Model = swApp.ActiveDoc


' Fetch current background state
BgState = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swColorsBackgroundAppearance)

' Toggle from Gradient <-> Plain
' If image or document scene are set, then no change will occur
If BgState = swColorsBackgroundAppearance_Gradient Then
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swColorsBackgroundAppearance, swColorsBackgroundAppearance_Plain
ElseIf BgState = swColorsBackgroundAppearance_Plain Then
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swColorsBackgroundAppearance, swColorsBackgroundAppearance_Gradient
End If

' Force rebuild to ensure new background is shown
Model.ForceRebuild3 (True)

End Sub
