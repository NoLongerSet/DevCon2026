Attribute VB_Name = "modSampleApp"
Option Compare Database
Option Explicit

Public Const APP_VERSION As String = "1.0.0"

Public Function GetAppVersion() As String
    GetAppVersion = APP_VERSION
End Function

' Run once from the Immediate Window: CreateVersionForm
Public Sub CreateVersionForm()
    Const FORM_NAME As String = "frmVersion"

    On Error Resume Next
    DoCmd.Close acForm, FORM_NAME
    DoCmd.DeleteObject acForm, FORM_NAME
    On Error GoTo 0

    Dim s As String
    Dim L As String
    L = vbCrLf

    ' Based on working Access export — all metadata preserved
    s = "Version =21" & L
    s = s & "VersionRequired =20" & L
    s = s & "PublishOption =1" & L
    s = s & "Begin Form" & L
    s = s & "    Caption =""GBLauncher Demo""" & L
    s = s & "    DividingLines = NotDefault" & L
    s = s & "    AllowDesignChanges = NotDefault" & L
    s = s & "    DefaultView =0" & L
    s = s & "    TabularFamily =0" & L
    s = s & "    PictureAlignment =2" & L
    s = s & "    DatasheetGridlinesBehavior =3" & L
    s = s & "    GridX =24" & L
    s = s & "    GridY =24" & L
    s = s & "    Width =4200" & L
    s = s & "    DatasheetFontHeight =11" & L
    s = s & "    ItemSuffix =5" & L
    s = s & "    DatasheetGridlinesColor =15263976" & L
    s = s & "    DatasheetFontName =""Aptos""" & L
    s = s & "    AllowDatasheetView =0" & L
    s = s & "    FilterOnLoad =0" & L
    s = s & "    ShowPageMargins =0" & L
    s = s & "    DisplayOnSharePointSite =1" & L
    s = s & "    DatasheetAlternateBackColor =15921906" & L
    s = s & "    DatasheetGridlinesColor12 =0" & L
    s = s & "    FitToScreen =1" & L
    s = s & "    DatasheetBackThemeColorIndex =1" & L
    s = s & "    BorderThemeColorIndex =3" & L
    s = s & "    ThemeFontIndex =1" & L
    s = s & "    ForeThemeColorIndex =0" & L
    s = s & "    AlternateBackThemeColorIndex =1" & L
    s = s & "    AlternateBackShade =95.0" & L
    s = s & "    NoSaveCTIWhenDisabled =1" & L
    ' Default label style
    s = s & "    Begin" & L
    s = s & "        Begin Label" & L
    s = s & "            BackStyle =0" & L
    s = s & "            TextFontFamily =0" & L
    s = s & "            FontSize =11" & L
    s = s & "            BorderColor =8355711" & L
    s = s & "            ForeColor =6710886" & L
    s = s & "            FontName =""Aptos""" & L
    s = s & "            GridlineColor =10921638" & L
    s = s & "            ThemeFontIndex =1" & L
    s = s & "            BackThemeColorIndex =1" & L
    s = s & "            BorderThemeColorIndex =0" & L
    s = s & "            BorderTint =50.0" & L
    s = s & "            ForeThemeColorIndex =0" & L
    s = s & "            ForeTint =60.0" & L
    s = s & "            GridlineThemeColorIndex =1" & L
    s = s & "            GridlineShade =65.0" & L
    s = s & "        End" & L
    ' Default textbox style
    s = s & "        Begin TextBox" & L
    s = s & "            AddColon = NotDefault" & L
    s = s & "            FELineBreak = NotDefault" & L
    s = s & "            TextFontFamily =0" & L
    s = s & "            BorderLineStyle =0" & L
    s = s & "            LabelX =-1800" & L
    s = s & "            FontSize =11" & L
    s = s & "            BorderColor =10921638" & L
    s = s & "            ForeColor =4210752" & L
    s = s & "            FontName =""Aptos""" & L
    s = s & "            AsianLineBreak =1" & L
    s = s & "            GridlineColor =10921638" & L
    s = s & "            BackThemeColorIndex =1" & L
    s = s & "            BorderThemeColorIndex =1" & L
    s = s & "            BorderShade =65.0" & L
    s = s & "            ThemeFontIndex =1" & L
    s = s & "            ForeThemeColorIndex =0" & L
    s = s & "            ForeTint =75.0" & L
    s = s & "            GridlineThemeColorIndex =1" & L
    s = s & "            GridlineShade =65.0" & L
    s = s & "        End" & L
    ' Detail section
    s = s & "        Begin Section" & L
    s = s & "            Height =1500" & L
    s = s & "            Name =""Detail""" & L
    s = s & "            AutoHeight =1" & L
    s = s & "            AlternateBackColor =15921906" & L
    s = s & "            AlternateBackThemeColorIndex =1" & L
    s = s & "            AlternateBackShade =95.0" & L
    s = s & "            BackThemeColorIndex =1" & L
    s = s & "            Begin" & L
    ' Version label
    s = s & "                Begin Label" & L
    s = s & "                    OverlapFlags =85" & L
    s = s & "                    Left =240" & L
    s = s & "                    Top =360" & L
    s = s & "                    Width =1440" & L
    s = s & "                    Height =300" & L
    s = s & "                    Name =""Label0""" & L
    s = s & "                    Caption =""Version #:""" & L
    s = s & "                End" & L
    ' Version textbox
    s = s & "                Begin TextBox" & L
    s = s & "                    Enabled = NotDefault" & L
    s = s & "                    Locked = NotDefault" & L
    s = s & "                    AllowAutoCorrect = NotDefault" & L
    s = s & "                    DecimalPlaces =0" & L
    s = s & "                    OverlapFlags =85" & L
    s = s & "                    IMESentenceMode =3" & L
    s = s & "                    Left =1800" & L
    s = s & "                    Top =360" & L
    s = s & "                    Width =2100" & L
    s = s & "                    Height =300" & L
    s = s & "                    Name =""Text3""" & L
    s = s & "                    ControlSource =""=GetAppVersion()""" & L
    s = s & "                    ShowDatePicker =0" & L
    s = s & "                End" & L
    s = s & "            End" & L
    s = s & "        End" & L
    s = s & "    End" & L
    s = s & "End" & L

    ' Write as UTF-16 LE (Access requires this encoding)
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\frmVersion.txt"

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "Unicode"
    stream.Open
    stream.WriteText s
    stream.SaveToFile tempPath, 2
    stream.Close

    Application.LoadFromText acForm, FORM_NAME, tempPath
    Kill tempPath

    MsgBox "Form created. Close and reopen to test.", vbInformation
End Sub
