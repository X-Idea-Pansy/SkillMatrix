Option Explicit
Sub ActiverEvents()
    Rem : Activate events
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub DésactiverEvents()
    Rem : Desactivate events
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub
''''
Sub AlleràWs(ws As Worksheet)
    Rem : Activate worksheet
    ws.Visible = True
    ws.Activate
End Sub
Sub RetourTOMenu()
    Rem : Go back to menu
     Call ActiverEvents
     ActiveWindow.SelectedSheets.Visible = False
     Dim wsPageAccueil As Worksheet
     Set wsPageAccueil = ThisWorkbook.Sheets("Menu")
     Call AlleràWs(wsPageAccueil)
End Sub
Sub CertifTOEmployeeData()
     Call ActiverEvents
     ActiveWindow.SelectedSheets.Visible = False
     Dim wsEmployeeData As Worksheet, wsCertifAluMachL3 As Worksheet
     Set wsCertifAluMachL3 = ThisWorkbook.Sheets("Employees certif L3XX")
     Set wsEmployeeData = ThisWorkbook.Sheets("Employees DATA")
     Call AlleràWs(wsEmployeeData)
End Sub
''''
Sub AccueilTOOperatorVisual()
     Call ActiverEvents
     Dim wsSkillMatrix As Worksheet, wsPageAccueil As Worksheet, ws As Worksheet
     Set wsPageAccueil = ThisWorkbook.Sheets("Menu")
     For Each ws In ThisWorkbook.Worksheets
        If InStr(1, CleanString(ws.Name), CleanString("Skill Matrix")) <> 0 And InStr(1, CleanString(ws.Name), CleanString("BLANK")) = 0 Then
            Set wsSkillMatrix = ws
            Call AlleràWs(wsSkillMatrix)
        End If
    Next ws
End Sub
Sub AccueilTOEmployeeData()
     Call ActiverEvents
     Dim wsEmployeeData As Worksheet, wsPageAccueil As Worksheet
     Set wsPageAccueil = ThisWorkbook.Sheets("Menu")
     Set wsEmployeeData = ThisWorkbook.Sheets("Employees DATA")
     Call AlleràWs(wsEmployeeData)
End Sub
Sub AccueilTODVP()
     Call ActiverEvents
     Dim wsDVP As Worksheet, wsPageAccueil As Worksheet
     Set wsPageAccueil = ThisWorkbook.Sheets("Menu")
     Set wsDVP = ThisWorkbook.Sheets("Developer")
     Call AlleràWs(wsDVP)
End Sub
Sub AccueilTOFind()
     Call ActiverEvents
     Dim wsFind As Worksheet, wsPageAccueil As Worksheet
     Set wsPageAccueil = ThisWorkbook.Sheets("Menu")
     Set wsFind = ThisWorkbook.Sheets("Find an Employee")
     Call AlleràWs(wsFind)
End Sub
Sub ProjectsTODVP()
     Call ActiverEvents
     Dim wsDVP As Worksheet, wsProjects As Worksheet
     ActiveWindow.SelectedSheets.Visible = False
     Set wsProjects = ThisWorkbook.Sheets("Projects")
     Set wsDVP = ThisWorkbook.Sheets("Developer")
     Call AlleràWs(wsDVP)
End Sub
Sub DVPToProjects()
     Call ActiverEvents
     Dim wsDVP As Worksheet, wsProjects As Worksheet
     Set wsProjects = ThisWorkbook.Sheets("Projects")
     Set wsDVP = ThisWorkbook.Sheets("Developer")
     Call AlleràWs(wsProjects)
End Sub
Sub EmployeeDataToCertif()
     Call ActiverEvents
     Dim wsEmployeeData As Worksheet, wsCertif As Worksheet, wsPageAccueil As Worksheet, ws As Worksheet
     Set wsPageAccueil = ThisWorkbook.Sheets("Menu")

     Set wsEmployeeData = ThisWorkbook.Sheets("Employees DATA")
     For Each ws In ThisWorkbook.Worksheets
        If InStr(1, CleanString(ws.Name), CleanString("Employees certif")) <> 0 And InStr(1, CleanString(ws.Name), CleanString("BLANK")) = 0 Then
            Set wsCertif = ws
            Call AlleràWs(wsCertif)
        End If
     Next ws
     Call AlleràWs(wsCertif)
End Sub

''''

Sub LineORProject()
    frmProjectORLine.Show
End Sub
Sub AddOrRemoveanEmployee()
    frmAddRemoveEmployee.Show
End Sub
Sub ModifyanEmployee()
    frmModifyEmployee.Show
End Sub
Sub AddaCertif()
    frmAddCertif.Show
End Sub
Sub SelectTeam()
    frmTeamSelect.Show
End Sub


''''
Sub GoSimiliFullScreen(ws As Worksheet)
    ws.Activate
    With Application
        .DisplayFormulaBar = False       ' Hide formula bar
    End With
    With ActiveWindow
        .DisplayHeadings = False         ' Hide row and column headings
    End With
End Sub
Sub GoFullScreen()
    With Application
        .DisplayFullScreen = True        ' Full screen mode
        .DisplayFormulaBar = False       ' Hide formula bar
        .DisplayStatusBar = False        ' Hide status bar
        .CommandBars("Ribbon").Visible = False  ' Hide ribbon
    End With
End Sub

Sub ExitFullScreen()
    With Application
        .DisplayFullScreen = False
        .DisplayFormulaBar = False
        .DisplayStatusBar = True
        .CommandBars("Ribbon").Visible = True
    End With
End Sub


Sub RefreshProjects()
    ActiveWorkbook.RefreshAll
End Sub

