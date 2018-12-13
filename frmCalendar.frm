VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Kalender"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3375
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'=============================================================================================================================
' Index
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 1.  UserForm_Initialize    -   Initialisierung UserForm
' 2.  pFillCmbMonth          -   Dropdown mit Monatsnamen füllen
' 3.  cmbMonth_Change        -   Refresh bei Monatswechsel
' 4.  cmbMonth_Exit          -   Refresh bei Monatswechsel
' 5.  udMonth_Change         -   Gleichschaltung Monat mit UpDown-Schalter (Buddy)      REMARKED IN VERS. 1.1 / 05.10.11
' 6.  txtYear_Change         -   Refresh bei Jahreswechsel
' 7.  setDataLabels          -   Tage in Raster verteilen
' 8.  pClearDataLabels       -   Raster zurücksetzen
' 9.  pReturnDate            -   Rückgabewert in globale Variable schreiben
' 10. pSetNewMonth           -   Raster mit neuem Monat befüllen bei Klick auf "graue" Datumsfelder
' 11 - 15.                   -   Schaltflächen Monat und Jahr + und -
' 16 - 56.                   -   Click-Events auf Labels (Kein Errorhandling)
'======================================================================================================================

Private Declare Function FindWindow Lib "user32" Alias _
      "FindWindowA" (ByVal lpClassName As String, ByVal _
      lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias _
      "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex _
      As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
      "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex _
      As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal _
      hwnd As Long) As Long
Private Const GWL_STYLE As Long = -16
Private Const WS_SYSMENU As Long = &H80000
Private hwndForm As Long
Private bCloseBtn As Boolean


'=============================================================================================================================
'Prozeduren
'=============================================================================================================================
' 1. UserForm_Initialize
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
On Error GoTo err_initialize
    g_bolInitialize = True              ' Während Initialisierung werden automatisch Refreshs (Change- und Exit-Events) unterbunden
        txtYear.Text = Year(Date)       ' Initialwert Jahr
        Call pFillCmbMonth              ' Aufruf Befüllung Dropdown "Monatsnamen"
        fSetMonthText (Month(Date))     ' Übersetzung der Inhalte des Dropdowns Monatsnamen
        Call setDataLabels(fChangeStrToInt(cmbMonth.Text), txtYear.Text)
    g_bolInitialize = False
Exit Sub
            'Info2
    Dim i As Integer

    If Val(Application.Version) >= 9 Then                           ' Sonderbereich zum Deaktivieren der Funktionen in der Titelleiste
        hwndForm = FindWindow("ThunderDFrame", Me.Caption)
    Else
        hwndForm = FindWindow("ThunderXFrame", Me.Caption)
    End If
    
    bCloseBtn = False
    SET_USERFORM_STYLE

' Errorhandling
err_initialize:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'UserForm_Initialize' in 'frmCalendar'. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)    ' nicht mit rotem X schließen

    If CloseMode = vbFormControlMenu Then
        Exit Sub
        Cancel = True
    End If
    
End Sub



Private Sub SET_USERFORM_STYLE()                                                      '

    Dim frmStyle As Long
    
    If hwndForm = 0 Then Exit Sub
    
    frmStyle = GetWindowLong(hwndForm, GWL_STYLE)
    
    If bCloseBtn Then
      frmStyle = frmStyle Or WS_SYSMENU
    Else
      frmStyle = frmStyle And Not WS_SYSMENU
    End If
    
    SetWindowLong hwndForm, GWL_STYLE, frmStyle
    DrawMenuBar hwndForm
    
End Sub



'=============================================================================================================================
' 2. pFillCmbMonth  -  Dropdown mit Monatsnamen füllen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub pFillCmbMonth()
On Error GoTo err_pFillCmbMonth
    cmbMonth.AddItem "Januar"
    cmbMonth.AddItem "Februar"
    cmbMonth.AddItem "März"
    cmbMonth.AddItem "April"
    cmbMonth.AddItem "Mai"
    cmbMonth.AddItem "Juni"
    cmbMonth.AddItem "Juli"
    cmbMonth.AddItem "August"
    cmbMonth.AddItem "September"
    cmbMonth.AddItem "Oktober"
    cmbMonth.AddItem "November"
    cmbMonth.AddItem "Dezember"
Exit Sub

' Errorhandling
err_pFillCmbMonth:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'err_pFillCmbMonth' in 'frmCalendar'. Dropdown konnte nicht befüllt werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
    
End Sub

'=============================================================================================================================
' 3. cmbMonth_Change    -  Refresh bei Monatswechsel
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Anmerkungen:      Bei Monatswechsel Refresh anstossen (ausser bei Initialisierung und beim Wechsel durch Klick auf "grauen" Datumsfeldern
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cmbMonth_Change()
On Error GoTo err_cmbMonthChange
    If g_bolInitialize = False And g_bolMonthChange = False Then
        If fPlaus(cmbMonth.Text, 1) = True Then                     ' Wenn Monatsname korrekt eingegeben wurde und...
            If lblk2d1.Caption <> "" Then                           ' das Labelraster Einträge enthält...
                Call pClearDataLabels                               ' Aufruf Prozedur um alle Einträge zu entfernen
            End If
            Call setDataLabels(cmbMonth.Text, txtYear.Text, fChangeStrToInt(cmbMonth.Text))     ' Falls Raster nicht beschrieben, Aufruf für die Befüllung des Monatsrasters
        End If
    End If
Exit Sub

' Errorhandling
err_cmbMonthChange:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'cmbMonth_Change' in 'frmCalendar'. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub
'=============================================================================================================================
' 4. cmbMonth_Exit      -  Refresh bei Monatswechsel
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Anmerkungen:      Bei Verlassen des Monatsfeldes Refresh anstossen (ausser bei Initialisierung und beim Wechsel durch Klick auf "grauen" Datumsfeldern
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cmbMonth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo err_cmbMonthExit
    If g_bolInitialize = False And g_bolMonthChange = False Then
        Call pClearDataLabels
        If IsNumeric(cmbMonth.Text) = True Then
            Call setDataLabels(cmbMonth.Text, txtYear.Text)
        Else
            Call setDataLabels(cmbMonth.Text, txtYear.Text, fChangeStrToInt(cmbMonth.Text))
        End If
    End If
Exit Sub

' Errorhandling
err_cmbMonthExit:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'cmbMonth_Exit' in 'frmCalendar'. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub

'REMARKED IN VERSION 1.1 / 05.10.11
'=============================================================================================================================
' 5. udMonth_Change -   Gleichschaltung Monat mit UpDown-Schalter (Buddy)
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Anmerkungen:      Bei Änderung des Monats via UpDown-Schaltfläche Refresh anstossen
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Private Sub udMonth_Change()
'On Error GoTo err_udMonthChange
'    If lblk2d1.Caption <> "" Then
'        Call pClearDataLabels
'    End If
'    fSetMonthText (cmbMonth.Value)
'    If g_bolMonthChange = False Then Call setDataLabels(cmbMonth.Value, txtYear.Text, fChangeStrToInt(cmbMonth.Value))
'Exit Sub
'
'' Errorhandling
'err_udMonthChange:
'    MsgBox "Error " & Err.Number & " (" & Err.Description _
'        & ") in der Prozedur 'udMonth_Change' in 'frmCalendar'. Monatswechsel konnte nicht verarbeitet werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
'    End
'End Sub

'=============================================================================================================================
' 6. txtYear_Change  -   Refresh bei Jahreswechsel
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Anmerkungen:      Bei Änderung des Jahres Refresh anstossen (ausser bei Initialisierung und beim Wechsel durch Klick auf "grauen" Datumsfeldern
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub txtYear_Change()
On Error GoTo err_txtYearChange
    If g_bolInitialize = False And g_bolMonthChange = False Then
        If lblk2d1.Caption <> "" Then
            Call pClearDataLabels
        End If
        If g_bolMonthChange = False Then
            If Len(txtYear.Text) = 4 And IsNumeric(txtYear.Text) = True And txtYear.Text > "1900" Then
                Call setDataLabels(cmbMonth.Value, txtYear.Text, fChangeStrToInt(cmbMonth.Text))
            End If
        End If
    End If
Exit Sub

' Errorhandling
err_txtYearChange:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'txtYear_Change' in 'frmCalendar'. Jahreswechsel konnte nicht verarbeitet werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub

'======================================================================================================================
' 7. setDataLabels - Tage in Raster verteilen
'======================================================================================================================
Private Sub setDataLabels(dMonth, dYear, Optional varMonth As Variant = "")
' Deklarationen
Dim strLabel            As String       'Name des jeweiligen Label-Steuerelements
Dim intCounter          As Integer      'Laufnummer für Tageszahl (1 - 28/29/30/31)
Dim intWeekCounter      As Integer      'Zähler für die angezeigten Wochen (6)
Dim intDayCounter       As Integer      'Zähler für die aktuelle Tagesnummer innerhalb der Wochen (1 - 7 | 1 = Montag)
Dim datLastDayMonth     As Date         'Letzter Tag im Monat als Datum
Dim intLastDayMonth     As Integer      'Letzter Tag im Monat als Ganzzahl
Dim datFirstDayOfMonth  As Date         'Erster Tag im Monat als Datum
Dim intStartKW          As Date         'Erste Kalenderwoche des angezeigten Zeitraums
Dim datActiveDate       As Date         'Aktuell bearbeitetes Datum
Dim bolPostActiveMonth  As Boolean      'Schalter, welcher auf True gesetzt wird, wenn die angezeigten Daten bereits zu nächsten Monat gehören
Dim strKW               As String       'Name des Labels für die Kalenderwoche
' Zähler und Variablen für Vormonat
Dim intVormonat         As Integer      'Zähler für die Vearbeitung jener angezeigten Daten, welche zum Vormonat gehören
Dim intVormonatTag      As Integer      'Zähler für die Vearbeitung jener angezeigten Daten, welche zum Vormonat gehören

' Initialisierung
bolPostActiveMonth = False
If varMonth <> "" Then dMonth = varMonth    ' Falls bei Aufruf der Prozedur mittels der optionalen Variable "varMonth" ein Monat (Zahl) übergeben wurde, diese verwenden
intCounter = 1
intDayCounter = Weekday("01." & dMonth & "." & dYear, vbMonday)
intStartKW = "01." & dMonth & "." & dYear
intWeekCounter = 1
datLastDayMonth = fLastDayInMonth("01." & dMonth & "." & dYear)
intLastDayMonth = Mid(datLastDayMonth, 1, 2)
intVormonat = 0

' Verarbeitung Vormonat
For intVormonatTag = intDayCounter - 1 To 1 Step -1
        strLabel = "lblk1d" & intVormonatTag
        If dMonth <> 1 Then
            Me.Controls(strLabel).Caption = Mid(fLastDayInMonth("01." & dMonth - 1 & "." & dYear), 1, 2) - intVormonat
        Else
            Me.Controls(strLabel).Caption = Mid(fLastDayInMonth("01." & dMonth + 11 & "." & dYear - 1), 1, 2) - intVormonat
        End If
        Me.Controls(strLabel).ForeColor = &H808080
        intVormonat = intVormonat + 1
Next intVormonatTag


' Verarbeitung aktiver Monat
For intWeekCounter = 1 To 6                                         ' Übergeordnete Schlaufe für jede angezeigte Woche
    For intDayCounter = intDayCounter To 7                          ' Schlaufe für jeden Tag innerhalb einer Woche
        strLabel = "lblk" & intWeekCounter & "d" & intDayCounter    ' Identifizierung Label
        Me.Controls(strLabel).Caption = intCounter                  ' Beschriftung Label
        If bolPostActiveMonth = True Then Me.Controls(strLabel).ForeColor = &H808080    ' Wenn Datum ausserhalb Betrachtungszeitraum Schriftfarbe auf Grau setzen
        datActiveDate = DateSerial(dYear, dMonth, intCounter)
        If datActiveDate = Date Then                                ' Falls Datum = Heute Schriftart Rot und Fett
                Me.Controls(strLabel).Font.Bold = True
                Me.Controls(strLabel).ForeColor = 255
        End If
        
        intCounter = intCounter + 1
           
        ' Verarbeitung Folgemonat
        If intCounter = intLastDayMonth + 1 Then
            intCounter = 1              ' Reset Monatstag auf 1
            dMonth = dMonth + 1         ' Monat = Monat + 1
            If dMonth = 13 Then         ' Falls Monat = 13, dann Jahr erhöhen und Monatsnummer auf 1 (Januar) setzen
                dMonth = 1
                dYear = dYear + 1
            End If
            bolPostActiveMonth = True   ' Flag für Vearbeitung Folgemonat gesetzt.
        End If
        
    Next intDayCounter
    
    ' Kalenderwoche pro Woche ermitteln
    strKW = "lblKW" & intWeekCounter
    Me.Controls(strKW).Caption = fGetKW(datActiveDate)
    
    intDayCounter = 1
Next intWeekCounter

bolPostActiveMonth = False

Exit Sub

' Errorhandling
err_udMonthChange:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'setDataLabels' in 'frmCalendar'. Kalender konnte nicht aufgebaut werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub

'======================================================================================================================
' 8. pClearDataLabels - Raster zurücksetzen
'======================================================================================================================
Sub pClearDataLabels()
Dim myCtl
On Error GoTo err_pClearDataLabels
    For Each myCtl In Me.Controls
        If myCtl.Name Like "lblk*" Then
            myCtl.Caption = ""
            myCtl.Font.Bold = False
            myCtl.ForeColor = &H80000012
        End If
    Next myCtl
Exit Sub

' Errorhandling
err_pClearDataLabels:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'pClearDataLabels' in 'frmCalendar'. Kalender konnte nicht zurückgesetzt werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub

'======================================================================================================================
' 9. pReturnDate - Rückgabewert in globale Variable schreiben
'======================================================================================================================
Public Sub pReturnDate(dTag, dMonat, dJahr)
On Error GoTo err_pReturnDate
        g_datCalendarDate = DateSerial(dJahr, dMonat, dTag)
        Unload Me
Exit Sub

' Errorhandling
err_pReturnDate:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'pReturnDate' in 'frmCalendar'. Rückgabewert konnte nicht übergeben werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub

'======================================================================================================================
' 10. pSetNewMonth - Raster mit neuem Monat befüllen bei Klick auf "graue" Datumsfelder
'======================================================================================================================
' Übergabeparameter:
'         intTypePrePost: 1 = Vormonat / 2 = Folgemonat
'----------------------------------------------------------------------------------------------------------------------

Private Sub pSetNewMonth(intTypePrePost)
On Error GoTo err_pSetNewMonth
g_bolMonthChange = True         ' Automatische Refreshs unterbinden

Call pClearDataLabels           ' Raster zurücksetzen

' Verarbeitung
Select Case intTypePrePost
        Case 1  'Vormonat
            If fChangeStrToInt(cmbMonth.Text) = 1 Then          ' Falls aktueller Monat = Januar, dann Vormonat auf Dezember und Jahreswechsel vornehmen
                    txtYear.Text = txtYear.Text - 1
                    Call fSetMonthText(12)
                    Call setDataLabels(12, txtYear.Text)
            Else
                    Call fSetMonthText(fChangeStrToInt(cmbMonth.Text) - 1)
                    Call setDataLabels(fChangeStrToInt(cmbMonth.Text), txtYear.Text)
            End If
                
        Case 2  'Folgemonat
            If fChangeStrToInt(cmbMonth.Text) = 12 Then         ' Falls aktueller Monat = Dezember, dann Folgemonat auf Januar und Jahreswechsel vornehmen
                    txtYear.Text = txtYear.Text + 1
                    Call fSetMonthText(1)
                    Call setDataLabels(1, txtYear.Text)
            Else
                    Call fSetMonthText(fChangeStrToInt(cmbMonth.Text) + 1)
                    Call setDataLabels(fChangeStrToInt(cmbMonth.Text), txtYear.Text)
            End If
End Select

g_bolMonthChange = False

Exit Sub

' Errorhandling
err_pSetNewMonth:
    MsgBox "Error " & Err.Number & " (" & Err.Description _
        & ") in der Prozedur 'pSetNewMonth' in 'frmCalendar'. Monatswechsel konnte nicht verarbeitet werden. Ausführung wird abgebrochen.", vbCritical, "Laufzeitfehler"
    End
End Sub


'=============================================================================================================================
' 11 - 14 : Schaltflächen Monat und Jahr + und -
'=============================================================================================================================
Private Sub cmdMonthUp_Click()
    If cmbMonth.ListIndex < 11 Then cmbMonth.Text = cmbMonth.List(cmbMonth.ListIndex + 1)
End Sub

Private Sub cmdMonthDown_Click()
    If cmbMonth.ListIndex > 0 Then cmbMonth.Text = cmbMonth.List(cmbMonth.ListIndex - 1)
End Sub

Private Sub cmdYearUp_Click()
    txtYear.Text = txtYear.Text + 1
End Sub
Private Sub cmdYearDown_Click()
    txtYear.Text = txtYear.Text - 1
End Sub


'======================================================================================================================
' 15 - 56 : Click-Events auf Labels
'======================================================================================================================
' Anmerkungen:
'       Woche 1:            Prüfung, ob Tag ausgegraut ist (Vormonat), wenn ja, Aufruf zum Setzen des Vormonats
'       Wochen 5 und 6:     Prüfung, ob Tag ausgegraut ist (Folgemonat), wenn ja, Aufruf zum Setzen des Folgemonats
'----------------------------------------------------------------------------------------------------------------------

' Woche 1
Private Sub lblk1d1_Click()
If lblk1d1.ForeColor = &H808080 And lblk1d1.Caption > 10 Then
    Call pSetNewMonth(1)
Else
  Call pReturnDate(lblk1d1.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk1d2_Click()
If lblk1d2.ForeColor = &H808080 And lblk1d2.Caption > 10 Then
       Call pSetNewMonth(1)
Else
  Call pReturnDate(lblk1d2.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk1d3_Click()
If lblk1d3.ForeColor = &H808080 And lblk1d3.Caption > 10 Then
      Call pSetNewMonth(1)
Else
  Call pReturnDate(lblk1d3.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk1d4_Click()
If lblk1d4.ForeColor = &H808080 And lblk1d3.Caption > 10 Then
      Call pSetNewMonth(1)
Else
  Call pReturnDate(lblk1d4.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk1d5_Click()
If lblk1d5.ForeColor = &H808080 And lblk1d5.Caption > 10 Then
    Call pSetNewMonth(1)
Else
  Call pReturnDate(lblk1d5.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk1d6_Click()
If lblk1d6.ForeColor = &H808080 And lblk1d6.Caption > 10 Then
    Call pSetNewMonth(1)
Else
  Call pReturnDate(lblk1d6.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk1d7_Click()
If lblk1d7.ForeColor = &H808080 And lblk1d7.Caption > 10 Then
     Call pSetNewMonth(1)
Else
   Call pReturnDate(lblk1d7.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk2d1_Click()
  Call pReturnDate(lblk2d1.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk2d2_Click()
  Call pReturnDate(lblk2d2.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk2d3_Click()
  Call pReturnDate(lblk2d3.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk2d4_Click()
  Call pReturnDate(lblk2d4.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk2d5_Click()
  Call pReturnDate(lblk2d5.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk2d6_Click()
  Call pReturnDate(lblk2d6.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk2d7_Click()
  Call pReturnDate(lblk2d7.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d1_Click()
  Call pReturnDate(lblk3d1.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d2_Click()
  Call pReturnDate(lblk3d2.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d3_Click()
  Call pReturnDate(lblk3d3.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d4_Click()
  Call pReturnDate(lblk3d4.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d5_Click()
  Call pReturnDate(lblk3d5.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d6_Click()
  Call pReturnDate(lblk3d6.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk3d7_Click()
  Call pReturnDate(lblk3d7.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d1_Click()
  Call pReturnDate(lblk4d1.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d2_Click()
  Call pReturnDate(lblk4d2.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d3_Click()
  Call pReturnDate(lblk4d3.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d4_Click()
  Call pReturnDate(lblk4d4.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d5_Click()
  Call pReturnDate(lblk4d5.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d6_Click()
  Call pReturnDate(lblk4d6.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub
Private Sub lblk4d7_Click()
  Call pReturnDate(lblk4d7.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End Sub

'=======================================================================
' Woche 5
Private Sub lblk5d1_Click()
If lblk5d1.Caption < 15 Then        'lblk5d1.ForeColor = &H808080
        Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d1.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk5d2_Click()
If lblk5d2.Caption < 15 Then
           Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d2.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk5d3_Click()
If lblk5d3.Caption < 15 Then
        Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d3.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk5d4_Click()
If lblk5d4.Caption < 15 Then
            Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d4.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk5d5_Click()
If lblk5d5.Caption < 15 Then
        Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d5.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk5d6_Click()
If lblk5d6.Caption < 15 Then
          Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d6.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk5d7_Click()
If lblk5d7.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk5d7.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
'=================================================================================
'Woche 6
Private Sub lblk6d1_Click()
If lblk6d1.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d1.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
End If
End Sub
Private Sub lblk6d2_Click()
If lblk6d2.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d2.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
  End If
End Sub
Private Sub lblk6d3_Click()
If lblk6d3.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d3.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
  End If
End Sub
Private Sub lblk6d4_Click()
If lblk6d4.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d4.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
  End If
End Sub
Private Sub lblk6d5_Click()
If lblk6d5.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d5.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
  End If
End Sub
Private Sub lblk6d6_Click()
If lblk6d6.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d6.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
  End If
End Sub
Private Sub lblk6d7_Click()
If lblk6d7.Caption < 15 Then
 Call pSetNewMonth(2)
Else
  Call pReturnDate(lblk6d7.Caption, fChangeStrToInt(cmbMonth.Text), txtYear.Text)
  End If
End Sub

Private Sub UserForm_Terminate()
    If g_datCalendarDate = "00:00:00" Then End
End Sub
