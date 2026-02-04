Attribute VB_Name = "Module1"
Option Explicit

Public Sub VirBU01_AUT()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim res As String

    Const OUT_COL As String = "AD"

    ' --- Add-in proof : on vise le classeur utilisateur + la feuille MEP
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets("MEP")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Feuille 'MEP' introuvable dans le classeur actif : " & wb.Name, vbExclamation
        Exit Sub
    End If

    Dim f As Range
' ws.Columns("A").Delete Shift:=xlToLeft
Set f = ws.Columns("A").Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, _
                             SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

If f Is Nothing Or f.Row < 2 Then
    MsgBox "Aucune donnée à traiter sur 'MEP' (colonne A vide).", vbInformation
    Exit Sub
End If

lastRow = f.Row

    ' --- Sauvegarde des paramètres Excel
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean

    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents

    ' --- Perf
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    For i = 2 To lastRow

        Dim aVal As String, cVal As String, sVal As String, tVal As String, uVal As String
        Dim pVal As Variant, rVal As Variant

        aVal = Trim$(CStr(ws.Cells(i, "A").Value))
        cVal = Trim$(CStr(ws.Cells(i, "C").Value))
        pVal = ws.Cells(i, "P").Value
        rVal = ws.Cells(i, "R").Value
        sVal = Trim$(CStr(ws.Cells(i, "S").Value))
        tVal = Trim$(CStr(ws.Cells(i, "T").Value))
        uVal = Trim$(CStr(ws.Cells(i, "U").Value))

        res = vbNullString

        ' A = AKAMAI
        If aVal = "AKAMAI" Then
            res = res & "Vérifier RIB / "
        End If

        ' C : TIT en début/fin OU 1er caractère ni chiffre ni lettre
        If IsFactureSuspecte(cVal) Then
            res = res & "Vérifier Numéro Facture / "
        End If

        ' P >= 800000
        If IsNumeric(pVal) Then
            If CDbl(pVal) >= 800000# Then res = res & ">=800K€ / "
        End If

        ' R : date passée
        If IsDate(rVal) Then
            If CDate(rVal) < Date Then res = res & "Vérifier Date passée / "
        End If

        ' S : fins sensibles
        If FinSensibleIBAN(sVal) Then
            res = res & "Mettre en PG18 IBAN / "
        End If

        ' T : BIC -> PG03
        If BIC_PG03(tVal) Then
            res = res & "Mettre en PG03 BIC / "
        End If

        ' T : liste "RIB bloqué"
        If BIC_RIBBloque(tVal) Then
            res = res & "Mettre RIB Bloqué / "
        End If

        ' U : pays OK
        If Not IsPaysOK(uVal) Then
            res = res & "PAYS / "
        End If

        ' OK si rien
        If Len(res) = 0 Then res = "OK"

        ws.Cells(i, OUT_COL).Value = res
    Next i

CleanExit:
    ' --- Restauration des paramètres Excel (même si erreur)
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    MsgBox "Vérifications terminées (" & (lastRow - 1) & " lignes).", vbInformation
    Exit Sub

CleanFail:
    MsgBox "Erreur : " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit

End Sub


' =========================
' === Fonctions utilitaires
' =========================

Private Function IsFactureSuspecte(ByVal cVal As String) As Boolean
    cVal = Trim$(cVal)

    If Len(cVal) = 0 Then
        IsFactureSuspecte = True
        Exit Function
    End If

    If Left$(cVal, 3) = "TIT" Or Right$(cVal, 3) = "TIT" Then
        IsFactureSuspecte = True
        Exit Function
    End If

    ' 1er caractère doit être chiffre ou lettre (A-Z / a-z)
    Dim ch As String
    ch = Left$(cVal, 1)

    IsFactureSuspecte = Not (ch Like "[0-9A-Za-z]")
End Function


Private Function FinSensibleIBAN(ByVal sVal As String) As Boolean
    sVal = Trim$(sVal)

    FinSensibleIBAN = (Len(sVal) >= 4 And Right$(sVal, 4) = "1623") Or _
                      (Len(sVal) >= 4 And Right$(sVal, 4) = "3310") Or _
                      (Len(sVal) >= 4 And Right$(sVal, 4) = "9742") Or _
                      (Len(sVal) >= 5 And Right$(sVal, 5) = "43840")
End Function


Private Function BIC_PG03(ByVal tVal As String) As Boolean
    ' ESTVIDE(T) OU GAUCHE(T;4)="TRPU" OU GAUCHE(T;8)="BDFEFRPP"
    tVal = UCase$(Trim$(tVal))

    If Len(tVal) = 0 Then
        BIC_PG03 = True
        Exit Function
    End If

    If Left$(tVal, 4) = "TRPU" Then
        BIC_PG03 = True
        Exit Function
    End If

    If Left$(tVal, 8) = "BDFEFRPP" Then
        BIC_PG03 = True
        Exit Function
    End If

    BIC_PG03 = False
End Function


Private Function BIC_RIBBloque(ByVal tVal As String) As Boolean
    tVal = UCase$(Trim$(tVal))

    BIC_RIBBloque = (Left$(tVal, 8) = "NORDFRPP") Or _
                    (Left$(tVal, 6) = "TARNFR") Or _
                    (Left$(tVal, 7) = "COURTFR") Or _
                    (Left$(tVal, 6) = "KOLBFR") Or _
                    (Left$(tVal, 6) = "BNUGFR") Or _
                    (Left$(tVal, 6) = "RAPLFR") Or _
                    (Left$(tVal, 6) = "SMCTFR") Or _
                    (Left$(tVal, 6) = "SGBTMC") Or _
                    (Left$(tVal, 7) = "SBGDFRP")
End Function


Private Function IsPaysOK(ByVal uVal As String) As Boolean

    Dim x As String, p As String

    x = UCase$(Trim$(uVal))
    x = Replace(x, " ", vbNullString)
    x = Replace(x, "-", vbNullString)
    x = Replace(x, ".", vbNullString)

    If Len(x) >= 2 Then
        p = Left$(x, 2)
    Else
        p = vbNullString
    End If

    Select Case p
        Case "FR", "RE", "MQ", "GP", "GF", "PF"
            IsPaysOK = True
        Case Else
            IsPaysOK = False
    End Select
End Function

