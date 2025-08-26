Attribute VB_Name = "Sheets"
Option Explicit

' ------------------------------------------------------------
' Funkcja: DeleteSheet
' Opis: Funkcja usuwająca wskazany arkusz
' Paramerty:
'   - sheetName: nazwa arkusza (String)
'   - targetWorkbook: opcjonalny - jeśli nie podano, używa aktywnego skoroszytu (Workbook)
'   - showAlerts: opcjonalny - określa czy wyświetlać potwierdzenie usunięcia (Boolean, domyślnie False)
' Zwraca:
'   - True przy powodzeniu,
'   - False, jeśli arkusz nie został znaleziony, jest chroniony lub wystąpił błąd
' Autor: github/barabasz
' Data utworzenia: 2025-08-26
' Data modyfikacji: 2025-08-26 14:17:44 UTC
' ------------------------------------------------------------
Function DeleteSheet(sheetName As String, Optional targetWorkbook As Workbook = Nothing, Optional showAlerts As Boolean = False) As Boolean
    Dim log As Logger: Set log = ToolkitAddin.CreateLogger
    log.SetCaller("DeleteSheet").ShowCaller(True).SetLevel(1)
    log.Start
    
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim currentAlertsState As Boolean
    
    ' Określ, który skoroszyt ma być użyty
    If targetWorkbook Is Nothing Then
        Set wb = Application.ActiveWorkbook
    Else
        Set wb = targetWorkbook
    End If

    ' Sprawdź, czy arkusz o podanej nazwie istnieje za pomocą funkcji SheetExists
    If Not SheetExists(sheetName, wb) Then
        log.Error "Arkusz o nazwie '" & sheetName & "' nie istnieje w skoroszycie " & wb.name & "."
        DeleteSheet = False
        GoTo CleanUp
    End If

    ' Pobierz arkusz (już wiemy, że istnieje)
    Set ws = wb.Sheets(sheetName)
    
    ' Sprawdź, czy skoroszyt ma więcej niż jeden arkusz
    If wb.Sheets.Count <= 1 Then
        log.Error "Nie można usunąć arkusza '" & sheetName & "', ponieważ jest to jedyny arkusz w skoroszycie " & wb.name & "."
        DeleteSheet = False
        GoTo CleanUp
    End If
    
    ' Sprawdź, czy arkusz jest chroniony przed usunięciem
    If ws.ProtectContents = True Then
        On Error Resume Next
        ' Próba usunięcia chronionego arkusza wywoła błąd, więc próbujemy najpierw sprawdzić ochronę
        log.Error "Arkusz '" & sheetName & "' jest chroniony i nie może zostać usunięty."
        DeleteSheet = False
        GoTo CleanUp
    End If
    
    ' Zapisz aktualny stan alertów
    currentAlertsState = Application.DisplayAlerts
    
    ' Ustaw alerty zgodnie z parametrem
    Application.DisplayAlerts = showAlerts
    
    ' Usuń arkusz
    ws.Delete
    
    ' Przywróć poprzedni stan alertów
    Application.DisplayAlerts = currentAlertsState
    
    log.Ok "Arkusz '" & sheetName & "' w skoroszycie " & wb.name & " został pomyślnie usunięty."
    DeleteSheet = True
    GoTo CleanUp

ErrorHandler:
    ' Przywróć poprzedni stan alertów w przypadku błędu
    Application.DisplayAlerts = currentAlertsState
    log.Exception "Wystąpił błąd podczas usuwania arkusza '" & sheetName & "'. Opis błędu: " & Err.Description & " (Numer błędu: " & Err.Number & ")"
    DeleteSheet = False

CleanUp:
    log.Done
    Set log = Nothing
End Function

' ------------------------------------------------------------
' Funkcja: SheetExists
' Opis: Funkcja sprawdzająca, czy istnieje arkusz o podanej nazwie
' Paramerty:
'   - sheetName: nazwa arkusza (String)
'   - targetWorkbook (opcjonalny - jeśli nie podano, używa aktywnego skoroszytu)
' Zwraca:
'   - Boolean: True, jeśli arkusz istnieje, False jeśli nie istnieje
' Autor: github/barabasz
' Data utworzenia: 2024-11-05
' Data modyfikacji: 2025-08-14 14:36:35 UTC
' ------------------------------------------------------------
Function SheetExists(sheetName As String, Optional targetWorkbook As Workbook = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Określ, który skoroszyt ma być użyty
    If targetWorkbook Is Nothing Then
        Set wb = Application.ActiveWorkbook
    Else
        Set wb = targetWorkbook
    End If
    
    ' Próba znalezienia arkusza o podanej nazwie
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    ' Sprawdź, czy arkusz został znaleziony
    If ws Is Nothing Then
        SheetExists = False
    Else
        SheetExists = True
    End If
    
    Exit Function
    
ErrorHandler:
    SheetExists = False
End Function

' ------------------------------------------------------------
' Funkcja: RefreshSheet
' Opis: Funkcja przeliczająca wskazany arkusz
' Paramerty:
'   - sheetName: nazwa arkusza (String)
'   - targetWorkbook (opcjonalny - jeśli nie podano, używa aktywnego skoroszytu)
' Zwraca:
'   - True przy powodzeniu,
'   - False, jeśli arkusz nie został znaleziony lub wystąpił błąd
' Autor: github/barabasz
' Data utworzenia: 2024-11-05
' Data modyfikacji: 2025-08-14 14:36:35 UTC
' ------------------------------------------------------------
Function RefreshSheet(sheetName As String, Optional targetWorkbook As Workbook = Nothing) As Boolean
    Dim log As Logger: Set log = ToolkitAddin.CreateLogger
    log.SetCaller("RefreshSheet").ShowCaller(True).SetLevel (4)
    
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wb As Workbook
    
    ' Określ, który skoroszyt ma być użyty
    If targetWorkbook Is Nothing Then
        Set wb = Application.ActiveWorkbook
    Else
        Set wb = targetWorkbook
    End If

    ' Sprawdź, czy arkusz o podanej nazwie istnieje za pomocą funkcji SheetExists
    If Not SheetExists(sheetName, wb) Then
        log.Error "Arkusz o nazwie '" & sheetName & "' nie istnieje w skoroszycie " & wb.name & "."
        RefreshSheet = False
        Exit Function
    End If

    ' Pobierz arkusz (już wiemy, że istnieje)
    Set ws = wb.Sheets(sheetName)

    ' Przeliczanie arkusza
    ws.Calculate
    log.Ok "Arkusz '" & sheetName & "' w skoroszycie " & wb.name & " został pomyślnie przeliczony."
    RefreshSheet = True
    Exit Function

ErrorHandler:
    log.Exception "Błąd: Wystąpił błąd podczas odświeżania arkusza '" & sheetName & "'. Opis błędu: " & Err.Description & " (Numer błędu: " & Err.Number & ")"
    RefreshSheet = False
End Function

