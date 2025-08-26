' ------------------------------------------------------------
' Funkcja: FunctionExample
' Opis: zwięzła informacja co robi dana funkcja/procedura
' Paramerty:
'   - param1: parametr 1 (typ parametru)
'   - param2: parametr 2 (typ parametru)
'   - targetWorkbook: opcjonalny skoroszyt docelowy; domyślnie: aktywny skoroszyt
' Zwraca:
'   - True: informaja, kiedy zwraca True
'   - False: informaja, kiedy zwraca False
' Przykład użycia:
'   - wynik = FunctionExample("tekst", 123)
' Wymagania: [ewentualne zależności: inne funkcje, klasy, itp.]
' Autor: github/barabasz
' Data utworzenia: 2025-08-25
' Data modyfikacji: 2025-08-26 08:38:00 UTC
' Ostatnia zmiana: [opcjonalny zwięzły opis ostatniej zmiany]
' ------------------------------------------------------------
Function FunctionExample(param1 As Variant, param2 As Variant, Optional targetWorkbook As Workbook = Nothing) As Boolean
    ' Inicjalizacja loggera
    Dim log As Logger: Set log = ToolkitAddin.CreateLogger("FunctionExample")
    log.SetLevel(1).ShowCaller(False).ShowTime(True).Start
    On Error GoTo ErrorHandler
    
    ' Walidacja obecności wymaganych parametrów wejściowych
    If IsMissing(param1) Or IsEmpty(param1) Then
        log.Error "Parametr param1 jest wymagany"
        FunctionExample = False
        GoTo CleanUp
    End If

    ' Przykład walidacji dla parametrów, które muszą spełniać określone warunki
    If TypeName(param1) <> "String" Or Len(param1) = 0 Then
        log.Error "Parametr param1 musi być niepustym ciągiem znaków"
        FunctionExample = False
        GoTo CleanUp
    End If

    ' opcjonalna sekcja do logowania parametrów funkcji
    log.Dbg "Parametry funkcji:"
    log.var "param1", param1
    log.var "param2", param2

    ' Określenie skoroszytu docelowego
    ' (dotyczy funkcji, które odwołują się do danych w skoroszycie)
    Dim wb As Workbook
    If targetWorkbook Is Nothing Then
        Set wb = ActiveWorkbook
    Else
        Set wb = targetWorkbook
    End If

    log.Info "Rozpoczynam główną operację funkcji..."
    ' ... kod główny ...
    log.Ok "Operacja zakończona sukcesem"
    
    ' Przypisanie wartości zwracanej w przypadku sukcesu
    FunctionExample = True
    
' opcjonalna sekcja do zwalniania zasobów
CleanUp:
    ' Zwolnienie zasobów
    log.Done
    Set log = Nothing
    Exit Function

' sekcja do obsługi błędów
ErrorHandler:
    ' Obsługa błędu
    log.Exception "Błąd " & Err.Description & " (Numer błędu: " & Err.Number & ")"
    FunctionExample = False
    Resume CleanUp
End Function
