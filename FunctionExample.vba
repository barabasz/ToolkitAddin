' ------------------------------------------------------------
' Funkcja: FunctionExample
' Opis: zwięzła informacja co robi dana funkcja/procedura
' Paramerty:
'   - param1: parametr 1 (typ parametru)
'   - param2: parametr 2 (typ parametru)
'   - targetWorkbook: opcjonalny skoroszyt docelowy; domyślnie: aktywny
' Zwraca:
'   - True: informaja, kiedy zwraca True
'   - False: informaja, kiedy zwraca False
' Przykład użycia:
'   - wynik = FunctionExample("tekst", 123)
' Wymagania: [ewentualne zależności: inne funkcje, klasy, itp.]
' Autor: github/barabasz
' Data utworzenia: 2025-08-26
' Data modyfikacji: 2025-08-26 08:38:00 UTC
' Ostatnia zmiana: pierwsza wersja
' ------------------------------------------------------------
Function FunctionExample(param1 As Variant, param2 As Variant, Optional targetWorkbook As Workbook = Nothing) As Boolean
    ' Inicjalizacja loggera
    Dim log As Logger: Set log = ToolkitAddin.CreateLogger("FunctionExample")
    log.SetLevel(1).ShowCaller(False).ShowTime(True).Start
    On Error GoTo ErrorHandler
    
    ' Walidacja parametrów wejściowych
    If IsMissing(param1) Or IsEmpty(param1) Then
        log.Error "Parametr param1 jest wymagany"
        FunctionExample = False
        GoTo CleanUp
    End If
    
    ' Określenie skoroszytu docelowego
    ' (dotyczy funkcji, które odwołują się do danych w skoroszycie)
    Dim wb As Workbook
    If targetWorkbook Is Nothing Then
        Set wb = ActiveWorkbook
    Else
        Set wb = targetWorkbook
    End If
    
    ' ... kod główny ...
    
    ' Przypisanie wartości zwracanej w przypadku sukcesu
    FunctionExample = True
    
CleanUp:
    ' Zwolnienie zasobów
    log.Done
    Set log = Nothing
    Exit Function
    
ErrorHandler:
    ' Obsługa błędu
    log.Exception "Błąd " & Err.Description & " (Numer błędu: " & Err.Number & ")"
    FunctionExample = False
    Resume CleanUp
End Function
