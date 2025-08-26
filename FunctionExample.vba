Option Explicit

' ------------------------------------------------------------
' Funkcja: FunctionExample
' Opis: zwięzła informacja co robi dana funkcja/procedura
' Paramerty:
'   - param1: parametr 1 (typ parametru)
'   - param2: parametr 2 (typ parametru)
' Zwraca:
'   - True: informaja, kiedy zwraca True
'   - False: informaja, kiedy zwraca False
' Wymagania: [ewentualne zależności: inne funkcje, klasy, itp.]
' Autor: github/barabasz
' Data utworzenia: 2025-08-01
' Data modyfikacji: 2025-08-26 10:33:48 UTC
' Ostatnia zmiana: zwięzły opis ostatniej modyfikacji
' ------------------------------------------------------------
Function FunctionExample(param1 As Variant, param2 As Variant) As Boolean
    ' inicjalizacja Loggera
    Dim log As Logger: Set log = ToolkitAddin.CreateLogger("FunctionExample")
    log.SetLevel(1).Start
    On Error GoTo ErrorHandler
    ' ... kod główny ...
    log.Done
    Exit Function
ErrorHandler:
    ' ... obsługa błędu ...
    log.Exception "Błąd " & Err.Description & " (Numer błędu: " & Err.Number & ")"
    log.Done
End Function
