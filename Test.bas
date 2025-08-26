Attribute VB_Name = "Test"
Option Explicit

Function TestLogger() As Boolean
    Dim log As Logger
    Set log = ToolkitAddin.CreateLogger("TestVariantToString")
    log.ShowCaller(False).SetLevel(0).Start
    log.Info "To jest informacja"
    log.Warn "To jest ostrzeżenie"
    log.Error "To jest błąd"
    log.Done
End Function
