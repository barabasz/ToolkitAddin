VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StopForm 
   Caption         =   "StopFormCaption"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "StopForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StopForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Zmienne modułowe
Private mCounterValue As Integer
Private mTimerRunning As Boolean
Private mCancelled As Boolean
Private mNextTime As Double  ' Przechowuje czas następnego wywołania timera
Private mLog As Logger      ' Logger do diagnostyki

' Zmienne przechowujące domyślne wartości
Private mDefaultCaption As String
Private mDefaultText As String
Private mDefaultTime As Integer

Private Sub UserForm_Initialize()
    ' Inicjalizacja loggera
    Set mLog = New Logger
    mLog.SetCaller("StopForm").SetLevel(2).Start
    
    ' Inicjalizacja domyślnych wartości
    mDefaultCaption = "Automatyczne uruchamianie"
    mDefaultText = "Za chwilę nastąpi automatyczne uruchomienie makra!"
    mDefaultTime = 5
    
    ' Ustawienie domyślnego tekstu i nagłówka
    Me.Caption = mDefaultCaption
    stopFormText.Caption = mDefaultText
    stopFormCounter.Caption = mDefaultTime
    
    ' Inicjalizacja zmiennych
    mTimerRunning = False
    mCancelled = False
End Sub

Private Sub UserForm_Terminate()
    ' Upewnij się, że timer jest zatrzymany
    StopTimer
    mLog.Done
    Set mLog = Nothing
End Sub

Private Sub stopFormCancel_Click()
    ' Przerwanie odliczania przez użytkownika
    mLog.Info "Anulowanie przez użytkownika - przycisk Cancel"
    StopTimer
    mCancelled = True
    ' Ustaw także flagę globalną
    SetStopFormCancelled
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Obsługa zamknięcia formularza przez "X" w prawym górnym rogu
    If CloseMode = vbFormControlMenu Then
        mLog.Info "Anulowanie przez użytkownika - przycisk X"
        mCancelled = True
        ' Ustaw flagę globalną
        SetStopFormCancelled
        StopTimer
        Me.Hide
    End If
End Sub

Public Sub StartTimer(Optional ByVal seconds As Integer = 0)
    ' Inicjalizacja i uruchomienie timera
    If seconds <= 0 Then
        mCounterValue = mDefaultTime
    Else
        mCounterValue = seconds
    End If
    
    stopFormCounter.Caption = mCounterValue
    mTimerRunning = True
    
    ' Uruchom timer używając Application.OnTime i procedury z modułu standardowego
    TimerTick
End Sub

Public Sub StopTimer()
    ' Zatrzymanie timera
    If mTimerRunning Then
        mTimerRunning = False
        On Error Resume Next
        Application.OnTime EarliestTime:=mNextTime, Procedure:="StopFormTimerTick", Schedule:=False
        On Error GoTo 0
    End If
End Sub

Public Sub TimerTick()
    ' Ta procedura zastępuje zdarzenie timera
    If Not mTimerRunning Then Exit Sub
    
    ' Zmniejsz licznik i zaktualizuj wyświetlanie
    mCounterValue = mCounterValue - 1
    stopFormCounter.Caption = mCounterValue
    
    ' Jeśli odliczanie doszło do zera, zatrzymaj timer i zamknij formularz
    If mCounterValue <= 0 Then
        mLog.Info "Odliczanie zakończone automatycznie"
        StopTimer
        Me.Hide
    Else
        ' Zaplanuj następne wywołanie za 1 sekundę używając procedury z modułu standardowego
        mNextTime = Now + TimeSerial(0, 0, 1)
        Application.OnTime EarliestTime:=mNextTime, Procedure:="StopFormTimerTick"
    End If
End Sub

Public Property Get Cancelled() As Boolean
    ' Właściwość wskazująca, czy użytkownik anulował odliczanie
    Cancelled = mCancelled
End Property

Public Sub ResetCancelled()
    ' Zresetuj flagę anulowania na początku
    mCancelled = False
End Sub

Public Property Get DefaultCaption() As String
    DefaultCaption = mDefaultCaption
End Property

Public Property Get DefaultText() As String
    DefaultText = mDefaultText
End Property

Public Property Get DefaultTime() As Integer
    DefaultTime = mDefaultTime
End Property

Public Sub SetText(ByVal newText As String)
    ' Ustawienie tekstu wyświetlanego w formularzu
    stopFormText.Caption = newText
End Sub

Public Sub SetCaption(ByVal newCaption As String)
    ' Ustawienie nagłówka formularza
    Me.Caption = newCaption
End Sub
