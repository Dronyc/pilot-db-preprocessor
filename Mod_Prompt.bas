Attribute VB_Name = "Mod_Prompt"
Option Explicit

' Каркас. Реализация в W4.

' ============================================================================
' Модуль промптов: формирование промптов для следующего шага и исправлений,
' чтение списка и текущей позиции шагов.
' ============================================================================

' Формирует промпт для выполнения следующего шага.
Public Sub MakeNextStepPrompt()
    ' TODO: реализация в W4.
End Sub

' Формирует промпт для исправления ошибок.
Public Sub MakeFixPrompt()
    ' TODO: реализация в W4.
End Sub

' Читает список шагов и возвращает его коллекцией.
Public Function ReadStepList() As Collection
    ' TODO: реализация в W4.
    Set ReadStepList = Nothing
End Function

' Возвращает номер текущего шага.
Public Function ReadCurrentStep() As String
    ' TODO: реализация в W4.
    ReadCurrentStep = ""
End Function

' Возвращает номер следующего шага.
Public Function ReadNextStep() As String
    ' TODO: реализация в W4.
    ReadNextStep = ""
End Function
