Attribute VB_Name = "Mod_Util"
Option Explicit

' Каркас. Реализация в W4.

' ============================================================================
' Модуль утилит: пути проекта, журналирование, чтение конфигурации.
' ============================================================================

' Возвращает корневую папку проекта.
Public Function GetProjectRoot() As String
    ' TODO: реализация в W4.
    GetProjectRoot = ""
End Function

' Возвращает путь к папке логов.
Public Function GetLogsFolder() As String
    ' TODO: реализация в W4.
    GetLogsFolder = ""
End Function

' Возвращает путь к папке результатов.
Public Function GetResultsFolder() As String
    ' TODO: реализация в W4.
    GetResultsFolder = ""
End Function

' Возвращает путь к папке промптов.
Public Function GetPromptsFolder() As String
    ' TODO: реализация в W4.
    GetPromptsFolder = ""
End Function

' Возвращает путь к папке шагов.
Public Function GetStepsFolder() As String
    ' TODO: реализация в W4.
    GetStepsFolder = ""
End Function

' Записывает сообщение в журнал (лог).
Public Sub WriteLog(ByVal message As String)
    ' TODO: реализация в W4.
End Sub

' Записывает ошибку с кодом в журнал.
Public Sub WriteError(ByVal code As String, ByVal message As String)
    ' TODO: реализация в W4.
End Sub

' Записывает предупреждение с кодом в журнал.
Public Sub WriteWarning(ByVal code As String, ByVal message As String)
    ' TODO: реализация в W4.
End Sub

' Возвращает значение параметра конфигурации по ключу.
Public Function GetConfigValue(ByVal key As String) As String
    ' TODO: реализация в W4.
    GetConfigValue = ""
End Function

' Открывает указанную папку в проводнике Windows.
Public Sub OpenFolderInExplorer(ByVal path As String)
    ' TODO: реализация в W4.
End Sub
