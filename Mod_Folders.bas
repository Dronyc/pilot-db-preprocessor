Attribute VB_Name = "Mod_Folders"
Option Explicit

' Каркас. Реализация в W4.

' ============================================================================
' Модуль работы с папками проекта: создание, проверка, папки шагов.
' ============================================================================

' Создаёт структуру папок проекта.
Public Sub CreateProjectFolders()
    ' TODO: реализация в W4.
End Sub

' Проверяет наличие исходных папок. Возвращает True, если всё на месте.
Public Function CheckSourceFolders() As Boolean
    ' TODO: реализация в W4.
    CheckSourceFolders = False
End Function

' Возвращает путь к папке указанного шага.
Public Function GetStepFolder(ByVal stepNumber As String) As String
    ' TODO: реализация в W4.
    GetStepFolder = ""
End Function

' Создаёт папку для шага и возвращает её путь.
Public Function CreateStepFolder(ByVal stepNumber As String, ByVal stepName As String) As String
    ' TODO: реализация в W4.
    CreateStepFolder = ""
End Function
