    Dim pythonPath As String
    Dim scriptPath As String
    Dim excelPath As String
    Dim command As String
    Dim shellCommand As String
    Dim workbookDir As String

    workbookDir = ThisWorkbook.Path

    pythonPath = ".\env\Scripts\python.exe"
    scriptPath = ".\main.py"
    excelPath = ThisWorkbook.FullName

    command = "cd '" & workbookDir & "'; " & pythonPath & " """ & scriptPath & """ """ & excelPath & """"
    ' Debug.Print command

    shellCommand = "powershell.exe -NoExit -Command " & Chr(34) & command & Chr(34)

    shell shellCommand, vbNormalFocus
