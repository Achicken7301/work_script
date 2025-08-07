    Dim pythonPath As String
    Dim scriptPath As String
    Dim excelPath As String
    Dim command As String
    Dim shellCommand As String
    Dim workbookDir As String

    ' L?y du?ng d?n thu m?c ch?a file Excel
    workbookDir = ThisWorkbook.Path

    ' Ðu?ng d?n tuong d?i tính t? thu m?c Excel
    pythonPath = ".\env\Scripts\python.exe"
    scriptPath = ".\main.py"
    excelPath = ThisWorkbook.FullName

    ' Câu l?nh PowerShell d?m b?o cd vào dúng thu m?c tru?c khi g?i Python
    command = "cd '" & workbookDir & "'; " & pythonPath & " """ & scriptPath & """ """ & excelPath & """"
    Debug.Print command

    shellCommand = "powershell.exe -NoExit -Command " & Chr(34) & command & Chr(34)

    shell shellCommand, vbNormalFocus