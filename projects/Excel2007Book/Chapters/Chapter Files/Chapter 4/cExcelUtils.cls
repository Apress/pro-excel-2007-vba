VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExcelUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'

Property Get ExcelVersion() As String
Dim xlApp As Object
Dim sExcelVersion As String
    
    Set xlApp = CreateObject("Excel.Application")
    sExcelVersion = xlApp.Version
    
    Set xlApp = Nothing
    ExcelVersion = sExcelVersion
End Property

Property Get ExcelVersionShort() As String
Dim xlApp As Object
Dim sExcelVersionShort As String
    
    Set xlApp = CreateObject("Excel.Application")
    sExcelVersionShort = Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)
    
    Set xlApp = Nothing
    ExcelVersionShort = sExcelVersionShort
End Property



