Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

Friend Module mExcelHelpers

    Public Application As Application = ExcelDnaUtil.Application

    Function Cells() As Range
        Cells = Application.Cells
    End Function
    Function Range(cell1 As Object) As Range
        Range = Application.Range(cell1)
    End Function
    Function Range(cell1 As Object, cell2 As Object) As Range
        Range = Application.Range(cell1, cell2)
    End Function
    Function Charts() As Sheets
        Charts = Application.Charts
    End Function
    Function ActiveChart() As Chart
        ActiveChart = Application.ActiveChart
    End Function
    Function Sheets() As Sheets
        Sheets = Application.Sheets
    End Function
    Function ActiveSheet() As Object
        ActiveSheet = Application.ActiveSheet
    End Function
    Function Selection() As Object
        Selection = Application.Selection
    End Function

End Module
