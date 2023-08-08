Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Integration
Imports RibbonBasics.My.Resources

<ComVisible(True)>
Public Class Ribbon
    Inherits ExcelRibbon

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Return RibbonResources.Ribbon ' The name here is the resource name that the ribbon xml has in the RibbonResources resource file
    End Function

    Public Overrides Function LoadImage(imageId As String) As Object
        ' This will return the image resource with the name specified in the image='xxxx' tag
        Return RibbonResources.ResourceManager.GetObject(imageId)
    End Function
    Public Sub OnSayHelloPressed(control As IRibbonControl)
        Dim app As Application
        Dim rng As Range

        app = ExcelDnaUtil.Application
        rng = app.Range("A1")
        rng.Value = "Hello from .NET!"

    End Sub

End Class
