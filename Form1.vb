Imports Inventor

Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim oApp As Inventor.Application
        oApp = GetObject(, "Inventor.Application")

        'Change this hard coded path to the ipt file on your system
        Dim oPartDoc As PartDocument
        oPartDoc = oApp.Documents.Open("C:\temp\test.ipt", True)

        'Set a reference to the component definition.
        Dim oPartCompDef As PartComponentDefinition
        oPartCompDef = oPartDoc.ComponentDefinition

        'get the TransientGeometry object
        Dim oTrans As TransientGeometry
        oTrans = oApp.TransientGeometry

        'call the method to create the work points
        WorkPoints.CreateWorkPoints(oPartCompDef, oTrans)

        'create work axes
        WorkAxes.CreateWorkAxes(oPartCompDef, oTrans)

        'create work planes
        WorkPlanes.CreateWorkPlanes(oPartCompDef, oTrans)

        'for more complex scenarios, detailed calculations may be required,for example:
        'create a work axis at intersection of two workplanes
        'Call CreateWorkAxisAtPlanes()
        Call WorkAxes.CreateWorkAxisAtPlanes(oPartCompDef, oTrans)
    End Sub
End Class
'End Namespace
