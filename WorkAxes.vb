
Imports Inventor

Module WorkAxes
    'Public Class Form1
    'Public Sub CreateWorkAxes()
    Public Sub CreateWorkAxes(ByVal oPartCompDef As PartComponentDefinition, ByVal oTrans As TransientGeometry)

        'WorkAxes.AddByFixed
        'create a WorkAxis passing through a point and along a direction vector
        Dim oWorkAxis1 As WorkAxis
        oWorkAxis1 = oPartCompDef.WorkAxes.AddFixed(oTrans.CreatePoint(0, 0, 0), oTrans.CreateUnitVector(1, 1, 1))

        'create a workaxis passing between two Points
        Dim oPoint1 As Point
        Dim oPoint2 As Point

        oPoint1 = oTrans.CreatePoint(4, -5, 6)
        oPoint2 = oTrans.CreatePoint(3, 3, 3)

        Dim oUnitVec As UnitVector
        oUnitVec = oTrans.CreateUnitVector((oPoint1.X - oPoint2.X), (oPoint1.Y - oPoint2.Y), (oPoint1.Z - oPoint2.Z))

        Dim oWorkAxis2 As WorkAxis
        oWorkAxis2 = oPartCompDef.WorkAxes.AddFixed(oPoint1, oUnitVec)

        'create a workaxis passing between two WorkPoints
        'the Point object can be queried from the WorkPoint object
        '    Dim oPoint3 As Point
        '    Set oPoint3 = oPartCompDef.WorkPoints.Item(1).Point
        '    Dim oPoint4 As Point
        '    Set oPoint4 = oPartCompDef.WorkPoints.Item(2).Point
        'now you have two Point objects, use the method to draw a work axis between two Point objects


        'create a workaxis between two sketch points
        'first create WorkPoints based on the sketch points
        'then create a work axis that passes between the two work points

        'create a workaxis normal to a plane at its midpoint (for example)
        Dim oMidPoint As Point = Nothing
        Dim oNormal As UnitVector = Nothing
        'just get a planar face, its midpoint and normal to the face
        Call GetPlanarFace(oMidPoint, oNormal, oPartCompDef, oTrans)

        Dim oWorkAxis3 As WorkAxis
        oWorkAxis3 = oPartCompDef.WorkAxes.AddFixed(oMidPoint, oNormal)

        'WorkAxes.AddByLine
        'can specify Edge,WorkAxis or SketchLine object

        'create a workaxis along an edge
        Dim oEdge As Edge = Nothing
        Call GetEdge(oEdge, oPartCompDef)

        Dim oWorkAxis4 As WorkAxis
        oWorkAxis4 = oPartCompDef.WorkAxes.AddByLine(oEdge)

        'create a workaxis along a sketchline
        Dim oWorkAxis5 As WorkAxis
        oWorkAxis5 = oPartCompDef.WorkAxes.AddByLine(oPartCompDef.Sketches(1).SketchLines(2))

        'WorkAxes.AddByRevolvedFace
        'get a cylindrical face
        Dim oRevolvedFace As Face = Nothing
        Call GetRevolvedFace(oRevolvedFace, oPartCompDef)

        Dim oWorkAxis6 As WorkAxis
        oWorkAxis6 = oPartCompDef.WorkAxes.AddByRevolvedFace(oRevolvedFace)


    End Sub

    Public Sub GetEdge(ByRef oEdge As Edge, ByVal oPartCompDef As PartComponentDefinition)

        'get the SurfaceBodies collection from the ComponentDefinition
        Dim oSurfaceBodies As SurfaceBodies
        oSurfaceBodies = oPartCompDef.SurfaceBodies

        'iterate through the surface bodies, faces, edge loops, edges and vertices
        'add them to the tree and save information when needed to the BrepComponents array
        Dim oSurfaceBody As SurfaceBody
        For Each oSurfaceBody In oSurfaceBodies

            'get the Faces collection from the SurfaceBody
            Dim oFaces As Faces
            oFaces = oSurfaceBody.Faces

            ' Iterate through the faces in the current body.
            Dim oFace As Face
            For Each oFace In oFaces

                'get the EdgeLoops collection from the Face
                Dim oEdgeLoops As EdgeLoops
                oEdgeLoops = oFace.EdgeLoops

                ' Iterate through the loops of the current face.
                Dim oLoop As EdgeLoop
                For Each oLoop In oEdgeLoops

                    'get the Edges collection from the EdgeLoop
                    Dim oEdges As Edges
                    oEdges = oLoop.Edges

                    ' Iterate through the edges of the current loop.
                    oEdge = oEdges.Item(1)
                Next
            Next
        Next


    End Sub

    Public Sub GetRevolvedFace(ByRef oRevolvedFace As Face, ByVal oPartCompDef As PartComponentDefinition)

        'get the SurfaceBodies collection from the ComponentDefinition
        Dim oSurfaceBodies As SurfaceBodies
        oSurfaceBodies = oPartCompDef.SurfaceBodies

        'iterate through the surface bodies, faces, edge loops, edges and vertices
        'add them to the tree and save information when needed to the BrepComponents array
        Dim oSurfaceBody As SurfaceBody
        For Each oSurfaceBody In oSurfaceBodies

            'get the Faces collection from the SurfaceBody
            Dim oFaces As Faces
            oFaces = oSurfaceBody.Faces

            ' Iterate through the faces in the current body.
            Dim oFace As Face
            For Each oFace In oFaces

                'check if the face is a cone, cylinder or torus
                If oFace.SurfaceType = Inventor.SurfaceTypeEnum.kConeSurface Or oFace.SurfaceType = Inventor.SurfaceTypeEnum.kCylinderSurface Or oFace.SurfaceType = Inventor.SurfaceTypeEnum.kTorusSurface Then
                    oRevolvedFace = oFace
                End If

            Next
        Next

    End Sub


    Public Sub GetPlanarFace(ByRef oMidPoint As Point, ByRef oNormal As UnitVector, ByVal oPartCompDef As PartComponentDefinition, ByVal oTrans As TransientGeometry)

        'get the SurfaceBodies collection from the ComponentDefinition
        Dim oSurfaceBodies As SurfaceBodies
        oSurfaceBodies = oPartCompDef.SurfaceBodies

        'iterate through the surface bodies, faces, edge loops, edges and vertices
        'add them to the tree and save information when needed to the BrepComponents array
        Dim oSurfaceBody As SurfaceBody
        For Each oSurfaceBody In oSurfaceBodies

            'get the Faces collection from the SurfaceBody
            Dim oFaces As Faces
            oFaces = oSurfaceBody.Faces

            ' Iterate through the faces in the current body.
            Dim oFace As Face
            For Each oFace In oFaces

                'check if the face is a cone, cylinder or torus
                If oFace.SurfaceType = Inventor.SurfaceTypeEnum.kPlaneSurface Then

                    'get the parameters representing the mid point on the planar face
                    'Dim Params(1 To 2) As Double
                    Dim Params(0 To 1) As Double
                    Params(0) = (oFace.Evaluator.ParamRangeRect.MaxPoint.X + oFace.Evaluator.ParamRangeRect.MinPoint.X) / 2.0#
                    Params(1) = (oFace.Evaluator.ParamRangeRect.MaxPoint.Y + oFace.Evaluator.ParamRangeRect.MinPoint.Y) / 2.0#

                    'get the actual 3d point coordinates with the help of the parametric representation
                    Dim Point(0 To 2) As Double
                    oFace.Evaluator.GetPointAtParam(Params, Point)

                    'create the actual Point object using the coordinates
                    ' oMidPoint = oTrans.CreatePoint(Point(1), Point(2), Point(3))
                    oMidPoint = oTrans.CreatePoint(Point(0), Point(1), Point(2))

                    'get the coordinates of the normal at the mid point
                    'Dim Normal(1 To 3) As Double
                    Dim Normal(0 To 2) As Double
                    oFace.Evaluator.GetNormal(Params, Normal)

                    'create the unitvector along the normal
                    'oNormal = oTrans.CreateUnitVector(Normal(1), Normal(2), Normal(3))
                    oNormal = oTrans.CreateUnitVector(Normal(0), Normal(1), Normal(2))

                End If

            Next
        Next

    End Sub

    Public Sub CreateWorkAxisAtPlanes(ByVal oPartCompDef As PartComponentDefinition, ByVal oTrans As TransientGeometry)
        'get an existing workplane
        Dim oWorkPlane1 As WorkPlane
        oWorkPlane1 = oPartCompDef.WorkPlanes.Item(2)

        'get an existing workplane
        Dim oWorkPlane2 As WorkPlane
        oWorkPlane2 = oPartCompDef.WorkPlanes.Item(3)

        Dim oNormal1 As UnitVector = Nothing
        Dim oRootPt1 As Point = Nothing
        Call GetPlaneData(oWorkPlane1, oNormal1, oRootPt1)

        Dim n1x As Double
        Dim n1y As Double
        Dim n1z As Double
        n1x = oNormal1.X
        n1y = oNormal1.Y
        n1z = oNormal1.Z

        Dim p1x As Double
        Dim p1y As Double
        Dim p1z As Double
        p1x = oRootPt1.X
        p1y = oRootPt1.Y
        p1z = oRootPt1.Z

        'calculate d1
        Dim d1 As Double
        d1 = n1x * p1x + n1y * p1y + n1z * p1z


        Dim oNormal2 As UnitVector = Nothing
        Dim oRootPt2 As Point = Nothing
        Call GetPlaneData(oWorkPlane2, oNormal2, oRootPt2)

        Dim n2x As Double
        Dim n2y As Double
        Dim n2z As Double
        n2x = oNormal2.X
        n2y = oNormal2.Y
        n2z = oNormal2.Z

        Dim p2x As Double
        Dim p2y As Double
        Dim p2z As Double
        p2x = oRootPt2.X
        p2y = oRootPt2.Y
        p2z = oRootPt2.Z

        'calculate d2
        Dim d2 As Double
        d2 = n2x * p2x + n2y * p2y + n2z * p2z


        'generate point that lies on both planes
        Dim Pt1x As Double
        Dim Pt1y As Double
        Dim Pt1z As Double
        Call GeneratePt(n1x, n1y, n1z, n2x, n2y, n2z, d1, d2, Pt1x, Pt1y, Pt1z)


        'determine the direction vector of line of intersection
        Dim oVec1x As Double
        oVec1x = n1y * n2z - n2y * n1z
        Dim oVec1y As Double
        oVec1y = n2x * n1z - n1x * n2z
        Dim oVec1z As Double
        oVec1z = n1x * n2y - n2x * n1y

        Dim det As Double
        det = Math.Sqrt(oVec1x * oVec1x + oVec1y * oVec1y + oVec1z * oVec1z)
        'create the direction vector
        Dim oUnitVec As UnitVector
        oUnitVec = oTrans.CreateUnitVector((oVec1x / det), (oVec1y / det), (oVec1z / det))

        'now create the work axis
        Dim oWorkAxis As WorkAxis
        oWorkAxis = oPartCompDef.WorkAxes.AddFixed(oTrans.CreatePoint(Pt1x, Pt1y, Pt1z), oUnitVec)

    End Sub


    Private Sub GetPlaneData(ByVal oWorkPlane As WorkPlane, ByRef oNormal As UnitVector, ByRef oRootPt As Point)

        oRootPt = oWorkPlane.Plane.RootPoint
        oNormal = oWorkPlane.Plane.Normal

    End Sub


    Private Sub GeneratePt(ByVal n1x As Double, ByVal n1y As Double, ByVal n1z As Double, ByVal n2x As Double, ByVal n2y As Double, ByVal n2z As Double, ByVal d1 As Double, ByVal d2 As Double, ByVal Pt1x As Double, ByVal Pt1y As Double, ByVal Pt1z As Double)

        Pt1x = 0.0#
        Pt1y = (d1 * n2z - d2 * n1z) / (n1y * n2z - n2y * n1z)
        Pt1z = (d1 * n2y - d2 * n1y) / (n1z * n2y - n2z * n1y)

    End Sub

End Module
