Imports Inventor

Module WorkPlanes
    'Public Sub CreateWorkPlanes()
    Public Sub CreateWorkPlanes(ByVal oPartCompDef As PartComponentDefinition, ByVal oTrans As TransientGeometry)

        'WorkPlanes.AddByThreePoints
        'create a workplane
        'create workplane passing through three sketch points

        Dim oSketchPt1 As SketchPoint
        oSketchPt1 = oPartCompDef.Sketches(1).SketchPoints.Add(oTrans.CreatePoint2d(5, 5))

        Dim oSketchPt2 As SketchPoint
        oSketchPt2 = oPartCompDef.Sketches(1).SketchPoints.Add(oTrans.CreatePoint2d(6, 7))

        Dim oSketchPt3 As SketchPoint
        oSketchPt3 = oPartCompDef.Sketches(1).SketchPoints.Add(oTrans.CreatePoint2d(8, 9))

        Dim oWorkPlane1 As WorkPlane
        oWorkPlane1 = oPartCompDef.WorkPlanes.AddByThreePoints(oSketchPt1, oSketchPt2, oSketchPt3)

        'create a workplane passing through three workpoints
        Dim oWorkPlane2 As WorkPlane
        oWorkPlane2 = oPartCompDef.WorkPlanes.AddByThreePoints(oPartCompDef.WorkPoints(1), oPartCompDef.WorkPoints(2), oPartCompDef.WorkPoints(3))

        'WorkPlanes.AddByPlaneAndOffset
        Dim oWorkPlane3 As WorkPlane
        'the offset value is a variant, so can also specify strings
        oWorkPlane3 = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes.Item(1), "0.5")
        'can also specify an existing workplane or sketch object
        'Set oWorkPlane3 = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.WorkPlanes(2), 1#)
        'Set oWorkPlane3 = oPartCompDef.WorkPlanes.AddByPlaneAndOffset(oPartCompDef.Sketches(1), 1#)

        'WorkPlanes.AddByFixed

        'can also create a workplane using three Vertex objects
        Dim oPoint As Point = Nothing
        'Call GetTangentPoint(oPoint)
        Call GetTangentPoint(oPoint, oPartCompDef, oTrans)

        Dim xAxis As UnitVector
        xAxis = oTrans.CreateUnitVector(1, 0, 0)

        Dim yAxis As UnitVector
        yAxis = oTrans.CreateUnitVector(0, 1, 0)

        Dim oWorkPlane4 As WorkPlane
        oWorkPlane4 = oPartCompDef.WorkPlanes.AddFixed(oPoint, xAxis, yAxis)
        'this method can also be used for more complex situations,e.g. creating a workplane passing through two edges
        'first some geometric calculations have to be done to determine the unitvectors along the edges and choose any point on the edges
        'then the above method can be used to draw the workplane

        'WorkPlanes.AddByLinePlaneandAngle
        'the first parameter is line through which the workplane passes, this could be a Edge, WorkAxis or SketchLine
        'the workplane is created at an angle (third parameter) to the plane (second parameter which could be a Face,WorkPlane or Sketch object)

        Dim oWorkPlane5 As WorkPlane
        oWorkPlane5 = oPartCompDef.WorkPlanes.AddByLinePlaneAndAngle(oPartCompDef.Sketches(1).SketchLines(2), oWorkPlane2, 10)
        'for more complex situations like a workplane passing through edge and tangent to a face:
        'get the Edge and Face objects by going through the Brep, specify the angle to be 0


    End Sub

    Private Sub GetTangentPoint(ByRef oPoint As Point, ByVal oPartCompDef As PartComponentDefinition, ByVal oTrans As TransientGeometry)

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

                If oFace.SurfaceType = Inventor.SurfaceTypeEnum.kCylinderSurface Then

                    'get the face evaluator
                    Dim oFaceEvaluator As SurfaceEvaluator
                    oFaceEvaluator = oFace.Evaluator

                    'get the Parametric range (box2d) from the evaluator
                    Dim oParamRange As Box2d
                    oParamRange = oFaceEvaluator.ParamRangeRect

                    'declare the parameter space coordinate variables
                    Dim u As Double
                    Dim v As Double

                    'iterate through the points in the parameter space and obtain the actual 3d coordinates and normals, tangents at the points on the face
                    For u = oParamRange.MinPoint.X To oParamRange.MaxPoint.X Step 0.01
                        For v = oParamRange.MinPoint.Y To oParamRange.MaxPoint.Y Step 0.01
                            Dim FaceParams As Array = New Double(1) {u, v}

                            'get the 3D point using the parameter space coordinates
                            Dim FacePoints As System.Array = New Double(1) {0, 0}
                            oFaceEvaluator.GetPointAtParam(FaceParams, FacePoints)


                            'get the normals
                            Dim FaceNormals As Array = New Double(1) {0, 0}
                            oFaceEvaluator.GetNormal(FaceParams, FaceNormals)

                            If FaceNormals.GetValue(0) < 0.1 And FaceNormals.GetValue(0) > -0.1 Then
                                If FaceNormals.GetValue(1) < 1.1 And FaceNormals.GetValue(1) > 0.9 Then

                                    'get the actual 3d point coordinates with the help of the parametric representation
                                    'Dim Point(1 To 3) As Double
                                    Dim Point(0 To 2) As Double
                                    oFace.Evaluator.GetPointAtParam(FaceParams, Point)

                                    'create the actual Point object using the coordinates
                                    oPoint = oTrans.CreatePoint(Point(0), Point(1) + 1.5, Point(2))

                                    Exit Sub

                                End If
                            End If
                        Next
                    Next
                End If
            Next
        Next
    End Sub

End Module

