Imports Inventor

Module WorkPoints
    Public Sub CreateWorkPoints(ByVal oPartCompDef As PartComponentDefinition, ByVal oTrans As TransientGeometry)

        'create some work points using AddFixed method
        Dim oWorkPoint1 As WorkPoint
        oWorkPoint1 = oPartCompDef.WorkPoints.AddFixed(oTrans.CreatePoint(5, 5, 5))

        Dim oWorkPoint2 As WorkPoint
        oWorkPoint2 = oPartCompDef.WorkPoints.AddFixed(oTrans.CreatePoint(5, -5, 5))

        Dim oWorkPoint3 As WorkPoint
        oWorkPoint3 = oPartCompDef.WorkPoints.AddFixed(oTrans.CreatePoint(5, 5, -5))

        'create workpoint using AddByPoint method,supply a Vertex object
        Dim oWorkPoint4 As WorkPoint
        Dim oVertex As Vertex = Nothing

        'this method just obtains the start vertex of the first edge in brep hierarchy
        Call GetVertex(oVertex, oPartCompDef) 'this is the wrong parameter wb
        oWorkPoint4 = oPartCompDef.WorkPoints.AddByPoint(oVertex)

        'create workpoint using AddByPoint method,supply a SketchPoint object
        Dim oSketch As PlanarSketch
        ' Create the sketch on the first workplane (xy plane).(1-yz,2-xz,3-xy)
        oSketch = oPartCompDef.Sketches.Add(oPartCompDef.WorkPlanes(3))

        'creation of corresponding sketch points using Sketch.SketchPoints.Add method, the second parameter specifies whether the point is a hole center
        Dim oSketchPt1 As SketchPoint
        oSketchPt1 = oSketch.SketchPoints.Add(oTrans.CreatePoint2d(-5, -5), False)

        Dim oWorkPoint5 As WorkPoint
        oWorkPoint5 = oPartCompDef.WorkPoints.AddByPoint(oSketchPt1)

    End Sub

    Public Sub GetVertex(ByRef oVertex As Vertex, ByVal oPartCompDef As PartComponentDefinition)

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
                    Dim oEdge As Edge
                    For Each oEdge In oEdges

                        'STARTVERTEX
                        'get the StartVertex of the edge
                        oVertex = oEdge.StartVertex

                    Next
                Next
            Next
        Next
    End Sub
End Module
