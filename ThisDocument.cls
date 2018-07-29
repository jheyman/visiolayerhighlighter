VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBox1_Change()

    Select Case ComboBox1.Value
        ' Format :
        ' highlight_shapes <layer_name>, <highlight_line_color>, <highlight_text_color>, <highlight_line_width>, <faded_line_color>, <faded_text_color>
    Case "layer1":
        highlight_shapes "layer1", "RGB(255, 0, 0)", "RGB(0, 0, 0)", "4 pt", "RGB(127, 127, 127)", "RGB(127, 127, 127)"
    Case "layer2":
        highlight_shapes "layer2", "RGB(0, 255, 0)", "RGB(0, 0, 0)", "4 pt", "RGB(127, 127, 127)", "RGB(127, 127, 127)"
    Case "layer3":
        highlight_shapes "layer3", "RGB(0, 0, 255)", "RGB(0, 0, 0)", "4 pt", "RGB(127, 127, 127)", "RGB(127, 127, 127)"
    Case "None":
        highlight_shapes "None", "<none>", "<none>", "<none>", "<none>", "<none>"
    End Select

End Sub

Private Sub highlight_shapes(layer As String, highlight_line_color As String, highlight_text_color As String, highlight_line_width As String, faded_line_color As String, faded_text_color As String)

    ' First of all, and in all cases, restore object properties to their original state:
    ' - this is appropriate when the target layer to be highlighted is "None"
    ' - this is also appropriate when selecting a target layer that is not "None", while the currently highlighted
    ' target layer is a different one: since displaying the previous layer may have modified shape properties (especially line
    ' width), they need to be restored first, and will be modified again for the new layer.
    For Each shp In findAllShapes
      ' Restore original color stored in custom property, and reinitialize backup prop
      If shp.CellExists("Prop.originalLineColorBackup", 0) Then
          shp.Cells("Linecolor") = shp.Cells("Prop.originalLineColorBackup")
          shp.DeleteRow visSectionProp, shp.CellsRowIndex("Prop.originalLineColorBackup")
      End If
      ' Restore original text color stored in custom property, and reinitialize backup prop
      If shp.CellExists("Prop.originalTextColorBackup", 0) Then
          shp.CellsSRC(visSectionCharacter, 0, visCharacterColor) = shp.Cells("Prop.originalTextColorBackup")
          shp.DeleteRow visSectionProp, shp.CellsRowIndex("Prop.originalTextColorBackup")
      End If
      ' Restore original line width stored in custom property, and reinitialize backup prop
      If shp.CellExists("Prop.originalLineWidthBackup", 0) Then
          shp.CellsSRC(visSectionObject, visRowLine, visLineWeight) = shp.Cells("Prop.originalLineWidthBackup")
          shp.DeleteRow visSectionProp, shp.CellsRowIndex("Prop.originalLineWidthBackup")
      End If
    Next shp

    ' Now proceed to modify shapes' color, text color and line width depending on whether the shape belongs to the target
    ' layer to be highlighted.
    If layer <> "None" Then
      ' parse all shapes, and highlight the ones that belong to the selected target layer
      Set lTarget = ActivePage.Layers.Item(layer)
      For Each shp In findAllShapes
      
        ' figure out if the user-defined property to store the original properties of the shape exists, else create them and
        ' backup the line color, text color, and line width
        ' Only do the backup if not done already (first time we transition from "None" to some specific layer.
        ' Otherwise, a transition from one layer to another layer will backup the highlighted colors, not the original color.
        If Not shp.CellExists("Prop.originalLineColorBackup", 0) Then
            shp.AddNamedRow visSectionProp, "originalLineColorBackup", visTagDefault
            shp.Cells("Prop.originalLineColorBackup") = shp.Cells("Linecolor")
        End If
        
        If Not shp.CellExists("Prop.originalTextColorBackup", 0) Then
            shp.AddNamedRow visSectionProp, "originalTextColorBackup", visTagDefault
            shp.Cells("Prop.originalTextColorBackup") = shp.CellsSRC(visSectionCharacter, 0, visCharacterColor)
        End If
        
        If Not shp.CellExists("Prop.originalLineWidthBackup", 0) Then
            shp.AddNamedRow visSectionProp, "originalLineWidthBackup", visTagDefault
            shp.Cells("Prop.originalLineWidthBackup") = shp.CellsSRC(visSectionObject, visRowLine, visLineWeight)
        End If
        
        ' Check if the shape belongs to the target layer to be highlighted
          If (shp.LayerCount > 0) Then
              ' by default, fade out the shape text & line colors
              shp.Cells("Linecolor").FormulaU = faded_line_color
              shp.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = faded_text_color
              ' now scan the layers this shape belongs to, and highlight the shape is the target layer is in this list
              For j = 1 To shp.LayerCount
                  Set l = shp.layer(j)
                  If l.NameU = lTarget.NameU Then
                      'Now change the color to highlight the element
                      shp.Cells("Linecolor").FormulaU = highlight_line_color
                      shp.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = highlight_text_color
                      shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = highlight_line_width
                  End If
              Next j
          Else
            'If the shp does not belong to any layer, color it to the faded colors
            shp.Cells("Linecolor").FormulaU = faded_line_color
            shp.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = faded_text_color
          End If
      Next shp
    End If
End Sub

Private Sub Document_DocumentOpened(ByVal Doc As IVDocument)

    ' Create the combox entries
    ComboBox1.AddItem ("layer1")
    ComboBox1.AddItem ("layer2")
    ComboBox1.AddItem ("layer3")
    ComboBox1.AddItem ("None")
    ComboBox1.Text = ComboBox1.List(3)

End Sub

' Utility/debug function for breakpoint-based inspection
Public Sub DEBUG_INSPECT_SHAPE()
    Dim vsoShape As Visio.Shape
    Dim selected As Visio.Selection
    Set selected = ActiveWindow.Selection
    Dim collection As collection
    
    ' Without this line, when selecting a sub-shape inside a group, the selection returns Nothing.
    selected.IterationMode = Visio.VisSelectMode.visSelModeSkipSuper
    ' place a breakpoint on this line, then add a spy on the variable.
    Set vsoShape = selected.PrimaryItem
    
    Set collection = findAllShapes
    
End Sub

' returns a collection of all shapes in the Page, including shapes grouped other shapes.
Public Function findAllShapes() As collection

Dim collection As New collection
Dim tempcollection As New collection
Dim x As Shape
Dim shp As Shape

    For Each shp In Visio.ActivePage.Shapes
        Set tempcollection = subShapes(shp)
        For Each tempShp In tempcollection
            collection.Add tempShp
        Next tempShp
    Next shp
    
    Set findAllShapes = collection
    
End Function

' recursively parse the subshapes of a shape, and make a collection of all shapes down from the top-level shape.
Public Function subShapes(ByVal shp As Shape) As collection

Dim collection As New collection
Dim tempcollection As New collection
    
     ' If shape is a group, need to parse grouped sub-shapes:
     If shp.Type = visTypeGroup Then
         ' add all the grouped shaped in the list
         For Each subshp In shp.Shapes
             Set tempcollection = subShapes(subshp)
             For Each tempShp In tempcollection
                collection.Add tempShp
             Next tempShp
         Next
         ' also add the group Shape itself in the list
         collection.Add shp
     Else
        ' shape is not a group: just add this shape only to the list.
         collection.Add shp
     End If
    
     Set subShapes = collection
    
End Function
