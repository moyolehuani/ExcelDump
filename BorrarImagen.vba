Sub borrarimagen()

'BoRrar imagenes primero
Dim img As Shape
'si existe alguna foto, la borro:
On Error Resume Next
For Each img In ActiveSheet.Shapes
  'El número 11 aplica para los JPGs
  img.Delete
Next

End Sub
