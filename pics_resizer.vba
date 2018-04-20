' It just resize all the pics in the doc. Simple as that.
Sub ResizeAllPics()
	Dim i As Long
	With ActiveDocument
	    For i = 1 To .InlineShapes.Count
	        With .InlineShapes(i)
	            .ScaleHeight = 50
	            .ScaleWidth = 50
	        End With
	    Next i
	End With
End Sub 
