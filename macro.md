' This is all a work in progress- use with caution

Sub HighlightTargets2()

	Dim range As range
	Dim i As Long
	Dim TargetList

	' long list of poor resume words
	TargetList = Array("Agreement", "Ignite", "terms")

	For i = 0 To UBound(TargetList)

	' make the entire document the range
	Set range = ActiveDocument.range

	With range.Find
	.Text = TargetList(i)
	.Format = True
	.MatchCase = false
	.MatchWholeWord = False
	.MatchWildcards = False
	.MatchSoundsLike = False
	.MatchAllWordForms = False

	Do While .Execute(Forward:=True) = True
		Then MsgBox "Not Found"
	range.HighlightColorIndex = wdYellow

	Loop

	End With
	Next

End Sub
