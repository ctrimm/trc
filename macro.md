###
This is all a work in progress- use with caution - written for Microsoft Office 2007 and above.

The words below are from CareerBuilder's survey of ~2,200 HR professionals in regards to the worst words to put on your resume.
###

    Sub HighlightPoorWords()

        Dim range As range
        Dim i As Long
        Dim TargetList

        ' long list of poor resume words
        TargetList = Array("Best of breed", "Go-getter", "go getter", "Think outside of the box", "outside the box", "Synergy", "Go-to person", "go-to person", "go to person", "Thought leadership", "Value add", "Results-driven", "results driven", "result driven", "Team player", "Bottom-line", "bottom line", "Hard worker", "Strategic thinker", "Dynamic", "Self-motivate", "self motivate", "Detail-oriented", "detail oriented", "Proactively", "Track record", "expert", "outstanding", "salary negotiable", "references available by request", "responsible for", "experience working in", "problem-solving skills", "hardworking", "team-player", "team player", "proactive", "objective", "multifaceted collaboration", "worked together")

        For i = 0 To UBound(TargetList)

        ' make the entire document the range
        Set range = ActiveDocument.range

        With range.Find
        .Text = TargetList(i)
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Forward:=True) = True
        range.HighlightColorIndex = wdYellow

        Loop

        End With
        Next

    End Sub




"Best of breed", "Go-getter", "go getter", "Think outside of the box", "Synergy", "Go-to person", "go to person", "Thought leadership", "Value add", "Results-driven", "results driven", "result driven", "Team player", "Bottom-line", "bottom line", "Hard worker", "Strategic thinker", "Dynamic", "Self-motivate", "self motivate", "Detail-oriented", "detail oriented", "Proactively", "Track record", "expert", "outstanding", "salary negotiable", "references available by request", "responsible for", "experience working in", "problem-solving skills", "hardworking", "team-player", "team player", "proactive", "objective", "multifaceted collaboration", "worked together"