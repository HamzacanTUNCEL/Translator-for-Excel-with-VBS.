Attribute VB_Name = "Module1"
Public Function Translate(WRD As Range, SL As String, TL As String)

Dim ie As New InternetExplorer
Dim doc As HTMLDocument

sltl = "sl=" & SL & "&" & "tl=" & TL & "&text=" & WRD & "&op=translate"
ie.navigate "https://translate.google.com/?hl=tr&" & sltl
ie.Visible = False


Do
DoEvents
Loop Until ie.readyState = READYSTATE_COMPLETE
Set doc = ie.document

Wait_Time (2)

Set cevap = doc.getElementsByClassName("ryNqvb")(0)
Translate = cevap.innerText

ie.Quit
Set ie = Nothing

End Function
Public Function Wait_Time(waitingtime As Double)
Start = Timer
Do
DoEvents
Loop Until (Timer - Start) >= waitingtime
End Function


