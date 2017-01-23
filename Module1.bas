Attribute VB_Name = "Module1"
Public Sub Sendkeys(text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys text, wait
   Set WshShell = Nothing
End Sub


