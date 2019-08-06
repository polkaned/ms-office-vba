Sub exportTab2Png()
  Dim tab0     As Visio.Page
  Dim tabs     As Visio.Pages
  Dim filename As String
  Dim tabname  As String
  Dim tabid    As Integer
  Set tabs = Application.ActiveDocument.Pages
  For tabid = 1 To tabs.Count
    Set tab0 = tabs(tabid)
    tabname = tab0.Name
    filename = Application.ActiveDocument.Path & tabname & ".png"
    tab0.Export filename
  Next tabid
End Sub
