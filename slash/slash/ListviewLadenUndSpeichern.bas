Attribute VB_Name = "ListviewLadenUndSpeichern"

Public Sub lvw_WriteData(lvw As ListView, _
  sDataFile As String)

  Dim F As Integer
  Dim intCols As Integer
  Dim i As Integer
  Dim U As Integer

  F = FreeFile
  Open sDataFile For Output As #F

  intCols = lvw.ColumnHeaders.Count

  With lvw.ListItems
    For i = 1 To .Count
      With lvw.ListItems(i)

        Print #F, .Key; vbTab;
        Print #F, .Text;

        For U = 1 To intCols - 1
          Print #F, vbTab; .SubItems(U);
        Next U

        Print #F, ""
      End With
    Next i
  End With

  Close #F
End Sub

Public Sub lvw_ReadData(lvw As ListView, _
  ByVal sDataFile As String, _
  Optional ByVal bAppend As Boolean = False)

  Dim F As Integer
  Dim sLine As String
  Dim sItem() As String
  Dim intCols As Integer
  Dim i As Integer
  Dim itemX As ListItem
  
  If Not bAppend Then lvw.ListItems.Clear
  
  If Dir(sDataFile, vbNormal) <> "" Then

    intCols = lvw.ColumnHeaders.Count

    F = FreeFile
    Open sDataFile For Input As #F
    
    With lvw.ListItems
      While Not EOF(F)
        Line Input #F, sLine
      
        sItem = Split(sLine, vbTab)
        
        If UBound(sItem) <> intCols Then
          ReDim Preserve sItem(intCols)
        End If
        
        Set itemX = .Add(, sItem(0), sItem(1))
        
        For i = 2 To intCols
          itemX.SubItems(i - 1) = sItem(i)
        Next i
      Wend
    End With
    
    Close #F
  End If
End Sub

