Attribute VB_Name = "ContextMenu"

Public Function ShellContextMenu(hWnd As Long, objLB As Control, X As Single, Y As Single, Shift As Integer, Filename As String) As String
  
  Dim pt As POINTAPI               ' screen location of the cursor
  Dim iItem As Integer                ' listbox index of the selected item (item under the cursor)
  Dim cItems As Integer             ' count of selected items
  Dim i As Integer                       ' counter
  Dim asPaths() As String           ' array of selected items' paths (zero based)
  Dim apidlFQs() As Long           ' array of selected items' fully qualified pidls (zero based)
  Dim isfParent As IShellFolder   ' selected items' parent shell folder
  Dim apidlRels() As Long           ' array of selected items' relative pidls (zero based)
  Dim commandrun As String
  
  pt.X = X \ Screen.TwipsPerPixelX
  pt.Y = Y \ Screen.TwipsPerPixelY
  Call ClientToScreen(objLB.hWnd, pt)
  
  ReDim asPaths(0)
  ReDim apidlFQs(0)
  ReDim apidlRels(0)
    asPaths = Split(Filename, Chr(0))
    cItems = UBound(asPaths) + 1
    
  If Len(asPaths(0)) Then
    For i = 0 To cItems - 1
      ReDim Preserve apidlFQs(i)
      apidlFQs(i) = GetPIDLFromPath(hWnd, asPaths(i))
    Next
    
    If apidlFQs(0) Then
      Set isfParent = GetParentIShellFolder(apidlFQs(0))
      If (isfParent Is Nothing) = False Then
        For i = 0 To cItems - 1
          ReDim Preserve apidlRels(i)
          apidlRels(i) = GetItemID(apidlFQs(i), GIID_LAST)
        Next
        
        If apidlRels(0) Then
          Call SubClass(hWnd, AddressOf WndProc)
          
          commandrun = ShowShellContextMenu(hWnd, isfParent, cItems, apidlRels(0), pt, False)
          If commandrun <> Empty Then
                ShellContextMenu = commandrun
          End If
          
          Call UnSubClass(hWnd)
        End If   ' apidlRels(0)

        For i = 0 To cItems - 1
          Call MemAllocator.Free(ByVal apidlRels(i))
        Next
      End If
      
      For i = 0 To cItems - 1
        Call MemAllocator.Free(ByVal apidlFQs(i))
      Next
      
    End If   ' apidlFQs(0)
  End If   ' Len(asPaths(0))
  
End Function
