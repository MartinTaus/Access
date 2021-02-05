# Access

Function acCompactRepair(ByVal pthfn As String, Optional doEnable As Boolean = True) As Boolean
'v2.02 2013-11-28 16:22
'if doEnable = True will compact and repair pthfn
'if doEnable = False will then disable auto compact on pthfn

On Error GoTo CompactFailed

Dim A As Object
Set A = CreateObject("Access.Application")
With A
    .OpenCurrentDatabase pthfn
    .SetOption "Auto compact", True
    .CloseCurrentDatabase
    If doEnable = False Then
        .OpenCurrentDatabase pthfn
        .SetOption "Auto compact", doEnable
    End If
    .Quit
End With
Set A = Nothing
acCompactRepair = True
Exit Function
CompactFailed:
End Function


'source: http://www.ammara.com/access_image_faq/recursive_folder_search.html
'tweaked slightly for error handling

Private Function RecursiveDir(colFiles As Collection, _
                             strFolder As String, _
                             strFileSpec As String, _
                             bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add files in strFolder matching strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
On Error Resume Next
    strTemp = ""
    strTemp = Dir(strFolder & strFileSpec)
On Error GoTo 0
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
On Error Resume Next
        strTemp = ""
        strTemp = Dir(strFolder, vbDirectory)
On Error GoTo 0
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders'
        For Each vFolderName In colFolders
            Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Function

Private Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function

Sub CompactRepair()
  Dim control As Office.CommandBarControl
  Set control = CommandBars.FindControl( Id:=2071 )
  control.accDoDefaultAction
End Sub
