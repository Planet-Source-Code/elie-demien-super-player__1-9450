VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "MOV PLAYER"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3330
   Icon            =   "FORM3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Brought to you by Brad Martinez
' http://members.aol.com/btmtz/vb

' Though this example has been optimized for speed,
' it's obviously not as efficient as it could be.
' Consider it a starting point...

' A liberal use of module level variables...
Dim PicHeight%, hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%

' These variables are allocated at the module level to save on
' stack space & on variable re-allocation time in SearchDirs().
' They could be declared as Static within their respective procs...
Dim WFD As WIN32_FIND_DATA, hItem&, hFile&

' SearchDirs() constants
Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46

Private Sub Form_Load()
    ScaleMode = vbPixels
   
    hLB& = List1.hWnd
    ' This speeds things a bit but will consume close to 6MB of memory...!!!
    SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
    Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
If Running% Then: Running% = False: Exit Sub
    
    Dim drvbitmask&, maxpwr%, pwr%
    On Error Resume Next
    
    FileSpec$ = "*.MOV"
    ' A parsing routine could be implemented here for
    ' multiple file spec searches, i.e. "*.bmp,*.wmf", etc.
    ' See the MS KB article Q130860 for information on how
    ' FindFirstFile() does not handle the "?" wildcard char correctly !!
    
    If Len(FileSpec$) = 0 Then Exit Sub
    
    MousePointer = 11
    Running% = True
    UseFileSpec% = True
    
    List1.Clear
    
    ' The following code block is used to demonstrate how
    ' to search every available drive on a system.
    ' See the "Browse for Folder" demo for an example of
    ' selecting individual drives or folders for a search.
    
    ' http://members.aol.com/btmtz/vb/browsdlg
    
    drvbitmask& = GetLogicalDrives()
    ' If GetLogicalDrives() succeeds, the return value is a bitmask representing
    ' the currently available disk drives. Bit position 0 (the least-significant bit)
    ' is drive A, bit position 1 is drive B, bit position 2 is drive C, and so on.
    ' If the function fails, the return value is zero.
    ' GetLogicalDriveStrings() could be used here instead,
    ' but it's string buffer would have to be parsed...
    If drvbitmask& Then
        
        ' Get & search each available drive
        maxpwr% = Int(Log(drvbitmask&) / Log(2))   ' a little math...
        For pwr% = 0 To maxpwr%
            If Running% And (2 ^ pwr% And drvbitmask&) Then _
                Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
        Next
    End If
    
    Running% = False
    UseFileSpec% = False
   
    MousePointer = 0

    
    Beep
     Form3.Caption = Form3.Caption & Space(1) & "found:" & List1.ListCount
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Cancels the search (Form1.KeyPreview = True)
    If KeyCode = vbKeyEscape And Running% Then Running% = False
End Sub

Private Sub Form_Resize()
    ' Much faster & cleaner than the Move Method...
    MoveWindow hLB&, 0, 0, ScaleWidth, ScaleHeight - PicHeight%, True
End Sub


' This is were it all happens...

' You can use the values in returned in the
' WIN32_FIND_DATA structure to virtually obtain any
' information you want for a particular folder or group of files.

' This recursive procedure is similar to the Dir$ function
' example found in the VB3 help file...

Private Sub SearchDirs(curpath$)  ' curpath$ is passed w/ trailing "\"

    ' These can't be static!!! They must be
    ' re-allocated on each recursive call.
    Dim dirs%, dirbuf$(), i%
    
    ' Display what's happening...
    ' A Timer could be used instead to display status at
    ' pre-defined intervals, saving on PictureBox redraw time...
    
    ' Allows the PictureBox to be redrawn
    ' & this proc to be cancelled by the user.
    ' It's not necessary to have this in the loop
    ' below since the loop works so fast...
    DoEvents
    If Not Running% Then Exit Sub
    
    ' This loop finds *every* subdir and file in the current dir
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        
        Do
            ' Tests for subdirs only...
            If (WFD.dwFileAttributes And vbDirectory) Then
                
                ' If not a  "." or ".." DOS subdir...
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    ' This is executed in the mnuFindFiles_Click()
                    ' call though it isn't used...
                    TotalDirs% = TotalDirs% + 1
                    ' This is the heart of a recursive proc...
                    ' Cache the subdirs of the current dir in the 1 based array.
                    ' This proc calls itself below for each subdir cached in the array.
                    ' (re-allocating the array only once every 10 itinerations improves speed)
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            
            ' File size and attribute tests can be used here, i.e:
            ' ElseIf (WFD.dwFileAttributes And vbHidden) = False Then  'etc...
            
            ' Get a total file count for mnuFolderInfo_Click()
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        
        ' Get the next subdir or file
        Loop While FindNextFile(hItem&, WFD)
        
        ' Close the search handle
        Call FindClose(hItem&)
    
    End If

    ' When UseFileSpec% is set mnuFindFiles_Click(),
    ' SearchFileSpec() is called & each folder must be
    ' searched a second time.
    If UseFileSpec% Then
        ' Turning off painting speeds things quite a bit...
        ' Speed also would be vastly improved if the redrawing
        ' & scrolling were placed in a Timer event...
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        ' Keeps the currently found items scrolled into view...
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    
    ' Recursively call this proc & iterate through each subdir cached above.
    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
  
End Sub

Private Sub SearchFileSpec(curpath$)   ' curpath$ is passed w/ trailing "\"
' This procedure *only*  finds files in the
' current folder that match the FileSpec$
    
    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        
        Do
            ' Use DoEvents here since we're loading a ListBox and
            ' there could be hundreds of files matching the FileSpec$
            DoEvents
            If Not Running% Then Exit Sub
            
            ' The ListBox's Sorted property is initially set to False.
            ' Set it to True and see how things slow down a bit...
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        
        ' Get the next file matching the FileSpec$
        Loop While FindNextFile(hFile&, WFD)
        
        ' Close the search handle
        Call FindClose(hFile&)
    
    End If

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If List1.ListCount = 0 Then
Exit Sub
End If

Form1.MediaPlayer1.FileName = List1.Text
Form1.MediaPlayer1.Play
End Sub


