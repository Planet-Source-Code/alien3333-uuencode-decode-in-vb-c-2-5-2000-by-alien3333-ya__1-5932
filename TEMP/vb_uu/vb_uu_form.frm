VERSION 5.00
Begin VB.Form vb_uu_form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UUencode/decode in VB (c) 2/5/2000 by Alien3333@yahoo.com"
   ClientHeight    =   3930
   ClientLeft      =   8835
   ClientTop       =   6390
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6.932
   ScaleMode       =   0  'User
   ScaleWidth      =   10.954
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UUdecode"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UUencode"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double click file to add !"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double click file to remove !"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "File(s) Selected :"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3120
      Width           =   3735
   End
End
Attribute VB_Name = "vb_uu_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
' UUencode/decode in VB (c) 2/5/2000 by Alien3333@yahoo.com
' A small utility that uuencode and decode with a easy standard M$ GUI
'
' Well, I put this together in my 2 day vacation to celebrate Chinese New Year !
' All codes are standard as possible, ignore some of my clumsy C-style codes =)
' This small application show how to use filelistbox, listbox, drivelistbox, dirlistbox,
' Reading/Writing binary files, Reading text file line by line, Reading file in small portions,
' uuencode/decode with VB way, not C way of bit shifting, character and string manipulation,
' and all you can name it ...
'
' I learn VB through the Internet and MSDN, so some of the codes can look very familiar.
' I also read through uuencode.c and uudecode.c in LINUX to verify correctness.
'
' Enjoy !!!
'==================================================================================
Dim filename As String

Private Sub Form_Load()
Label1.Caption = Dir1.Path   ' Show path in Label.
File1.ReadOnly = True
File1.Archive = True
File1.Normal = True
File1.System = True
File1.Hidden = True
Label3.Caption = List1.ListCount
Label3.Refresh
File1.Refresh
End Sub


Private Sub Command1_Click()
Dim i As Long

Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

If List1.ListCount <> 0 Then
For i = 0 To List1.ListCount - 1
Label4.Caption = "UUencoding " + List1.List(i) + " ..."
Label4.Refresh
uuencode List1.List(i), List1.List(i) + ".uue"
Next

Label4.Caption = "Done UUencoding (" + CStr(List1.ListCount) + ") file(s) !"
Label4.Refresh
Command3_Click
End If


Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
File1.Refresh
End Sub

Private Sub Command2_Click()
Dim i As Long

Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

If List1.ListCount <> 0 Then
For i = 0 To List1.ListCount - 1
Label4.Caption = "UUdecoding " + List1.List(i) + " ..."
Label4.Refresh
uudecode List1.List(i)
Next

Label4.Caption = "Done UUdecoding (" + CStr(List1.ListCount) + ") file(s) !"
Label4.Refresh
Command3_Click
End If

Command1.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
File1.Refresh
End Sub

Private Sub Command3_Click()

Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False

' Empty list box.
List1.Clear
Label3.Caption = List1.ListCount
Label3.Refresh

Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
File1.Refresh
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive   ' Set directory path.
Label3.Caption = List1.ListCount
Label3.Refresh
File1.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path   ' Set file path.
Label3.Caption = List1.ListCount
Label3.Refresh
File1.Refresh
End Sub

Private Sub File1_PathChange()
Label1.Caption = Dir1.Path  ' Show path in Label.
Label3.Caption = List1.ListCount
Label3.Refresh
File1.Refresh
End Sub

Private Sub File1_Click()

' Display the selected filename when double-clicked.

If Right(Dir1.Path, 1) = "\" Then
filename = Dir1.Path + File1.filename
Else
filename = Dir1.Path + "\" + File1.filename
End If

Label1.Caption = filename
Label1.Refresh

Label3.Caption = List1.ListCount
Label3.Refresh
End Sub

Private Sub File1_DblClick()

' Display the selected filename when double-clicked.

If Right(Dir1.Path, 1) = "\" Then
filename = Dir1.Path + File1.filename
Else
filename = Dir1.Path + "\" + File1.filename
End If

Label1.Caption = filename
Label1.Refresh

If inlist(filename) = False Then
List1.AddItem filename
End If

Label3.Caption = List1.ListCount
Label3.Refresh
End Sub

Private Sub List1_DblClick()
   Dim Ind As Integer
   Ind = List1.ListIndex   ' Get index.
   ' Make sure list item is selected.
   If Ind >= 0 Then
      ' Remove it from list box.
      List1.RemoveItem Ind
      ' Display number.
      Label3.Caption = List1.ListCount
      Label3.Refresh
      Else
      Beep
      End If
      
Label3.Caption = List1.ListCount
Label3.Refresh
End Sub


Function inlist(ByVal filename As String) As Boolean
Dim i As Long
inlist = False
For i = 0 To List1.ListCount
If (StrComp(filename, List1.List(i))) = 0 Then
inlist = True
End If
Next
End Function

Function plain_filename(ByVal filename As String) As String
Dim length As Long
Dim temp_str As String
Dim i As Long
Dim done As Boolean
Dim left_str As String
Dim right_str As String

done = False

temp_str = filename

length = Len(temp_str)

i = length

While done = False And i <> 0
right_str = Right(temp_str, 1)
If right_str = "\" Then
done = True
End If
temp_str = Left(temp_str, i - 1)
i = i - 1
Wend

plain_filename = Right(filename, length - Len(temp_str) - 1)

'MsgBox plain_filename

End Function


Function result_filename(ByVal filename As String) As String
Dim length As Long
Dim temp_str As String
Dim i As Long
Dim done As Boolean
Dim left_str As String
Dim right_str As String

done = False

temp_str = filename

length = Len(temp_str)

i = length

While done = False And i <> 0
right_str = Right(temp_str, 1)
If right_str = " " Then
done = True
End If
temp_str = Left(temp_str, i - 1)
i = i - 1
Wend

result_filename = Right(filename, length - Len(temp_str) - 1)

End Function

Function result_dirname(ByVal filename As String) As String
Dim length As Long
Dim temp_str As String
Dim i As Long
Dim done As Boolean
Dim left_str As String
Dim right_str As String

done = False

temp_str = filename

length = Len(temp_str)

i = length

While done = False And i <> 0
right_str = Right(temp_str, 1)
If right_str = "\" Then
done = True
End If
temp_str = Left(temp_str, i - 1)
i = i - 1
Wend

result_dirname = Left(filename, Len(temp_str) + 1)

End Function


' 3 character into 4 characters
Private Function Encode(ByVal instring As String) As String
Dim outstring As String
Dim i As Integer

Dim y0 As Integer
Dim y1 As Integer
Dim y2 As Integer
Dim y3 As Integer


Dim x0 As Integer
Dim x1 As Integer
Dim x2 As Integer

' Very Important pad 3 byte to make 3 multiple
' This can add 1 or 2 extra NULL character to the end of the file
' Resulting a different file size, but no harm, for easier implementation

If Len(instring) Mod 3 <> 0 Then
instring = instring & String(3 - Len(instring) Mod 3, Chr$(0))
End If


For i = 1 To Len(instring) Step 3
x0 = Asc(Mid(instring, i, 1))
x1 = Asc(Mid(instring, i + 1, 1))
x2 = Asc(Mid(instring, i + 2, 1))

'MsgBox "x0=" + CStr(x0) + ", " + "x1=" + CStr(x1) + ", " + "x2=" + CStr(x2)

y0 = (x0 \ 4 + 32)
y1 = ((x0 Mod 4) * 16) + (x1 \ 16 + 32)
y2 = ((x1 Mod 16) * 4) + (x2 \ 64 + 32)
y3 = (x2 Mod 64) + 32

If (y0 = 32) Then y0 = 96
If (y1 = 32) Then y1 = 96
If (y2 = 32) Then y2 = 96
If (y3 = 32) Then y3 = 96

'MsgBox "y0=" + CStr(y0) + ", " + "y1=" + CStr(y1) + ", " + "y2=" + CStr(y2) + ", " + "y3=" + CStr(y3)

outstring = outstring + Chr(y0) + Chr(y1) + Chr(y2) + Chr(y3)

Next i
Encode = outstring

End Function


' 4 character into 3 characters
Private Function Decode(ByVal instring As String) As String
Dim outstring As String

Dim i As Integer

Dim x0 As Integer
Dim x1 As Integer
Dim x2 As Integer

Dim y0 As Integer
Dim y1 As Integer
Dim y2 As Integer
Dim y3 As Integer



For i = 1 To Len(instring) Step 4
y0 = Asc(Mid(instring, i, 1))
y1 = Asc(Mid(instring, i + 1, 1))
y2 = Asc(Mid(instring, i + 2, 1))
y3 = Asc(Mid(instring, i + 3, 1))

If (y0 = 96) Then y0 = 32
If (y1 = 96) Then y1 = 32
If (y2 = 96) Then y2 = 32
If (y3 = 96) Then y3 = 32

'MsgBox "y0=" + CStr(y0) + ", " + "y1=" + CStr(y1) + ", " + "y2=" + CStr(y2) + ", " + "y3=" + CStr(y3)

x0 = ((y0 - 32) * 4) + ((y1 - 32) \ 16)
x1 = ((y1 Mod 16) * 16) + ((y2 - 32) \ 4)
x2 = ((y2 Mod 4) * 64) + (y3 - 32)

'MsgBox "x0=" + CStr(x0) + ", " + "x1=" + CStr(x1) + ", " + "x2=" + CStr(x2)

outstring = outstring + Chr(x0) + Chr(x1) + Chr(x2)
Next i

Decode = outstring
End Function


Private Sub uuencode(ByVal filename1 As String, ByVal filename2 As String)
Dim portion_size As Long

portion_size = 45

'open the original file as binary read
Open (filename1) For Binary Access Read Shared As #1

'open the target file as binary write
Open (filename2) For Binary Access Write As #2

'for standard uuencode compatibility
Put #2, , "begin 644 " + plain_filename(filename1) + vbCrLf

'total number of full sized portion with "portion_size" bytes
        total = LOF(1) \ portion_size

'remain hold the remaining bytes toward end of file
        remain = LOF(1) Mod portion_size

'prepare instring to read "portion_size" bytes at a time
        instring$ = String(portion_size, 0)

'current file position
        current = 1

'for loop to read the portion one by one
         
        For i = 1 To total
          Get #1, current, instring$

'use the ENC() for standard uuencode compatibility, pad "M"
              
                Put #2, , ENC(portion_size) + Encode(instring$) + vbCrLf
                  current = current + portion_size
        Next
        
        instring = String(remain, 0)
       
'get the remaining bytes toward end of the file
        Get #1, current, instring$
        
'get the remaining bytes size and calculate ENC() for the last line
        
        Put #2, , ENC(LOF(1) - current + 1) + Encode(instring$) + vbCrLf
                
        Close #1
 
'put "end" for standard uuencode compatibility
       
        Put #2, , ENC(0) + vbCrLf + "end" + vbCrLf
        Close #2
End Sub


Private Sub uudecode(ByVal filename1 As String)
Dim instring As String
Dim outstring As String
Dim has_begin As Boolean
Dim filename2 As String


has_begin = False

Open (filename1) For Input As #1 ' file opened for reading

While Not EOF(1) And has_begin = False
Line Input #1, instring

If Left(instring, 6) = "begin " Then
has_begin = True
End If

Wend

If has_begin = True Then
'MsgBox result_filename(instring)
'MsgBox result_dirname(filename1)
Else
MsgBox filename1 + " : No begin line !"
End If
                                
If has_begin = True Then
filename2 = result_dirname(filename1) + result_filename(instring)
Open (filename2) For Binary Access Write Shared As #2

While Not EOF(1)
Line Input #1, instring



outstring = Right(instring, Len(instring) - 1)
'MsgBox Len(outstring)

If instring <> "end" Then
If Len(outstring) > 2 Then
Put #2, , Decode(outstring)
End If
End If

Wend

Close #2
End If
                                
'Print #2, outstring
Close #1

End Sub

Function ENC(ByVal c As Integer) As String
If c = 0 Then
ENC = "`"
Else
c = c + 32
ENC = Chr(c)
End If
End Function
