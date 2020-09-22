VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQsort 
   Caption         =   "Quick Sort"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   Icon            =   "frmQsort.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Write Output"
      Height          =   315
      Left            =   360
      TabIndex        =   17
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2940
      TabIndex        =   13
      Top             =   2880
      Width           =   795
   End
   Begin VB.TextBox Text7 
      Height          =   255
      Left            =   3540
      TabIndex        =   9
      Top             =   1260
      Width           =   555
   End
   Begin VB.TextBox Text6 
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1260
      Width           =   555
   End
   Begin VB.TextBox Text5 
      Height          =   255
      Left            =   3540
      TabIndex        =   7
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox Text4 
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   3540
      TabIndex        =   5
      Top             =   540
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   540
      Width           =   555
   End
   Begin MSComDlg.CommonDialog getfile 
      Left            =   180
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   2280
      Width           =   6315
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Field 3"
      Height          =   195
      Left            =   1980
      TabIndex        =   16
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Field 2"
      Height          =   195
      Left            =   1980
      TabIndex        =   15
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Field 1"
      Height          =   195
      Left            =   1980
      TabIndex        =   14
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "End"
      Height          =   195
      Left            =   3600
      TabIndex        =   12
      Top             =   300
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Start"
      Height          =   195
      Left            =   2820
      TabIndex        =   11
      Top             =   300
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Columns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2820
      TabIndex        =   10
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "File to Sort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
   End
End
Attribute VB_Name = "frmQsort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sortArray(500000) As String, strArray(500000) As String
Dim Nrec As Long, sttime As Date, endtime As Date
Private Sub cmdopen_Click()

Dim intFile As Integer, I As Long
I = 1
ChDir "C:\My Documents"
getfile.ShowOpen
intFile = FreeFile
If getfile.FileName = "" Then Exit Sub
Text1.Text = getfile.FileName
Open getfile.FileName For Input As #intFile
While Not EOF(intFile)
    Line Input #intFile, strArray(I)
    I = I + 1
Wend
Close #intFile
Nrec = I - 1
'Debug.Print Nrec
End Sub

Private Sub cmdSort_Click()
Dim I As Long, s1, s2, s3, e1, e2, e3 As Integer
Dim mesg As String
s1 = Val(Text2.Text)
s2 = Val(Text4.Text)
s3 = Val(Text6.Text)
e1 = Val(Text3.Text)
e2 = Val(Text5.Text)
e3 = Val(Text7.Text)

If s1 = 0 Then
    For I = 1 To Nrec
        sortArray(I) = strArray(I)
    Next I
End If
If s1 > 0 And e1 = 0 Then
    For I = 1 To Nrec
        sortArray(I) = Mid(strArray(I), s1)
    Next I
ElseIf e1 > 0 And s1 > 0 Then
    For I = 1 To Nrec
        sortArray(I) = Mid(strArray(I), s1, e1 - s1 + 1)
    Next I
End If

If s2 > 0 And e2 = 0 Then
    For I = 1 To Nrec
        sortArray(I) = sortArray(I) + Mid(strArray(I), s2)
    Next I
ElseIf e2 > 0 And s2 > 0 Then
    For I = 1 To Nrec
        sortArray(I) = sortArray(I) + Mid(strArray(I), s2, e2 - s2 + 1)
    Next I
End If

If s3 > 0 And e3 = 0 Then
    For I = 1 To Nrec
        sortArray(I) = sortArray(I) + Mid(strArray(I), s3)
    Next I
ElseIf e3 > 0 And s3 > 0 Then
    For I = 1 To Nrec
        sortArray(I) = sortArray(I) + Mid(strArray(I), s3, e3 - s3 + 1)
    Next I
End If
sttime = Time

'Call QuickSort(1, Nrec)
Call Bubble_Sort(1, Nrec)

endtime = Time
mesg = CStr(sttime) + " to " + CStr(endtime) + " Sort Complete"
MsgBox mesg
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Dim I As Long
Open Text1.Text For Output As #9
For I = 1 To Nrec
    Print #9, strArray(I)
    
Next I
Close #9
MsgBox "Write Complete"
End Sub

Private Sub Bubble_Sort(S As Long, E As Long)
Dim I, J As Long
Dim sortTempItem, strTempItem As String
Me.AutoRedraw = True
For I = S To E - 1
Debug.Print I, E
    For J = I + 1 To E
        If sortArray(I) > sortArray(J) Then
            sortTempItem = sortArray(J)
            sortArray(J) = sortArray(I)
            sortArray(I) = sortTempItem
            strTempItem = strArray(J)
            strArray(J) = strArray(I)
            strArray(I) = strTempItem
        End If
    Next J
Next I
End Sub
Private Sub QuickSort(lonLower, lonUpper)

Dim lonRandomPivot As Long
Dim lonTempLower As Long
Dim lonTempUpper As Long
Dim strLastItem As String
Dim strTempItem As String
Dim sortTempItem As String
Dim sortLastItem As String
Randomize Timer

' Only sort if the lower boundary is lesser
' than the upper boundary.
If lonLower < lonUpper Then
    If lonUpper - lonLower = 1 Then
        ' Switch the upper and lower array items
        ' if the lower is greater than the upper.
        If sortArray(lonLower) > sortArray(lonUpper) Then
            sortTempItem = sortArray(lonUpper)
            sortArray(lonUpper) = sortArray(lonLower)
            sortArray(lonLower) = sortTempItem
            strTempItem = strArray(lonUpper)
            strArray(lonUpper) = strArray(lonLower)
            strArray(lonLower) = strTempItem
        End If
    Else
        ' Pick a random "pivot" item.
        lonRandomPivot = Int(Rnd _
            * (lonUpper - lonLower + 1)) + lonLower
        ' Switch the upper array item with the
        ' pivot item.
        sortTempItem = sortArray(lonUpper)
        sortArray(lonUpper) = sortArray(lonRandomPivot)
        sortArray(lonRandomPivot) = sortTempItem
        strTempItem = strArray(lonUpper)
        strArray(lonUpper) = strArray(lonRandomPivot)
        strArray(lonRandomPivot) = strTempItem

        ' Store the upper array item.
        sortLastItem = sortArray(lonUpper)
        strLastItem = strArray(lonUpper)
        Do
            ' Define the temporary upper and
            ' lower boundaries.
            lonTempUpper = lonUpper
            lonTempLower = lonLower
            ' Move down towards the pivot item,
            ' looping until the pivot item is greater
            ' than or equal to the temporary
            ' lower boundary.
            Do While (lonTempLower < lonTempUpper) And _
                (sortArray(lonTempLower) <= sortLastItem)
                    lonTempLower = lonTempLower + 1
            Loop
            ' Move up towards the pivot item, looping
            ' until the pivot item is less than
            ' or equal to the temporary upper
            ' boundary.
            Do While (lonTempUpper > lonTempLower) And _
                (sortArray(lonTempUpper) >= sortLastItem)
                    lonTempUpper = lonTempUpper - 1
            Loop
            ' If the pivot item hasn't been
            ' reached, then two of the items on
            ' either side of the pivot are out
            ' of order. If so, then switch them.
            If lonTempLower < lonTempUpper Then
                sortTempItem = sortArray(lonTempUpper)
                sortArray(lonTempUpper) = sortArray(lonTempLower)
                sortArray(lonTempLower) = sortTempItem
                strTempItem = strArray(lonTempUpper)
                strArray(lonTempUpper) = strArray(lonTempLower)
                strArray(lonTempLower) = strTempItem
               
            End If
        Loop While (lonTempLower < lonTempUpper)
        ' Switch the temporary lower boundary
        ' item with the original upper boundary
        ' item.
        sortTempItem = sortArray(lonTempLower)
        sortArray(lonTempLower) = sortArray(lonUpper)
        sortArray(lonUpper) = sortTempItem
        strTempItem = strArray(lonTempLower)
        strArray(lonTempLower) = strArray(lonUpper)
        strArray(lonUpper) = strTempItem
                
        ' Call the QuickSort routine again
        ' recursively, using new upper and
        ' lower boundaries.
        If (lonTempLower - lonLower) < (lonUpper - lonTempLower) Then
            QuickSort lonLower, lonTempLower - 1
            QuickSort lonTempLower + 1, lonUpper
        Else
            QuickSort lonTempLower + 1, lonUpper
            QuickSort lonLower, lonTempLower - 1
        End If
    End If
End If

End Sub


