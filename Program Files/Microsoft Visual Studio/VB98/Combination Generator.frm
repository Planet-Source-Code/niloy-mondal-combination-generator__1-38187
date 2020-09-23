VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Combination Generator"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Input String:-"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
MsgBox "Author:- Niloy Mondal. Email - niloygk@yahoo.com" & vbCrLf & "Special Thanks to Dhaval Faria for showing how to split a string."
End Sub

Private Sub cmdGenerate_Click()
Dim bytBYTE(100), strString As String
Dim length, i, j As Long
length = Len(Text1.Text)
List1.Clear
'Split the input sting in each single character and store in bytBYTE array
For i = 0 To length - 1
    Text1.SelStart = i
    Text1.SelLength = 1
    bytBYTE(i) = Text1.SelText
Next i
'This loops generates the combinations and adds them to list1
'The number of combinations is equal to (2^length)-1
For i = 1 To (2 ^ length) - 1
    strString = strString & i & ")   "
    For j = 0 To length
        If i And 2 ^ j Then strString = strString & bytBYTE(j)
    Next j
    List1.AddItem strString
    strString = ""
Next i
End Sub
