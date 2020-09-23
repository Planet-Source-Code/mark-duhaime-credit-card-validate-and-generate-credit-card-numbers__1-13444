VERSION 5.00
Begin VB.Form Validate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validate and generate credit card number"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "Validate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   520
      Left            =   3150
      ScaleHeight     =   465
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Card Type to generate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
      Begin VB.OptionButton opt 
         Caption         =   "Discover"
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
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         Caption         =   "American Express"
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         Caption         =   "Mastercard"
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton opt 
         Caption         =   "Visa"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.TextBox txt_Verify 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txt_Number 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Enter credit card number here to verify, or click generate to generate a random credit card number."
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Generate a valid visa credit card number."
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Verify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter credit card number here to verify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "Validate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dim index for option buttons
Dim vbIndex As Integer

'the user pressed the verify button
'perform some checks
'and pass the number to the verify function
Private Sub cmdCheck_Click()
    Dim vbNumber As String
    Dim vbInstr As Integer
    Dim vbTemp As String
    Dim vbNumber2 As String
    
    'If the text box is blank msg and exit
    If IsNull(txt_Number) Then
        MsgBox "Please enter a number.", vbExclamation + vbOKOnly, "Enter a number."
        Exit Sub
    Else ' otherwise assign it to variable string for verification
        vbNumber = txt_Number
    End If
    'turn the picture control on
    Picture1.Visible = True
    'vbInstr will check for "-" in the string and parse it out
    vbInstr = 1
    'initialize our temp variable string to null
    vbTemp = ""
    While vbInstr > 0 'parse entire string for "-"
        vbInstr = InStr(vbNumber, "-")
        If vbInstr > 0 Then
            'end of string
            vbNumber2 = Left$(vbNumber, vbInstr - 1)
        Else ' vbNumber2 string will hold parsed string
            vbNumber2 = vbNumber
        End If
        'assign the string
        vbNumber = Right$(vbNumber, Len(vbNumber) - vbInstr)
        vbTemp = vbTemp + vbNumber2
    Wend
    'if vbTemp has a length then assign it
    If Len(vbTemp) > 1 Then
        vbNumber = vbTemp
    End If
    'load and display the appropiate picture
    Select Case Left$(vbNumber, 1)
        Case "4" ' Visa
            Picture1.Picture = LoadPicture(App.Path + "\visa.gif")
        Case "5" 'Mastercard
            Picture1.Picture = LoadPicture(App.Path + "\mcard.gif")
        Case "6" 'Discover
            Picture1.Picture = LoadPicture(App.Path + "\discover.gif")
        Case "3" ' American Express
            Picture1.Picture = LoadPicture(App.Path + "\amex.gif")
        Case Else 'None
            Picture1.Visible = False
    End Select
    
    txt_Verify.SetFocus
    'verify the number
    If CheckCard(vbNumber) = False Then
        txt_Verify.Text = "Invalid Number."
    Else
        txt_Verify.Text = "Valid Number."
    End If
    
End Sub

'******************************************
'
'   function to perform LUHN formula
'   on the number
'   Returns: True is valid #
'             or False if Invalid #
'   Pass Number as the number to verify
'
'******************************************
Function CheckCard(CCNumber As String) As Boolean
    Dim vbCounter As Integer
    Dim vbInt As Integer
    Dim vbAnswer As Integer

    vbCounter = 1 ' variable to count the digits
    vbInt = 0 'temp sum variable

    'loop until all digits are calculated
    While vbCounter <= Len(CCNumber)
        'Perform LUHN check
        vbInt = Val(Mid$(CCNumber, vbCounter, 1))
        'check for odd position
        If Not (vbCounter Mod 2) Then
            vbInt = vbInt * 2
            If vbInt > 9 Then vbInt = vbInt - 9
        End If
        vbAnswer = vbAnswer + vbInt
        vbCounter = vbCounter + 1
    Wend

    vbAnswer = vbAnswer Mod 10 'divide by 10

    If vbAnswer = 0 Then ' card valid
        CheckCard = True
    Else
        CheckCard = False ' card invalid
    End If

End Function

'User pressed generate button
'do a random number then call
'verify to verify it.
Private Sub cmdGenerate_Click()
    Dim vbCounter As Integer
    Dim vbInt As Integer
    Dim vbAnswer As String
    Dim vbStart As String
    Dim vbFirst(11) As String
    Dim vbBool As Boolean
    Dim vbSet As Integer
    Dim vbLength As Integer
    
    vbBool = False

While vbBool = False
    Randomize
    vbInt = 0
    Picture1.Visible = True
    Select Case vbIndex
        Case 1 'Visa setup banks
            vbFirst(1) = "4032"
            vbFirst(2) = "4128"
            vbFirst(3) = "4250"
            vbFirst(4) = "4312"
            vbFirst(5) = "4421"
            vbFirst(6) = "4539"
            vbFirst(7) = "4556"
            vbFirst(8) = "4673"
            vbFirst(9) = "4722"
            vbFirst(10) = "4800"
            vbFirst(11) = "4833"
            vbSet = Int(11 * Rnd) + 1
            vbStart = vbFirst(vbSet)
            vbLength = 16
            Picture1.Picture = LoadPicture(App.Path + "\visa.gif")
        Case 2 'Mastercard
            vbFirst(1) = "510813" ' Bank
            vbSet = 1
            vbStart = vbFirst(vbSet)
            vbLength = 16
            Picture1.Picture = LoadPicture(App.Path + "\mcard.gif")
        Case 3 'American Express Banks
            vbFirst(1) = "372034"
            vbFirst(2) = "372407"
            vbFirst(3) = "372861"
            vbFirst(4) = "373227"
            vbSet = Int(4 * Rnd) + 1
            vbStart = vbFirst(vbSet)
            vbLength = 15
            Picture1.Picture = LoadPicture(App.Path + "\amex.gif")
        Case 4 'Discover
            vbFirst(1) = "601100" 'Bank
            vbSet = 1
            vbStart = vbFirst(vbSet)
            vbLength = 16
            Picture1.Picture = LoadPicture(App.Path + "\discover.gif")
    End Select
    
    While Len(vbStart) < vbLength
        vbInt = Int((9 * Rnd) + 1)
        If Not (vbCounter Mod 2) Then
            vbInt = vbInt * 2
            If vbInt > 9 Then
                vbInt = vbInt - 9
            End If
        ElseIf (vbCounter Mod 2) Then
            vbInt = vbInt * 2
            If vbInt > 9 Then
                vbInt = vbInt - 9
            End If
        End If
        
        vbStart = vbStart + LTrim$(Str$(vbInt))
        vbCounter = vbCounter + 1
    Wend
      
    If CheckCard(vbStart) = False Then
        vbInt = 0
        txt_Verify.SetFocus
        txt_Verify.Text = "Invalid Number."
    Else
        vbBool = True
        txt_Number.SetFocus
        txt_Number.Text = vbStart
        txt_Verify.SetFocus
        txt_Verify.Text = "Valid Number."
    End If
    
Wend
    
End Sub

Private Sub Form_Load()
    'Initialize Index
    vbIndex = 1
End Sub

'set index for selected option
'visa, mastercard, american express
'or discover
Private Sub opt_Click(Index As Integer)
    vbIndex = Index
End Sub
