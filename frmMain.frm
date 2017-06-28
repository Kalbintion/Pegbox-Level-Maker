VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rockbox Pegbox Level Maker"
   ClientHeight    =   6255
   ClientLeft      =   7755
   ClientTop       =   4065
   ClientWidth     =   5205
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   Begin VB.CheckBox chkGrid 
      BackColor       =   &H00000000&
      Caption         =   "Grid"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Height          =   1695
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMain.frx":0000
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Line Line18 
      X1              =   296
      X2              =   296
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line17 
      X1              =   272
      X2              =   272
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line16 
      X1              =   248
      X2              =   248
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line15 
      X1              =   224
      X2              =   224
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line14 
      X1              =   200
      X2              =   200
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line13 
      X1              =   176
      X2              =   176
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line12 
      X1              =   152
      X2              =   152
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line11 
      X1              =   128
      X2              =   128
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line10 
      X1              =   104
      X2              =   104
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line9 
      X1              =   80
      X2              =   80
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line8 
      X1              =   56
      X2              =   56
      Y1              =   40
      Y2              =   232
   End
   Begin VB.Line Line7 
      X1              =   32
      X2              =   320
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Line Line6 
      X1              =   32
      X2              =   320
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line5 
      X1              =   32
      X2              =   320
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line4 
      X1              =   32
      X2              =   320
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line3 
      X1              =   32
      X2              =   320
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line Line2 
      X1              =   32
      X2              =   320
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line Line1 
      X1              =   32
      X2              =   320
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   7
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   95
      Left            =   4440
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   94
      Left            =   4080
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   93
      Left            =   3720
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   92
      Left            =   3360
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   91
      Left            =   3000
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   90
      Left            =   2640
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   89
      Left            =   2280
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   88
      Left            =   1920
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   87
      Left            =   1560
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   86
      Left            =   1200
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   85
      Left            =   840
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   84
      Left            =   480
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   83
      Left            =   4440
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   82
      Left            =   4080
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   81
      Left            =   3720
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   80
      Left            =   3360
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   79
      Left            =   3000
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   78
      Left            =   2640
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   77
      Left            =   2280
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   76
      Left            =   1920
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   75
      Left            =   1560
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   74
      Left            =   1200
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   73
      Left            =   840
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   72
      Left            =   480
      Top             =   2760
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   71
      Left            =   4440
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   70
      Left            =   4080
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   69
      Left            =   3720
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   68
      Left            =   3360
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   67
      Left            =   3000
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   66
      Left            =   2640
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   65
      Left            =   2280
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   64
      Left            =   1920
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   63
      Left            =   1560
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   62
      Left            =   1200
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   61
      Left            =   840
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   60
      Left            =   480
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   59
      Left            =   4440
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   58
      Left            =   4080
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   57
      Left            =   3720
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   56
      Left            =   3360
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   55
      Left            =   3000
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   54
      Left            =   2640
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   53
      Left            =   2280
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   52
      Left            =   1920
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   51
      Left            =   1560
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   50
      Left            =   1200
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   49
      Left            =   840
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   48
      Left            =   480
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   47
      Left            =   4440
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   46
      Left            =   4080
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   45
      Left            =   3720
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   44
      Left            =   3360
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   43
      Left            =   3000
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   42
      Left            =   2640
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   41
      Left            =   2280
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   40
      Left            =   1920
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   39
      Left            =   1560
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   38
      Left            =   1200
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   37
      Left            =   840
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   36
      Left            =   480
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   35
      Left            =   4440
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   34
      Left            =   4080
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   33
      Left            =   3720
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   32
      Left            =   3360
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   31
      Left            =   3000
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   30
      Left            =   2640
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   29
      Left            =   2280
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   28
      Left            =   1920
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   27
      Left            =   1560
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   26
      Left            =   1200
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   25
      Left            =   840
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   24
      Left            =   480
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   23
      Left            =   4440
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   22
      Left            =   4080
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   21
      Left            =   3720
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   20
      Left            =   3360
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   19
      Left            =   3000
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   18
      Left            =   2640
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   17
      Left            =   2280
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   16
      Left            =   1920
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   15
      Left            =   1560
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   14
      Left            =   1200
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   13
      Left            =   840
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   12
      Left            =   480
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   11
      Left            =   4440
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   10
      Left            =   4080
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   9
      Left            =   3720
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   8
      Left            =   3360
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   7
      Left            =   3000
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   6
      Left            =   2640
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   5
      Left            =   2280
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   4
      Left            =   1920
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   3
      Left            =   1560
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   2
      Left            =   1200
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   1
      Left            =   840
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgField 
      Height          =   360
      Index           =   0
      Left            =   480
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   6
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   5
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   4
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   2
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   1
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgSel 
      Height          =   360
      Left            =   720
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgTool 
      Height          =   360
      Index           =   0
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Toolbox:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkGrid_Click()
    For Each obj In Me.Controls
        If VarType(obj) = 11 And Left$(obj.Name, 4) = "Line" Then
            obj.Visible = chkGrid.Value
        End If
    Next
End Sub

Private Sub cmdAbout_Click()
    MsgBox "This tool was developed to make the easy of creating levels for the " & vbCrLf & _
           "Pegbox game in Rockbox firmware as easy as possible without having to" & vbCrLf & _
           "memorize the numbers and what shape they represent.", vbOKOnly, "About"
End Sub

Private Sub cmdClear_Click()
    ' Load Default Selection
    imgSel.Picture = imgTool(0).Picture
    imgSel.Tag = 0

    ' Load Default Floor
    For i = 0 To 95
        imgField(i).Picture = imgTool(0).Picture
        imgField(i).Tag = 0
    Next
End Sub

Private Sub cmdParse_Click()
    Dim s As String, n() As String
    
    s = txtCode.Text
    s = Replace(s, ",", " ")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    s = Replace(s, vbCrLf, "  ")
    While InStr(1, s, "  ")
        s = Replace(s, "  ", " ")
    Wend
    s = Mid$(s, 2, Len(s) - 2)
    s = Replace(s, " ", ",")
    
    Clipboard.Clear
    Clipboard.SetText s
    
    n = Split(s, ",")
    If UBound(n) <> 95 Then
        MsgBox "Error! Invalid Amount of locations.", vbOKOnly + vbCritical, "Error!"
        Exit Sub
    End If
    
    For i = 0 To 95
        If Val(n(i)) < 0 Or Val(n(i)) > 7 Then
            MsgBox "Error! Invalid tool type at index " & i & ". Expected 0 through 7. Encountered " & Val(n(i)) & "."
            
            ' Load Default Floor
            For e = 0 To 95
                imgField(e).Picture = imgTool(0).Picture
                imgField(e).Tag = 0
            Next
            
            Exit Sub
        End If
        imgField(i).Picture = imgTool(Val(n(i))).Picture
        imgField(i).Tag = Val(n(i))
    Next
    UpdateCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode >= vbKey1 And KeyCode <= vbKey8 Then
        imgTool_Click KeyCode - vbKey1
    ElseIf KeyCode >= vbKeyNumpad1 And KeyCode <= vbKeyNumpad8 Then
        imgTool_Click KeyCode - vbKeyNumpad1
    End If
End Sub

Private Sub Form_Load()
    ' Load Toolbox
    For i = 0 To 7
        imgTool(i).Picture = LoadPicture(App.Path & "/imgs/" & i & ".kdk")
        imgTool(i).ToolTipText = "Hotkey: " & Chr(i + vbKey1)
    Next
    
    ' Load Default Selection
    imgSel.Picture = imgTool(0).Picture
    imgSel.Tag = 0

    ' Load Default Floor
    For i = 0 To 95
        imgField(i).Picture = imgTool(0).Picture
        imgField(i).Tag = 0
    Next
End Sub

Private Sub imgField_Click(Index As Integer)
    imgField(Index).Picture = imgSel.Picture
    imgField(Index).Tag = imgSel.Tag
    Call UpdateCode
End Sub

Private Sub imgTool_Click(Index As Integer)
    imgSel.Picture = imgTool(Index).Picture
    imgSel.Tag = Index
End Sub

Private Sub UpdateCode()
    Dim s As String, c As Integer
    
    s = "    {{"
    For i = 0 To 95
        s = s & imgField(i).Tag
        c = c + 1
        If c = 12 And i <> 95 Then
            c = 0
            s = s & ",}," & vbCrLf & "     {"
        ElseIf i = 95 Then
            s = s & ",}},"
        Else
            s = s & ", "
        End If
    Next
    txtCode.Text = s
End Sub
