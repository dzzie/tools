VERSION 5.00
Begin VB.UserControl pBar 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   ScaleHeight     =   750
   ScaleWidth      =   3270
   ToolboxBitmap   =   "pBar.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   165
      ScaleHeight     =   435
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   105
      Width           =   2970
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblPercent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -15
         TabIndex        =   2
         Top             =   75
         Width           =   2910
      End
      Begin VB.Label lblPb 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Height          =   435
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3000
      End
   End
End
Attribute VB_Name = "pBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private showPercent As Boolean

Public Property Let ShowPercentage(t As Boolean)
    showPercent = t
End Property
Public Property Get ShowPercentage() As Boolean
    ShowPercentage = showPercent
End Property


Private Sub UserControl_Initialize()
 Call ResetPb
End Sub

Private Sub UserControl_Resize()
    Picture1.Width = UserControl.Width - 200
    Picture1.Height = UserControl.Height - 200
    lblPb.Height = Picture1.Height
    lblPb.Width = Picture1.Width
    lblPercent.Width = Picture1.Width
    lblPercent.Top = Picture1.Top + ((0.5 * Picture1.Height) - 250)
End Sub

Public Sub ResetPb()
    lblPb.Width = 0
    lblPb.Visible = False
    lblPercent.Caption = Empty
End Sub

Public Sub SetPbPercent(percent, Optional ShowMessage As String = Empty)
        DoEvents
        On Error Resume Next
        If Not lblPb.Visible Then lblPb.Visible = True
        If InStr(percent, ".") Then percent = Round(percent, 0)
        If showPercent Or ShowMessage <> Empty Then
            If Not lblPercent.Visible Then lblPercent.Visible = True
            If ShowMessage <> Empty Then lblPercent = ShowMessage _
              Else lblPercent = percent & " %"
            lblPercent.Refresh
        End If
        If Len(percent) = 1 Then
            decmilpercent = Round(".0" & percent, 2)
        ElseIf Len(percent) = 2 Then
            decmilpercent = Round("." & percent, 2)
        End If
        lblPb.Width = Int(Picture1.Width * decmilpercent)
        lblPb.Refresh
        DoEvents
End Sub

Public Sub SetPbDecimal(decimalpercent, Optional ShowMessage As String = Empty)
        DoEvents
        On Error Resume Next
        If decimalpercent > 1 Then Exit Sub
        lblPb.Visible = True
        If showPercent Or ShowMessage <> Empty Then
            lblPercent.Visible = True
            If ShowMessage <> Empty Then lblPercent = ShowMessage _
              Else lblPercent = (decimalpercent * 100) & " %"
            lblPercent.Refresh
        End If
        lblPb.Width = Int(Picture1.Width * decimalpercent)
        lblPb.Refresh
        DoEvents
End Sub

