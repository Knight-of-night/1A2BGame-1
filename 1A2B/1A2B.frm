VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "1A2B������"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6840
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "��ʼ"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2340
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   150
   End
   Begin VB.Label Label2 
      Caption         =   "��ʷ��¼��"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ӭ����1A2B������~"
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 4) As Integer

Private Sub Command1_Click()
    Dim b(1 To 4) As Integer
    Dim i, j, m, n As Integer
    For i = 1 To 4
        b(i) = Val(Mid(Text1.Text, i, 1))
    Next i
    For i = 1 To 4
        For j = 1 To 4
            If a(i) = b(j) Then
                If i = j Then
                    m = m + 1
                Else
                    n = n + 1
                End If
            End If
        Next j
    Next i
    Label4.Caption = Str(m) & "A" & Trim(Str(n)) & "B"
    List1.AddItem Text1.Text & "=" & Label4.Caption
    If m = 4 Then
        Label3.Caption = "̫���ˣ�"
        MsgBox "�������������̫�����ˡ�"
    End If
    If m = 0 Then
        Label3.Caption = "������̫����~"
    Else
        Label3.Caption = "�ܽӽ��ˣ������롣"
    End If
End Sub

Private Sub Command2_Click()
    Dim i, j As Integer
    Command2.Caption = "���¿�ʼ"
    Command1.Enabled = True
    Label1.Caption = "��������Ҫ�µ�����"
    Text1.Enabled = True
    Label3.Caption = "������һ�����ظ����ֵ���λ�����쿪ʼ�°�~"
    For i = 1 To 4
        Randomize
        a(i) = Int(Rnd * 9)
        If i >= 2 Then
            For j = 1 To i - 1
                If a(i) = a(j) Then i = i - 1
            Next j
        End If
    Next i
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
