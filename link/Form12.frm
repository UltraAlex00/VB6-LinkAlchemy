VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Linker"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox check_cmd 
      Caption         =   "Run in CMD"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton btn_remove 
      Caption         =   "-"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton btn_add 
      Caption         =   "+"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton btn_exit 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton btn_ok 
      Caption         =   "Compile"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ListBox args_list 
      Height          =   1410
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      Style           =   1  'Kontrollkästchen
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub args_list_DblClick()
    
    Dim index As Long
    Dim result As String

    
    index = args_list.ListIndex
    If index < 0 Then Exit Sub
    args_list.Selected(index) = Not args_list.Selected(index)
    
    result = Trim$(InputBox("Parameter:", , args_list.List(index)))
    
    If LenB(result) > 0 Then
        args_list.List(index) = result
    End If
    
End Sub

Private Sub btn_add_Click()
    
    Dim result As String
    
    result = Trim$(InputBox("Parameter:"))
    If LenB(result) > 0 Then
        args_list.AddItem result
    End If
    
End Sub

Private Sub btn_remove_Click()
    
    Dim index As Long
    
    index = args_list.ListIndex
    If index <> -1 Then
        args_list.RemoveItem index
    End If
    
End Sub

Private Sub btn_exit_Click()
    Unload Me
End Sub

Private Sub btn_ok_Click()
    
    Dim i As Long
    Dim j As Long
    Dim args() As String
    Dim command As String
    
    For i = 0 To args_list.ListCount - 1
        If args_list.Selected(i) Then
            ReDim Preserve args(j)
            args(j) = args_list.List(i)
            j = j + 1
        End If
    Next i
    
    command = App.Path & "\LINK2.EXE " & Join(args)
    
    If check_cmd.Value = 1 Then
        Shell Environ$("ComSpec") & " /c " & command & " & pause", vbNormalFocus
    Else
        Shell command
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim i As Long
    Dim args() As String
    
    If LenB(command$) = 0 Then Exit Sub
    
    args = SplitParams(Fix_Spaces(command$))
    
    For i = 0 To UBound(args)
        args_list.AddItem args(i)
        args_list.Selected(i) = True
    Next i
    
    args_list.ListIndex = 0
    
End Sub

Public Function SplitParams(ByVal Text As String) As String()
    Dim result() As String
    Dim current As String
    Dim i As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim count As Long
    
    ReDim result(0)
    
    For i = 1 To Len(Text)
        ch = Mid$(Text, i, 1)
        
        Select Case ch
            
            Case """"
                inQuotes = Not inQuotes
                current = current & ch   ' <<< KEEP the quote
                
            Case " "
                If inQuotes Then
                    current = current & ch
                Else
                    If Len(current) > 0 Then
                        result(count) = current
                        count = count + 1
                        ReDim Preserve result(0 To count)
                        current = ""
                    End If
                End If
                
            Case Else
                current = current & ch
                
        End Select
    Next i
    
    ' Add final token
    If Len(current) > 0 Then
        result(count) = current
    ElseIf count > 0 Then
        ReDim Preserve result(0 To count - 1)
    End If
    
    SplitParams = result
End Function

Public Function Fix_Spaces(ByVal str As String) As String
    
    Dim i As Long
    Dim j As Long
    Dim isStr As Boolean
    
    Const SPACE_CHAR = " "
    
    Fix_Spaces = Space$(Len(str))
    j = 1
    
    For i = 1 To Len(str)
        If (Mid$(str, i, 1) = SPACE_CHAR) And Not isStr Then
            If Mid$(str, i + 1, 1) <> SPACE_CHAR Then
                Mid$(Fix_Spaces, j, 1) = SPACE_CHAR
                j = j + 1
            End If
        ElseIf Mid$(str, i, 1) = """" Then
            isStr = Not isStr
        
            Mid$(Fix_Spaces, j, 1) = Mid$(str, i, 1)
            j = j + 1
        Else
            Mid$(Fix_Spaces, j, 1) = Mid$(str, i, 1)
            j = j + 1
        End If
    Next i
    
    Fix_Spaces = RTrim$(Fix_Spaces)
    
End Function




















