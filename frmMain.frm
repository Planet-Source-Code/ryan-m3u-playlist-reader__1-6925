VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "M3U Playlist Loader"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   4275
      TabIndex        =   2
      Top             =   885
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5910
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   3300
      Left            =   0
      TabIndex        =   0
      Top             =   735
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   5821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   7232
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load M3U"
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   180
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "The path of each file in the m3u is stored in the listview items tag property"
      Height          =   600
      Left            =   1530
      TabIndex        =   4
      Top             =   60
      Width           =   2595
   End
   Begin VB.Label label1 
      Height          =   690
      Left            =   90
      TabIndex        =   3
      Top             =   4260
      Width           =   3990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
listview1.ListItems.Clear
Dim FilePath$, tmpString$, I%, FindComma%

CommonDialog1.Filter = "M3U Playlists (*.m3u)|*.m3u"
CommonDialog1.ShowOpen

FilePath$ = Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
Open CommonDialog1.FileName For Input As #1

Do
Line Input #1, thedata$
List1.AddItem thedata$
Loop While Not (EOF(1))
Close #1

For I% = 1 To List1.ListCount - 1
    If Left(List1.List(I%), 7) = "#EXTINF" Then 'see if a title was left
        FindComma% = InStr(1, List1.List(I%), ",")
        tmpString$ = Right(List1.List(I%), Len(List1.List(I%)) - FindComma%)
        If InStr(1, List1.List(I% + 1), "\") = 0 Then
            tmpString2$ = FilePath$ & List1.List(I% + 1)
        Else
            tmpString2$ = List1.List(I% + 1)
        End If
        listview1.ListItems.Add , , tmpString$
        listview1.ListItems.Item(listview1.ListItems.Count).Tag = tmpString2$
        I% = I% + 1
    Else 'no title, use filename
        listview1.ListItems.Add , , List1.List(I%)
        listview1.ListItems.Item(listview1.ListItems.Count).Tag = FilePath$ & List1.List(I%)
    End If
Next

End Sub

Private Sub ListView1_Click()
On Error Resume Next
label1 = listview1.SelectedItem.Tag
End Sub
