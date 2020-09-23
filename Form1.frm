VERSION 5.00
Begin VB.Form msWordDemo 
   Caption         =   "Microsoft Word in a Box"
   ClientHeight    =   10575
   ClientLeft      =   2160
   ClientTop       =   1635
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   13485
   Begin VB.CommandButton Command19 
      Caption         =   "Load Test File"
      Height          =   255
      Left            =   60
      TabIndex        =   61
      Top             =   420
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Show Hide Ruler"
      Height          =   255
      Left            =   7620
      TabIndex        =   59
      Top             =   1860
      Width           =   1665
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Hide Header-Footer"
      Height          =   255
      Left            =   7620
      TabIndex        =   58
      Top             =   1590
      Width           =   1665
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Show Header-Footer"
      Height          =   255
      Left            =   7620
      TabIndex        =   57
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bookmarks (Replaceable fields)"
      Height          =   2535
      Left            =   9300
      TabIndex        =   10
      Top             =   1050
      Width           =   4065
      Begin VB.ComboBox cbBookMarks 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   540
         Width           =   3915
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   0
         Left            =   1470
         TabIndex        =   14
         Text            =   "Makis Charalambous"
         Top             =   930
         Width           =   2505
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   1470
         TabIndex        =   13
         Text            =   "Limassol, Cyprus"
         Top             =   1170
         Width           =   2505
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   1470
         TabIndex        =   12
         Text            =   "123-1234567"
         Top             =   1410
         Width           =   2505
      End
      Begin VB.CommandButton cmdDoFilling 
         Caption         =   "Fill Fields"
         Height          =   315
         Left            =   210
         TabIndex        =   11
         Top             =   1980
         Width           =   3795
      End
      Begin VB.Label Label3 
         Caption         =   "Predefined Bookmarks"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   90
         TabIndex        =   56
         Top             =   300
         Width           =   3585
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   930
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   1170
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Telephone"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1410
         Width           =   1305
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Add Various "
      Height          =   1485
      Left            =   5550
      TabIndex        =   51
      Top             =   1050
      Width           =   2025
      Begin VB.CommandButton Command18 
         Caption         =   "Insert Page"
         Height          =   255
         Left            =   60
         TabIndex        =   60
         Top             =   1050
         Width           =   1875
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Blank line"
         Height          =   255
         Left            =   60
         TabIndex        =   54
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Insert Paragraph"
         Height          =   255
         Left            =   60
         TabIndex        =   53
         Top             =   510
         Width           =   1875
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Insert Picture"
         Height          =   255
         Left            =   60
         TabIndex        =   52
         Top             =   780
         Width           =   1875
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Manual Bookmarks"
      Height          =   1005
      Left            =   4470
      TabIndex        =   38
      Top             =   60
      Width           =   4785
      Begin VB.CommandButton Command8 
         Caption         =   "Bookmarks Count"
         Height          =   255
         Left            =   90
         TabIndex        =   45
         Top             =   690
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add Bookmark"
         Height          =   255
         Left            =   90
         TabIndex        =   44
         Top             =   210
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Set Bookmark"
         Height          =   255
         Left            =   90
         TabIndex        =   43
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txtBKlabel 
         Height          =   285
         Left            =   3930
         TabIndex        =   42
         ToolTipText     =   "Name of Bookmark to appear in document"
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtBk 
         Height          =   285
         Left            =   2550
         TabIndex        =   41
         ToolTipText     =   "Name of Bookmark"
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtSBK 
         Height          =   285
         Left            =   2550
         TabIndex        =   40
         ToolTipText     =   "Name of Bookmark"
         Top             =   450
         Width           =   795
      End
      Begin VB.TextBox txtSBKtext 
         Height          =   285
         Left            =   3930
         TabIndex        =   39
         ToolTipText     =   "Text that will replace bookmark"
         Top             =   450
         Width           =   795
      End
      Begin VB.Label lblBC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1920
         TabIndex        =   50
         Top             =   690
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Label"
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   49
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Index           =   5
         Left            =   1950
         TabIndex        =   48
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   225
         Index           =   6
         Left            =   1950
         TabIndex        =   47
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Text"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   46
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Add Table"
      Height          =   1005
      Left            =   2100
      TabIndex        =   32
      Top             =   60
      Width           =   2355
      Begin VB.CommandButton Command11 
         Caption         =   "Add Table"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2085
      End
      Begin VB.TextBox txtRow 
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Text            =   "2"
         Top             =   510
         Width           =   525
      End
      Begin VB.TextBox txtCol 
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Text            =   "2"
         Top             =   510
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Rows"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Cols"
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   36
         Top             =   540
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Files I/O"
      Height          =   1485
      Left            =   0
      TabIndex        =   27
      Top             =   1050
      Width           =   2025
      Begin VB.CommandButton Command1 
         Caption         =   "Load File"
         Height          =   255
         Left            =   90
         TabIndex        =   31
         Top             =   210
         Width           =   1785
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Insert File"
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   480
         Width           =   1785
      End
      Begin VB.CommandButton btnRsAction 
         Caption         =   "&Print"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   29
         ToolTipText     =   "Print Selected File"
         Top             =   1020
         Width           =   1785
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Save File"
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   750
         Width           =   1785
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "View Layouts"
      Height          =   1005
      Left            =   9270
      TabIndex        =   22
      Top             =   60
      Width           =   4095
      Begin VB.CommandButton cmbLayout 
         Caption         =   "Outline Layout"
         Height          =   255
         Index           =   4
         Left            =   1710
         TabIndex        =   55
         Top             =   450
         Width           =   1605
      End
      Begin VB.CommandButton cmbLayout 
         Caption         =   "Normal Layout"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   450
         Width           =   1605
      End
      Begin VB.CommandButton cmbLayout 
         Caption         =   "Web Layout"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   210
         Width           =   1605
      End
      Begin VB.CommandButton cmbLayout 
         Caption         =   "Print Layout"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   690
         Width           =   1605
      End
      Begin VB.CommandButton cmbLayout 
         Caption         =   "Reading Layout"
         Height          =   255
         Index           =   3
         Left            =   1710
         TabIndex        =   23
         Top             =   210
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Insert Text"
      Height          =   1035
      Left            =   30
      TabIndex        =   19
      Top             =   2550
      Width           =   9225
      Begin VB.CommandButton Command7 
         Caption         =   "Insert text ->"
         Height          =   765
         Left            =   120
         TabIndex        =   21
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         Height          =   795
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Top             =   210
         Width           =   7485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find and replace"
      Height          =   1485
      Left            =   2100
      TabIndex        =   3
      Top             =   1050
      Width           =   3405
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   270
         Width           =   2325
      End
      Begin VB.TextBox txtReplace 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   570
         Width           =   2325
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   870
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find and Replace"
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   1110
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "Find"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Replace"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   570
         Width           =   795
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Unload msWord"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Load msWord"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   1935
   End
   Begin Project1.msWord msWord1 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   3630
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   10398
   End
End
Attribute VB_Name = "msWordDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbBookMarks_Click()
  msWord1.AddBookMark cbBookMarks.Text, cbBookMarks.Text
End Sub

Private Sub cmbLayout_Click(Index As Integer)

   Select Case Index
     
     Case 0: msWord1.View msWeb
     Case 1: msWord1.View msNormal
     Case 2: msWord1.View msPrint
     Case 3: msWord1.View msReading
     Case 4: msWord1.View msOutline
     
   End Select
End Sub

Private Sub cmdDoFilling_Click()
  With msWord1
      .SetBookMark "Name", txtFields(0).Text
      .SetBookMark "Address", txtFields(1).Text
      .SetBookMark "Tel1", txtFields(2).Text
  End With
End Sub

Private Sub Command1_Click()
    
   Dim sFileName As String
   msWord1.LoadFile sFileName ' empty sFileName makes control open a dialog to ask for it

End Sub

Private Sub Command10_Click()
   msWord1.AddBookMark txtBk.Text, txtBKlabel.Text
End Sub

Private Sub Command11_Click()
   msWord1.AddTable Val(txtRow.Text), Val(txtCol.Text)
End Sub
Private Sub Command12_Click()
  msWord1.LoadWord
End Sub

Private Sub Command13_Click()
  msWord1.UnloadWord
End Sub

Private Sub Command14_Click()
  
  msWord1.SaveDocument "", wdFormatDocument
  
End Sub

Private Sub Command15_Click()
   msWord1.ShowHeaderFooter
End Sub

Private Sub Command16_Click()
 msWord1.HideHeaderFooter
End Sub

Private Sub Command17_Click()
  msWord1.ShowHideRuler
End Sub

Private Sub Command18_Click()
  msWord1.InsertPage
End Sub

Private Sub Command19_Click()
   Dim sFileName As String
   msWord1.LoadFile App.Path & "\Test.doc"

End Sub

Private Sub Command2_Click()
   msWord1.FindAndReplace txtFind.Text, txtReplace.Text, Check1.Value
End Sub

Private Sub Command3_Click()
   msWord1.InsertEmptyLine
End Sub

Private Sub Command4_Click()
   Dim sFileName As String
   msWord1.InsertFile sFileName  ' empty sFileName makes control open a dialog to ask for it

End Sub

Private Sub Command5_Click()
   Dim sFileName As String
   msWord1.InsertPicture sFileName ' empty sFileName makes control open a dialog to ask for it
End Sub

Private Sub Command6_Click()
   msWord1.InsertParagraph
End Sub

Private Sub Command7_Click()
   msWord1.TypeText Text1.Text
End Sub

Private Sub btnRsAction_Click(Index As Integer)
   msWord1.DocPrint
End Sub
Private Sub Command8_Click()
   lblBC.Caption = msWord1.BookmarksCount
End Sub

Private Sub Command9_Click()
   msWord1.SetBookMark txtSBK.Text, txtSBKtext.Text
End Sub

Private Sub Form_Load()
   
   cbBookMarks.AddItem "Name"
   cbBookMarks.AddItem "Balance"
   cbBookMarks.AddItem "Address"
   cbBookMarks.AddItem "Tel1"
   cbBookMarks.AddItem "Mobile"
   
End Sub

