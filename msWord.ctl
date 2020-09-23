VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl msWord 
   Alignable       =   -1  'True
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   ScaleHeight     =   6885
   ScaleWidth      =   8355
   ToolboxBitmap   =   "msWord.ctx":0000
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5865
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   7905
      ExtentX         =   13944
      ExtentY         =   10345
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9870
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "msWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim oDocument As Object

Enum Formats
        wdFormatDocument = 0    ' Saves as a Word document.
        wdFormatHTML = 1        ' Saves the file in HTML format (a Web page) so that it can be viewed in a Web browser.
        wdFormatTemplate = 2    ' Saves as a Word template.
        wdFormatText = 3        ' Text Only: Saves text without its formatting. Converts all section breaks, page breaks, and new line characters to paragraph marks. Uses the ANSI character set. Select this format only if the destination program cannot read any of the other available file formats.
        wdFormatDOSText = 4     ' MS-DOS Text: Converts files the same way as Text only format (wdFormatText). Uses the extended ASCII character set, which is the standard for MS-DOS-based programs. Use this format to share documents between Word and non-Windows-based programs.
        wdFormatTextLineBreaks = 5 ' Text only with line breaks: Saves text without formatting. Converts all line breaks, section breaks, and page breaks to paragraph marks. Use this format when you want to maintain line breaks; for example, when transferring documents to an electronic mail system.
        wdFormatDOSTextLineBreaks = 6 ' MS-DOS text only with line breaks: Saves text without formatting. Converts all line breaks, section breaks, and page breaks to paragraph marks. Use this format when you want to maintain line breaks; for example, when transferring documents to an electronic mail system.
        wdFormatRTF = 7         ' Rich Text Format (RTF): Saves all formatting. Converts formatting to instructions that other programs, including compatible Microsoft programs, can read and interpret.
        wdFormatUnicodeText = 8 ' Saves as a Unicode text file: Converts text between common character encoding standards, including Unicode 2.0, Mac OS, Windows, EUC, and ISO-8859 series.
End Enum

Enum Views
    msWeb = 0
    msPrint = 1
    msReading = 2
    msNormal = 3
    msOutline = 4
End Enum
Enum FindAndReplace
     SingleReplace = 0
     ReplaceAll = 1
End Enum
Private Sub UserControl_Initialize()

   On Error Resume Next
   
   Screen.MousePointer = vbHourglass
  
  ' Set oDocument = Nothing
   
  ' wb.Navigate app.path & "\blank.doc" ' So that webbrowser starts in 'msWord mode' and empty document
   
   wb.Navigate "about:blank" ' So that webbrowser starts with a blank screen without loading word.
   
   With CommonDialog1
      .Filter = "Office Documents " & _
      "(*.doc, *.dot)|*.doc;*.dot"
      .FilterIndex = 1
      .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
   End With
   
   Screen.MousePointer = vbNormal
   
   If Err Then
    MsgBox "I did not find the blank.doc file"
   End If
    
    
End Sub
Sub InsertPicture(sPictureName As String)
   
   On Error Resume Next
   
   ' If the user didn't provide a name then ask for one
   If Len(sPictureName) = 0 Then
       With CommonDialog1
          .Filter = "Office Pictures " & _
          "(*.gif,*.jpg,*.bmp)|*.gif;*.jpg;*.bmp"
          .FilterIndex = 1
          .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
          
          .FileName = ""
          .ShowOpen
          sPictureName = .FileName
       End With
   End If
   
 ' If the user didn't cancel, open the file...
  If Len(sPictureName) Then
        With wb.Document.Application
          
          If Val(.Version) >= 9 Then
                      'OFFICE 2000+
                      .Selection.InlineShapes.AddPicture FileName:=sPictureName
          Else
                      'OFFICE 97
                      .ChangeFileOpenDirectory App.Path
                      .ActiveDocument.Shapes.AddPicture _
                             Anchor:=Selection.Range _
                            , FileName:=sPictureName _
                            , LinkToFile:=False _
                            , SaveWithDocument:=True
          End If
         
          .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
          
         End With
    End If
End Sub
Sub InsertParagraph()
  
  On Error Resume Next
  With wb.Document.Application
      .Selection.TypeParagraph
  End With

End Sub
Sub DocPrint()
 
  On Error Resume Next
  With wb
    .ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
  End With

End Sub
Sub TypeText(sText As String)

   On Error Resume Next
   With wb.Document.Application.Selection
    .TypeText Text:=sText
   End With
End Sub

Sub LoadFile(sFileName As String)
   
' If the user didn't provide a name then ask for one
   If Len(sFileName) = 0 Then
       With CommonDialog1
          .Filter = "Office Documents " & _
          "(*.doc, *.dot)|*.doc;*.dot"
          .FilterIndex = 1
          .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
   
          .FileName = ""
          .ShowOpen
          sFileName = .FileName
       End With
   End If
 
 ' If the user didn't cancel, open the file...
   If Len(sFileName) Then
       wb.Navigate sFileName
   End If

End Sub


Public Sub InsertEmptyLine()
    On Error Resume Next
    With wb.Document.Application
       .Selection.TypeParagraph
    End With
End Sub

Public Sub InsertFile(sFileName As String)
   
   On Error Resume Next
   
   ' If the user didn't provide a name then ask for one
   If Len(sFileName) = 0 Then
       With CommonDialog1
          .Filter = "Office Documents " & _
          "(*.doc, *.dot)|*.doc;*.dot"
          .FilterIndex = 1
          .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
          
          .FileName = ""
          .ShowOpen
          sFileName = .FileName
       End With
   End If
 ' If the user didn't cancel, open the file...
   If Len(sFileName) Then
        With wb.Document.Application
            .Selection.InsertFile FileName:=sFileName, ConfirmConversions:=False
            .Selection.Collapse Direction:=wdCollapseEnd
        End With
   End If
 
End Sub

Private Sub UserControl_Resize()
   wb.Move 30, 30, UserControl.Width - 30, UserControl.Height - 60
End Sub


Private Sub wb_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
   
   On Error Resume Next
  ' Set oDocument = pDisp.Document
   
   wb.ExecWB OLECMDID_HIDETOOLBARS, OLECMDEXECOPT_DONTPROMPTUSER
   
End Sub
Public Function Find(strFind As String) As Boolean
    
    With wb.Document.Application.Selection.Find
        .ClearFormatting
        .Forward = True
        .Wrap = Word.WdFindWrap.wdFindContinue
        .Text = strFind
        If .Execute = True Then
            Find = True
        Else
            Find = False
        End If
    End With

End Function
Public Function FindAndReplace(strFind As String, sReplace As String, iMode As FindAndReplace) As Boolean
        
    On Error Resume Next
    With wb.Document.Application.Selection.Find
        
        .ClearFormatting
        .Forward = True
        .Wrap = Word.WdFindWrap.wdFindContinue
        
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Text = strFind
        
        With .Replacement
            .ClearFormatting
            .Text = sReplace
        End With
        
        If iMode = SingleReplace Then
           .Execute Replace:=Word.WdReplace.wdReplaceOne
        Else
           .Execute Replace:=Word.WdReplace.wdReplaceAll
        End If
        
        
    End With
      
 '       .ClearFormatting
 '       .Forward = True
 '       .Wrap = Word.WdFindWrap.wdFindContinue
 '       .Text = strFind
 '
 '       If iMode = SingleReplace Then
 '            If Find(strFind) Then
 '                TypeText sReplace
 '                iFoundSomething = True
 '            End If
 '        Else
 '            Do While Find(strFind)
 '                TypeText sReplace
 '                iFoundSomething = True
 '            Loop
 '        End If
    
   
End Function
Public Function BookmarksCount() As Integer
   On Error Resume Next
   With wb.Document.Bookmarks
        BookmarksCount = .Count
   End With
End Function
Public Sub SetBookMark(sBookMarkName As String, sText As String)
   On Error Resume Next
   With wb.Document.Bookmarks
        .Item(sBookMarkName).Range.Text = sText
   End With
End Sub
Public Sub AddBookMark(sName As String, sText As String)
   On Error Resume Next
   
   With wb.Document.Application.Selection

        .TypeText Text:=sText
        
        .MoveLeft Unit:=wdCharacter, Count:=Len(sText)
        .MoveRight Unit:=wdCharacter, Count:=Len(sText), Extend:=wdExtend
        
         wb.Document.Bookmarks.Add Range:=Selection.Range, Name:=sName
         wb.Document.Bookmarks.DefaultSorting = wdSortByName
         wb.Document.Bookmarks.ShowHidden = False
                  
    End With

End Sub
Public Sub AddTable(iNumRows As Integer, iNumColumns As Integer)
  
    On Error GoTo ErrTable
    
    With wb.Document.Application.Selection
          
          wb.Document.Tables.Add _
                    Range:=.Range, _
                    NumRows:=iNumRows, _
                    NumColumns:=iNumColumns, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitFixed
           
            If .Tables(1).Style <> "Table Grid" Then
               .Tables(1).Style = "Table Grid"
            End If
           
           .Tables(1).ApplyStyleHeadingRows = True
           .Tables(1).ApplyStyleLastRow = True
           .Tables(1).ApplyStyleFirstColumn = True
           .Tables(1).ApplyStyleLastColumn = True
         
    End With
TableExit:
    Exit Sub

ErrTable:
      MsgBox "Error " & Err & ": " & Error
      Resume TableExit
End Sub
Public Sub View(vv As Views)
  
  On Error Resume Next
  With wb.Document
    Select Case vv
       
       Case Views.msWeb
            .ActiveWindow.View.Type = wdWebView
       
       Case Views.msPrint
            If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
               .ActiveWindow.ActivePane.View.Type = wdPrintView
            Else
                .ActiveWindow.View.Type = wdPrintView
            End If
       
       Case Views.msReading
            .ActiveWindow.View.ReadingLayout = Not .ActiveWindow.View.ReadingLayout
       
       Case Views.msNormal
            If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
               .ActiveWindow.ActivePane.View.Type = wdNormalView
            Else
                .ActiveWindow.View.Type = wdNormalView
            End If
        Case Views.msOutline
             .ActiveWindow.ActivePane.View.Type = wdMasterView
      End Select

  End With
End Sub
Public Sub LoadWord()
   
 ' --- Instead of having Blank.doc in the directory you can load it from a resource file
 '
 '  Dim DataArray() As Byte
 '  DataArray = LoadResData(101, "CUSTOM") ' This is a blanc document embeded in a resource file
 '
 '  FileNum = FreeFile
 '  Open App.Path & "\blank.doc" For Binary As #FileNum ' We save the file
 '  Put #FileNum, 1, DataArray()
 '  Close #FileNum
   
   
   ' and we load it in the webbrowser control. This load msword because of the type (.doc) for the file
   wb.Navigate App.Path & "\blank.doc"   ' So that webbrowser starts in 'msWord mode' and empty document
   
 '  Erase DataArray ' Cleanup our memory

End Sub
Public Sub UnloadWord()
   wb.Navigate "about:blank" ' So that webbrowser starts in 'msWord mode' and empty document
End Sub
Public Sub SaveDocument(sFileName As String, iFormat As Formats)
    
    
    Dim iSaveFormat As Integer
    
    On Error Resume Next
    
    Select Case iFormat
            
            Case Formats.wdFormatDocument: iSaveFormat = wdFormatDocument
            Case Formats.wdFormatHTML: iSaveFormat = wdFormatHTML
            Case Formats.wdFormatTemplate: iSaveFormat = wdFormatTemplate
            Case Formats.wdFormatText: iSaveFormat = wdFormatText
            Case Formats.wdFormatDOSText: iSaveFormat = wdFormatDOSText
            Case Formats.wdFormatTextLineBreaks: iSaveFormat = wdFormatTextLineBreaks
            Case Formats.wdFormatDOSTextLineBreaks: iSaveFormat = wdFormatDOSTextLineBreaks
            Case Formats.wdFormatRTF: iSaveFormat = wdFormatRTF
            Case Formats.wdFormatUnicodeText: iSaveFormat = wdFormatUnicodeText
            
    End Select
    
   ' If the user didn't provide a name then ask for one
   
   If Len(sFileName) = 0 Then
       With CommonDialog1
          .Filter = "Office Documents " & _
          "(*.*)|*.*"
          .FilterIndex = 1
          .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
          .FileName = ""
          .ShowSave
          sFileName = .FileName
       End With
   End If
 ' If the user didn't cancel, open the file...
   If Len(sFileName) Then
        With wb.Document.Application
             .ActiveDocument.SaveAs sFileName, iSaveFormat
        End With
   End If

End Sub
Public Sub ShowHeaderFooter()
 On Error Resume Next
 With wb.Document
    If .ActiveWindow.View.SplitSpecial <> wdPaneNone Then
       .ActiveWindow.Panes(2).Close
    End If
    If .ActiveWindow.ActivePane.View.Type = wdNormalView Or .ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        .ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
  End With
End Sub
Public Sub HideHeaderFooter()
    On Error Resume Next
    wb.Document.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Public Sub ShowHideRuler()
  
  On Error Resume Next
  wb.Document.ActiveWindow.ActivePane.DisplayRulers = Not ActiveWindow.ActivePane.DisplayRulers
    
End Sub
Public Sub InsertPage()
   On Error Resume Next
   With wb.Document.Application
       .Selection.InsertBreak Type:=wdPageBreak
   End With
End Sub


'Public Sub msSetup()
'
'  With wb.Document
'    .Application.DisplayStatusBar = True
'    .Application.ShowWindowsInTaskbar = True
'    .Application.ShowStartupDialog = True
'
'    With wb.Document.ActiveWindow
'        .DisplayHorizontalScrollBar = True
'        .DisplayVerticalScrollBar = True
'        .DisplayLeftScrollBar = False
'        .StyleAreaWidth = InchesToPoints(0)
'        .DisplayVerticalRuler = True
'        .DisplayRightRuler = False
'        .DisplayScreenTips = True
'
'        With wb.Document.ActiveWindow.View
'            .ShowAnimation = True
'            .Draft = False
'            .WrapToWindow = False
'            .ShowPicturePlaceHolders = False
'            .ShowFieldCodes = False
'            .ShowBookmarks = True
'            .FieldShading = wdFieldShadingWhenSelected
'            .ShowTabs = False
'            .ShowSpaces = False
'            .ShowParagraphs = False
'            .ShowHyphens = False
'            .ShowHiddenText = False
'            .ShowAll = False
'            .ShowDrawings = True
'            .ShowObjectAnchors = False
'            .ShowTextBoundaries = False
'            .ShowHighlight = True
'            .DisplayPageBoundaries = True
'            .DisplaySmartTags = True
'        End With
'    End With
'   End With
'End Sub
   
