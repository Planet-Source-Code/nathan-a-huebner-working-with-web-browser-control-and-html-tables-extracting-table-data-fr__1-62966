VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Working with the Web Browser Control's Tags!"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView TableData 
      Height          =   3405
      Left            =   0
      TabIndex        =   4
      Top             =   2850
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   6006
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Row #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MyName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "City"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Extract"
      Height          =   375
      Left            =   5310
      TabIndex        =   3
      Top             =   690
      Width           =   1485
   End
   Begin VB.TextBox BodyText 
      Height          =   1665
      Left            =   5250
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reload Page"
      Height          =   375
      Left            =   5310
      TabIndex        =   1
      Top             =   180
      Width           =   1485
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   2835
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   5001
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Don't forget to vote! :) I need votes ;("
      Height          =   1575
      Left            =   5340
      TabIndex        =   5
      Top             =   1170
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TriggerNewRow As Boolean
Dim CurColumn As Integer
Dim ColumnCount As Long
Dim RowCount As Long
Dim LastSaveMark As Long

Private Sub AutoSave_Timer()
    

End Sub

Private Sub Command1_Click()
Browser.Navigate App.Path & "\htmld.html"
End Sub

Private Sub Command2_Click()
On Error Resume Next

Dim NextRow As Long
Dim CountToSave As Long
Dim CountSets As Long
Dim x As Long
Dim endx As Long
Dim ReturnedText As String
Dim FinalOutput As String
Dim CurRow As Long
Dim Q1 As Long

If TableData.ListItems.Count > 0 Then
' Table data exists.
    Q1 = MsgBox("Would you like to clear the data first?", vbYesNoCancel + vbInformation, "Clear data?")
        Select Case Q1
            Case 6
            ' Yes
            TableData.ListItems.Clear
            Case 7
            ' No
            Case 2
            ' Cancel
            MsgBox "Aborted", vbCritical, "Aborted"
            Exit Sub
        End Select
End If


ColumnCount = 4 'This is how many columns you pull back (important)
CountToSave = 0
LastSaveMark = 1

' This counts how many TD tags are in the document
' You can count any tag with this statement
endx = Browser.Document.body.All.tags("td").length

RowCount = CountRows

NextRow = 0
CurColumn = 0
CurRow = 1

TriggerNewRow = True ' Start off with a new row (Row 1)

For x = 0 To endx
DoEvents
    ReturnedText = ""
    ' This returns the column/row text, however
    ' the function ReturnAfterRemoval removes column headers
    ' BE SURE TO TWEAK THAT FUNCTION, IT COULD CAUSE DATA PROBLEMS
    ' IF YOUR DATA HAS THE SAME TEXT POSSIBLY AS YOUR COLUMN HEADERS
    ReturnedText = Trim(ReturnAfterRemoval(Browser.Document.body.All.tags("td").Item(x).innerText))
    ' You may want to remove CHR(13) and CHR(10) because it may have some extra stuff
    FinalOutput = ReturnedText
    
       
        If FinalOutput = "" Then
        Else
        
            ' Output found.. Verify its correct position to match Row
        CurColumn = CurColumn + 1
        
        ' Create index for adding another Row
        
        If TriggerNewRow = True Then
        NextRow = TableData.ListItems.Count + 1
        TableData.ListItems.Add
        TableData.ListItems.Item(NextRow).Text = CurRow
        TriggerNewRow = False
        End If
        
        ' Add data, check if the next row is coming.
        
            If CurColumn >= ColumnCount Then
            ' At the end of the column. Lets get ready for the next row
            TableData.ListItems.Item(NextRow).SubItems(CurColumn) = FinalOutput
            CurRow = CurRow + 1
            CurColumn = 0
            TriggerNewRow = True ' Tell program a new row is coming
            Else
            ' Append data to remaining...
            TableData.ListItems.Item(NextRow).SubItems(CurColumn) = FinalOutput
            End If
        
        End If

Next x

MsgBox "Table data has been placed into VB!", vbInformation, ":) Don't forget to vote!"
End Sub

Function ReturnAfterRemoval(Btxt As String) As String
' Removes headers, so you don't grab the column headers.
' You may need to tweak this, it could interfere with actual data.
If InStr(Btxt, "Name") Then Exit Function
If InStr(Btxt, "Phone") Then Exit Function
If InStr(Btxt, "City") Then Exit Function ' This may cause issues for cities that have the word CITY in them.
If InStr(Btxt, "State") Then Exit Function
ReturnAfterRemoval = Btxt
End Function

Function CountRows() As Long
On Error Resume Next
Dim CurCount As Long

CurCount = Browser.Document.body.All.tags("tr").length

CountRows = CurCount - 1 ' Don't count the header row

End Function

Sub AutoSave()
On Error Resume Next
    Dim x As Long
    Dim CompileRow As String
    x = LastSaveMark
    
    On Error Resume Next
    Open "I:\!AUTOSAVE\" & Replace(Time, ":", ".") & ".txt" For Output As #2
    For x = x To TableData.ListItems.Count
    CompileRow = ""
    CompileRow = TableData.ListItems(x).Text & Chr(9)
    CompileRow = CompileRow & TableData.ListItems(x).SubItems(1) & Chr(9)
    CompileRow = CompileRow & TableData.ListItems(x).SubItems(2) & Chr(9)
    CompileRow = CompileRow & TableData.ListItems(x).SubItems(3) & Chr(9)
    CompileRow = CompileRow & TableData.ListItems(x).SubItems(4) & Chr(9)
    CompileRow = CompileRow & TableData.ListItems(x).SubItems(5) & Chr(9)
    CompileRow = CompileRow & TableData.ListItems(x).SubItems(6) & Chr(9)
    
        Print #2, CompileRow
    Next x
    LastSaveMark = x
    Close #2
    
End Sub

Private Sub Form_Load()
Browser.Navigate App.Path & "\htmld.html"

End Sub

Private Sub TableData_BeforeLabelEdit(Cancel As Integer)
' cancel = 1  ' this would disable editing...
' Editing is enabled if you do not add cancel=1
End Sub
