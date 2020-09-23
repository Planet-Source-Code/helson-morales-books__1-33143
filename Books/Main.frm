VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Books"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5085
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Book Information"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4845
      Begin VB.TextBox txtEdition 
         BackColor       =   &H8000000F&
         DataField       =   "Edition"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MouseIcon       =   "Main.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H8000000F&
         DataField       =   "Author"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MouseIcon       =   "Main.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtReference 
         BackColor       =   &H8000000F&
         DataField       =   "Reference"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MouseIcon       =   "Main.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   680
         Width           =   3375
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H8000000F&
         DataField       =   "Title"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MouseIcon       =   "Main.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   300
         Width           =   3375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Edition:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Author:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Reference:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   345
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VB6_Projects\Access dbase projects\Books\books.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VB6_Projects\Access dbase projects\Books\books.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bookstable"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add new entry"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save entry"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Remove record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "report"
            Object.ToolTipText     =   "Open report"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Edit "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F32
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":125A
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1582
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":189E
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":217A
            Key             =   "report"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":24A2
            Key             =   "find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2D7E
            Key             =   "exit"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Books Version 1.0.0"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuitemadd 
         Caption         =   "Add New"
      End
      Begin VB.Menu mnuitemsave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuitemdelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuitemedit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuitemline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuitemexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "Tools"
      Begin VB.Menu mnuitemreport 
         Caption         =   "Report"
      End
      Begin VB.Menu mnuitemline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuitemabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuprevious 
      Caption         =   "Previous"
   End
   Begin VB.Menu mnunext 
      Caption         =   "Next"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Created by Helson Morales (Moe's Software And Nets)
'All Rights Reserved (r)2002, this small application is
'for tutorial use only, you may add or update the full
'project, but if you would like to distribute it, please
'contact me first, some of my tutorials are part of
'active sold applications to schools and clients.
'Thanks!
'               **********EMAIL********
'               moesoftware@hotmail.com
'               herminia@centennialpr.net
'               ***********************

Private Sub Command1_Click()
'Code will help user navigate to a previous
'or back record
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command2_Click()
'Code will allow user to move foward to
'a next available record
Adodc1.Recordset.MoveNext
End Sub

Private Sub mnuitemabout_Click()
'A messege dialog that will show programmers info.
'and applications version
MsgBox "Books Version 1.0.0 by Moe's Software And Nets", vbInformation, "Books"
End Sub

Private Sub mnuitemadd_Click()
'code will set the database to receive a new
'record input from user, this code will enable
'current textbox's and will set focus on the
'fron txtbox in this case txttitle.
Adodc1.Recordset.AddNew

txtTitle.Enabled = True
txtReference.Enabled = True
txtAuthor.Enabled = True
txtEdition.Enabled = True

    txtTitle.BackColor = vbWhite
    txtReference.BackColor = vbWhite
    txtAuthor.BackColor = vbWhite
    txtEdition.BackColor = vbWhite
    
txtTitle.SetFocus
End Sub

Private Sub mnuitemdelete_Click()
'Code will remove a record from the database
'and will update the full database after that
'record has been removed
Adodc1.Recordset.Delete adAffectCurrent
End Sub



Private Sub mnuitemedit_Click()

txtTitle.Enabled = True
txtReference.Enabled = True
txtAuthor.Enabled = True
txtEdition.Enabled = True

    txtTitle.BackColor = vbWhite
    txtReference.BackColor = vbWhite
    txtAuthor.BackColor = vbWhite
    txtEdition.BackColor = vbWhite

End Sub

Private Sub mnuitemexit_Click()
'Will terminate application and return to windows
Unload Me
End Sub

Private Sub mnuitemreport_Click()
'Will show the report on the screen or monitor
DataReport1.Show
End Sub

Private Sub mnuitemsave_Click()
'Code will update existing database by
'updating a new file.
'Code will return all text box's to a
'disable state until a new record is input.
Adodc1.Recordset.Save

txtTitle.Enabled = False
txtReference.Enabled = False
txtAuthor.Enabled = False
txtEdition.Enabled = False

    txtTitle.BackColor = vbButtonFace
    txtReference.BackColor = vbButtonFace
    txtAuthor.BackColor = vbButtonFace
    txtEdition.BackColor = vbButtonFace
End Sub

Private Sub mnunext_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub mnuprevious_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Code will set the select case statement in order to
'initialize the toolbar and its contents
Select Case Button.Key

Case "add"
'Code will set the application to receive a new record
'input and will enable all text box's
    Adodc1.Recordset.AddNew
        
        txtTitle.Enabled = True
        txtReference.Enabled = True
        txtAuthor.Enabled = True
        txtEdition.Enabled = True

    txtTitle.BackColor = vbWhite
    txtReference.BackColor = vbWhite
    txtAuthor.BackColor = vbWhite
    txtEdition.BackColor = vbWhite
    
    txtTitle.SetFocus
    
Case "save"
 'Will save all new data and disable all avilable
 'text boxes, this will protect the data
    Adodc1.Recordset.Save
    
        txtTitle.Enabled = False
        txtReference.Enabled = False
        txtAuthor.Enabled = False
        txtEdition.Enabled = False

    txtTitle.BackColor = vbButtonFace
    txtReference.BackColor = vbButtonFace
    txtAuthor.BackColor = vbButtonFace
    txtEdition.BackColor = vbButtonFace
    
      
Case "delete"
'Code will delete a record or file and move to
'the next avilable record
    Adodc1.Recordset.Delete adAffectCurrent
Case "report"
'Code will show the avilable report
    DataReport1.Show
    
Case "edit"
'Will close the application a little mixed up
'with the declaration of the buttons in the toolbar
'but you understand.
    Unload Me

Case "exit"
'PS...Same mixed up but is working fine.

txtTitle.Enabled = True
txtReference.Enabled = True
txtAuthor.Enabled = True
txtEdition.Enabled = True

    txtTitle.BackColor = vbWhite
    txtReference.BackColor = vbWhite
    txtAuthor.BackColor = vbWhite
    txtEdition.BackColor = vbWhite

End Select
End Sub
