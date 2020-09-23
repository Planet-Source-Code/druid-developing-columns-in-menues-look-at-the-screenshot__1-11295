VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Columns in menues!"
   ClientHeight    =   2220
   ClientLeft      =   4065
   ClientTop       =   3345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4230
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click for Popup menu"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   7
      Left            =   3960
      Picture         =   "Form1.frx":0000
      Top             =   960
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   6
      Left            =   3960
      Picture         =   "Form1.frx":02AA
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   5
      Left            =   3960
      Picture         =   "Form1.frx":0554
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   4
      Left            =   3960
      Picture         =   "Form1.frx":07FE
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   3
      Left            =   3960
      Picture         =   "Form1.frx":0AA8
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   2
      Left            =   3720
      Picture         =   "Form1.frx":0D52
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":0FFC
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgMenu 
      Height          =   210
      Index           =   0
      Left            =   3720
      Picture         =   "Form1.frx":12A6
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Menu columnmenu 
      Caption         =   "ColumnMenu"
      Begin VB.Menu menuitem 
         Caption         =   "New"
         Index           =   0
      End
      Begin VB.Menu menuitem 
         Caption         =   "Open"
         Index           =   1
      End
      Begin VB.Menu menuitem 
         Caption         =   "Save"
         Index           =   2
      End
      Begin VB.Menu menuitem 
         Caption         =   "Save as"
         Index           =   3
      End
      Begin VB.Menu menuitem 
         Caption         =   "Print"
         Index           =   4
      End
      Begin VB.Menu menuitem 
         Caption         =   "Undo"
         Index           =   5
      End
      Begin VB.Menu menuitem 
         Caption         =   "Redo"
         Index           =   6
      End
      Begin VB.Menu menuitem 
         Caption         =   "Cut"
         Index           =   7
      End
      Begin VB.Menu menuitem 
         Caption         =   "Copy"
         Index           =   8
      End
      Begin VB.Menu menuitem 
         Caption         =   "Paste"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************
'* API Declarations  *
'*********************
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

'*********************
'* Consts            *
'*********************
Const MF_BYPOSITION = &H400&

Private Sub Form_Load()
    'Dim Variables
    Dim hMenu&, hSubMenu&
    Dim mnuItemCnt&, mnuItemID&, mnuItemText$
    Dim X%, Result&, Buffer$
    'Get the hwnd of the menu
    hMenu = GetMenu(Me.hwnd)
    'Get the hwnd of the submenu
    hSubMenu = GetSubMenu(hMenu, 0)
    'Apply the pictures to the menu items to make them look better
    Call SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, imgMenu(0).Picture, imgMenu(0).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 2, MF_BYPOSITION, imgMenu(1).Picture, imgMenu(1).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 4, MF_BYPOSITION, imgMenu(2).Picture, imgMenu(2).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 5, MF_BYPOSITION, imgMenu(3).Picture, imgMenu(3).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 6, MF_BYPOSITION, imgMenu(4).Picture, imgMenu(4).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 7, MF_BYPOSITION, imgMenu(5).Picture, imgMenu(5).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 8, MF_BYPOSITION, imgMenu(6).Picture, imgMenu(6).Picture)
    Call SetMenuItemBitmaps(hSubMenu, 9, MF_BYPOSITION, imgMenu(7).Picture, imgMenu(7).Picture)
    mnuItemCnt = GetMenuItemCount(hSubMenu)
    'The Step is the number of items in one column
    For X = 6 To mnuItemCnt Step 5
        Buffer = Space$(256)
        Result = GetMenuString(hSubMenu, X - 1, Buffer, Len(Buffer), &H400&)
        mnuItemText = Left$(Buffer, Result)
        mnuItemID = GetMenuItemID(hSubMenu, X - 1)
        'Modify the menu to a column menu
        Call ModifyMenu(hSubMenu, X - 1, &H400& Or &H20&, mnuItemID, mnuItemText)
    Next X
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
    'Show popup menu
    Me.PopupMenu columnmenu, , X, Y
    End If
End Sub

'The procedure that is called when a menu item is clicked
Private Sub menuitem_Click(Index As Integer)
    MsgBox "Click on Item '" & menuitem(Index).Caption & "'", vbInformation, "Click event"
End Sub
