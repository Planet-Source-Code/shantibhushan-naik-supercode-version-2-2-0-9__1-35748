VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   4980
   ClientTop       =   3360
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.TextBox Text1 
      Height          =   2865
      Left            =   98
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0442
      Top             =   165
      Width           =   4485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2070
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0672
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BB4
            Key             =   "uponelevel"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CC6
            Key             =   "font"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1260
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17A2
            Key             =   "find"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18B4
            Key             =   "help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DF6
            Key             =   "new"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2338
            Key             =   "open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":287A
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":298C
            Key             =   "paint"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2ECE
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3410
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3952
            Key             =   "print"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E94
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":43D6
            Key             =   "save"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4918
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4CAE
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51F0
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5732
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5844
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5956
            Key             =   "right"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5A68
            Key             =   "center"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B7A
            Key             =   "left"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5C8C
            Key             =   "spelling"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":61CE
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":62E0
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":63F2
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6504
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6616
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69A3
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F3D
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":74D7
            Key             =   "arngico"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7865
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7977
            Key             =   "level1"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":79D6
            Key             =   "level2"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A39
            Key             =   "level3"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7A9F
            Key             =   "level4"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B08
            Key             =   "level5"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B72
            Key             =   "level6"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7BE2
            Key             =   "level7"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7C53
            Key             =   "level8"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7CC4
            Key             =   "level9"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7D38
            Key             =   "level10"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSideBar 
         Caption         =   "File..."
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "Save Options"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save &All"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "Print Layout Options"
      End
      Begin VB.Menu mnuFilePgSetup 
         Caption         =   "Page Set&up..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrntPrvw 
         Caption         =   "Print Pre&view..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep0 
         Caption         =   "Clipboard Options"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "Search Options"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditProp 
         Caption         =   "Proper&ties"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "For&mat"
      Begin VB.Menu mnuFormatSideBar 
         Caption         =   "Format Options..."
      End
      Begin VB.Menu mnuFmtFont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuFmtSplChk 
         Caption         =   "&Spell Check"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuFmtSep0 
         Caption         =   "Text Format Options"
      End
      Begin VB.Menu mnuFmtBold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFmtItalic 
         Caption         =   "Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFmtUndrln 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFmtSort 
         Caption         =   "Sort"
         Begin VB.Menu mnuFmtSortAsc 
            Caption         =   "Ascending"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuFmtSortDesc 
            Caption         =   "Descending"
            Shortcut        =   ^{F2}
         End
      End
      Begin VB.Menu mnuFmtSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFrmtAlgn 
         Caption         =   "Align"
         Begin VB.Menu mnuFrmtAlgnSideBar 
            Caption         =   "Align Options"
         End
         Begin VB.Menu mnuFrmtAlgnOptn 
            Caption         =   "&Left"
            Index           =   0
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuFrmtAlgnOptn 
            Caption         =   "&Right"
            Index           =   1
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuFrmtAlgnOptn 
            Caption         =   "&Center"
            Index           =   2
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuFrmtAlgnOptn 
            Caption         =   "&Justify"
            Enabled         =   0   'False
            Index           =   3
         End
      End
      Begin VB.Menu mnuFmtSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFmtPaint 
         Caption         =   "Paint"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuNstMnu 
      Caption         =   "&Nested Menu"
      Begin VB.Menu mnuNstSubMnu1 
         Caption         =   "Sub Menu 1"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuNstSubMnu2 
         Caption         =   "Sub Menu 2"
         Begin VB.Menu mnuNstSubMnu3 
            Caption         =   "Sub Menu 3"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuNstSubMnu4 
            Caption         =   "Sub Menu 4"
            Begin VB.Menu mnuNstSubMnu5 
               Caption         =   "Sub Menu 5"
               Shortcut        =   {F5}
            End
            Begin VB.Menu mnuNstSubMnu6 
               Caption         =   "Sub Menu 6"
               Begin VB.Menu mnuNstSubMnu7 
                  Caption         =   "Sub Menu 7"
                  Shortcut        =   {F7}
               End
               Begin VB.Menu mnuNstSubMnu8 
                  Caption         =   "Sub Menu 8"
                  Begin VB.Menu mnuNstSubMnu8SideBar 
                     Caption         =   "Last"
                  End
                  Begin VB.Menu mnuNstSubMnu9 
                     Caption         =   "Sub Menu 9"
                     Shortcut        =   {F9}
                  End
                  Begin VB.Menu mnuNstSubMnu10 
                     Caption         =   "Sub Menu 10"
                     Shortcut        =   {F11}
                  End
               End
            End
         End
      End
   End
   Begin VB.Menu mnuWnd 
      Caption         =   "&Window"
      Begin VB.Menu mnuWndSidebar 
         Caption         =   "Window"
      End
      Begin VB.Menu mnuWndwNew 
         Caption         =   "New Window"
      End
      Begin VB.Menu mnuWndSep0 
         Caption         =   "Arrange Windows"
      End
      Begin VB.Menu mnuWndPos 
         Caption         =   "Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuWndPos 
         Caption         =   "Tile Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuWndPos 
         Caption         =   "Tile Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuWndPos 
         Caption         =   "Arrange Icons"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSystem                                                     As SystemInteroperatability.System

Private Sub Form_Load()
  Rem /* Just look at the Object Browser to have a look at the *numerous* features this SystemInteroperatability
  Rem  * DLL has to offer. No .bas modules anymore. Work purely in an object oriented way and what more, just
  Rem  * the plain old way with VB with almost no learning curve.
  Rem  */
  Call System.Window.Init(Me.hWnd, Me)  ' /* Make the system aware which window we want to use. */
  
  Call setMenuBitmaps                   ' /* Load pure Office Style menus. */
  Call setToolTips                      ' /* Load pure Win2K/WinXP Style tooltips. */
  Call addToSystemTray                  ' /* Load system tray icon. Now as easy as that !. */
  
  Rem /* Load transparently. As many of us know, translucency effects came around Win2K (really WinNT5.0..!)
  Rem  * and later. So developers now can use these *special* effects in no time! Just place the following
  Rem  * in any form that you want to animate. (Not only transparent effects, do have a look at other
  Rem  * effects too.
  Rem  * Notice how easily I can just have an if condition to compare *OS*, with almost no code.
  Rem  */
  If (System.OS > WINDOWSNT400) Then Call System.Window.Animate
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Let Me.ScaleMode = VBRUN.ScaleModeConstants.vbPixels
  Select Case X
    Case MOUSEMESSAGES.LEFTBUTTONUP:
      Call VBA.Interaction.MsgBox("See how I communicated with VB code !")
    Case Else:
  End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Rem /* See more translucent effects. You'll definitely like it! */
'  If (System.OS > WINDOWSNT400) Then Call System.Window.Animate(DEACTIVATE_SLIDE_FADE_TRANSITION)            ' /* Fade delay 300ms */
  If (System.OS > WINDOWSNT400) Then Call System.Window.Animate(DEACTIVATE_SLIDE_FADE_TRANSITION, 1000)      ' /* Fade delay 1000ms */
  
  Rem /* Unload any objects. */
  If Not (System Is Nothing) Then Set System = Nothing
End Sub

Private Sub mnuFileExit_Click()
  Rem /* You can continue working with your *normal* events */
  Call VB.Unload(Me)
End Sub

Private Property Get System() As SystemInteroperatability.System
  Rem /* Check if we are accessing this member for the first time */
  If (mSystem Is Nothing) Then Set mSystem = New SystemInteroperatability.System
  
  Set System = mSystem
End Property

Private Property Set System(ByVal value As SystemInteroperatability.System)
  Set mSystem = value
End Property

Private Sub mnuFmtBold_Click()
  Rem /* Toggle the state. */
  Let mnuFmtBold.Checked = Not mnuFmtBold.Checked
  
  Rem /* Set the appropriate text formatting options. */
  Let Text1.Font.Bold = mnuFmtBold.Checked
End Sub

Private Sub mnuFmtItalic_Click()
  Rem /* Toggle the state. */
  Let mnuFmtItalic.Checked = Not mnuFmtItalic.Checked
  
  Rem /* Set the appropriate text formatting options. */
  Let Text1.Font.Italic = mnuFmtItalic.Checked
End Sub

Private Sub mnuFmtUndrln_Click()
  Rem /* Toggle the state. */
  Let mnuFmtUndrln.Checked = Not mnuFmtUndrln.Checked
  
  Rem /* Set the appropriate text formatting options. */
  Let Text1.Font.Underline = mnuFmtUndrln.Checked
End Sub

Private Sub mnuFrmtAlgnOptn_Click(Index As Integer)
  Dim item       As Long: Let item = 0
  Rem /* Toggle the state for the item. */
  Let mnuFrmtAlgnOptn(Index).Checked = True
  Rem /* Set the appropriate text alignment options. */
  Let Text1.Alignment = Index
  Rem /* Only *one* out of the current alignment options has to be checked. */
  For item = mnuFrmtAlgnOptn.UBound To 0 Step -1
    If Not (item = Index) Then Let mnuFrmtAlgnOptn(item).Checked = False ' /* Uncheck any other item. */
  Next item
End Sub

Private Sub mnuWndPos_Click(Index As Integer)
  Dim item       As Long: Let item = 0
  
  Rem /* Toggle the state for the item. */
  Let mnuWndPos(Index).Checked = True
  Rem /* Only *one* out of the current window arrange options has to be checked. */
  For item = mnuWndPos.UBound To 0 Step -1
    If Not (item = Index) Then Let mnuWndPos(item).Checked = False ' /* Uncheck any other item. */
  Next item
End Sub

Rem /* The Office Menus. Don't they look *real* good. Well for a VB programmer this is not too far
Rem  * away. Just as simple as the following code and you can have good looking menus in your VB
Rem  * applications too.
Rem  */
Private Sub setMenuBitmaps()
  Dim Images      As MSComctlLib.ListImages: Set Images = ImageList1.ListImages
  
  Rem /* That's it, now we shall assign some bitmaps to our own menus. */
  With System.Window.Menu
'     Let .EnableDisabledMenuSelection = True           ' /* See this property is you want to allow
                                                       '  * selection over disabled menu items too.
                                                       '  */
    
    Rem /* #File Menu. */
    Rem /* Wow now gradient sidebars for your menus too. And as easy as that. Only one caution,
    Rem  * sidebar menus have to be the very *first* menu items. Only that !
    Rem  */
    Let .SideBar(mnuFileSideBar) = True
    Set .BitmapImage(mnuFileNew) = Images("new").ExtractIcon
    Set .BitmapImage(mnuFileOpen) = Images("open").ExtractIcon
    Let .Include(mnuFileClose) = True                              ' /* Simply include it. [Optional] You can comment it and see the difference. */
    
    Rem /* A new feature for VB programmers. Though not new, but have you imagined that a menu sepera
    Rem  * tor can now be used for a more productive task, say to indicate a group! Yes, now VB Developer's
    Rem  * have menu seperators that are just not *simple* seperators, but menu seperators with
    Rem  * informative text. Simply create a normal menu and pass it! That's it. The menu caption will
    Rem  * be the menu seperator informative text. Don't believe it! Just see those *fabulous* menus..!!
    Rem  */
    Let .MenuSeperator(mnuFileSep0) = True
    
    Set .BitmapImage(mnuFileSave) = Images("save").ExtractIcon
    Set .BitmapImage(mnuFileSaveAll) = Images("saveall").ExtractIcon
    Let .Include(mnuFileSaveAs) = True
    Let .MenuSeperator(mnuFileSep3) = True
    Set .BitmapImage(mnuFilePrntPrvw) = Images("preview").ExtractIcon
    Let .Include(mnuFilePgSetup) = True
    Set .BitmapImage(mnuFilePrint) = Images("print").ExtractIcon
    Set .BitmapImage(mnuFileExit) = Images("exit").ExtractIcon
    
    Rem /* #Edit Menu. */
    Set .BitmapImage(mnuEditUndo) = Images("undo").ExtractIcon
    Set .BitmapImage(mnuEditRedo) = Images("redo").ExtractIcon
    Let .MenuSeperator(mnuEditSep0) = True
    Set .BitmapImage(mnuEditCut) = Images("cut").ExtractIcon
    Set .BitmapImage(mnuEditCopy) = Images("copy").ExtractIcon
    Set .BitmapImage(mnuEditPaste) = Images("paste").ExtractIcon
    Let .MenuSeperator(mnuEditSep1) = True
    Set .BitmapImage(mnuEditFind) = Images("find").ExtractIcon
    
    Set .BitmapImage(mnuEditProp) = Images("properties").ExtractIcon
    
    Rem /* #Format Menu. */
    Let .SideBar(mnuFormatSideBar) = True
    Set .BitmapImage(mnuFmtFont) = Images("font").ExtractIcon
    Set .BitmapImage(mnuFmtSplChk) = Images("spelling").ExtractIcon
    
    Let .MenuSeperator(mnuFmtSep0) = True
    Let .Checked(mnuFmtBold, Images("bold").ExtractIcon) = True
    Let mnuFmtBold.Checked = False            ' /* The text is not bold yet. */
    Let .Checked(mnuFmtItalic, Images("italic").ExtractIcon) = True
    Let mnuFmtItalic.Checked = False          ' /* The text is not italicized yet. */
    Let .Checked(mnuFmtUndrln, Images("underline").ExtractIcon) = True
    Let mnuFmtUndrln.Checked = False          ' /* The text is not underlined yet. */
    
    Set .BitmapImage(mnuFmtSortAsc) = Images("sortasc").ExtractIcon
    Set .BitmapImage(mnuFmtSortDesc) = Images("sortdesc").ExtractIcon
    
    Rem /* Wow now gradient sidebars for your menus too. And as easy as that. Only one caution,
    Rem  * sidebar menus have to be the very *first* menu items in *any* submenu. Don't worry
    Rem  * just keep the sidebar item the first item in submenu of *any* deep nesting. The search
    Rem  * algorithm will correctly find it out 99% of time. To have a feel just have a look where
    Rem  * the sidebar is the "Nested Menu" and you'll get convinced that you had to do almost
    Rem  * nothing.
    Rem  */
    Let .SideBar(mnuFrmtAlgnSideBar) = True
    
    Rem /* Now even *radio* bitmapped images are simple. Also look at the click event too.
    Rem  * Just supply the optional bitmap for the checked property and a bitmap will show
    Rem  * up instead of the default checkmark.
    Rem  */
    Let .Checked(mnuFrmtAlgnOptn(0), Images("left").ExtractIcon) = True     ' /* Left Align. */
    Let .Checked(mnuFrmtAlgnOptn(1), Images("right").ExtractIcon) = True    ' /* Right Align */
    Let .Checked(mnuFrmtAlgnOptn(2), Images("center").ExtractIcon) = True   ' /* Center Align */
    Let .Checked(mnuFrmtAlgnOptn(3), Images("justify").ExtractIcon) = True  ' /* Justify */
    Call mnuFrmtAlgnOptn_Click(0)         ' /* Select the left justify option. */
    
    Set .BitmapImage(mnuFmtPaint) = Images("paint").ExtractIcon
    
    Rem /* #Nesting Menu. */
    Rem /* The most important thing to notice here is the *nesting* of this menu. Notice how deep
    Rem  * this menu item is placed in the hierarchy. No problem, how deep the menu item may be.
    Rem  * There is a *very* efficient search algorithm inside the DLL that *can* search virtually
    Rem  * infinite deep levels of nesting within the given menu.
    Rem  * EVEN THIS MENU IS THE *LAST* LEVEL THAT CAN BE CREATED USING VB'S MENU EDITOR !!!
    Rem  */
    Set .BitmapImage(mnuNstSubMnu1) = Images("level1").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu2) = Images("level2").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu3) = Images("level3").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu4) = Images("level4").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu5) = Images("level5").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu6) = Images("level6").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu7) = Images("level7").ExtractIcon
    Let .SideBar(mnuNstSubMnu8SideBar) = True
    Set .BitmapImage(mnuNstSubMnu8) = Images("level8").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu9) = Images("level9").ExtractIcon
    Set .BitmapImage(mnuNstSubMnu10) = Images("level10").ExtractIcon
      
    Rem /* #Window Menu. */
    Rem /* Wow now gradient sidebars for your menus too. And as easy as that. Only one caution,
    Rem  * sidebar menus have to be the very *first* menu items. Only that !
    Rem  */
    Let .SideBar(mnuWndSidebar) = True
    
    Rem /* You can have 3D checkmarks exactly like Office or VS has. Too cool. */
    Let .Checked(mnuWndwNew) = True: Let mnuWndwNew.Checked = True
    
    Let .MenuSeperator(mnuWndSep0) = True
    Let .Checked(mnuWndPos(0), Images("cascade").ExtractIcon) = True        ' /* Cascade. */
    Let .Checked(mnuWndPos(1), Images("tilehoriz").ExtractIcon) = True      ' /* Tile Horizontal. */
    Let .Checked(mnuWndPos(2), Images("tilevert").ExtractIcon) = True       ' /* Tile Vertical. */
    Let .Checked(mnuWndPos(3), Images("arngico").ExtractIcon) = True        ' /* Arrange Icons. */
    Call mnuWndPos_Click(0)               ' /* Select the cascade window option. */
  End With
End Sub

Private Sub mnuWndwNew_Click()
  Rem /* Simply toggle back the state. */
  Let mnuWndwNew.Checked = Not mnuWndwNew.Checked
End Sub

Private Sub Text1_GotFocus()
  Rem /* Select the entire text present. */
  With Text1
    Let .SelStart = &H0&                        ' /* Position the cursor to the first position. */
    Let .SelLength = VBA.Strings.Len(.Text)     ' /* Drag past the entire text. */
  End With
End Sub

Private Sub setToolTips()
  Rem /* Now VB developer's can have those *elegant* tooltips. Forget the days when VB Developer's
  Rem  * had to contend with those *simple* tooltips. Now even a VB Developer can have very good
  Rem  * tooltips of *various* styles. Say, how about a balloon tooltip? Want one, just look at the
  Rem  * following code to see how *easily* you could create a balloon tooltip for the textbox!
  Rem  */
  With System.ToolTips
    Call .Add(Form:=Me, _
              Control:=Text1, _
              ToolTipText:=Text1.Text, _
              Key:=Text1.Name, _
              tooltiptitle:="Balloon Tooltips in VB too", _
              tooltipstyle:=TOOLTIP_STYLE.BALLOON, _
              tooltipicon:=SystemInteroperatability.Icon.INFO)
  End With
End Sub

Private Sub addToSystemTray()
  Rem /* System Trays. Am I demonstrating something new? No. Believe me, there's nothing new here!
  Rem  * Then what's new? Allow me to tell you. Now VB Developer's do not have to add something to
  Rem  * the system tray and keep quiet. How about a balloon tooltip popping up (saying your app has
  Rem  * minimized to tray!) Wow! How gratified the user will be. Most of us probably saw this feature
  Rem  * when a LAN network connects or say MSN Messenger is minimized. Doesn't it remind that it has
  Rem  * been minimized to tray and you can access it there? Now VB Developer's are no far away. With
  Rem  * just a little code shown below, even VB Developer's can have *those* tray icons. Have a look..!!
  Rem  */
  With System.SystemTrays
    Call .Add(hWnd:=Me.hWnd, _
              Icon:=Me.Icon, _
              Style:=TRAY_STYLE.BALLOON, _
              message:="Even a system tray is accessible from VB with a simple call" & VBA.Constants.vbNewLine _
                     & "Even multiline messages are easy to use." & VBA.Constants.vbNewLine _
                     & "Click on me to see how easily a system tray communicates" & VBA.Constants.vbNewLine _
                     & "with a VB code", _
              Title:="System Trays for VB too !", _
              messageicon:=TRAY_MESSAGEICON.INFO, _
              traycallbackmessage:=MOUSEOVER)
  End With
End Sub
