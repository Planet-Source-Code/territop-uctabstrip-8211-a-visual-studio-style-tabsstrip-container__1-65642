VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucTabStrip - Test Harness Example"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin prjucTabStrip.ucTabStrip ucTabStrip1 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      GradientStyle   =   1
      TabCount        =   4
      Begin VB.ComboBox cmbGradientStyle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -7550
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "Tab2"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.ListBox lstPropertyItems 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   -10430
         TabIndex        =   25
         Tag             =   "Tab2"
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbTabBackStyle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -1555
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "Tab1"
         Top             =   1200
         Width           =   1255
      End
      Begin VB.CheckBox chkSeparators 
         Caption         =   "Use Tab Separators"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -5155
         TabIndex        =   14
         Tag             =   "Tab1"
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.ComboBox cmbCaptionAlign 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -1555
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tab1"
         Top             =   2400
         Width           =   1255
      End
      Begin VB.ComboBox cmbActiveTab 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -4195
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Tab1"
         Top             =   1200
         Width           =   1255
      End
      Begin VB.TextBox txtActiveTabHeight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -1555
         TabIndex        =   11
         Tag             =   "Tab1"
         Top             =   1800
         Width           =   1255
      End
      Begin VB.ComboBox cmbTabCount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -4195
         TabIndex        =   10
         Tag             =   "Tab1"
         Top             =   1800
         Width           =   1255
      End
      Begin VB.ComboBox cmbTabStyle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -4195
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Tab1"
         Top             =   2400
         Width           =   1255
      End
      Begin VB.ComboBox cmbBorderStyle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -1555
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Tab1"
         Top             =   600
         Width           =   1255
      End
      Begin VB.ComboBox cmbAppearance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -4195
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Tab1"
         Top             =   600
         Width           =   1255
      End
      Begin VB.TextBox txtToolTipText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -1555
         TabIndex        =   6
         Tag             =   "Tab1"
         Top             =   3000
         Width           =   1260
      End
      Begin prjucTabStrip.ucPickBox pbColor 
         Height          =   315
         Left            =   -7550
         TabIndex        =   27
         Tag             =   "Tab2"
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         Filters         =   "Supported files|*.*|All Files (*.*)"
         FolderFlags     =   0
         Printer         =   "False"
         ToolTipText3    =   "Click Here to Locate File"
         ToolTipText4    =   "Click Here to Locate Printer"
      End
      Begin prjucTabStrip.ucPickBox pbFont 
         Height          =   315
         Left            =   -7550
         TabIndex        =   28
         Tag             =   "Tab2"
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DialogType      =   2
         Filters         =   "Supported files|*.*|All Files (*.*)"
         FolderFlags     =   0
         Printer         =   "False"
         ToolTipText3    =   "Click Here to Locate File"
         ToolTipText4    =   "Click Here to Locate Printer"
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Left            =   -15345
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Paul R. Territo, Ph.D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -14145
         TabIndex        =   38
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "v1.0.125 (5/12/2006 11:16:18 PM)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -14145
         TabIndex        =   37
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ucTabStrip.ctl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -14145
         TabIndex        =   36
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblInfoCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Developer:"
         Height          =   255
         Index           =   2
         Left            =   -15225
         TabIndex        =   35
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblInfoCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Build Info:"
         Height          =   255
         Index           =   1
         Left            =   -15225
         TabIndex        =   34
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblInfoCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Control:"
         Height          =   255
         Index           =   0
         Left            =   -15225
         TabIndex        =   33
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblItemColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -7550
         TabIndex        =   32
         Tag             =   "Tab2"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblItemColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Gradient Style:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -7550
         TabIndex        =   31
         Tag             =   "Tab2"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblItemColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -10430
         TabIndex        =   30
         Tag             =   "Tab2"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblItemColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Selected Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -7550
         TabIndex        =   29
         Tag             =   "Tab2"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "TabBackStyle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -2635
         TabIndex        =   24
         Tag             =   "Tab1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "CaptionAlign:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -2635
         TabIndex        =   23
         Tag             =   "Tab1"
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "ActiveTab:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -5155
         TabIndex        =   22
         Tag             =   "Tab1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "TabCount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -5155
         TabIndex        =   21
         Tag             =   "Tab1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "TabStyle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -5155
         TabIndex        =   20
         Tag             =   "Tab1"
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "TabHeight:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -2635
         TabIndex        =   19
         Tag             =   "Tab1"
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "BorderStyle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -2635
         TabIndex        =   18
         Tag             =   "Tab1"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -5155
         TabIndex        =   17
         Tag             =   "Tab1"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTipText:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -2635
         TabIndex        =   16
         Tag             =   "Tab1"
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "click here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4125
         MouseIcon       =   "frmMain.frx":1CCA
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Tag             =   "Tab0"
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to ucTabStrip!"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Tag             =   "Tab0"
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":1FD4
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "Tab0"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Image imWelcomePic 
         Height          =   1545
         Left            =   3450
         Picture         =   "frmMain.frx":2120
         Stretch         =   -1  'True
         Tag             =   "Tab0"
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label lblAuthorMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To provide constructive feedback on this control, please                 ...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Tag             =   "Tab0"
         Top             =   3120
         Width           =   5055
      End
   End
   Begin VB.CheckBox chkEnabled 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled (ucTabStrip && Contained Controls)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3495
      Value           =   1  'Checked
      Width           =   5175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+  File Description:
'       ucTabStrip - Simple Visual Studio Styled TabStrip Container
'
'   Product Name:
'       ucTabStrip.ctl
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Fred.cpp - OffsetColor, TranslateColor, APILine, DrawVGradient, DrawHGradient)
'           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=61476&lngWId=1
'       (Evan Todder - Tab Functionality)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57642&lngWId=1
'           Note: This submission was removed by the author and is not currently on PCS
'       (Nahum III Betancourt Marquez - Circular Gradient)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54561&lngWId=1
'
'   Legal Copyright & Trademarks:
'       Copyright © 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       sold for use as a license in accordance with the terms of the License
'       Agreement in the accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       01Apr06 - Initial UserControl Build
'       02Apr06 - Fixed Separator Alignment Bug
'               - Fixed Row Offset for new controls if they exceed the maximum width of the
'                 the control surface.
'       11May06 - Added TabHeight Property to control the button size
'               - Added tabStyle property to the control
'               - Fixed Tab Control Offset Bug by adding ActiveTab = 0 to Usercontrol Terminate Event
'               - Added Font Property to control the Tab label font style
'               - Fixed bug in the Refresh sub which painted the control recursively
'       12May06 - Added DrawCGradient, DrawDGradient, DrawHGradient, and DrawVGradient
'                 methods and associated routines for gradient background painting
'               - Fixed bug with TabStyle when set to Flat and the separators ramained 3D
'               - Added GradientStart, GradientEnd properties
'               - Added hWnd and hDC properties
'               - Added TabStyle Appearance functionality
'               - Added TabBackStyle functionality
'       13May06 - Added bInitGradient flag to prevent unwanted redrawing of the gradient
'                 on each refresh to improve performance..
'       14May06 - Fixed Drawing bug with DrawRGradient which resulted in the final rectangle
'                 not being drawn when X=X2 and/or Y=Y2 so the delta = 0.
'               - Added all public events expected in a UserControl plus some custom ones...
'               - Added Accelerator Keys functionality
'               - Added TabHeight Auto Adjustment for font sizes which exceed the current
'                 TabHeight and updated the TabHeight property on change. This is achieved
'                 through setting the AutoSize property on the Label to True in the Refresh
'                 method and then setting this to False once we call the MoveControls method.
'               - Fixed bug with the Font property which did not "set" the m_Font private variable.
'               - Fixed bug in MoveControls method which did not loop through all controls.
'       19May06 - Added All API method for DrawCGradient
'       11Jun06 - Added Additional In-Line comments for clarity
'
'   Force Declarations
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'   Link URL address which searches for our control submission on PCS
Const sLink As String = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&?lngWId=1&grpCategories=&txtMaxNumberOfEntriesPerPage=10&optSort=Alphabetical&chkThoroughSearch=&blnTopCode=False&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=1&intLastRecordOnPage=10&intMaxNumberOfEntriesPerPage=10&intLastRecordInRecordset=499&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=ucTabStrip"

Private bInternal As Boolean
Private m_Caption(3) As String

Private Sub AddCaptions()
    Dim i As Long
    
    On Error Resume Next
    
    With Me
        For i = 0 To 3
            If .ucTabStrip1.TabCount - 1 >= i Then
                .ucTabStrip1.Caption(i) = m_Caption(i)
            End If
        Next i
    End With
    
    On Error GoTo 0

End Sub

Private Sub chkEnabled_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.Enabled = IIf(.chkEnabled.Value = vbChecked, True, False)
        End If
    End With
End Sub

Private Sub chkSeparators_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.Separators = IIf(.chkSeparators.Value = vbChecked, True, False)
        End If
    End With
End Sub

Private Sub cmbActiveTab_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.ActiveTab = .cmbActiveTab.ListIndex
        End If
    End With
End Sub

Private Sub cmbAppearance_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.Appareance = .cmbAppearance.ListIndex
        End If
    End With
End Sub

Private Sub cmbBorderStyle_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.BorderStyle = .cmbBorderStyle.ListIndex
        End If
    End With
End Sub

Private Sub cmbCaptionAlign_Click()
    With Me
        If Not bInternal Then
            Select Case .cmbCaptionAlign.ListIndex
                Case 0
                    .ucTabStrip1.CaptionAlign = vbLeftJustify
                Case 1
                    .ucTabStrip1.CaptionAlign = vbRightJustify
                Case Else
                    .ucTabStrip1.CaptionAlign = vbCenter
            End Select
        End If
    End With
End Sub

Private Sub cmbGradientStyle_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.GradientStyle = .cmbGradientStyle.ListIndex
        End If
    End With
End Sub

Private Sub cmbTabBackStyle_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.TabBackStyle = .cmbTabBackStyle.ListIndex
        End If
    End With
End Sub

Private Sub cmbTabCount_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.TabCount = .cmbTabCount.ListIndex + 1
            Call RebuildList
            Call AddCaptions
        End If
    End With
End Sub

Private Sub cmbTabCount_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If Not bInternal Then
            If KeyCode = 13 Then
                If IsNumeric(.cmbTabCount.Text) And (.cmbTabCount.Text > 0) Then
                    .ucTabStrip1.TabCount = .cmbTabCount.Text
                    Call RebuildList
                    Call AddCaptions
                End If
            End If
        End If
    End With
End Sub

Private Sub cmbTabCount_LostFocus()
    If Not bInternal Then
        Call cmbTabCount_KeyDown(13, 0)
    End If
End Sub

Private Sub cmbTabStyle_Click()
    With Me
        If Not bInternal Then
            .ucTabStrip1.TabStyle = .cmbTabStyle.ListIndex
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    With Me
        bInternal = True
        With .ucTabStrip1
            'Our control array is "Zero" based ;-D
            .TabCount = 4
            .ActiveTab = 0
            m_Caption(0) = " &Welcome... "
            m_Caption(1) = " &Settings "
            m_Caption(2) = " &Colors "
            m_Caption(3) = " &Version "
            '   Add the captions
            Call AddCaptions
            .CaptionAlign = vbCenter
            '   These are the defaults, but we will enforce them
            '   anyhow just to make sure it loads correctly...
            .BackColor = vbButtonFace
            .ActiveForeColor = &H0
            .InActiveForeColor = &H80000011
            .HoverColor = &H8000000C
            .TabHeight = 20
            .TabStyle = ts3D
            .Separators = True
        End With
        '   Fill the comboboxes
        With .cmbAppearance
            .AddItem "Flat"
            .AddItem "3D"
            .ListIndex = 0
        End With
        With .cmbBorderStyle
            .AddItem "tsNone"
            .AddItem "tsFixedSingle"
            .ListIndex = 0
        End With
        With .cmbCaptionAlign
            .AddItem "vbLeftJustify"
            .AddItem "vbRightJustify"
            .AddItem "vbCenter"
            .ListIndex = 2
        End With
        With .cmbTabBackStyle
            .AddItem "tsTransparent"
            .AddItem "tsOpaque"
            .ListIndex = 0
        End With
        With .cmbTabCount
            For i = 1 To 20
                .AddItem i
            Next i
            .ListIndex = 3
        End With
        With .cmbActiveTab
            For i = 0 To 3
                .AddItem i
            Next i
            .ListIndex = 0
        End With
        With .cmbTabStyle
            .AddItem "Flat"
            .AddItem "3D"
            .ListIndex = 1
        End With
        With .cmbGradientStyle
            .AddItem "tsNoGradient"
            .AddItem "tsCircular"
            .AddItem "tsDiagonalNWSE"
            .AddItem "tsDiagonalSWNE"
            .AddItem "tsHorizontal"
            .AddItem "tsRectangular"
            .AddItem "tsVertical"
            .ListIndex = 2
        End With
        With .lstPropertyItems
            .AddItem "ActiveForeColor"
            .AddItem "BackColor"
            .AddItem "GradientEnd"
            .AddItem "GradientStart"
            .AddItem "HoverColor"
            .AddItem "InActiveForeColor"
            .AddItem "SeparatorColor"
            .AddItem "TabBackColor"
            .ListIndex = -1
        End With
        With .txtActiveTabHeight
            .Text = ucTabStrip1.TabHeight
        End With
        '   Fill the current verison data
        .lblInfo(1).Caption = .ucTabStrip1.Version(True)
        
        bInternal = False
    End With
End Sub

Private Sub lblAuthorMessage_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
    End With
End Sub

Private Sub lblLink_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
    End With
End Sub

Private Sub lstPropertyItems_Click()
    With Me
        Select Case .lstPropertyItems.ListIndex
            Case 0  'ActiveForeColor
                .pbColor.Color = .ucTabStrip1.ActiveForeColor
            Case 1  'BackColor
                .pbColor.Color = .ucTabStrip1.BackColor
            Case 2  'GradientEnd
                .pbColor.Color = .ucTabStrip1.GradientEnd
            Case 3  'GradientStart
                .pbColor.Color = .ucTabStrip1.GradientStart
            Case 4  'HoverColor
                .pbColor.Color = .ucTabStrip1.HoverColor
            Case 5  'InActiveForeColor
                .pbColor.Color = .ucTabStrip1.InActiveForeColor
            Case 6  'SeparatorColor
                .pbColor.Color = .ucTabStrip1.SeparatorColor
            Case 7  'TabBackColor
                .pbColor.Color = .ucTabStrip1.TabBackColor
        End Select
    End With
End Sub

Private Sub pbColor_Click()
    With Me
        Select Case .lstPropertyItems.ListIndex
            Case 0  'ActiveForeColor
                .ucTabStrip1.ActiveForeColor = .pbColor.Color
            Case 1  'BackColor
                .ucTabStrip1.BackColor = .pbColor.Color
            Case 2  'GradientEnd
                .ucTabStrip1.GradientEnd = .pbColor.Color
            Case 3  'GradientStart
                .ucTabStrip1.GradientStart = .pbColor.Color
            Case 4  'HoverColor
                .ucTabStrip1.HoverColor = .pbColor.Color
            Case 5  'InActiveForeColor
                .ucTabStrip1.InActiveForeColor = .pbColor.Color
            Case 6  'SeparatorColor
                .ucTabStrip1.SeparatorColor = .pbColor.Color
            Case 7  'TabBackColor
                .ucTabStrip1.TabBackColor = .pbColor.Color
        End Select
    End With
End Sub

Private Sub pbFont_Click()
    With Me
        If Not bInternal Then
            '   Set the selected font
            Set .ucTabStrip1.Font = .pbFont.Font
            .txtActiveTabHeight.Text = .ucTabStrip1.TabHeight
        End If
    End With
End Sub

Private Sub RebuildList()
    Dim i As Long
    
    With Me
        .cmbActiveTab.Clear
        For i = 0 To .ucTabStrip1.TabCount - 1
            .cmbActiveTab.AddItem i
        Next i
        Call ucTabStrip1_TabSwitch(1)
    End With
End Sub

Private Sub txtActiveTabHeight_Change()
    With Me
        If Not bInternal Then
            If .txtActiveTabHeight.Text = "" Then Exit Sub
            If IsNumeric(.txtActiveTabHeight.Text) And (.txtActiveTabHeight.Text > 0) Then
                .ucTabStrip1.TabHeight = .txtActiveTabHeight.Text
            End If
        End If
    End With
End Sub

Private Sub txtToolTipText_Change()
    With Me
        If Not bInternal Then
            .ucTabStrip1.ToolTipText = .txtToolTipText.Text
        End If
    End With
End Sub


Private Sub ucTabStrip1_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
    Debug.Print "BeforeTabSwitch: " & iNewActiveTab & ", " & bCancel
End Sub

Private Sub ucTabStrip1_Click()
    Debug.Print "Click: "
End Sub

Private Sub ucTabStrip1_DblClick()
    Debug.Print "DblClick: "
End Sub

Private Sub ucTabStrip1_GotFocus()
    Debug.Print "GotFocus: "
End Sub

Private Sub ucTabStrip1_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyDown: KeyCode(" & KeyCode & "), Shift(" & Shift & ")"
End Sub

Private Sub ucTabStrip1_KeyPress(KeyAscii As Integer)
    Debug.Print "KeyPress: KeyAscii(" & KeyAscii & ")"
End Sub

Private Sub ucTabStrip1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyUp: KeyCode(" & KeyCode & "), Shift(" & Shift & ")"
End Sub

Private Sub ucTabStrip1_LostFocus()
    Debug.Print "LostFocus: "
End Sub

Private Sub ucTabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseDown: Button(" & Button & "), Shift(" & Shift & "), X(" & X & "), Y(" & Y & ")"
End Sub

Private Sub ucTabStrip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseMove: Button(" & Button & "), Shift(" & Shift & "), X(" & X & "), Y(" & Y & ")"
End Sub

Private Sub ucTabStrip1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseUp: Button(" & Button & "), Shift(" & Shift & "), X(" & X & "), Y(" & Y & ")"
End Sub

Private Sub ucTabStrip1_TabDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "TabDown: Index(" & Index & "), Button(" & Button & "), Shift(" & Shift & "), X(" & X & "), Y(" & Y & ")"
End Sub

Private Sub ucTabStrip1_TabHover(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "TabHover: Index(" & Index & "), Button(" & Button & "), Shift(" & Shift & "), X(" & X & "), Y(" & Y & ")"
End Sub

Private Sub ucTabStrip1_TabSwitch(iLastActiveTab As Integer)
    With Me
        bInternal = True
        If .cmbActiveTab.ListCount > 1 Then
            .cmbActiveTab.ListIndex = .ucTabStrip1.ActiveTab
        End If
        bInternal = False
        Debug.Print "TabSwitch: LastActiveTab(" & iLastActiveTab & ")"
    End With
End Sub

Private Sub ucTabStrip1_TabUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "TabUp: Index(" & Index & "), Button(" & Button & "), Shift(" & Shift & "), X(" & X & "), Y(" & Y & ")"
End Sub


