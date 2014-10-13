VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#4.0#0"; "HTE_TabView6.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   10065
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4875
      Left            =   165
      TabIndex        =   40
      Top             =   480
      Visible         =   0   'False
      Width           =   9770
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   7455
         TabIndex        =   51
         Top             =   0
         Width           =   7455
         Begin VB.OptionButton optDirection 
            Caption         =   "Desc"
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   45
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optDirection 
            Caption         =   "Asc"
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   44
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.ComboBox cmbSortList 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton cmdUndo 
            Caption         =   "&Undo"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4920
            TabIndex        =   50
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "&New"
            Height          =   495
            Left            =   5760
            TabIndex        =   46
            Top             =   0
            Width           =   810
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   495
            Left            =   6600
            TabIndex        =   47
            Top             =   0
            Width           =   810
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4080
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4700
         Left            =   165
         ScaleHeight     =   4695
         ScaleWidth      =   9555
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   120
         Width           =   9560
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   4215
            Left            =   0
            TabIndex        =   42
            Top             =   405
            Width           =   9465
            _cx             =   16695
            _cy             =   7435
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmMain.frx":164A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   1
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   1
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   0
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label4 
            Caption         =   "Alias Configuration"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   2415
         End
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4875
      Left            =   165
      TabIndex        =   29
      Top             =   480
      Visible         =   0   'False
      Width           =   9770
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   165
         ScaleHeight     =   225
         ScaleWidth      =   9555
         TabIndex        =   37
         Top             =   120
         Width           =   9560
         Begin VB.CheckBox chkAutoDetail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Auto-Scroll"
            ForeColor       =   &H80000008&
            Height          =   235
            Left            =   7800
            TabIndex        =   21
            Top             =   20
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkLockDetail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Lock-View"
            ForeColor       =   &H80000008&
            Height          =   235
            Left            =   5640
            TabIndex        =   20
            Top             =   20
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Detail Messages"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   3015
         End
      End
      Begin HTE_TabView6.TabControl tcDetail 
         Height          =   4215
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7435
         TabAlign        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483629
         ShowCloseButton =   0   'False
         UnpinnedWidth   =   9255
         Begin VB.TextBox txtDetail 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Text            =   "frmMain.frx":1789
            Top             =   0
            Visible         =   0   'False
            Width           =   9590
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4875
      Left            =   165
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   9770
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4700
         Left            =   165
         ScaleHeight     =   4695
         ScaleWidth      =   9555
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
         Width           =   9560
         Begin VB.ListBox lstProcesses 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            Height          =   2760
            Left            =   2520
            OLEDropMode     =   1  'Manual
            TabIndex        =   10
            Top             =   720
            Width           =   3015
         End
         Begin VB.ListBox lstChain 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            Height          =   2760
            Left            =   6240
            OLEDropMode     =   1  'Manual
            TabIndex        =   15
            Top             =   720
            Width           =   3015
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   18
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5640
            TabIndex        =   11
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   18
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5640
            TabIndex        =   12
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdUp 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   18
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5640
            TabIndex        =   13
            Top             =   1920
            Width           =   495
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   18
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5640
            TabIndex        =   14
            Top             =   2520
            Width           =   495
         End
         Begin VB.ListBox lstRoutes 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            Height          =   1785
            Left            =   0
            TabIndex        =   5
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmMain.frx":1795
            Left            =   960
            List            =   "frmMain.frx":179F
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddRoute 
            Caption         =   "Add"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Add Route"
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemoveRoute 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            ToolTipText     =   "Remove Route"
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton cmdProperties 
            Caption         =   "Properties"
            Height          =   375
            Left            =   8040
            TabIndex        =   16
            ToolTipText     =   "Properties"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkServer 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Route is server process?"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Message Routing"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Available Processes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   34
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Process Chain"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   6240
            TabIndex        =   33
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Routes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Inbound:"
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   2640
            Width           =   1455
         End
      End
      Begin VB.Label lblWhatAmI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         TabIndex        =   28
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label lblInstance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Label lblProgID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6360
         TabIndex        =   26
         Top             =   3600
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4875
      Left            =   165
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   9770
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4700
         Left            =   165
         ScaleHeight     =   4695
         ScaleWidth      =   9555
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   120
         Width           =   9560
         Begin VB.TextBox txtProcessed 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4215
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "frmMain.frx":17AF
            Top             =   405
            Width           =   9570
         End
         Begin VB.CheckBox chkLock 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Lock-View"
            ForeColor       =   &H80000008&
            Height          =   235
            Left            =   5640
            TabIndex        =   17
            Top             =   20
            Width           =   1695
         End
         Begin VB.CheckBox chkAuto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Auto-Scroll"
            ForeColor       =   &H80000008&
            Height          =   235
            Left            =   7800
            TabIndex        =   18
            Top             =   20
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Messages Processed"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   2415
         End
      End
   End
   Begin VB.CheckBox chkSysTray 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Show in System Tray?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   5550
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Restart"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CheckBox chkAllowExit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Allow Exit Menu?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   5550
      Width           =   2295
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9340
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Configuration"
            Key             =   "ts01"
            Object.Tag             =   "Configuration"
            Object.ToolTipText     =   "Process and routes configuration"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Messages"
            Key             =   "ts02"
            Object.Tag             =   "Messages"
            Object.ToolTipText     =   "Current messages being processed"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Detail"
            Key             =   "ts03"
            Object.Tag             =   "Detail"
            Object.ToolTipText     =   "By process message being processed"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Aliasing"
            Key             =   "ts04"
            Object.Tag             =   "Aliasing"
            Object.ToolTipText     =   "Configure how various endpoints will be described"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ILStatusList 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2239
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":277B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListGlyph 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CBD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrator"
      Visible         =   0   'False
      Begin VB.Menu mnuAConfigure 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuANevermind 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuASep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRoutes 
      Caption         =   "&Routes"
      Begin VB.Menu mnuRAddRoute 
         Caption         =   "Add Route"
      End
      Begin VB.Menu mnuRRemoveRoute 
         Caption         =   "Remove Route"
      End
      Begin VB.Menu mnuRSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRSaveRoute 
         Caption         =   "Save Changes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRUndoRoute 
         Caption         =   "Undo Changes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuRExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuProcesses 
      Caption         =   "Pr&ocesses"
      Begin VB.Menu mnuPSAddProcess 
         Caption         =   "Add"
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "&Process"
      Begin VB.Menu mnuPProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuPSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRemoveProcess 
         Caption         =   "Remove Process"
      End
      Begin VB.Menu mnuPProcessUp 
         Caption         =   "Move Process Up"
      End
      Begin VB.Menu mnuPProcessDown 
         Caption         =   "Move Process Down"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements HTE_GPS.AppStatus 'Application status callback
Implements HTE_GPS.PropPageCallback 'PropertyPage callback - for new uninitialized processes

Private gApplication As HTE_GPS.Application 'Main application object
Attribute gApplication.VB_VarHelpID = -1
Private Const cModuleName = "HTE_GPSInterface.frmMain"
Private curHostStatus As HTE_GPS.GPS_HOST_STATUS  'Added to allow Admin to show in task tray, keep last status in memory for display
Private WithEvents sysTray As frmSysTray
Attribute sysTray.VB_VarHelpID = -1
Private WithEvents propPage As frmProperties
Attribute propPage.VB_VarHelpID = -1
Private Const cWindowTitle = "SunGard Public Sector GPS Configuration Application"
Private appSettings As cApplication
Private availProcesses As cProcesses
Private m_Admin As Boolean, m_Exit As Boolean
Private m_Server As Boolean 'allows aliasing feature to be displayed
Private cHeader As String
'Functions
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW As Long = 5
Private Enum reasonClosing
    rcValidRequest = 0
    rcNotSoleInstance = 1
    rcCosmetic = 2
End Enum
Private peCloseReason As reasonClosing

'XP from manifest
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
Private IsRunningService As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private ds As HTE_GPSData.DataSource 'used when alias frame becomes active
Private rs As ADODB.Recordset
Private m_ModifiedData As Boolean 'used to determine if compact/repair is required
'faster than vb6 strcomp for equality (even with binarycompare)
'http://www.xbeat.net/vbspeed/c_IsSameText.htm
'usage: fRet = (lstrcmpi(sDum1, sDum2) = 0)
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Function initializeForXP() As Boolean
Dim iccex As tagInitCommonControlsEx
On Error Resume Next
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   initializeForXP = (Err.Number <> 0)
End Function

Private Sub AppStatus_Processed(ByVal Message As HTE_GPS.GPSMessage, ByVal currentRoute As String)
Dim b() As Byte
Dim sMsg As String
Dim i As Long
    Select Case UCase$(TabStrip1.Tag)
    Case "FRAME2"
        If Me.Visible And Not -(chkLock.Value) Then
            DisplayProcessedMessage txtProcessed, Message, currentRoute
            If -chkAuto.Value Then txtProcessed.SelStart = Len(txtProcessed.Text)
        End If
    Case "FRAME3"
        If Not -chkLockDetail.Value Then
             For i = 0 To txtDetail.Count - 1
                 If InStr(1, currentRoute, txtDetail(i).Tag, vbTextCompare) > 0 Then
                    DisplayProcessedMessage txtDetail(i), Message, currentRoute
                    If -chkAutoDetail.Value Then txtDetail(i).SelStart = Len(txtDetail(i).Text)
                    Exit For
                End If
            Next
        End If
    End Select
End Sub

Private Function DisplayProcessedMessage(ByRef thisText As TextBox, ByVal Message As HTE_GPS.GPSMessage, ByVal currentRoute As String) As String
Dim b() As Byte
Dim sMsg As String
Dim sMessage As String
    If Message.rawMessage = vbNullString Then
        sMsg = " Total Bytes = 0"
    Else
        b = StrConv(Message.rawMessage, vbFromUnicode)
        sMsg = HexDump(VarPtr(b(0)), UBound(b) + 1)
    End If
    sMessage = cHeader & _
            Space$(1) & currentRoute & _
            "; Type - " & appSettings.typeDescription(Message.Type) & _
            "; MessageStatus - " & messageStatusDesc(Message.MessageStatus) & _
            vbCrLf & _
            cHeader & _
            sMsg & _
            vbCrLf
    If Not thisText Is Nothing Then
        thisText.Text = thisText.Text & sMessage
        If Len(thisText.Text) >= 65535 Then thisText.Text = vbNullString
    End If
    DisplayProcessedMessage = sMessage
    
    
End Function

Private Sub AppStatus_StatusChange(statusCode As HTE_GPS.GPS_HOST_STATUS)
Dim i As Long
    SetStatus statusCode
    FindAllDetailStatus
    If Not propPage Is Nothing Then
        If propPage.Visible Then
            i = getCurrentProcessImage
            propPage.imgStatus.Picture = propPage.ILStatusList.ListImages(i).Picture
        End If
    End If
End Sub

Private Sub chkAllowExit_Click()
    appSettings.AllowExit = -chkAllowExit.Value
    m_Exit = appSettings.AllowExit
    EnableActions
End Sub

Private Sub chkServer_Click()
    If lstRoutes.ListIndex <> -1 Then
        If -chkServer.Value <> appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).ServerProcess Then
'''            appSettings.flagChanges
            appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).ServerProcess = -chkServer.Value
            appSettings.commitChanges
        End If
    End If
    EnableActions
End Sub

Private Sub chkSysTray_Click()
    If appSettings.ShowInTray <> -chkSysTray.Value Then
'''        appSettings.flagChanges
        appSettings.ShowInTray = -chkSysTray.Value
        appSettings.commitChanges
    End If
    SysTrayDest
    If appSettings.ShowInTray Then SysTrayInit
    EnableActions
End Sub

Private Sub cmbSortList_Click()
'Dim ds As HTE_GPSData.DataSource 'HTE_GPSData.Identities '
    Screen.MousePointer = vbHourglass
    'If ds Is Nothing Then Set ds = New HTE_GPSData.DataSource 'HTE_GPSData.DataSource
    With VSFlexGrid1
        .Editable = flexEDKbdMouse
        .Redraw = flexRDNone
        .DataMode = flexDMBoundImmediate 'flexDMBoundBatch
        .FocusRect = flexFocusSolid
'        .FlexDataSource = ds
        Set rs = ds.GridView(True, cmbSortList.ItemData(cmbSortList.ListIndex), (optDirection(0).Value))
        Set .DataSource = rs 'ds.GridView(True, cmbSortList.ItemData(cmbSortList.ListIndex), (optDirection(0).Value))
        .Redraw = flexRDBuffered
        cmdDelete.Enabled = (VSFlexGrid1.Rows > 1) And VSFlexGrid1.Row > -1 'cmdNew.Enabled
        If VSFlexGrid1.Visible Then VSFlexGrid1.SetFocus
    End With
'    Set ds = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbType_Click()
'''    appSettings.flagChanges
    If lstRoutes.ListIndex <> -1 Then
        If appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).InboundType <> cmbType.ItemData(cmbType.ListIndex) Then
'''            appSettings.flagChanges
            appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).InboundType = cmbType.ItemData(cmbType.ListIndex)
            appSettings.commitChanges
        End If
    End If
    EnableActions
End Sub

Private Sub cmdAdd_Click()
    'Add New Process to Process Chain
    If lstProcesses.ListIndex <> -1 And lstRoutes.ListIndex <> -1 Then
        AddProcess
'''        appSettings.flagChanges
        appSettings.commitChanges
        EnableActions
    End If
End Sub

Private Sub AddProcess()
Dim maxVal As Long, sMaxVal As String, sInstance As String
Dim oProcess As cProcess, oRoute As cRoute, newProcess As cProcess
    If lstProcesses.ListIndex <> -1 And lstRoutes.ListIndex <> -1 Then
        'Find the maximum instance identifier and increment by 1 - for readability and uniqueness
        Screen.MousePointer = vbHourglass
        Set oRoute = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex))
        For Each oProcess In oRoute 'lstChain.ListCount - 1
            sMaxVal = Right$(oProcess.InstanceID, 2)
            If IsNumeric(sMaxVal) Then
                If CLng(sMaxVal) > maxVal Then maxVal = CLng(sMaxVal)
            End If
        Next
        sInstance = "R" & Right$(oRoute.RouteID, 2) & Format$(maxVal + 1, "00")
        Set newProcess = New cProcess
        Set oProcess = availProcesses.Item(lstProcesses.List(lstProcesses.ListIndex))
        'CREATE A NEW INSTANCE OTHERWISE MODIFIES EXISTING INSTANCE!!!
        newProcess.friendlyName = oProcess.friendlyName
        newProcess.progID = oProcess.progID
        newProcess.InstanceID = sInstance
        lstChain.AddItem newProcess.friendlyName
        oRoute.Add newProcess
        'Prepare the instance identifier
        lstChain.ItemData(lstChain.NewIndex) = getInstanceValue(newProcess.InstanceID)
        lstChain.ListIndex = lstChain.NewIndex
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdAddRoute_Click()
    Screen.MousePointer = vbHourglass
    AddRoute
    EnableActions
    Screen.MousePointer = vbDefault
End Sub

Private Sub AddRoute()
Dim oRoute As cRoute
Dim lMax As Long
    
    For Each oRoute In appSettings
        If CLng(Right$(oRoute.RouteID, 2)) > lMax Then
            lMax = CLng(Right$(oRoute.RouteID, 2))
        End If
    Next
    Set oRoute = New cRoute
    oRoute.RouteID = "ROUTE" & Format$(CStr(lMax + 1), "00")
'''    appSettings.flagChanges
    appSettings.Add oRoute
    appSettings.commitChanges
    lstRoutes.AddItem oRoute.RouteID
    lstRoutes.ListIndex = lstRoutes.NewIndex
    lstRoutes_Click

End Sub

Private Sub cmdClose_Click()
Dim Cancel As Integer
    If appSettings.IsDirty Then appSettings.commitChanges
    Form_QueryUnload Cancel, vbFormControlMenu
End Sub

Private Sub cmdDelete_Click()
Dim curRow As Long, curCol As Long
    Debug.Print "cmdDelete_Click Enter"
    If VSFlexGrid1.Row = -1 Then Exit Sub
    If MsgBox("Delete currently selected record(s). " & vbCrLf & _
                    "Are you sure you want to do this?", _
                    vbYesNo Or vbQuestion Or vbDefaultButton2, "Delete Selected?") = vbYes Then
        curRow = VSFlexGrid1.Row
        curCol = VSFlexGrid1.Col
'   On Error Resume Next
        rs.Delete
        rs.UpdateBatch
        With VSFlexGrid1
            .DataRefresh
            .SetFocus
            SendKeys "{DOWN}"
            cmdDelete.Enabled = (.Rows > 1) And VSFlexGrid1.Row > -1 And cmdNew.Enabled
        End With
        NotifyOfAliasChanges
'        If VSFlexGrid1.Rows > 1 Then
'            VSFlexGrid1.Select VSFlexGrid1.TopRow, curCol
'            VSFlexGrid1.EditCell
'            VSFlexGrid1.TextMatrix(VSFlexGrid1.TopRow, 0) = 0
'        End If
    End If
    Debug.Print "cmdDelete_Click Exit" & VSFlexGrid1.Row
End Sub

Private Sub cmdDown_Click()
    Screen.MousePointer = vbHourglass
    MoveProcess False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNew_Click()
    Dim rs As ADODB.Recordset
    Set rs = VSFlexGrid1.DataSource
    If rs.State = adStateOpen Then
        With VSFlexGrid1
            rs.AddNew
            rs.UpdateBatch
            .SetFocus
            .DataRefresh
            cmdDelete.Enabled = VSFlexGrid1.Rows > 1 And VSFlexGrid1.Row > -1
            NotifyOfAliasChanges
        End With
    End If
    'cmbSortList_Click
End Sub

Private Sub cmdProperties_Click()
    If lstChain.ListIndex > -1 Then
        lstChain_DblClick
    End If
End Sub

Private Sub cmdRemove_Click()
'Remove current process from chain
    Screen.MousePointer = vbHourglass
    RemoveProcess True
'''    appSettings.flagChanges
    appSettings.commitChanges
    TerminateApplication
    InitializeApplication
    getInstanceDescriptions
    EnableActions
    Screen.MousePointer = vbDefault
End Sub

Private Sub RemoveRoute()
    appSettings.Remove lstRoutes.List(lstRoutes.ListIndex)
    lstRoutes.RemoveItem lstRoutes.ListIndex
    If lstRoutes.ListCount > 0 Then lstRoutes.ListIndex = 0
End Sub

Private Sub cmdRemoveRoute_Click()
Dim i As Long
    'Remove Current Route from config file
    Screen.MousePointer = vbHourglass
    If lstRoutes.ListIndex <> -1 Then
        For i = 0 To lstChain.ListCount - 1
            lstChain.ListIndex = 0
            RemoveProcess
        Next
        RemoveRoute
'''        appSettings.flagChanges
        appSettings.commitChanges
        TerminateApplication
        EnableActions
        InitializeApplication
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub InitializeApplication()
On Local Error Resume Next
Const retries As Integer = 10
Dim X As Integer
    For X = 0 To retries
        Set gApplication = New HTE_CLIENTUTILITIES.Application
        If Not gApplication Is Nothing Then
            gApplication.StatusCallback Me
            Exit For
        Else
            Sleep 175
            DoEvents
        End If
    Next
End Sub

Private Sub TerminateApplication(Optional ByVal ReasonForClosure As reasonClosing = rcValidRequest)
On Local Error Resume Next
    If Not gApplication Is Nothing Then gApplication.CleanUp
    If Err.Number <> 0 Then UEH_LogError cModuleName, "TerminateApplication", Err
    Set gApplication = Nothing
    'Commented out it was put in place to "clean-up" after MCS but never belonged here!!
    'If ReasonForClosure = reasonClosing.rcValidRequest Then KillApp "HTE_ClientUtilities.exe"
End Sub

Private Sub cmdSave_Click()
Dim rs As ADODB.Recordset
    cmdSave.Enabled = False
    cmdUndo.Enabled = False
    cmbSortList.Enabled = True
    optDirection(0).Enabled = True
    optDirection(1).Enabled = True
    cmdDelete.Enabled = (VSFlexGrid1.Rows > 1) And VSFlexGrid1.Row > -1
    cmdNew.Enabled = True
On Error Resume Next
    Set rs = VSFlexGrid1.DataSource
    rs.UpdateBatch ' try to update the recordset (it may fail integrity rules)
    If Err <> 0 Then 'And Err <> -2147217885 Then 'not record deleted message
        MsgBox "Update was incomplete, possibly because some changes failed database integrity rules." & vbCrLf & _
               "Any valid changes have been committed.", vbInformation, "Aliasing Update"
        cmbSortList_Click
    End If
    NotifyOfAliasChanges
End Sub
Private Sub NotifyOfAliasChanges()
    'If ds Is Nothing Then Set ds = New HTE_GPSData.DataSource
    ds.NotifyOfChanges
    m_ModifiedData = True
End Sub
Private Sub cmdStartStop_Click()
    Screen.MousePointer = vbHourglass
    If appSettings.IsDirty Then appSettings.commitChanges
    StopServiceIfApplies
    TerminateApplication
    InitializeApplication
    If Not gApplication Is Nothing Then gApplication.ShowLastMessage = (TabStrip1.SelectedItem.index <> 1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUndo_Click()
Dim rs As ADODB.Recordset
    cmdSave.Enabled = False
    cmdUndo.Enabled = False
    cmbSortList.Enabled = True
    optDirection(0).Enabled = True
    optDirection(1).Enabled = True
    cmdDelete.Enabled = (VSFlexGrid1.Rows > 1) And VSFlexGrid1.Row > -1
    cmdNew.Enabled = True
    Set rs = VSFlexGrid1.DataSource
    rs.CancelBatch
End Sub

Private Sub cmdUp_Click()
    Screen.MousePointer = vbHourglass
    MoveProcess
    Screen.MousePointer = vbDefault
End Sub

Private Sub MoveProcess(Optional ByVal bUp As Boolean = True)
Dim sText As String, lItemData As Long
Dim FromIndex As Long, ToIndex As Long
Dim oProcess As cProcess, oRoute As cRoute
Dim progInstance As Collection
    FromIndex = lstChain.ListIndex
    ToIndex = FromIndex + IIf(bUp, -1, 1)
    If FromIndex < 0 Then FromIndex = lstChain.ListIndex
    ' exit if argument not in range
    If ToIndex < 0 Or ToIndex > lstChain.ListCount - 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    LockWindowUpdate lstChain.hWnd
    With lstChain
        sText = .List(FromIndex) ' save text of the current item
        lItemData = .ItemData(FromIndex) ' save data of the current item
        .RemoveItem FromIndex ' remove the item
        .AddItem sText, ToIndex ' re-add the item text
        .ItemData(.NewIndex) = lItemData ' re-add the item data
        .ListIndex = ToIndex ' select the new item
    End With
    'Remove All Routes
On Error Resume Next
    Set oRoute = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex))
    If Not oRoute Is Nothing Then
        Set progInstance = New Collection
        For Each oProcess In oRoute
            progInstance.Add oProcess.progID, oProcess.InstanceID
            oRoute.Remove oProcess.InstanceID
        Next
        'Readd Routes in correct order
        For ToIndex = 0 To lstChain.ListCount - 1
            Set oProcess = New cProcess
            oProcess.friendlyName = lstChain.List(ToIndex)
            oProcess.InstanceID = getInstanceID(lstChain.ItemData(ToIndex))
            oProcess.progID = progInstance.Item(oProcess.InstanceID)
            oRoute.Add oProcess
        Next
'''        appSettings.flagChanges
        appSettings.commitChanges
    End If
    LockWindowUpdate 0&
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
    initializeForXP
End Sub
Private Sub StopServiceIfApplies()
    If IsInstalledService Then
        Select Case GetServiceStatus
            Case SERVICE_RUNNING, SERVICE_START_PENDING
                IsRunningService = True
                'stop service to continue
                Do Until StopNTService = 0
                    Sleep 500
                Loop
            Case Else
        End Select
    End If
End Sub
Private Sub Form_Load()
Dim m_Status As HTE_GPS.GPS_HOST_STATUS
Dim oObj As HTE_GPS.Application
Dim prevhWnd As Long
    
    m_Admin = InStr(1, Command(), "-M", vbTextCompare) > 0 Or InStr(1, Command(), "/M", vbTextCompare) > 0
    m_Server = InStr(1, Command(), "-Server", vbTextCompare) > 0 Or InStr(1, Command(), "/Server", vbTextCompare) > 0
    IsRunningService = False
    StopServiceIfApplies
    If SoleInstance Then
        MaintainDb False
        m_Exit = True 'In case of a catastrophic event...allow them to shutdown!!!
        App.OleRequestPendingMsgTitle = "Global Positioning Configuration Interface"
        App.OleServerBusyMsgTitle = "Global Positioning Configuration Interface"
        App.OleServerBusyMsgText = "The Global Positioning Configuration Interface is busy trying to complete a task." & vbCrLf & "If the situation persists, report the problem to your system administrator."
        App.OleRequestPendingTimeout = 120000
        App.OleServerBusyTimeout = 240000
        App.OleRequestPendingMsgText = "The Global Positioning Configuration Interface is waiting for a task to complete." & vbCrLf & "If the situation persists, report the problem to your system administrator."
        App.OleServerBusyRaiseError = True 'This should nullify any ActiveX message box error, and raise an error in the corresponding routine
        App.TaskVisible = False
        InitializeApplication
        Me.Caption = cWindowTitle 'set caption AFTER we recognize we are sole instance
        Set appSettings = New cApplication
        Set availProcesses = New cProcesses
        m_Exit = appSettings.AllowExit
        cHeader = Space$(1) & String(86, "=") & vbCrLf
        chkSysTray.Value = Abs(appSettings.ShowInTray)
        TabStrip1.Tabs(1).Selected = True
        InitializeFrame (1)
        If lstRoutes.ListCount > 0 Then lstRoutes.ListIndex = 0
        ILStatusList.MaskColor = vbMagenta: ILStatusList.UseMaskColor = True
    On Error Resume Next
        tcDetail.ImageList = Me.ILStatusList
'        If Not AreThemesActive Then
'            cmbType.Appearance = 0
'            FixFlatComboboxes Me, False
'        End If
        'NO ALIASING FOR NON-SERVER CONFIGURATIONS
        If Not m_Server Then
            RemoveAliasing
        Else
            'If we "started" in non-server mode allow
            AddAliasing
        End If
        'For installation set-up....
        If InStr(1, Command(), "-S", vbTextCompare) > 0 Or InStr(1, Command(), "/S", vbTextCompare) > 0 Then
            Me.Show
        Else
            Me.Hide
        End If
    Else
        'Already running...see if we started with admin privledges
        If m_Admin Then
            'try and activate previous instance for our friend here
            prevhWnd = FindWindow(vbNullString, cWindowTitle)
            If prevhWnd Then
                If Not IsWindowVisible(prevhWnd) Then ShowWindow prevhWnd, SW_SHOW
                SetForegroundWindow prevhWnd
                BringWindowToTop prevhWnd
            End If
            If m_Server Then AddAliasing
        End If
        peCloseReason = rcNotSoleInstance
        Unload Me
    End If
    peCloseReason = rcValidRequest
End Sub

Private Sub RemoveAliasing()
On Local Error Resume Next
    If TabStrip1.Tabs.Count = 4 Then TabStrip1.Tabs.Remove 4
    Frame4.Visible = False
End Sub

Private Sub AddAliasing()
Dim tb As MSComctlLib.Tab
On Local Error Resume Next
    If TabStrip1.Tabs.Count = 3 Then
        Set tb = TabStrip1.Tabs.Add(4, "ts04", "Aliasing")
        tb.ToolTipText = "Configure how various endpoints will be described"
        tb.Tag = tb.Caption
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Local Error Resume Next
    Cancel = (UnloadMode = vbFormControlMenu)
    If Not propPage Is Nothing Then
        If propPage.Visible Then propPage.Hide
    End If
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim bRtn As Boolean
    SysTrayDest
    CleanupDBResources
    TerminateApplication peCloseReason
    If Not propPage Is Nothing Then
        Unload propPage
        Set propPage = Nothing
    End If
    StopServiceIfApplies 'they may have started after tool was launched, flag and restart.
    'everything is off compact now or never!
    MaintainDb True
    If peCloseReason <> rcNotSoleInstance Then
        'Already launched Admin shortcut - don't teardown!
        KillApp "HTE_ClientUtilities.exe" 'this is to ensure that one UTILITIES is running
        Sleep 500
    End If
    If IsRunningService Then
        Do Until StartNTService = 0
            Sleep 500
        Loop
    End If
    EndApp
End Sub

Private Sub SysTrayInit()
Set sysTray = New frmSysTray
    With sysTray
        .Initialize Me.hWnd
        'Initialize with last known status of the host
        SetStatus (curHostStatus) 'SetStatus (GPS_HOST_UNINITIALIZED)
    End With
End Sub

Private Sub SysTrayDest()
    If Not sysTray Is Nothing Then Unload sysTray
    Set sysTray = Nothing
End Sub

Private Sub SetStatus(ByVal statusCode As HTE_GPS.GPS_HOST_STATUS)
    If Not sysTray Is Nothing Then
        sysTray.IconHandle = sysTray.ILStatusList.ListImages(statusCode + 1).ExtractIcon.Handle
        sysTray.ToolTip = Choose(statusCode + 1, "GPS Uninitialized", "GPS Catastrophic", "GPS Warning", "GPS Active")
   End If
   curHostStatus = statusCode
End Sub

Private Sub lstChain_Click()
    getInstanceDescriptions
    EnableActions
End Sub

Private Sub getInstanceDescriptions()
Dim sInst As String
Dim oProcess As cProcess
    lblProgID.Caption = vbNullString
    lblInstance.Caption = vbNullString
    If lstRoutes.ListIndex <> -1 And lstChain.ListIndex <> -1 Then
        sInst = getInstanceID(lstChain.ItemData(lstChain.ListIndex))
        Set oProcess = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).Item(sInst)
        If Not oProcess Is Nothing Then
            lblProgID.Caption = oProcess.progID
            lblInstance.Caption = sInst
        End If
    End If
End Sub

Private Function getInstanceValue(ByVal inst As String) As Long
    getInstanceValue = Mid$(inst, 2)
End Function

Private Function getInstanceID(ByVal inst As Long) As String
    getInstanceID = "R" & LeftPad(inst, 4)
End Function

Private Function LeftPad(ByVal strSource As String, ByVal intPadTo As Integer, Optional strPadChar As String = "0")
    'pre-pend the string with the specifid number of specified characters, then return the rightmost specified number of characters
    LeftPad = Right$(String$(intPadTo, strPadChar) & strSource, intPadTo)
End Function

Private Function getCurrentProcessImage(Optional ByRef psStatus As String, Optional ByRef sInstance As String) As Long
Dim rStatus As HTE_GPS.GPS_PROCESSOR_STATUS
        sInstance = getInstanceID(lstChain.ItemData(lstChain.ListIndex)) 'appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).Item(getInstanceID(lstChain.ItemData(lstChain.ListIndex)))  'Split(m_InstanceArray(m_ChainIdx), cSep)(cChnInstID)
        If Not gApplication Is Nothing Then
            rStatus = gApplication.ProcessStatus(sInstance)
            psStatus = processorStatusDesc(rStatus)
            Select Case rStatus
                Case GPS_STAT_READYANDWILLING
                    getCurrentProcessImage = 4
                Case GPS_STAT_WARNING
                    getCurrentProcessImage = 3
                Case GPS_STAT_ERROR, GPS_STAT_BAD_INTERFACE, GPS_STAT_HOST_UNSUPPORTED
                    getCurrentProcessImage = 2
                Case Else
                    getCurrentProcessImage = 1
            End Select
        End If
End Function

Private Sub lstChain_DblClick()
Dim oProcess As cProcess, availProcess As cProcess
Dim oRoute As cRoute, bReload As Boolean
Dim i As Long, sInst As String, sStatusDesc As String, sProgID As String
Dim oObj As HTE_GPS.Process, oProp As HTE_GPS.PropertyPage

On Error GoTo err_lstChainDblClick
    Screen.MousePointer = vbHourglass
'''    'Added this line to circumvent disaster...if the config OCX opens a form, we as the Interface can't unload it
'''    'so disable the exit menu, if not already visible no big deal...if so then close the PropPage before exiting
    mnuAExit.Enabled = False
    If lstRoutes.ListIndex <> -1 And lstChain.ListIndex <> -1 Then
        Set propPage = New frmProperties
        Load propPage
        With propPage
            Set oRoute = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex))
            Set oProcess = oRoute.Item(getInstanceID(lstChain.ItemData(lstChain.ListIndex)))
            .lblProcess.Caption = oProcess.progID
            i = getCurrentProcessImage(sStatusDesc, sInst)
            .lblStatusDesc.Caption = sStatusDesc
            .lblInstance.Caption = sInst
            .imgStatus.Picture = .ILStatusList.ListImages(i).Picture
            For Each availProcess In availProcesses
                If InStr(1, availProcess.progID, .lblProcess.Caption, vbTextCompare) > 0 Then
                    propPage.Caption = availProcess.friendlyName
                    propPage.frProperties.Caption = Space$(3) & availProcess.friendlyName
                    propPage.txtFriendlyDesc = lstChain.List(lstChain.ListIndex)
                    Set oObj = CreateObject(availProcess.progID)
                    If Not gApplication.PropertyPage(sInst) Is Nothing Then
                        If StrComp(gApplication.PropertyPage(sInst).Name, oObj.PropertyPage.Name, vbTextCompare) = 0 Then
                            Set propPage.propPage = gApplication.PropertyPage(sInst)
                            sProgID = Replace(availProcess.progID, ".PROCESS", vbNullString, , , vbTextCompare)
                            bReload = True
                        Else
                            Set oProp = oObj.PropertyPage
                            If Not oProp Is Nothing Then
                                oProp.PropertyCallback = Me 'override callback to self- since not in scheme
                                sProgID = Replace(oProcess.progID, ".PROCESS", vbNullString, , , vbTextCompare)
                                oProp.Settings = appSettings.ProcessSettings(sProgID, sInst)
                                Set propPage.propPage = oProp
                            End If
                        End If
                    Else 'Just added create from new
                        Set oProp = oObj.PropertyPage
                        If Not oProp Is Nothing Then
                            oProp.PropertyCallback = Me 'override callback to self- since not in scheme
                            sProgID = Replace(oProcess.progID, ".PROCESS", vbNullString, , , vbTextCompare)
                            appSettings.LetProcessSettings sProgID, sInst, vbNullString
                            oProp.Settings = appSettings.ProcessSettings(sProgID, sInst)
                            Set propPage.propPage = oProp
                        End If
                    End If
                End If
            Next
        End With
ShowForm:
        Screen.MousePointer = vbDefault
'''        appSettings.flagChanges
        propPage.Show vbModal, Me
        appSettings.Reload
'        If bReload Then
'            'appSettings.LetProcessSettings sProgID, sInst, gApplication.PropertyPage(sInst).Settings
'            appSettings.Reload
'        Else
'            appSettings.LetProcessSettings sProgID, sInst, oProp.Settings
'        End If
    End If
Exit_lstChainDblClick:
    mnuAExit.Enabled = True
    Set propPage = Nothing
    Set oProp = Nothing
    Set oObj = Nothing
    Exit Sub
err_lstChainDblClick:
    UEH_LogError "frmMain", "lstChainDblClick", Err
    GoTo ShowForm
End Sub

Private Sub lstChain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstChain.OLEDrag
End Sub

Private Sub lstChain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strList As String

    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strList = Data.GetData(vbCFText)
    If strList = lstProcesses.Text Then
        cmdAdd_Click
    End If
End Sub

Private Sub lstChain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim strList As String

    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strList = Data.GetData(vbCFText)
    If strList = lstProcesses.Text Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub lstChain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    If cmdRemove.Enabled Then
        AllowedEffects = vbDropEffectMove Or vbDropEffectNone
        Data.SetData lstChain
    End If
End Sub

Private Sub lstProcesses_Click()
    getProcessDescription
    EnableActions
End Sub

Private Sub getProcessDescription()
    If lstProcesses.ListIndex <> -1 Then
        lblWhatAmI.Caption = availProcesses.Item(lstProcesses.List(lstProcesses.ListIndex)).friendlyName
    End If
End Sub

Private Sub lstProcesses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstProcesses.OLEDrag
End Sub

Private Sub lstProcesses_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strList As String
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strList = Data.GetData(vbCFText)
    If strList = lstChain.Text Then
        cmdRemove_Click
    End If
End Sub

Private Sub lstProcesses_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Dim strList As String

    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strList = Data.GetData(vbCFText)
    If strList = lstChain.Text Then
        Effect = vbDropEffectMove
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub lstProcesses_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    If cmdAdd.Enabled Then
        AllowedEffects = vbDropEffectCopy Or vbDropEffectNone
        Data.SetData lstProcesses
    End If
End Sub

Private Sub lstRoutes_Click()
Dim oRoute As cRoute
Dim i As Long
    If lstRoutes.ListIndex <> -1 Then
        Set oRoute = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex))
        chkServer.Value = Abs(oRoute.ServerProcess)
        For i = 0 To cmbType.ListCount
            If oRoute.InboundType = cmbType.ItemData(i) Then
                cmbType.ListIndex = i
                Exit For
            End If
        Next
        LoadProcessChain oRoute
    End If
    EnableActions
End Sub

Private Sub LoadProcessChain(ByRef activeRoute As cRoute)
Dim oProcess As cProcess
    lstChain.Clear
    For Each oProcess In activeRoute
        lstChain.AddItem oProcess.friendlyName
        lstChain.ItemData(lstChain.NewIndex) = CLng(Mid$(oProcess.InstanceID, 2))
    Next
    'need to re-initialize list to reflect where we are!!!!
    If lstChain.ListCount > 0 Then
        lstChain.ListIndex = 0
        lstChain_Click
    End If
    EnableActions
End Sub

Private Sub mnuAConfigure_Click()
On Local Error Resume Next
'Its altogether possible that a property page control would display an additional form hence the resume next statement
    If Not propPage Is Nothing Then
        If Not propPage.Visible Then
            Me.Show
        Else
            AppActivate Me.Caption
        End If
    Else
        Me.Show
    End If
End Sub

Private Sub mnuAExit_Click()
On Local Error Resume Next
    Unload Me
End Sub

Private Sub mnuANevermind_Click()
Dim Cancel As Integer
    Form_QueryUnload Cancel, vbFormControlMenu
End Sub

Private Sub mnuPProcessDown_Click()
    cmdDown_Click
End Sub

Private Sub mnuPProcessUp_Click()
    cmdUp_Click
End Sub

Private Sub mnuPProperties_Click()
    cmdProperties_Click
End Sub

Private Sub mnuPRemoveProcess_Click()
    cmdRemove_Click
End Sub

Private Sub mnuPSAddProcess_Click()
    cmdAdd_Click
End Sub

Private Sub mnuRAddRoute_Click()
    cmdAddRoute_Click
End Sub

Private Sub mnuRClose_Click()
    cmdClose_Click
End Sub

Private Sub mnuRExit_Click()
    If appSettings.IsDirty Then appSettings.commitChanges
    mnuAExit_Click
End Sub

Private Sub mnuRRemoveRoute_Click()
    cmdRemoveRoute_Click
End Sub

Private Sub optDirection_Click(index As Integer)
    cmbSortList_Click
End Sub

Private Sub propPage_DescriptionChange(ByVal NewDesc As String)
Dim oProcess As cProcess
Dim sInst As String
    Screen.MousePointer = vbHourglass
    If lstChain.ListIndex <> -1 And lstRoutes.ListIndex <> -1 Then
        sInst = getInstanceID(lstChain.ItemData(lstChain.ListIndex))
        Set oProcess = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).Item(sInst)
        oProcess.friendlyName = NewDesc
        lstChain.List(lstChain.ListIndex) = NewDesc
        If TabStrip1.SelectedItem.index = 3 Then 'Detail view
            tcDetail.Tabs.Item(lstChain.ListIndex + 1).Caption = NewDesc
        End If
'''        appSettings.flagChanges
        appSettings.commitChanges
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub propPage_SaveChanges(ByVal XMLDOMNode As String)
    PropPageCallback_SaveChanges XMLDOMNode
End Sub

Private Sub PropPageCallback_Exit()
   
End Sub

Private Function PropPageCallback_SaveChanges(ByVal XMLDOMNode As String) As Boolean

Dim oProcess As cProcess
Dim sInst As String
    If lstChain.ListIndex <> -1 And lstRoutes.ListIndex <> -1 Then
        sInst = getInstanceID(lstChain.ItemData(lstChain.ListIndex))
        Set oProcess = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).Item(sInst)
        If Not oProcess Is Nothing Then
            PropPageCallback_SaveChanges = appSettings.LetProcessSettings(oProcess.progID, oProcess.InstanceID, XMLDOMNode)
        Else
            UEH_Log cModuleName, "PropPageCallback_SaveChanges", "Process returned is inactive!", logWarning
        End If
    End If
End Function

Private Sub sysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    If eButton = vbLeftButton Then
        If m_Admin Then
            mnuAExit.Visible = m_Exit
            mnuASep.Visible = m_Exit
            mnuAConfigure_Click
        End If
    End If
End Sub

Private Sub sysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If m_Admin Then
        If eButton = vbRightButton Then
            mnuAExit.Visible = m_Exit
            mnuASep.Visible = m_Exit
            PopupMenu mnuAdmin, , , , mnuANevermind
        End If
    End If
End Sub

Private Sub TabStrip1_Click()
Dim sFrame As String
    ' Show the frame for the tab we are displaying.
    sFrame = TabStrip1.Tag
    If sFrame = "Frame4" Then
        termAliasing
        CleanupDBResources
    End If
    If sFrame <> vbNullString And "Frame" & CStr(TabStrip1.SelectedItem.index) <> TabStrip1.Tag Then
        Controls(sFrame).Visible = False
        Controls(sFrame).Enabled = False
    End If
    sFrame = "Frame" & CStr(TabStrip1.SelectedItem.index)
On Local Error Resume Next
    If Not gApplication Is Nothing Then gApplication.ShowLastMessage = (TabStrip1.SelectedItem.index <> 1)
    InitializeFrame TabStrip1.SelectedItem.index
    Controls(sFrame).Visible = True
    Controls(sFrame).Enabled = True
    TabStrip1.Tag = sFrame
    EnableActions
End Sub
Private Sub VerifyAliasChanges()
    If cmdSave.Enabled Then
        ' get confirmation from user
'        If MsgBox("There are edits pending. " & vbCrLf & _
'                    "Do you want to commit them?", _
'                    vbYesNo Or vbQuestion, "Commit Changes?") = vbYes Then
'            cmdSave_Click
'        Else
'            cmdUndo_Click
'            cmdSave.Enabled = False
'            cmdUndo.Enabled = False
'            cmbSortList.Enabled = True
'            optDirection(0).Enabled = True
'            optDirection(1).Enabled = True
'            cmdDelete.Enabled = (VSFlexGrid1.Rows > 1) And True
'            cmdNew.Enabled = True
'        End If
    End If
End Sub
Private Sub CleanupDBResources()
'Dim rs As ADODB.Recordset
'Dim ds As HTE_GPSData.DataSource

    VerifyAliasChanges
    'Set rs = VSFlexGrid1.DataSource
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
            Set rs = Nothing
            Set VSFlexGrid1.DataSource = Nothing
        End If
    End If
    Set ds = Nothing
    'we are the only one's with an "active" connection we should maintain
    MaintainDb True
'''    If m_ModifiedData Then
'''        If ds Is Nothing Then Set ds = New HTE_GPSData.DataSource
'''        m_ModifiedData = Not ds.MaintainDB
'''        'Debug.Assert bRtn
'''    End If
    
End Sub

Private Sub InitializeFrame(ByVal index As Long)

    txtProcessed.Text = vbNullString
    chkLock.Value = vbUnchecked
    chkLockDetail.Value = vbUnchecked
    lstRoutes_Click
    Select Case index
        Case 1
            initTypesCombo
            initAllowExit
            initProcesses
            initRoutes
            Picture1.backColor = Frame1.backColor
        Case 2
            Picture2.backColor = Frame2.backColor
        Case 3
            Picture3.backColor = Frame3.backColor
            initIndividualTabs
        Case 4
            Picture4.backColor = Frame4.backColor
            initAliasing
            
    End Select
End Sub
Private Sub initAliasing()
Dim i As Long
    If ds Is Nothing Then Set ds = New HTE_GPSData.DataSource
    Set rs = New ADODB.Recordset
    With cmbSortList
        .Clear
        For i = 2 To VSFlexGrid1.Cols - 1 'first col is blank, second is unique key
            .AddItem VSFlexGrid1.Cell(flexcpTextDisplay, 0, i)
            .ItemData(.NewIndex) = i - 1
        Next
        .ListIndex = 3
    End With
End Sub
Private Sub termAliasing()
    
End Sub
Private Sub initIndividualTabs()
Dim i As Long
Dim imgID As Long
Dim sToolTip As String
Dim tabX As cTab
Dim oProcess As cProcess
Dim oRoute As cRoute
    With tcDetail
        .Tabs.Clear
        For i = txtDetail.Count - 1 To 1 Step -1
            Unload txtDetail(i)
        Next
        If lstRoutes.ListIndex <> -1 Then
            i = 0
            Set oRoute = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex))
            For Each oProcess In oRoute
                imgID = FindDetailStatus(oProcess.InstanceID, sToolTip)
                Set tabX = .Tabs.Add(oProcess.InstanceID, , oProcess.friendlyName, imgID - 1)
                If i > 0 Then Load txtDetail(i)
                Set tabX.Panel = txtDetail(i)
                txtDetail(i).Text = vbNullString
                txtDetail(i).Tag = tabX.Key
                i = i + 1
            Next
        End If
    End With
End Sub

Private Sub FindAllDetailStatus()
Dim tabX As cTab
Dim i As Long, imgID As Long
Dim sTool As String
    If TabStrip1.SelectedItem.index = 3 Then 'Detail View
        With tcDetail.Tabs
        For i = 1 To .Count
            Set tabX = .Item(i)
            imgID = FindDetailStatus(tabX.Key, sTool)
            tabX.IconIndex = imgID - 1
            tabX.ToolTipText = sTool
        Next
        End With
    End If
End Sub

Private Function FindDetailStatus(ByVal sInstance As String, Optional ByRef sToolTip As String) As Long
Dim rStatus As HTE_GPS.GPS_PROCESSOR_STATUS
Dim imgIndex As Long
    If TabStrip1.SelectedItem.index = 3 Then 'Detail View
        If Not gApplication Is Nothing Then
            rStatus = gApplication.ProcessStatus(sInstance)
            sToolTip = processorStatusDesc(rStatus)
                Select Case rStatus
                    Case GPS_STAT_READYANDWILLING
                        imgIndex = 4
                    Case GPS_STAT_WARNING
                        imgIndex = 3
                    Case GPS_STAT_ERROR, GPS_STAT_BAD_INTERFACE, GPS_STAT_HOST_UNSUPPORTED
                        imgIndex = 2
                    Case Else
                        imgIndex = 1
                End Select
                FindDetailStatus = imgIndex
        End If
    End If
End Function

Private Sub initTypesCombo()
Dim iNode As MSXML2.IXMLDOMNode
Dim i As Long
    Set iNode = appSettings.typesNode
    cmbType.Clear
    For i = 0 To iNode.childNodes.Length - 1
        cmbType.AddItem iNode.childNodes(i).nodeName
        cmbType.ItemData(cmbType.NewIndex) = iNode.childNodes(i).nodeTypedValue
    Next
End Sub

Private Sub initAllowExit()
    chkAllowExit.Value = Abs(appSettings.AllowExit)
End Sub

Private Sub initRoutes()
Dim oRoute As cRoute
Dim oProcess As cProcess
Dim lastIndex As Long
    lastIndex = lstRoutes.ListIndex
    lstRoutes.Clear
    For Each oRoute In appSettings
        lstRoutes.AddItem oRoute.RouteID
    Next
    lstRoutes.ListIndex = lastIndex
    If lstRoutes.ListIndex <> -1 Then lstRoutes_Click
End Sub

Private Sub initProcesses()
Dim oProcess As cProcess
Dim lastIndex As Long
    lastIndex = lstProcesses.ListIndex
    lstProcesses.Clear
    For Each oProcess In availProcesses
        lstProcesses.AddItem oProcess.progID
    Next
    lstProcesses.ListIndex = lastIndex
End Sub

Private Sub EnableActions()
    mnuRExit.Visible = m_Exit
    chkServer.Enabled = lstRoutes.ListCount > 0: 'If Not chkServer.Enabled Then chkServer.Value = vbUnchecked
    cmbType.Enabled = lstRoutes.ListCount > 0
    cmdRemoveRoute.Enabled = lstRoutes.ListCount > 0: mnuRRemoveRoute.Enabled = cmdRemoveRoute.Enabled
    cmdAdd.Enabled = lstProcesses.ListIndex <> -1 And lstRoutes.ListIndex <> -1: mnuPSAddProcess.Enabled = cmdAdd.Enabled
    cmdRemove.Enabled = lstChain.ListIndex <> -1 And lstRoutes.ListIndex <> -1: mnuPRemoveProcess.Enabled = cmdRemove.Enabled
    cmdUp.Enabled = lstChain.ListIndex > 0 And lstRoutes.ListIndex <> -1: mnuPProcessUp.Enabled = cmdUp.Enabled
    cmdDown.Enabled = lstChain.ListIndex <> lstChain.ListCount - 1 And lstRoutes.ListIndex <> -1: mnuPProcessDown.Enabled = cmdDown.Enabled
    cmdProperties.Enabled = lstChain.ListIndex <> -1 And lstRoutes.ListIndex <> -1: mnuPProperties.Enabled = cmdProperties.Enabled
End Sub

Private Function RemoveProcess(Optional ByVal bPerm As Boolean = True)
Dim oProcess As cProcess
    Screen.MousePointer = vbHourglass
    Set oProcess = appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).Item(getInstanceID(lstChain.ItemData(lstChain.ListIndex)))
    appSettings.Item(lstRoutes.List(lstRoutes.ListIndex)).Remove oProcess.InstanceID
    appSettings.DeleteProcessSettings oProcess.progID, oProcess.InstanceID
    lstChain.RemoveItem lstChain.ListIndex
    If lstChain.ListCount > 0 Then lstChain.ListIndex = 0
    Screen.MousePointer = vbDefault
End Function

Private Sub tcDetail_TabDoubleClick(theTab As HTE_TabView6.cTab)
    If theTab.index <= lstChain.ListCount Then
        lstChain.ListIndex = theTab.index - 1
        lstChain_DblClick
    End If
End Sub

Private Sub VSFlexGrid1_AfterDataRefresh()
    Debug.Print "VSFlexGrid1_AfterDataRefresh"
    With VSFlexGrid1
        .AutoSize 1, .Cols - 1
        ' use another imagelist to show db-like cursor glyph
        .ColWidth(0) = .RowHeight(0)
        .ColImageList(0) = imgListGlyph.hImageList
        .ColAlignment(0) = flexAlignCenterCenter
        .FrozenCols = 1
    End With
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Debug.Print "VSFlexGrid1_AfterEdit"
    NotifyOfAliasChanges
'''    cmdSave.Enabled = True
'''    cmdUndo.Enabled = True
'''    cmbSortList.Enabled = False
'''    optDirection(0).Enabled = False
'''    optDirection(1).Enabled = False
'''    cmdDelete.Enabled = False 'VSFlexGrid1.Rows > 1
'''    cmdNew.Enabled = False
End Sub

Private Sub VSFlexGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Debug.Print "VSFlexGrid1_AfterRowColChange"
    With VSFlexGrid1
        .TextMatrix(OldRow, 0) = ""
        .TextMatrix(NewRow, 0) = 0
    End With
End Sub

Private Sub VSFlexGrid1_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Debug.Print "VSFlexGrid1_AfterSelChange"
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Debug.Print "VSFlexGrid1_BeforeEdit"
End Sub

Private Sub VSFlexGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Debug.Print "VSFlexGrid1_BeforeMouseDown"
End Sub

Private Sub VSFlexGrid1_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Debug.Print "VSFlexGrid1_BeforeMoveColumn"
End Sub

Private Sub VSFlexGrid1_BeforeMoveRow(ByVal Row As Long, Position As Long)
    Debug.Print "VSFlexGrid1_BeforeMoveRow"
End Sub

Private Sub VSFlexGrid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Debug.Print "VSFlexGrid1_BeforeRowColChange"
    With VSFlexGrid1
        .TextMatrix(OldRow, 0) = ""
        .TextMatrix(NewRow, 0) = 0
    End With
    If cmdSave.Enabled Then
        If OldRow <> NewRow Then
            VerifyAliasChanges
        End If
    End If
'    ' update DB-like cursor glyph
'    With VSFlexGrid1
'        .TextMatrix(OldRow, 0) = ""
'        .TextMatrix(NewRow, 0) = 0
'    End With
End Sub

Private Sub VSFlexGrid1_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Debug.Print "VSFlexGrid1_BeforeRowColChange"

End Sub

Private Sub VSFlexGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Debug.Print "VSFlexGrid1_CellButtonClick"
End Sub

Private Sub VSFlexGrid1_ChangeEdit()
    Debug.Print "VSFlexGrid1_ChangeEdit"
End Sub

Private Sub VSFlexGrid1_Click()
    Debug.Print "VSFlexGrid1_Click"
End Sub

Private Sub VSFlexGrid1_EnterCell()
    Debug.Print "VSFlexGrid1_EnterCell"
End Sub

Private Sub VSFlexGrid1_LeaveCell()
    Debug.Print "VSFlexGrid1_LeaveCell"
    
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With VSFlexGrid1
        If Button = vbRightButton Then
            If .MouseRow > -1 Then
                .Select .MouseRow, .MouseCol
                .EditCell
            Else
                If cmdNew.Enabled Then cmdNew_Click
            End If
        End If
    End With
End Sub

Private Sub VSFlexGrid1_RowColChange()
    With VSFlexGrid1
        If .RowSel <> 0 Then
            .TextMatrix(.RowSel, 0) = 0
        End If
    End With
End Sub

Private Sub VSFlexGrid1_SelChange()
    Debug.Print "VSFlexGrid1_SelChange"
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Debug.Print "VSFlexGrid1_StartEdit"
End Sub

Private Sub VSFlexGrid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim strField As String, strValue As String
Dim validExp As VBScript_RegExp_55.RegExp
Dim mc As VBScript_RegExp_55.MatchCollection
Dim duhText As String
Const IP_OCTET As String = "(2[5][012345])|(2[01234]\d)|(1\d\d)|(\d\d)|(\d)"
    Debug.Print "VSFlexGrid1_ValidateEdit"
    Set validExp = New VBScript_RegExp_55.RegExp
    With validExp
        .MultiLine = False
        .IgnoreCase = True
        .Global = True
        strField = LCase$(VSFlexGrid1.ColKey(Col))
        strValue = Trim$(VSFlexGrid1.EditText) 'VSFlexGrid1.Cell(flexcpVariantValue, Row, Col))
        VSFlexGrid1.EditText = strValue
        Select Case strField
            
            Case "aliasid"
                Cancel = True 'don't modify auto-generated key
            
            Case "physicallookup"
                duhText = "Invalid format - use 0-9, A-F in the following format xx-xx-xx-xx-xx-xx"
                'regex - mac address - no duplicates, zero-len ok
                If Len(strValue) > 0 Then
                    .Pattern = "^([0-9a-fA-F][0-9a-fA-F][-]){5}([0-9a-fA-F][0-9a-fA-F])$"
                    Set mc = .Execute(strValue)
                    If mc.Count = 1 Then
                        Cancel = Not (Len(strValue) = mc.Item(0).Length)
                        'TODO: check for unique
                        If Not Cancel Then
                            Cancel = Not IsUniqueValue(Col, Row, strValue)
                            If Cancel Then duhText = "The " & VSFlexGrid1.Cell(flexcpText, 0, Col) & " value [" & strValue & "] already exists and should be unique."
                        End If
                    Else
                        Cancel = True
                    End If
                End If
                
            Case "addresslookup"
                duhText = "Invalid format - use a set of 4 octets from 0-255 seperated by a ."
                If Len(strValue) > 0 Then
                    .Pattern = "((" & IP_OCTET & ")\.){3}(" & IP_OCTET & ")"
                    Set mc = .Execute(strValue)
                    If mc.Count = 1 Then
                        Cancel = Not (Len(strValue) = mc.Item(0).Length)
                        'TODO: check for unique
                        If Not Cancel Then
                            Cancel = Not IsUniqueValue(Col, Row, strValue)
                            If Cancel Then duhText = "The " & VSFlexGrid1.Cell(flexcpText, 0, Col) & " value [" & strValue & "] already exists and should be unique."
                        End If
                    Else
                        Cancel = True
                    End If
                End If
                'regex - ip address - no duplicates
            
            Case "applookup", "device"
                duhText = "Maximum length permitted for " & strField & " is 16."
                'max-len(16) lookup - duplicates ok
                Cancel = Len(strValue) > 16
                'TODO: check for unique
                If Not Cancel Then
                    Cancel = Not IsUniqueValue(Col, Row, strValue)
                    If Cancel Then duhText = "The " & VSFlexGrid1.Cell(flexcpText, 0, Col) & " value [" & strValue & "] already exists and should be unique."
                End If
            
            Case "alias"
                'max-len(32) lookup - duplicates ok
                duhText = "Maximum length permitted for " & strField & " is 32."
                Cancel = Len(strValue) > 32
            
            Case "comments"
                'max length (250)
                duhText = "Maximum length permitted for " & strField & " is 250."
                Cancel = Len(strValue) > 250
            Case Else
        End Select
        If Cancel Then
            VSFlexGrid1.ToolTipText = duhText
        Else
            VSFlexGrid1.ToolTipText = vbNullString
            If StrComp(strValue, VSFlexGrid1.Cell(flexcpText, Row, Col), vbTextCompare) <> 0 Then NotifyOfAliasChanges
        End If
        
    End With
    Set validExp = Nothing
End Sub

Private Function IsUniqueValue(ByVal Col As Long, Row As Long, ValueToCheck As String) As Boolean
Dim i As Long
Dim bRtn As Boolean
    'Debug.Print ValueToCheck
    IsUniqueValue = True
    If Len(ValueToCheck) > 0 Then
        For i = 0 To VSFlexGrid1.Rows - 1
            If i <> Row Then
                'Debug.Print "Is " & VSFlexGrid1.Cell(flexcpText, i, Col) & " = " & ValueToCheck
                If (lstrcmpi(VSFlexGrid1.Cell(flexcpText, i, Col), ValueToCheck) = 0) Then
                    IsUniqueValue = False
                    Exit For
                End If
            End If
        Next
    End If
End Function

Private Function MaintainDb(Optional OnlyIfChanged As Boolean = True)
On Local Error Resume Next
    If OnlyIfChanged Then
        If m_ModifiedData Then
            If ds Is Nothing Then Set ds = New HTE_GPSData.DataSource
            m_ModifiedData = Not ds.MaintainDb
        End If
    Else
        If ds Is Nothing Then Set ds = New HTE_GPSData.DataSource
        m_ModifiedData = Not ds.MaintainDb
        Debug.Print "Compact/repair occurred:" & CStr(Not m_ModifiedData)
    End If
End Function
