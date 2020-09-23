VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00D05C28&
   Caption         =   "Multiple Flexgrid Example"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   7185
   Begin VB.Frame fraBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6465
      Begin VB.PictureBox pctBack 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6420
         TabIndex        =   19
         Top             =   5880
         Width           =   6420
         Begin VB.Line lnBack 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            Visible         =   0   'False
            X1              =   0
            X2              =   6465
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Image imgAdd 
            Height          =   285
            Index           =   4
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblCInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remote"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   20
            Top             =   90
            Width           =   675
         End
         Begin VB.Image imgAdd_HI 
            Height          =   285
            Index           =   4
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide_HI 
            Height          =   285
            Index           =   4
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide 
            Height          =   285
            Index           =   4
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgPicTitle 
            Height          =   240
            Index           =   4
            Left            =   240
            Picture         =   "Form1.frx":038A
            Top             =   67
            Width           =   240
         End
      End
      Begin VB.PictureBox pctBack 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6420
         TabIndex        =   17
         Top             =   6360
         Width           =   6420
         Begin VB.Line lnBack 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            Visible         =   0   'False
            X1              =   0
            X2              =   6465
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Image imgAdd 
            Height          =   285
            Index           =   5
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblCInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calibration"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   840
            TabIndex        =   18
            Top             =   90
            Width           =   915
         End
         Begin VB.Image imgAdd_HI 
            Height          =   285
            Index           =   5
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide_HI 
            Height          =   285
            Index           =   5
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide 
            Height          =   285
            Index           =   5
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgPicTitle 
            Height          =   240
            Index           =   5
            Left            =   240
            Picture         =   "Form1.frx":0714
            Top             =   67
            Width           =   240
         End
      End
      Begin VB.PictureBox pctBack 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6420
         TabIndex        =   15
         Top             =   5400
         Width           =   6420
         Begin VB.Image imgHide 
            Height          =   285
            Index           =   3
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide_HI 
            Height          =   285
            Index           =   3
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgAdd_HI 
            Height          =   285
            Index           =   3
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblCInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   16
            Top             =   90
            Width           =   465
         End
         Begin VB.Image imgAdd 
            Height          =   285
            Index           =   3
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Line lnBack 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            Visible         =   0   'False
            X1              =   0
            X2              =   6465
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Image imgPicTitle 
            Height          =   240
            Index           =   3
            Left            =   240
            Picture         =   "Form1.frx":0A9E
            Top             =   67
            Width           =   240
         End
      End
      Begin VB.PictureBox pctBack 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6420
         TabIndex        =   13
         Top             =   0
         Width           =   6420
         Begin VB.Image imgAdd 
            Height          =   285
            Index           =   0
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblCInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   14
            Top             =   90
            Width           =   555
         End
         Begin VB.Image imgAdd_HI 
            Height          =   285
            Index           =   0
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide_HI 
            Height          =   285
            Index           =   0
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide 
            Height          =   285
            Index           =   0
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Line lnBack 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   6465
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Image imgPicTitle 
            Height          =   240
            Index           =   0
            Left            =   240
            Picture         =   "Form1.frx":0E28
            Top             =   67
            Width           =   240
         End
      End
      Begin VB.PictureBox pctBack 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6420
         TabIndex        =   11
         Top             =   2295
         Width           =   6420
         Begin VB.Image imgHide 
            Height          =   285
            Index           =   1
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide_HI 
            Height          =   285
            Index           =   1
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgAdd_HI 
            Height          =   285
            Index           =   1
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblCInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   12
            Top             =   90
            Width           =   645
         End
         Begin VB.Image imgAdd 
            Height          =   285
            Index           =   1
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Line lnBack 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   6465
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Image imgPicTitle 
            Height          =   240
            Index           =   1
            Left            =   240
            Picture         =   "Form1.frx":11B2
            Top             =   67
            Width           =   240
         End
      End
      Begin VB.PictureBox pctBack 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6420
         TabIndex        =   9
         Top             =   4950
         Width           =   6420
         Begin VB.Line lnBack 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            Visible         =   0   'False
            X1              =   0
            X2              =   6465
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Image imgAdd 
            Height          =   285
            Index           =   2
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblCInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Battery"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   10
            Top             =   90
            Width           =   645
         End
         Begin VB.Image imgAdd_HI 
            Height          =   285
            Index           =   2
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide_HI 
            Height          =   285
            Index           =   2
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgHide 
            Height          =   285
            Index           =   2
            Left            =   6060
            Tag             =   "0"
            Top             =   45
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Image imgPicTitle 
            Height          =   240
            Index           =   2
            Left            =   240
            Picture         =   "Form1.frx":153C
            Top             =   67
            Width           =   240
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
         Height          =   1935
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Tag             =   "0"
         Top             =   2640
         Visible         =   0   'False
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   100
         BackColor       =   16774384
         ForeColor       =   0
         BackColorFixed  =   15244408
         ForeColorFixed  =   16777215
         BackColorSel    =   16764603
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
         Height          =   3135
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Tag             =   "1"
         Top             =   365
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   100
         BackColor       =   16774384
         ForeColor       =   0
         BackColorFixed  =   15244408
         ForeColorFixed  =   16777215
         BackColorSel    =   16764603
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
         Height          =   6255
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Tag             =   "0"
         Top             =   5325
         Visible         =   0   'False
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   11033
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   100
         BackColor       =   16774384
         BackColorFixed  =   15244408
         ForeColorFixed  =   16777215
         BackColorSel    =   16764603
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
         Height          =   480
         Index           =   3
         Left            =   0
         TabIndex        =   24
         Tag             =   "0"
         Top             =   5775
         Visible         =   0   'False
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   100
         BackColor       =   16774384
         BackColorFixed  =   15244408
         ForeColorFixed  =   16777215
         BackColorSel    =   16764603
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
         Height          =   6990
         Index           =   5
         Left            =   0
         TabIndex        =   25
         Tag             =   "0"
         Top             =   6720
         Visible         =   0   'False
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   12330
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   100
         BackColor       =   16774384
         BackColorFixed  =   15244408
         ForeColorFixed  =   16777215
         BackColorSel    =   16764603
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
         Height          =   480
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Tag             =   "0"
         Top             =   6120
         Visible         =   0   'False
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   100
         BackColor       =   16774384
         BackColorFixed  =   15244408
         ForeColorFixed  =   16777215
         BackColorSel    =   16764603
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.VScrollBar vsbCInfo 
      CausesValidation=   0   'False
      Height          =   3640
      LargeChange     =   1800
      Left            =   6430
      SmallChange     =   300
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Value           =   15
      Width           =   255
   End
   Begin VB.PictureBox pctHide 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":18C6
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctHide_Press 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":1D84
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctHide_Hi 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":2242
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctAdd 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":2700
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctAdd_Press 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":2BBE
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctAdd_Hi 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":307C
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctHand 
      Height          =   345
      Left            =   6570
      Picture         =   "Form1.frx":353A
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const TotFlex = 5   ' Number of flexgrids - 1 (5 = Six Flexgrids)

' Set default fixed values
Sub InitSingleCol(Flex As MSFlexGrid, Value, Col, Row, Width)
    Flex.Row = Row
    Flex.Col = Col
    Flex.ColAlignment(Col) = flexAlignCenterCenter
    Flex.Text = Value
    Flex.ColWidth(Col) = Width
End Sub

' Initialize the Flexgrids
Private Sub InitializeGrid()
Dim i As Integer
    For i = 0 To TotFlex
        InitSingleCol fxgCInfo(i), "Value", 0, 0, 1000
        InitSingleCol fxgCInfo(i), "Range", 1, 0, 1000
        InitSingleCol fxgCInfo(i), "Description", 2, 0, 4267
        fxgCInfo(i).Rows = 1
    Next i
End Sub

Private Sub Form_Load()
Dim i As Integer
    On Error Resume Next
    
    ' Initialize flexgrids titles and column width
    InitializeGrid
    
    For i = 0 To TotFlex
        ' Load command pictures
        imgHide(i).Picture = pctHide.Picture
        imgAdd(i).Picture = pctAdd.Picture
        imgHide_HI(i).Picture = pctHide_Hi.Picture
        imgAdd_HI(i).Picture = pctAdd_Hi.Picture
        imgHide_HI(i).MouseIcon = pctHand.Picture
        imgHide_HI(i).MousePointer = 99
        imgAdd_HI(i).MouseIcon = pctHand.Picture
        imgAdd_HI(i).MousePointer = 99
        lblCInfo(i).MouseIcon = pctHand.Picture
        lblCInfo(i).MousePointer = 99
        If fxgCInfo(i).Tag = 0 Then
            imgHide(i).Visible = False
            imgAdd(i).Visible = True
            lnBack(i).Visible = True
        Else
            imgHide(i).Visible = True
            imgAdd(i).Visible = False
            lnBack(i).Visible = False
        End If
    Next i
    
    ' Set some fake value
    Populate
     
End Sub

' Add some fake values into the flexgrids
Private Sub Populate()
Dim i As Integer, j As Integer
    Randomize                               ' Initialize random numbers
    For j = 0 To TotFlex                    ' For all flexgrids
        For i = 0 To Int((10 * Rnd) + 1)    ' Add random numbers
            AddRowInFlex i, Int((50 * Rnd) + 1), "1..50", "Description " & i & " " & j, j
        Next i
    Next j
End Sub

Private Sub Form_Resize()
Dim i As Integer

    On Error Resume Next
    
    ' Change Height/Width of the controls according to the form size
    vsbCInfo.Height = Me.Height - 410
    vsbCInfo.Left = Me.Width - 380
    fraBack.Width = Me.Width - 345
    For i = 0 To 5
        fxgCInfo(i).Width = Me.Width - 390
        fxgCInfo(i).ColWidth(2) = fxgCInfo(i).Width - 2000
        pctBack(i).Width = Me.Width - 390
        lnBack(i).X2 = pctBack(i).Width
        imgHide(i).Left = pctBack(i).Width - 360
        imgHide_HI(i).Left = pctBack(i).Width - 360
        imgAdd(i).Left = pctBack(i).Width - 360
        imgAdd_HI(i).Left = pctBack(i).Width - 360
    Next i
    
    ArrangeFlexes
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Index
    ' Hide the hilighted images
    For Index = 0 To TotFlex
        ' If the Tag property is "0" then it means that the highlight picture
        ' is not visible, so there is no need to hide it again
        If imgAdd_HI(Index).Tag = 1 Then
            imgAdd_HI(Index).Tag = 0
            imgAdd_HI(Index).Visible = False
            lblCInfo(Index).ForeColor = &HFFFFFF
        End If
        If imgHide_HI(Index).Tag = 1 Then
            imgHide_HI(Index).Tag = 0
            imgHide_HI(Index).Visible = False
            lblCInfo(Index).ForeColor = &HFFFFFF
        End If
    Next Index
End Sub

Private Sub fxgCInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Call the same routine to hide the highlighted images
    Form_MouseMove 0, 0, 0, 0
End Sub

Private Sub imgAdd_HI_Click(Index As Integer)
    imgAdd_HI(Index).Tag = 0
    imgHide_HI(Index).Tag = 1
    imgAdd(Index).Visible = False
    imgAdd_HI(Index).Visible = False
    imgHide(Index).Visible = True
    imgHide_HI(Index).Visible = True
    imgHide_HI(Index).ZOrder
    fxgCInfo(Index).Visible = True      ' Shows the flexgrid
    fxgCInfo(Index).Tag = 1             ' Remember that the flexgrid is shown
    lnBack(Index).Visible = False       ' Hide the divisory line
    ArrangeFlexes                       ' Re-Arrange the flexgrids
End Sub

Private Sub imgAdd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When the mouse is over an image, I show the highlight one
    ' If the Tag property is "1" then it means that the highlight picture
    ' is already visible, so there is no need to show it again (so it does not flicker)
    If imgAdd_HI(Index).Tag = 0 Then
        imgAdd_HI(Index).Tag = 1
        imgAdd_HI(Index).Visible = True
        lblCInfo(Index).ForeColor = &HFFE3D9
        imgAdd_HI(Index).ZOrder
    End If
End Sub

Private Sub imgHide_HI_Click(Index As Integer)
    imgAdd_HI(Index).Tag = 1
    imgHide_HI(Index).Tag = 0
    imgHide(Index).Visible = False
    imgHide_HI(Index).Visible = False
    imgAdd(Index).Visible = True
    imgAdd_HI(Index).Visible = True
    imgAdd_HI(Index).ZOrder
    fxgCInfo(Index).Visible = False     ' Hides the flexgrid
    fxgCInfo(Index).Tag = 0             ' Remember that the flexgrid is hidden
    lnBack(Index).Visible = True        ' Show a divisor line
    ArrangeFlexes                       ' Re-Arrange the flexgrids
End Sub

Private Sub imgHide_HI_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When an image is pressed I show the proper image
    imgHide_HI(Index).Picture = pctHide_Press.Picture
End Sub

Private Sub imgHide_HI_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' The mouse button is up: remove the pressed image
    imgHide_HI(Index).Picture = pctHide_Hi.Picture
End Sub

Private Sub imgAdd_HI_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When an image is pressed I show the proper image
    imgAdd_HI(Index).Picture = pctAdd_Press.Picture
End Sub

Private Sub imgAdd_HI_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' The mouse button is up: remove the pressed image
    imgAdd_HI(Index).Picture = pctAdd_Hi.Picture
End Sub

Private Sub imgHide_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When the mouse is over an image, I show the highlight one
    ' If the Tag property is "1" then it means that the highlight picture
    ' is already visible, so there is no need to show it again (so it does not flicker)
    If imgHide_HI(Index).Tag = 0 Then
        imgHide_HI(Index).Tag = 1
        imgHide_HI(Index).Visible = True
        lblCInfo(Index).ForeColor = &HFFE3D9
        imgHide_HI(Index).ZOrder
    End If
End Sub

Private Sub lblCInfo_Click(Index As Integer)
    ' When the user clicks on the label it happens the same thing when he
    ' clicks on the hide/show images
    If fxgCInfo(Index).Tag = 1 Then
        imgHide_HI_Click Index
    Else
        imgAdd_HI_Click Index
    End If
End Sub

Private Sub lblCInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When the user moves the mouse on the label it happens the same thing when he
    ' moves the mouse on the hide/show images
    If fxgCInfo(Index).Tag = 1 Then
        imgHide_MouseMove Index, 0, 0, 0, 0
    Else
        imgAdd_MouseMove Index, 0, 0, 0, 0
    End If
End Sub

Private Sub pctBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Call the same routine to hide the highlighted images
    Form_MouseMove 0, 0, 0, 0
End Sub

' This function changes the flexgrids position when a flexgrid is explosed
' or reduced, or when the form size is changed
' It also set the scrollbar min. and max. values
Private Sub ArrangeFlexes()
Dim i As Integer
Dim j As Integer
Dim TotTop
    TotTop = 0
    For i = 0 To TotFlex
        pctBack(i).Top = TotTop
        fxgCInfo(i).Top = pctBack(i).Top + pctBack(i).Height
        TotTop = TotTop + pctBack(i).Height
        If fxgCInfo(i).Tag = 1 Then
            TotTop = TotTop + fxgCInfo(i).Height
        End If
    Next i
    fraBack.Height = TotTop
    TotTop = TotTop + 375
    ' If the height of the flexgrids is higher than the form height then
    ' the scrollbar will be enabled
    If TotTop > Me.Height Then
        vsbCInfo.Enabled = True
        vsbCInfo.Max = TotTop - Me.Height
        vsbCInfo.Value = 0
    Else
        vsbCInfo.Enabled = False
        fraBack.Top = 0
    End If
End Sub

Private Sub vsbCInfo_Change()
    ' All the flexgrids are placed on a frame, so to scroll them i just change
    ' the frame height
    fraBack.Top = -vsbCInfo.Value
End Sub

Public Sub AddRowInFlex(ByVal Par As Variant, ByVal Value As Variant, ByVal Range As String, ByVal Description As String, Index As Integer)
Dim CurrRow As Integer
       
    ' Set the proper height (so it does not automatically show the flexgrid srollbar)
    fxgCInfo(Index).Height = (240 * (fxgCInfo(Index).Rows + 1)) + 15
    
    ' Creates a new row
    fxgCInfo(Index).Rows = fxgCInfo(Index).Rows + 1
    
    ' Get the row where to put data
    CurrRow = fxgCInfo(Index).Rows - 1
    
    fxgCInfo(Index).TextMatrix(CurrRow, 0) = Value
    fxgCInfo(Index).TextMatrix(CurrRow, 1) = Range
    fxgCInfo(Index).TextMatrix(CurrRow, 2) = Description

End Sub


