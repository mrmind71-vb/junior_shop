VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form online_ship_grd1frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "»Ê«·’ «·‘Õ‰"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   1995
      ScaleWidth      =   20370
      TabIndex        =   17
      Top             =   8490
      Width           =   20370
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   16785
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   0
         Width           =   3435
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„»«⁄…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   225
            Width           =   735
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   225
            Width           =   690
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "„› ÊÕ…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   225
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   16785
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   585
         Width           =   3435
         Begin VB.OptionButton optShip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2655
            RightToLeft     =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton optShip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "» «—ÌŒ ‘Õ‰"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optShip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   11340
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   0
         Width           =   5415
         Begin VB.TextBox xdate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   900
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Tag             =   "filter"
            Top             =   225
            Width           =   1365
         End
         Begin VB.CheckBox chkDay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox xdate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2385
            TabIndex        =   37
            Tag             =   "COLOR"
            Top             =   270
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ «·ÿ·»Ì…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   15
            Left            =   4230
            TabIndex        =   36
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   11220
         Begin VB.TextBox xOnline_doc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   225
            Width           =   1635
         End
         Begin VB.TextBox xship_no 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   7830
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   225
            Width           =   2085
         End
         Begin VB.TextBox xdoc_no 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3150
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   225
            Width           =   3030
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            Height          =   240
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·»Ê·Ì’…"
            Height          =   240
            Index           =   0
            Left            =   9990
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   240
            Left            =   6300
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   7155
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1215
         Visible         =   0   'False
         Width           =   5415
         Begin VB.OptionButton optMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰ "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   270
            Width           =   735
         End
         Begin VB.OptionButton optMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "» «—ÌŒ «Ì„Ì·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   4455
            RightToLeft     =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   645
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   11340
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   540
         Width           =   5415
         Begin VB.TextBox xdate_ship1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox chkDayShip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox xdate_ship2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   945
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Tag             =   "filter"
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ «·‘Õ‰"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   4275
            TabIndex        =   31
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2340
            TabIndex        =   30
            Tag             =   "COLOR"
            Top             =   270
            Width           =   420
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1215
         Visible         =   0   'False
         Width           =   5370
         Begin VB.TextBox xdate_mail1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Tag             =   "filter"
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox chkdayMail 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·ÌÊ„"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox xdate_mail2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   945
            MaxLength       =   300
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Tag             =   "filter"
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   " «—ÌŒ «Ì„Ì·"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   4275
            TabIndex        =   43
            Tag             =   "COLOR"
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Õ Ì"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   2430
            TabIndex        =   42
            Tag             =   "COLOR"
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   585
         Width           =   6900
         Begin VB.TextBox xPhone 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   225
            Width           =   2445
         End
         Begin VB.TextBox xName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3555
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   225
            Width           =   2670
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«· ·Ì›Ê‰"
            Height          =   240
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·≈”„ "
            Height          =   240
            Left            =   6255
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   12645
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1215
         Width           =   4110
         Begin Threed.SSCommand cmdExit 
            Height          =   465
            Left            =   45
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   820
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "shipping_grd1.frx":0000
            Caption         =   "///"
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdGo 
            Height          =   465
            Left            =   2700
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   820
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "shipping_grd1.frx":2323
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdExcel 
            Height          =   465
            Left            =   1305
            TabIndex        =   71
            Top             =   135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   820
            _Version        =   196610
            ForeColor       =   0
            BackColor       =   16777215
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "shipping_grd1.frx":51D6
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   16785
         TabIndex        =   48
         Top             =   1215
         Width           =   3435
         Begin MSDataListLib.DataCombo xYear 
            Height          =   330
            Left            =   1665
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   180
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo xMonth 
            Height          =   330
            Left            =   90
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   180
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·”‰…"
            Height          =   285
            Left            =   2835
            TabIndex        =   52
            Tag             =   "Color"
            Top             =   225
            Width           =   510
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·‘Â—"
            Height          =   330
            Left            =   855
            TabIndex        =   51
            Tag             =   "Color"
            Top             =   225
            Width           =   555
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   585
         Width           =   4290
         Begin Threed.SSCommand cmdShip 
            Height          =   420
            Left            =   2115
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   180
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   741
            _Version        =   196610
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "‘—ﬂ… «·‘Õ‰"
            TagVariant      =   "‘—ﬂ… «·‘Õ‰"
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdType 
            Height          =   420
            Left            =   45
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   180
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   741
            _Version        =   196610
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "‰Ê⁄ «·ÿ·»Ì…"
            TagVariant      =   "‰Ê⁄ «·»Ê·Ì’…"
            ButtonStyle     =   3
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   20175
      _cx             =   35586
      _cy             =   14473
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   8454143
      ForeColorSel    =   128
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   -1575
      Top             =   3015
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4455
      Top             =   1575
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2475
      Top             =   1125
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   270
      Top             =   1350
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10680
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   476
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   3
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   4
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   5
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   6
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2850
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3825
      Visible         =   0   'False
      Width           =   6225
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   4740
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰  «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   270
            Width           =   1725
         End
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "» «—ÌŒ «” ·«„"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton optDelivery 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3825
            RightToLeft     =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»›« Ê—…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "»œÊ‰ ›« Ê—…"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   270
            Width           =   1230
         End
         Begin VB.OptionButton optDoc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "«·ﬂ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   270
            Value           =   -1  'True
            Width           =   780
         End
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   10485
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "online_ship_grd1frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Public myForm As Form
Public sFilter As String, sSup As String
Dim bStopCell As Boolean
'Dim rsShip As New ADODB.Recordset
Dim bIgClick As Boolean
Dim clist1 As String, cList2 As String
Dim oSearchInv As New Search_abd
Dim oSearchship As New Search_abd
Dim oSearchType As New Search_abd
Dim con As New ADODB.Connection
Private Sub myLoadGrd()
Dim strSql As String
Dim cWhere As String
Dim aFields(12)
aFields(0) = "[SHIP_NO]"
aFields(1) = "[TYPE_DESCA]"
aFields(2) = "[DATE_SHIP]"
aFields(3) = "[SHIP_DESCA]"
aFields(4) = "[DATE_PAY]"
aFields(5) = "[COD]"
aFields(6) = "[FEE]"
aFields(7) = "[ORDER_DATE]"
aFields(8) = "[ORDER_NO]"
aFields(9) = "[INV_NO]"
aFields(10) = "[INV_DATE]"
aFields(11) = "[TOTAL]"
aFields(12) = "[NAME]"

strSql = "SELECT " & arString(aFields, ",") & _
          " FROM vw_online_ship v"

cWhere = retWhere

If cWhere <> "" Then
    strSql = strSql & " WHERE " & cWhere
End If

strSql = strSql & " ORDER BY v.order_date,v.type"

Dim db As New clsDb
Set grid1.DataSource = db.myRs(strSql)
Set db = Nothing
Fixgrd

grid1.ExplorerBar = flexExSort
End Sub
Private Sub myload()
myLoadGrd
End Sub

Private Sub chkDay_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub
Private Sub chkDuplicate_Click()
myload
End Sub

Private Sub chkNoNum_Click()
myload
End Sub

Private Sub chkNoQuant_Click()
myload
End Sub

Private Sub chkOrderReceipt_Click()
myload
End Sub
Private Sub chkDay_ship_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub chkday_Return_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub chkdaymail_Click()
If bIgClick Then Exit Sub
myLoadGrd
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim i As Long
Dim aRow As Variant
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 19)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
Me.MousePointer = 0
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub

Private Sub cmdShip_Click()
ShipLookup Me, oSearchship, , , IIf(cmdShip.Tag <> "", "ﬂ· «·‘—ﬂ« ", "")
End Sub
Private Sub cmdType_Click()
online_TypeLook Me, oSearchType, , , IIf(cmdType.Tag <> "", "ﬂ· «‰Ê«⁄ «·ÿ·»Ì« ", "")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Tag <> "filter" Then
        KeyCode = 0
        grid1.SetFocus
    End If
End If
End Sub
Private Sub Form_Resize()
Dim nHeight As Long
nHeight = Me.Height - (grid1.Top + Picture1.Height + panel1(0).Height + 700)
grid1.Height = IIf(nHeight < 2000, 2000, nHeight)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set online_ship_grd1frm = Nothing
Err.Clear
End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
For i = 2 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i - 1
Next
End Sub

Private Sub grid1_DblClick()
If grid1.Row < 2 Then Exit Sub
 
If grid1.Col = 10 Then
    sales_onlinefrm.bEdit = True
    sales_onlinefrm.sDoc_no = grid1.TextMatrix(grid1.Row, 10)
    sales_onlinefrm.Show 1
Else
    Clipboard.Clear
    Clipboard.SetText grid1.TextMatrix(grid1.Row, grid1.Col)
    Inform " „ «·‰”Œ"
End If
End Sub
Private Sub Form_Load()

Dim db As New clsDb
Set xYear.RowSource = db.myRs(YearsString(Year(Date)))
xYear.ListField = "CODE"
xYear.BoundColumn = "CODE"

Set xMonth.RowSource = db.myRs(MonthString(Month(Date)))
xMonth.ListField = "CODE"
xMonth.BoundColumn = "CODE"
Set db = Nothing
Fixgrd

End Sub
Private Sub Fixgrd()
With grid1
.TextMatrix(0, 0) = "„”·”·"
.TextMatrix(0, 1) = "—ﬁ„ «·»Ê·Ì’…"
.TextMatrix(0, 2) = "‰Ê⁄ «·»Ê·Ì’…"
.TextMatrix(0, 3) = " «—ÌŒ «·‘Õ‰"
.TextMatrix(0, 4) = "‘—ﬂ… «·‘Õ‰"
.TextMatrix(0, 5) = " «—ÌŒ «· Õ’Ì·"
.TextMatrix(0, 6) = " Õ’Ì· «·‘Õ‰"
.TextMatrix(0, 7) = "„’«—Ì› «·‘Õ‰"
.TextMatrix(0, 8) = " «—ÌŒ «·ÿ·»Ì…"
.TextMatrix(0, 9) = "—ﬁ„ «·ÿ·»Ì…"
.TextMatrix(0, 10) = "—ﬁ„ «·›« Ê—…"
.TextMatrix(0, 11) = " «—ÌŒ «·›« Ê—…"
.TextMatrix(0, 12) = "ﬁÌ„… «·›« Ê—…"
.TextMatrix(0, 13) = "≈”„ «·⁄„Ì·"

.RowHeight(0) = 600

.ColWidth(0) = 800
.ColWidth(1) = 1500
.ColWidth(2) = 1300
.ColWidth(3) = 1300
.ColWidth(4) = 1500
.ColWidth(5) = 1300
.ColWidth(6) = 1500
.ColWidth(7) = 1500
.ColWidth(8) = 1300
.ColWidth(9) = 1300
.ColWidth(10) = 1800
.ColWidth(11) = 1300
.ColWidth(12) = 1400
.ColWidth(13) = 2000

.ColDataType(4) = flexDTDate
.ColDataType(5) = flexDTDate
.ColDataType(8) = flexDTDate
.ColDataType(11) = flexDTDate

.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(12) = flexDTDouble

For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 0) = i
Next


For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
    If .ColDataType(i) = flexDTDate Then
        .ColFormat(i) = "yyyy/m/d"
    ElseIf .ColDataType(i) = flexDTDouble Then
        .Subtotal flexSTSum, -1, i, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    End If
Next
End With
End Sub
Private Sub optclose_Click(index As Integer)
If bIgClick Then Exit Sub
myload
End Sub
Private Sub optmail_Click(index As Integer)
If bIgClick Then Exit Sub
myLoadGrd
End Sub
Private Function retWhere() As String
Dim cWhere As New ChilkatStringBuilder

If IsDate(xdate1.text) Then
    cWhere.Append " AND " & "ORDER_DATE >= " & addDate(xdate1.text)
End If

If IsDate(xdate2.text) Then
    cWhere.Append " AND " & "ORDER_DATE <= " & addDate(xdate2.text)
End If

If chkDay.Value = 1 Then
    cWhere.Append " AND " & "ORDER_DATE = " & addDate(Date)
End If

If optType(1).Value Then
    'cWhere.Append " AND v.ORDER_NO IN (SELECT ORDER_NO FROM vw_online_orders_open)"
    cWhere.Append " AND CLOSED = 0"
ElseIf optType(2).Value Then
    cWhere.Append " AND CLOSED = 1"
    'cWhere.Append " AND v.ORDER_NO IN (SELECT ORDER_NO FROM vw_online_orders_closed)"
End If

If optShip(1).Value Then
    cWhere.Append " AND DATE_SHIP IS NOT NULL"
ElseIf optShip(2).Value Then
    cWhere.Append " AND DATE_SHIP IS NULL"
End If

If Trim(xdoc_no.text) <> "" Then
    cWhere.Append " AND " & "v.INV_NO = " & MyParn(xdoc_no.text)
End If

If Trim(xOnline_doc.text) <> "" Then
    cWhere.Append " AND " & "v.ORDER_NO = " & MyParn(xOnline_doc.text)
End If

If Trim(xship_no.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xship_no.text, "SHIP_NO")
End If

If optMail(1).Value Then
    cWhere.Append " AND (DATE_MAIL IS NOT NULL)"
ElseIf optMail(2).Value Then
    cWhere.Append " AND (DATE_MAIL IS NULL)"
End If

If IsDate(xdate_mail1.text) Then
    cWhere.Append " AND " & "v.DATE_MAIL >= " & addDate(xdate_mail1.text)
End If

If chkdayMail.Value = 1 Then
    cWhere.Append " AND " & "v.DATE_MAIL = " & addDate(Date)
End If


If IsDate(xdate_mail2.text) Then
    cWhere.Append " AND " & "v.DATE_MAIL <= " & addDate(xdate_mail2.text)
End If

If IsDate(xdate_ship1.text) Then
    cWhere.Append " AND " & "v.date_ship >= " & addDate(xdate_ship1.text)
End If

If IsDate(xdate_ship2.text) Then
    cWhere.Append " AND " & "v.date_ship <= " & addDate(xdate_ship2.text)
End If

If chkDayShip.Value = 1 Then
    cWhere.Append " AND " & "v.DATE_SHIP = " & addDate(Date)
End If



If Trim(xName.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xName.text, "v.NAME")
End If

If Trim(xPhone.text) <> "" Then
    cWhere.Append " AND " & MyParnAnd(xPhone.text, "v.PHONE")
End If

If cmdType.Tag <> "" Then
    cWhere.Append " AND " & "v.TYPE = " & cmdType.Tag
End If

If cmdShip.Tag <> "" Then
    cWhere.Append " AND " & "v.SHIP = " & MyParn(cmdShip.Tag)
End If
retWhere = Mid(cWhere.GetAsString, 6)
End Function
Public Sub myProc()
If ActiveControl.Name = cmdShip.Name Then
    If oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 0) = "" Then
        cmdShip.Tag = ""
        cmdShip.Caption = cmdShip.TagVariant
    Else
        cmdShip.Tag = oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 0)
        cmdShip.Caption = oSearchship.grid1.TextMatrix(oSearchship.grid1.Row, 1)
    End If
    myload
    oSearchship.Hide
ElseIf ActiveControl.Name = cmdType.Name Then
    If oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0) = "" Then
        cmdType.Tag = ""
        cmdType.Caption = cmdType.TagVariant
    Else
        cmdType.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
        cmdType.Caption = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1)
    End If
    myload
    oSearchType.Hide
End If
End Sub
Private Sub optShip_Click(index As Integer)
myLoadGrd
End Sub
Private Sub xYear_Change()
If bIgClick Then Exit Sub
myLoadGrd
End Sub
Private Sub xMonth_Change()
If bIgClick Then Exit Sub
If xYear.text = "" Then Exit Sub
myLoadGrd
End Sub
Private Sub xYear_GotFocus()
myGotFocus xYear
End Sub
Private Sub xYear_LostFocus()
myLostFocus xYear
If Not xYear.MatchedWithList Then xYear.BoundText = ""
End Sub
Private Sub xMonth_GotFocus()
myGotFocus xMonth
End Sub
Private Sub xMonth_LostFocus()
myLostFocus xMonth
If Not xMonth.MatchedWithList Then xMonth.BoundText = ""
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub XONLINE_DOC_GotFocus()
myGotFocus xOnline_doc
End Sub
Private Sub XONLINE_DOC_LostFocus()
myLostFocus xOnline_doc
End Sub
Private Sub xship_no_GotFocus()
myGotFocus xship_no
End Sub
Private Sub xship_no_LostFocus()
myLostFocus xship_no
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
End Sub
Private Sub xName_GotFocus()
myGotFocus xName
End Sub
Private Sub xname_LostFocus()
myLostFocus xName
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub XPHONE_LostFocus()
myLostFocus xPhone
End Sub
Private Sub xdate_ship2_GotFocus()
myGotFocus xdate_ship2
End Sub
Private Sub xdate_ship2_LostFocus()
myLostFocus xdate_ship2
myValidDate xdate_ship2
End Sub
Private Sub xdate_ship1_GotFocus()
myGotFocus xdate_ship1
End Sub
Private Sub xdate_ship1_LostFocus()
myLostFocus xdate_ship1
myValidDate xdate_ship1
End Sub
Private Sub xdate_Mail1_GotFocus()
myGotFocus xdate_mail1
End Sub
Private Sub xdate_Mail1_LostFocus()
myLostFocus xdate_mail1
myValidDate xdate_mail1
End Sub
Private Sub xdate_Mail2_GotFocus()
myGotFocus xdate_mail2
End Sub
Private Sub xdate_Mail2_LostFocus()
myLostFocus xdate_mail2
myValidDate xdate_mail2
End Sub

