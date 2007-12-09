VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame objFrameJobs 
      Caption         =   "Jobs and Log"
      Height          =   6375
      Left            =   240
      TabIndex        =   86
      Top             =   600
      Width           =   11055
      Begin VB.CommandButton objCmdRemoveAll 
         Caption         =   "Remove all"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton objCmdRemove 
         Caption         =   "Remove job"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton objCmdAdd 
         Caption         =   "Add job"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton objCmdRun 
         Caption         =   "Run batch"
         Height          =   255
         Left            =   8520
         TabIndex        =   8
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton objCmdStop 
         Caption         =   "Stop batch"
         Height          =   255
         Left            =   9720
         TabIndex        =   9
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton objCmdAddJoblist 
         Caption         =   "Add joblist"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton objCmdClear 
         Caption         =   "Clear log"
         Height          =   255
         Left            =   9720
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox objTextLog 
         Height          =   5175
         Left            =   5520
         TabIndex        =   3
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   9128
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0442
      End
      Begin MSComctlLib.ListView objJobList 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   9128
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "objImageList"
         SmallIcons      =   "objImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label objLabelBatch 
         Caption         =   "Batch list (0 jobs)"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label objLabelLog 
         Caption         =   "Log"
         Height          =   255
         Left            =   5640
         TabIndex        =   87
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame objFrameLog 
      Caption         =   "Regenerator Log file settings"
      Height          =   6375
      Left            =   240
      TabIndex        =   85
      Top             =   600
      Width           =   11055
      Begin VB.CheckBox objCheckVerboseLog 
         Caption         =   "Verbose Log"
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox objTextSaveLogPath 
         Height          =   285
         Left            =   240
         TabIndex        =   105
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CommandButton objCmdLogFolderBrws 
         Caption         =   "browse"
         Height          =   255
         Left            =   4320
         TabIndex        =   106
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label objLabelSaveLogPath 
         Caption         =   "Log file save path (folder only)"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   840
         Width           =   3855
      End
   End
   Begin VB.Frame objFrameSettings 
      Caption         =   "Advanced Settings"
      Height          =   6375
      Left            =   240
      TabIndex        =   54
      Top             =   600
      Width           =   11055
      Begin VB.Frame Frame3 
         Caption         =   "Misc Fix"
         Height          =   1215
         Left            =   120
         TabIndex        =   82
         Top             =   3720
         Width           =   4935
         Begin VB.CheckBox objCheckAddCss 
            Caption         =   "Add default CSS"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   720
            Width           =   4575
         End
         Begin VB.CheckBox objCheckPb2kLayoutFix 
            Caption         =   "PB2K/TK NCC Layout Fix"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame objFrameAdvancedFix 
         Caption         =   "Advanced Fix"
         Height          =   6015
         Left            =   5280
         TabIndex        =   83
         Top             =   240
         Width           =   5655
         Begin VB.Frame Frame5 
            Caption         =   "Broken NCC/Content Doc Links"
            Height          =   615
            Left            =   120
            TabIndex        =   110
            Top             =   1320
            Width           =   5415
            Begin VB.CheckBox objCheckDisableXhtmlLinks 
               Caption         =   "Disable broken links"
               Height          =   195
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   2655
            End
            Begin VB.CheckBox objCheckEstimateXhtmlLinks 
               Caption         =   "Attempt estimation first"
               Height          =   195
               Left            =   2880
               TabIndex        =   69
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Point SMIL targets to..."
            Height          =   615
            Left            =   120
            TabIndex        =   109
            Top             =   2040
            Width           =   5415
            Begin VB.OptionButton objRadioSmilTargetText 
               Caption         =   "<text>"
               Height          =   255
               Left            =   3840
               TabIndex        =   72
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton objRadioSmilTargetPar 
               Caption         =   "<par>"
               Height          =   195
               Left            =   2040
               TabIndex        =   71
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton objRadioSmilTargetNoChange 
               Caption         =   "no change"
               Height          =   315
               Left            =   120
               TabIndex        =   70
               Top             =   180
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin VB.CheckBox objCheckMakeTrueNccOnly 
            Caption         =   "Make True NCC only"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   930
            Width           =   5295
         End
         Begin VB.Frame Frame2 
            Caption         =   "Phrase Merge"
            Height          =   3135
            Left            =   120
            TabIndex        =   84
            Top             =   2760
            Width           =   5415
            Begin VB.CheckBox objCheckMergeShort 
               Caption         =   "Merge short first phrases:"
               Height          =   255
               Left            =   240
               TabIndex        =   73
               Top             =   240
               Width           =   3975
            End
            Begin VB.TextBox objTextMergeIfLower 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3840
               TabIndex        =   75
               Text            =   "10000"
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox objTextMergeIfShorter2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3840
               TabIndex        =   77
               Text            =   "10000"
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox objTextMergeAndNextIsShorter 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3840
               TabIndex        =   79
               Text            =   "10000"
               Top             =   2040
               Width           =   615
            End
            Begin VB.TextBox objTextClipEndBeginSpan 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3840
               TabIndex        =   81
               Text            =   "10000"
               Top             =   2640
               Width           =   615
            End
            Begin MSComctlLib.Slider objSliderMergeIfLower 
               Height          =   255
               Left            =   240
               TabIndex        =   74
               Top             =   840
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   100
               SmallChange     =   10
               Max             =   10000
               TickFrequency   =   500
            End
            Begin MSComctlLib.Slider objSliderMergeIfShorter2 
               Height          =   255
               Left            =   240
               TabIndex        =   76
               Top             =   1440
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   100
               SmallChange     =   10
               Max             =   10000
               TickFrequency   =   500
            End
            Begin MSComctlLib.Slider objSliderClipEndBeginSpan 
               Height          =   255
               Left            =   240
               TabIndex        =   80
               Top             =   2640
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   100
               SmallChange     =   10
               Max             =   10000
               TickFrequency   =   500
            End
            Begin MSComctlLib.Slider objSliderMergeAndNextIsShorter 
               Height          =   255
               Left            =   240
               TabIndex        =   78
               Top             =   2040
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   100
               SmallChange     =   10
               Max             =   10000
               TickFrequency   =   500
            End
            Begin VB.Label objLabelMergeIfLower 
               Caption         =   "Merge with next if clip is shorter than:"
               Height          =   255
               Left            =   240
               TabIndex        =   102
               Top             =   600
               Width           =   3975
            End
            Begin VB.Label objLabelMergeIfShorter2 
               Caption         =   "Merge with next if clip is shorter than:"
               Height          =   255
               Left            =   240
               TabIndex        =   101
               Top             =   1200
               Width           =   3975
            End
            Begin VB.Label objLabelMergeAndNextIsShorter 
               Caption         =   "... and next clip is shorter than:"
               Height          =   255
               Left            =   240
               TabIndex        =   100
               Top             =   1800
               Width           =   3375
            End
            Begin VB.Label objLabelClipEndBeginSpan 
               Caption         =   "Allowed span between clip-end and clip-begin"
               Height          =   255
               Left            =   240
               TabIndex        =   99
               Top             =   2400
               Width           =   3615
            End
         End
         Begin VB.CheckBox objCheckMangleLinks 
            Caption         =   "Rebuild link structure"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   600
            Width           =   5175
         End
         Begin VB.CheckBox objCheckFixPar 
            Caption         =   "Adjust invalid smil par elements"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Value           =   1  'Checked
            Width           =   5295
         End
      End
      Begin VB.Frame objMiscGeneralSettings 
         Caption         =   "Misc Settings"
         Height          =   1575
         Left            =   120
         TabIndex        =   58
         Top             =   2040
         Width           =   5055
         Begin VB.CommandButton objCmdBrowseProgram 
            Caption         =   "Browse"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   62
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox objTextLaunchProgram 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   61
            Top             =   960
            Width           =   2775
         End
         Begin VB.CheckBox objCheckLaunchProgram 
            Caption         =   "Launch program after each job done"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   600
            Width           =   3975
         End
         Begin VB.CheckBox objCheckHalt 
            Caption         =   "Halt on regeneration / file rendering error"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame objFrameJobDefault 
         Caption         =   "Default Paths"
         Height          =   1695
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   5055
         Begin VB.TextBox objTextDefaultMetaPath 
            Height          =   285
            Left            =   240
            TabIndex        =   57
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox objTextDefaultSavepath 
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label objLabelDefaultMetaPath 
            Caption         =   "Default Meta Location"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label objLabelDefaultSavepath 
            Caption         =   "Default Destination Directory"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   360
            Width           =   3975
         End
      End
   End
   Begin VB.Frame objFrameJobProperties 
      Caption         =   "Properties for job"
      Height          =   6375
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   11055
      Begin VB.CommandButton objCmdSetAll 
         Caption         =   "Set all jobs to these settings"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   5880
         Width           =   2655
      End
      Begin VB.CommandButton objCmdRestore 
         Caption         =   "Restore factory settings"
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Frame objFrameMeta 
         Caption         =   "Metadata Handling"
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   5175
         Begin VB.CommandButton objCmdMetaBrws 
            Caption         =   "Browse"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4080
            TabIndex        =   20
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox objTextMetaFile 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   3735
         End
         Begin VB.OptionButton objRadioMetaImp 
            Caption         =   "Import"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton objRadioMetaPres 
            Caption         =   "Preserve"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame objFrameIBP 
         Caption         =   "Input DTB properties"
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   5175
         Begin VB.ComboBox objComboIanaCS 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2760
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   1080
            Width           =   1695
         End
         Begin VB.ComboBox objComboCharset 
            Height          =   315
            ItemData        =   "frmMain.frx":04C4
            Left            =   1200
            List            =   "frmMain.frx":04D7
            TabIndex        =   14
            Text            =   "Western -- utf-8"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox objComboDTBType 
            Height          =   315
            ItemData        =   "frmMain.frx":0503
            Left            =   1200
            List            =   "frmMain.frx":0519
            TabIndex        =   13
            Text            =   "audioNcc"
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label objLabelCharset 
            Caption         =   "Input charset"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label objLabelDTBType 
            Caption         =   "DTB Type"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame objFrameFilenames 
         Caption         =   "Alter filenames"
         Height          =   1815
         Left            =   5400
         TabIndex        =   23
         Top             =   360
         Width           =   5535
         Begin VB.CheckBox objCheckSeqRen 
            Caption         =   "Sequential rename"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.TextBox objTextPrefix 
            Height          =   285
            Left            =   840
            TabIndex        =   26
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox objCheckNumeric 
            Caption         =   "Use numeric portion of ID"
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label objLabelPrefix 
            Caption         =   "Prefix"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   615
         End
      End
      Begin VB.Frame objFrameSave 
         Caption         =   "Output DTB Destination Directory"
         Height          =   1935
         Left            =   5400
         TabIndex        =   27
         Top             =   2280
         Width           =   5535
         Begin VB.CheckBox objCheckMoveBook 
            Caption         =   "Move book"
            Height          =   255
            Left            =   3240
            TabIndex        =   30
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton objCmdFolderBrws 
            Caption         =   "Browse"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4440
            TabIndex        =   32
            Top             =   1320
            Width           =   855
         End
         Begin VB.OptionButton objRadioSameFolder 
            Caption         =   "Same folder (replace old book)"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton objRadioNewFolder 
            Caption         =   "New folder"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox objTextFoldername 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label Label1 
            Caption         =   "Destination Path"
            Height          =   210
            Left            =   240
            TabIndex        =   0
            Top             =   1080
            Width           =   2535
         End
      End
   End
   Begin VB.Frame objFrameValidator 
      Caption         =   "Validation settings"
      Height          =   6375
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Width           =   11055
      Begin VB.Frame objFrameValidatorLog 
         Caption         =   "Validator Log settings"
         Height          =   2775
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   4935
         Begin VB.CommandButton objCmdValReportFolderBrws 
            Caption         =   "browse"
            Height          =   255
            Left            =   3840
            TabIndex        =   103
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox objTextStandalonePath 
            Height          =   285
            Left            =   240
            TabIndex        =   42
            Top             =   2280
            Width           =   3375
         End
         Begin VB.CheckBox objCheckCreateStandalone 
            Caption         =   "Create standalone validator report for each job"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   3855
         End
         Begin VB.CheckBox objCheckIncludeAdvancedADTD 
            Caption         =   "Include advanced ADTD information in log"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   3975
         End
         Begin VB.CheckBox objCheckIncludeNCErrors 
            Caption         =   "Include non-critical errors in log"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   3975
         End
         Begin VB.CheckBox objCheckIncludeWarnings 
            Caption         =   "Include warnings in log"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label objLabelStandalonePath 
            Caption         =   "Validator report savepath"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2040
            Width           =   3975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Validator general settings"
         Height          =   4095
         Left            =   5160
         TabIndex        =   43
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton objCmdValTmpFolderBrws 
            Caption         =   "browse"
            Height          =   255
            Left            =   4680
            TabIndex        =   53
            Top             =   3600
            Width           =   855
         End
         Begin VB.CommandButton objCmdValVtmFolderBrws 
            Caption         =   "browse"
            Height          =   255
            Left            =   4680
            TabIndex        =   51
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton objCmdValExtFolderBrws 
            Caption         =   "browse"
            Height          =   255
            Left            =   4680
            TabIndex        =   49
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox objTextTimeFluctuation 
            Height          =   300
            Left            =   4680
            TabIndex        =   47
            Top             =   1560
            Width           =   735
         End
         Begin VB.CheckBox objCheckValLightMode 
            Caption         =   "Validator Light Mode"
            Height          =   300
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox objTextTempPath 
            Height          =   285
            Left            =   240
            TabIndex        =   52
            Top             =   3600
            Width           =   4215
         End
         Begin VB.TextBox objTextExtPath 
            Height          =   285
            Left            =   240
            TabIndex        =   48
            Top             =   2160
            Width           =   4215
         End
         Begin VB.TextBox objTextVTMPath 
            Height          =   285
            Left            =   240
            TabIndex        =   50
            Top             =   2880
            Width           =   4215
         End
         Begin VB.CheckBox objCheckSyncWValidator 
            Caption         =   "Synchronize settings with validator software"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   3975
         End
         Begin MSComctlLib.Slider objSliderTimeFluctuation 
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1560
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   100
            SmallChange     =   10
            Max             =   1500
            TickFrequency   =   100
         End
         Begin VB.Label objLabelTempPath 
            Caption         =   "Temp directory"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   3360
            Width           =   3975
         End
         Begin VB.Label objLabelTimeFluct 
            Caption         =   "Allowed time fluctuation (ms)"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label objLabelTimeFluctMin 
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label objLabelTimeFLuctMax 
            Caption         =   "1500"
            Height          =   255
            Left            =   4080
            TabIndex        =   95
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label objLabelExtPath 
            Caption         =   "Externals directory"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1920
            Width           =   3975
         End
         Begin VB.Label objLabelVTMPath 
            Caption         =   "VTM directory"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   2640
            Width           =   3975
         End
      End
      Begin VB.CheckBox objCheckValidateJob 
         Caption         =   "Validate job after regeneration"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   $"frmMain.frx":056E
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   4095
      End
   End
   Begin MSComctlLib.TabStrip objTabstrip 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12303
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Jobs"
            Key             =   "tab1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Job properties"
            Key             =   "tab2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Validation settings"
            Key             =   "tab4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced settings"
            Key             =   "tab3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log settings"
            Key             =   "tab5"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList objImageList 
      Left            =   6120
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0637
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0989
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":102D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer objTmrUpdate 
      Interval        =   500
      Left            =   6720
      Top             =   -120
   End
   Begin MSComDlg.CommonDialog objCommonDialog 
      Left            =   7200
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Begin VB.Menu mnuRunBatch 
         Caption         =   "Run Batch"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuStopBatch 
         Caption         =   "Stop Batch"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuJob 
      Caption         =   "Job"
      Begin VB.Menu mnuAddJob 
         Caption         =   "Add Job..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAddJobList 
         Caption         =   "Add Joblist..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuRemoveJob 
         Caption         =   "Remove Job"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuFocus 
      Caption         =   "Focus"
      Begin VB.Menu mnuJobsTabFocus 
         Caption         =   "Jobs Tab"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuJobPropTabFocus 
         Caption         =   "Job Properties Tab"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuValSettingsTabFocus 
         Caption         =   "Validation Settings Tab"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuAdvancedSettingsTabFocus 
         Caption         =   "Advanced Settings Tab"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu LogSettingsTabFocus 
         Caption         =   "Log Settings Tab"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuJobWindowFocus 
         Caption         =   "Job Queue Window"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuLogWinFocus 
         Caption         =   "Log Window"
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "Regenerator Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpDevManual 
         Caption         =   "Developer Manual"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Daisy 2.02 Regenerator Batch UI
' Copyright (C) 2003 Daisy Consortium
'
'    This file is part of Daisy 2.02 Regenerator.
'
'    Daisy 2.02 Regenerator is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Daisy 2.02 Regenerator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Daisy 2.02 Regenerator; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Option Explicit

' This flag is indicating wheter the UI is updating itself or not (to prevent the
' different controls events from trying to update while it's currently updating)
Private bolOwnChange As Boolean, bolprivBusy As Boolean
Private sLastLogMessage As String

Public Sub subVBFriendlyEvent( _
    isEvent As String, _
    vParam1 As Variant, _
    vParam2 As Variant)
  
  Select Case isEvent
    Case "Regenerator.AddLog"
      fncAddLog CStr(vParam1), True
    
    Case "ValidatorEngine.ErrorLog"
      fncAddLog CStr(vParam1), True

    Case "ValidatorEngine.SucceededTest"
    
    Case "ValidatorEngine.FailedTest"
    
    Case "ValidatorEngine.ProgressChanged"
            
    Case "ValidatorEngine.Log"
      fncAddLog CStr(vParam1), True
    Case Else
      'Stop
  End Select
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuExit_Click()
  Form_Unload 0
  Unload Me
End Sub

Private Sub mnuHelpManual_Click()
  ShellExecute frmMain.hwnd, "Open", App.Path & _
    "\manual\regenerator_manual.html", "", "", vbNormalFocus
End Sub

Private Sub mnuHelpDevManual_Click()
  ShellExecute frmMain.hwnd, "Open", App.Path & _
    "\manual\regenerator_developer.html", "", "", vbNormalFocus
End Sub


Private Sub objCheckMoveBook_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Public Sub subUpdateIanaCharactersets()
' Stuff the IANA combo with charsets
  Do Until objComboIanaCS.ListCount = 0
    objComboIanaCS.RemoveItem (0)
  Loop
  
  Dim lCounter As Long
  For lCounter = 0 To lCharsetCount - 1
    objComboIanaCS.AddItem aIanaCharset(lCounter)
  Next lCounter
End Sub

Private Sub Form_Activate()
  subRefreshInterface
End Sub

Private Sub Form_Initialize()
  Form_Activate
End Sub

' Do some initialization
Private Sub Form_Load()
  objJobList.ColumnHeaders.Add , , "Path", ((objJobList.Width / 100) * 50)
  objJobList.ColumnHeaders.Add , , "Reg. Status", ((objJobList.Width / 100) * 24)
  objJobList.ColumnHeaders.Add , , "Val. Status", ((objJobList.Width / 100) * 24)
  Caption = objRegeneratorUserControl.fncGetDllVersion
  
' Enable the UI so it will work with all languages
  subApplyUserLcid
  
  Show
End Sub

' Unloading
Private Sub Form_Unload(Cancel As Integer)
' Cancel unload if the program is regenerating
  If bolRegenerating Then
    MsgBox ("Stop regenerating before exit")
    Cancel = 1
    Exit Sub
  End If
  
  fncDeinitValidator
' Save registry settings
  fncSaveRegistrySettings
  'fncAddMemLog "Exiting program."
End Sub

' The following functions are just updating the interface and the settings / job
' properties

Private Sub objCheckCreateStandalone_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckHalt_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckIncludeAdvancedADTD_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckIncludeNCErrors_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckIncludeWarnings_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckLaunchProgram_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckMangleLinks_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckFixPar_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
End Sub

Private Sub objCheckDisableXhtmlLinks_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckAddCss_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckEstimateXhtmlLinks_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckMakeTrueNccOnly_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckMergeShort_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckNumeric_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Private Sub objCheckSeqRen_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Private Sub objCheckVerboseLog_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckSyncWValidator_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckValLightMode_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCmdClear_Click()
  objTextLog.Text = ""
End Sub

Private Sub objCmdSetAll_Click()
  Dim objResult As VbMsgBoxResult, lCounter As Long
  frmApplySettings.Show vbModal
End Sub

Private Sub objCmdValReportFolderBrws_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Validator Report Save directory", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  objTextStandalonePath.Text = objFolder.Self.Path
  If Not Left$(objTextStandalonePath.Text, 1) = "\" Then objTextStandalonePath.Text = _
    objTextStandalonePath.Text & "\"

End Sub

Private Sub objComboCharset_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Private Sub objComboDTBType_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  fncSetJobProperties
End Sub

Private Sub objComboIanaCS_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  fncSetJobProperties
End Sub

Private Sub objCheckValidateJob_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objRadioMetaImp_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Private Sub objRadioMetaPres_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Private Sub objRadioNewFolder_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub

Private Sub objRadioSameFolder_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
  fncSetJobProperties
End Sub


Private Sub objSliderClipEndBeginSpan_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objSliderClipEndBeginSpan_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If bolOwnChange Then Exit Sub
    fncUpdateInfoFromUI
    subRefreshInterface
  End If
End Sub

Private Sub objSliderMergeAndNextIsShorter_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objSliderMergeAndNextIsShorter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If bolOwnChange Then Exit Sub
    fncUpdateInfoFromUI
    subRefreshInterface
  End If
End Sub

Private Sub objSliderMergeIfLower_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objSliderMergeIfLower_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If bolOwnChange Then Exit Sub
    fncUpdateInfoFromUI
    subRefreshInterface
  End If
End Sub

Private Sub objSliderMergeIfShorter2_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objSliderMergeIfShorter2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If bolOwnChange Then Exit Sub
    fncUpdateInfoFromUI
    subRefreshInterface
  End If
End Sub

Private Sub objSliderTimeFluctuation_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objTabstrip_Click()
  subRefreshInterface
End Sub

Private Sub objTextDefaultMetaPath_change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
End Sub

Private Sub objTextDefaultSavepath_change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
End Sub

'Private Sub objCheckPointTargetsToPar_Click()
'  If bolOwnChange Then Exit Sub
'  fncUpdateInfoFromUI
'  subRefreshInterface
'End Sub zzz

Private Sub objRadioSmilTargetNoChange_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
End Sub

Private Sub objRadioSmilTargetPar_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
End Sub

Private Sub objRadioSmilTargetText_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
End Sub



Private Sub objTextExtPath_LostFocus()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objTextFoldername_change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  fncSetJobProperties
End Sub

Private Sub objTextLaunchProgram_LostFocus()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objCheckPb2kLayoutFix_Click()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub


'*************************************
'added by mg 20030215 +

Private Sub objTextTimeFluctuation_LostFocus()
 If IsNumeric(objTextTimeFluctuation.Text) Then
   If (CLng(objTextTimeFluctuation.Text) < 1500) Or _
    (CLng(objTextTimeFluctuation.Text) > 0) Then
     objSliderTimeFluctuation.Value = CLng(objTextTimeFluctuation.Text)
     fncUpdateInfoFromUI
     subRefreshInterface
   Else
     objTextTimeFluctuation.Text = objSliderTimeFluctuation.Value
   End If
 Else
   objTextTimeFluctuation.Text = "0"
 End If

End Sub

Private Sub objTextMergeAndNextIsShorter_LostFocus()
  If IsNumeric(objTextMergeAndNextIsShorter.Text) Then
    If CLng(objTextMergeAndNextIsShorter.Text) > 10000 Then objTextMergeAndNextIsShorter.Text = 10000
    objSliderMergeAndNextIsShorter.Value = CLng(Trim$(objTextMergeAndNextIsShorter.Text))
  Else
    objTextMergeAndNextIsShorter.Text = "1000"
  End If
End Sub

Private Sub objTextMergeIfLower_LostFocus()
 If IsNumeric(objTextMergeIfLower.Text) Then
   If CLng(objTextMergeIfLower.Text) > 10000 Then objTextMergeIfLower.Text = 10000
   objSliderMergeIfLower.Value = CLng(Trim$(objTextMergeIfLower.Text))
 Else
   objTextMergeIfLower.Text = "100"
 End If
End Sub

Private Sub objTextMergeIfShorter2_LostFocus()
 If IsNumeric(objTextMergeIfShorter2.Text) Then
   If CLng(objTextMergeIfShorter2.Text) > 10000 Then objTextMergeIfShorter2.Text = 10000
   objSliderMergeIfShorter2.Value = CLng(Trim$(objTextMergeIfShorter2.Text))
 Else
   objTextMergeIfShorter2.Text = "500"
 End If
End Sub

Private Sub objTextClipEndBeginSpan_Lostfocus()
 If IsNumeric(objTextClipEndBeginSpan.Text) Then
   If CLng(objTextClipEndBeginSpan.Text) > 10000 Then objTextClipEndBeginSpan.Text = 10000
   objSliderClipEndBeginSpan.Value = CLng(Trim$(objTextClipEndBeginSpan.Text))
 Else
   objTextClipEndBeginSpan.Text = "50"
 End If
End Sub

'***********************************

Private Sub objTextMetaFile_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  fncSetJobProperties
End Sub

Private Sub objTextPrefix_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  fncSetJobProperties
End Sub

Private Sub objTextSaveLogPath_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objTextSaveLogPath_LostFocus()
  If InStr(1, objTextSaveLogPath.Text, "*") Then
    MsgBox ("path variables are not allowed here")
  End If
End Sub

Private Sub objCmdValTmpFolderBrws_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose directory for Validator temporary file storage", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  objTextTempPath.Text = objFolder.Self.Path
  If Not Left$(objTextTempPath.Text, 1) = "\" Then objTextTempPath.Text = _
    objTextTempPath.Text & "\"

'  fncSetJobProperties
End Sub

Private Sub objCmdValVtmFolderBrws_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose directory of Validator 'vtm.xml' file", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  objTextVTMPath.Text = objFolder.Self.Path
  If Not Left$(objTextVTMPath.Text, 1) = "\" Then objTextVTMPath.Text = _
    objTextVTMPath.Text & "\"

'  fncSetJobProperties
End Sub

Private Sub objCmdValExtFolderBrws_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Validator Externals directory", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  objTextExtPath.Text = objFolder.Self.Path
  If Not Left$(objTextExtPath.Text, 1) = "\" Then objTextExtPath.Text = _
    objTextExtPath.Text & "\"

End Sub

Private Sub objTextStandalonePath_Change()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objTextTempPath_LostFocus()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

Private Sub objTextVTMPath_LostFocus()
  If bolOwnChange Then Exit Sub
  fncUpdateInfoFromUI
  subRefreshInterface
End Sub

' This control is adding a job to the joblist
Private Sub objCmdAdd_Click()
  Dim sFileName As String
  If Not _
    fncOpenFile("NCC file (ncc.htm*)|ncc.htm*", "ncc.html", True, sFileName) Then _
    Exit Sub
  fncInsertJob sFileName
  fncUpdateJobList
  subRefreshInterface
End Sub

' This control is adding a joblist file to the joblist
Private Sub objCmdAddJoblist_Click()
  MousePointer = 11
  Dim sFile As String
  If fncOpenFile("XML file (*.xml)|*.xml", "", True, sFile) Then _
    fncAddJobList (sFile)
  fncUpdateJobList
  MousePointer = 0
End Sub

' This control is the folder browser for the destination folder of regeneration
Private Sub objCmdFolderBrws_Click()

  Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose save folder", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)
  
  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub
  
  objTextFoldername.Text = objFolder.Self.Path
  If Not Left$(objTextFoldername.Text, 1) = "\" Then objTextFoldername.Text = _
    objTextFoldername.Text & "\"
  
  fncSetJobProperties
End Sub

' This control is the folder browser for the destination folder of regenerator.log
Private Sub objCmdLogFolderBrws_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose save folder for log file", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)
  
  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  objTextSaveLogPath.Text = objFolder.Self.Path
  If Not Left$(objTextSaveLogPath.Text, 1) = "\" Then objTextSaveLogPath.Text = _
    objTextSaveLogPath.Text & "\"
  
  fncSetJobProperties
End Sub


' This function is browsing for a meta import file
Private Sub objCmdMetaBrws_Click()
  Dim sFile As String
  If fncOpenFile("XML file (*.xml)|*.xml", "", True, sFile) Then _
    objTextMetaFile.Text = sFile
  fncSetJobProperties
End Sub

' This control removes the selected job
Private Sub objCmdRemove_Click()
  If objJobList.SelectedItem Is Nothing Then Exit Sub
  
  Dim lCounter As Long
  
  For lCounter = 1 To lJobCount
    If aJobItems(lCounter).sPath = objJobList.SelectedItem.Text Then Exit For
  Next lCounter
  
  If lCounter > lJobCount Then Exit Sub
  
  Set aJobItems(lCounter) = Nothing
  
  For lCounter = lCounter To lJobCount - 1
    Set aJobItems(lCounter) = aJobItems(lCounter + 1)
  Next lCounter
  
  objJobList.ListItems.Remove lJobCount
  
  Set aJobItems(lJobCount) = Nothing
  
  lJobCount = lJobCount - 1
  ReDim Preserve aJobItems(lJobCount)
  
  fncUpdateJobList
  subRefreshInterface
End Sub

' This control removes all jobs
Private Sub objCmdRemoveAll_Click()
  lJobCount = 0
  ReDim aJobItems(0)
  
  Do Until objJobList.ListItems.Count = 0
    objJobList.ListItems.Remove (1)
  Loop
  
  fncUpdateJobList
  subRefreshInterface
End Sub

' This control sets all job properties to default values
Private Sub objCmdRestore_Click()
  Dim sPath As String
  
  If fncParsePathWithConstants(sDefaultSavePath, sPath, lJobCount) Then _
    objTextFoldername.Text = sPath
  If fncParsePathWithConstants(sDefaultMetaPath, sPath, lJobCount) Then _
    objTextMetaFile.Text = sPath & "metadata.xml"
  
  objComboDTBType.ListIndex = 1
  objComboCharset.ListIndex = 0
  objComboIanaCS.ListIndex = -1
  objRadioMetaPres.Value = True
  objCheckSeqRen.Value = vbChecked
  objCheckNumeric.Value = vbUnchecked
  objTextPrefix.Text = ""
  objRadioSameFolder.Value = False
  fncSetJobProperties
  subRefreshInterface
End Sub

' This control runs all jobs in the jobqueue
Private Sub objCmdRun_Click()
  'subRunJobQueue
  fncRunJobQueue
End Sub

' This control aborts all regenerations
Private Sub objCmdStop_Click()
  If bolRegenerating Then bolAbort = True
End Sub

' This control selects a job in the jobqueue
Private Sub objJobList_Click()
  If objJobList.SelectedItem Is Nothing Then Exit Sub
  
  If Not bolRegenerating Then lCurrentJob = objJobList.SelectedItem.Index
  
  objJobList.ListItems.Item(lCurrentJob).Selected = True
  With aJobItems(lCurrentJob)
    lDtbType = .lDtbType
    lCharset = .lInputCharset
    lIANACharset = .lCharsetOther
    bolPreserveMeta = .bolPreserveMeta
    sMetaFile = .sMetaImport
    bolSeqRename = .bolSeqRename
    bolUseNumeric = .bolUseNumeric
    sPrefix = .sPrefix
    bolSameFolder = .bolSaveSame
    sSavePath = .sNewFolder
  End With
  
  subRefreshInterface
End Sub

Public Function fncAddLog( _
    ByVal sLogItem As String, _
    ByVal bolIncludeInFileLog As Boolean _
    )
Dim sUiText As String 'trims the xml data for display

  'This function adds a log to the log RTF edit box

  If (sLogItem <> "") And (sLogItem <> sLastLogMessage) Then
    sUiText = fncTrimXml(sLogItem)
    If sUiText <> "" Then
      objTextLog.Text = objTextLog.Text & sUiText & vbCrLf
      objTextLog.SelStart = Len(objTextLog.Text)
      ' Save the log in the current job items personal log
      objCurrentLog.fncAddToLog sUiText
    End If
    
    sLastLogMessage = sLogItem
    On Error Resume Next
    If bolIncludeInFileLog Then
      If InStr(1, sLogItem, "<status>") < 1 Then '<status> messages dont need to go into filelog
        objLogFile.writeline (sLogItem)
      End If
    End If
  End If

' If the log text box has more than 32000 characters, clear the textbox
' If not, the program will be exteremely slow
  If (Len(objTextLog.Text) >= 32000) And (bolRegenerating) Then objTextLog.Text = ""
  
  DoEvents
End Function

Private Function fncTrimXml(sXml As String) As String
Dim lGt As Long
Dim lLt As Long
  
  'gets the text node from a nonempty element

  'find the first occurence of ">"
  'find the first occurence of "<" after ">"
  'return the chars between these two positions
  On Error GoTo errH
  lGt = InStr(1, sXml, ">")
  If lGt = 0 Then 'this is not an xml string
    fncTrimXml = sXml
    Exit Function
  End If
  lLt = InStr(lGt, sXml, "<")
  If lGt > 1 Then
    If lLt > lGt Then
      fncTrimXml = Mid(sXml, lGt + 1, lLt - lGt - 1)
      Exit Function
    End If
  End If
  fncTrimXml = ""
  Exit Function
errH:
  fncTrimXml = sXml
End Function

' This function takes care of the events coming from the timer TmrUpdate and updates
' info on the screen
Private Sub objTmrUpdate_Timer()
  Dim sProgress As String

  lProgress = lProgress + 1
  Select Case lProgress
    Case 1
      sProgress = "|"
    Case 2
      sProgress = "/"
    Case 3
      sProgress = "-"
    Case 4
      sProgress = "\"
      lProgress = 0
  End Select

  If bolAbort Then
    Caption = "Regenerator " & objRegeneratorUserControl.fncGetDllVersion & " " & sProgress
  ElseIf bolRegenerating Then
    Caption = "Regenerator " & objRegeneratorUserControl.fncGetDllVersion & " " & sProgress
  Else
    Caption = "Regenerator " & objRegeneratorUserControl.fncGetDllVersion
  End If
  
  If bolAbort Then Caption = Caption & " [aborting]"
End Sub
'!!!
' This function updates the current settings from the choices made in the UI
Public Function fncUpdateInfoFromUI()
  Dim bolTemp As Boolean
  
  If objCheckNumeric.Value = vbChecked Then
    bolUseNumeric = True
  Else
    bolUseNumeric = False
  End If
  If objCheckSeqRen.Value = vbChecked Then
    bolSeqRename = True
  Else
    bolSeqRename = False
  End If
  
  lCharset = objComboCharset.ListIndex
  lDtbType = objComboDTBType.ListIndex
  lIANACharset = objComboIanaCS.ListIndex
  bolPreserveMeta = objRadioMetaPres.Value
  bolSameFolder = objRadioSameFolder.Value
  bolMoveBook = fncCheck2Bol(objCheckMoveBook) Or (Not bolSameFolder)
  sSavePath = objTextFoldername.Text
  sMetaFile = objTextMetaFile.Text
  sPrefix = objTextPrefix.Text

  sDefaultSavePath = objTextDefaultSavepath.Text
  sDefaultMetaPath = objTextDefaultMetaPath.Text
  
  If objCheckHalt.Value = vbChecked Then
    bolHalt = True
  Else
    bolHalt = False
  End If
  

'  If objCheckLaunchProgram.Value = vbChecked Then
'    bolLaunchProgram = True
'  Else
'    bolLaunchProgram = False
'  End If
'  sLaunchProgram = fncStripIdAddPath(objTextLaunchProgram.Text, sAppPath)
  
  If objCheckPb2kLayoutFix.Value = vbChecked Then
    bolPb2kLayoutFix = True
  Else
    bolPb2kLayoutFix = False
  End If
  
  If objCheckAddCss.Value = vbChecked Then
    bolAddCss = True
  Else
    bolAddCss = False
  End If
  
  
  If objCheckFixPar.Value = vbChecked Then
    bolFixPar = True
  Else
    bolFixPar = False
  End If
  
  If objCheckMangleLinks.Value = vbChecked Then
    bolRebuildLinkStructure = True
  Else
    bolRebuildLinkStructure = False
  End If
  
  If objCheckDisableXhtmlLinks.Value = vbChecked Then
    bolDisableBrokenXhtmlLinks = True
  Else
    bolDisableBrokenXhtmlLinks = False
  End If
    
  If objCheckEstimateXhtmlLinks.Value = vbChecked Then
    bolEstimateBrokenXhtmlLinks = True
  Else
    bolEstimateBrokenXhtmlLinks = False
  End If
  
  If objCheckMakeTrueNccOnly.Value = vbChecked Then
    bolMakeTrueNccOnly = True
  Else
    bolMakeTrueNccOnly = False
  End If
  
'  If objCheckPointTargetsToPar.Value = vbChecked Then
'    bolPointTargetsToPar = True
'  Else
'    bolPointTargetsToPar = False
'  End If zzz
  
  If objRadioSmilTargetNoChange.Value = True Then
    lSmilTarget = SMILTARGET_NOCHANGE
  ElseIf objRadioSmilTargetPar.Value = True Then
    lSmilTarget = SMILTARGET_PAR
  Else
    lSmilTarget = SMILTARGET_TEXT
  End If
  
  If objCheckVerboseLog.Value = vbChecked Then
    bolDoVerboseLog = True
  Else
    bolDoVerboseLog = False
  End If

  If objCheckMergeShort.Value = vbChecked Then
    bolMergeShortPhrases = True
  Else
    bolMergeShortPhrases = False
  End If
    
'  objRegeneratorUserControl.fncSetMergeShort bolMergeShortPhrases
'  bolMergeShortPhrases = objRegeneratorUserControl.fncGetMergeShort


  lClipLessThan = objSliderMergeIfLower.Value
'  lClipShort = objSliderMergeIfLower.Value
'  objRegeneratorUserControl.fncSetClipLessThan lClipShort
'  lClipShort = objRegeneratorUserControl.fncGetClipLessThan
  
   lFirstClipLessThan = objSliderMergeIfShorter2.Value
'  lClipShort2 = objSliderMergeIfShorter2.Value
'  objRegeneratorUserControl.fncSetFirstClipLessThan lClipShort2
'  lClipShort2 = objRegeneratorUserControl.fncGetFirstClipLessThan
  
  lNextClipLessThan = objSliderMergeAndNextIsShorter.Value
'  lNextShort = objSliderMergeAndNextIsShorter.Value
'  objRegeneratorUserControl.fncSetNextClipLessThan lNextShort
'  lNextShort = objRegeneratorUserControl.fncGetNextClipLessThan
  
  'bug - fix mg 20030215:
  'lClipSpan = objSliderMergeAndNextIsShorter.Value
  lClipSpan = objSliderClipEndBeginSpan.Value
'  objRegeneratorUserControl.fncSetClipSpan lClipSpan
'  lClipSpan = objRegeneratorUserControl.fncGetClipSpan
  
  If objCheckValidateJob.Value = vbChecked Then
    bolUseValidator = True
  Else
    bolUseValidator = False
  End If
  
  If objCheckIncludeNCErrors.Value = vbChecked Then
    bolIncludeNCErrors = True
  Else
    bolIncludeNCErrors = False
  End If
  
  If objCheckIncludeWarnings.Value = vbChecked Then
    bolIncludeWarnings = True
  Else
    bolIncludeWarnings = False
  End If
  
  If objCheckIncludeAdvancedADTD.Value = vbChecked Then
    bolIncludeADVADTD = True
  Else
    bolIncludeADVADTD = False
  End If
  
  If objCheckCreateStandalone.Value = vbChecked Then
    bolCreateStandalone = True
  Else
    bolCreateStandalone = False
  End If
  
  sStandalonePath = objTextStandalonePath.Text
  
  If objCheckValLightMode.Value = vbChecked Then
    bolValidatorLightMode = True
  Else
    bolValidatorLightMode = False
  End If
  'mg20030325:
  If Not objValidatorUserControl Is Nothing Then
    objValidatorUserControl.fncSetLightMode bolValidatorLightMode
  End If
  
  If objCheckSyncWValidator.Value = vbChecked Then
    bolSyncWValidator = True
    
    sBaseKey = "Software\DaisyWare\Validator\Misc"
    fncLoadRegistryData "Dtd_AdtdPath", sExtPath, HKEY_CURRENT_USER, , ""
    fncLoadRegistryData "AppPath", sVTMPath, HKEY_CURRENT_USER, , ""
    fncLoadRegistryData "TempPath", sTempPath, HKEY_CURRENT_USER, , ""
    sBaseKey = "Software\DaisyWare\Validator\Settings"

    fncLoadRegistryData "TimeFluctuation", lTimeFluctuation, HKEY_CURRENT_USER, , 0
  Else
    bolSyncWValidator = False
    lTimeFluctuation = objSliderTimeFluctuation.Value
    sExtPath = fncStripIdAddPath(objTextExtPath.Text, sAppPath)
    sVTMPath = fncStripIdAddPath(objTextVTMPath.Text, sAppPath)
    sTempPath = fncStripIdAddPath(objTextTempPath.Text, sAppPath)
  End If

'  mg removed this feature 20030202
'  If objRadioSaveXHTML.Value = True Then bolSaveXHTML = True Else _
'    bolSaveXHTML = False

  sLogPath = objTextSaveLogPath.Text

  subRefreshInterface
End Function

' This function updates the settings and job properties in the UI from the currently
' saved settings and the selected jobs properties. It also disables / enables controls
' depending on the settings made.

Public Sub subRefreshInterface()
  Dim bolJobTabEnabled As Boolean
  Dim bolFrameJobs As Boolean, bolFrameJobProp As Boolean
  Dim bolFrameSettings As Boolean, bolFrameValidator As Boolean
  Dim bolFrameLog As Boolean, bolFrameDebug As Boolean

' Select tab

  Select Case objTabstrip.SelectedItem.Key
    Case "tab1"
      bolFrameJobs = True
    Case "tab2"
      bolFrameJobProp = True
    Case "tab3"
      bolFrameSettings = True
    Case "tab4"
      bolFrameValidator = True
    Case "tab5"
      bolFrameLog = True
  End Select
  objFrameJobs.Visible = bolFrameJobs
  objFrameJobProperties.Visible = bolFrameJobProp
  objFrameSettings.Visible = bolFrameSettings
  objFrameValidator.Visible = bolFrameValidator
  objFrameLog.Visible = bolFrameLog
    
' Disable job tab if no job selected

  If objJobList.SelectedItem Is Nothing Then
    bolJobTabEnabled = False
    objFrameJobProperties.Caption = "Properties for job"
  Else
    bolJobTabEnabled = True
    objFrameJobProperties.Caption = "Properties for job @ " & _
      objJobList.SelectedItem.Text
  End If
  
' Update screen information
  bolOwnChange = True

  objComboDTBType.ListIndex = lDtbType
  objComboCharset.ListIndex = lCharset
  If Not lIANACharset = 0 Then objComboIanaCS.ListIndex = lIANACharset
  objRadioMetaPres.Value = bolPreserveMeta
  objRadioMetaImp.Value = Not bolPreserveMeta
  objTextMetaFile.Text = sMetaFile
  If bolSeqRename Then objCheckSeqRen.Value = vbChecked Else _
    objCheckSeqRen.Value = vbUnchecked
  If bolUseNumeric Then objCheckNumeric.Value = vbChecked Else _
    objCheckNumeric.Value = vbUnchecked
  objRadioSameFolder.Value = bolSameFolder
  objRadioNewFolder.Value = Not bolSameFolder
  objTextPrefix.Text = sPrefix
  
  If objRadioNewFolder Then bolMoveBook = False
  
  If bolMoveBook Then
    objCheckMoveBook.Value = vbChecked
  Else
   objCheckMoveBook.Value = vbUnchecked
  End If
  
  objTextFoldername.Text = sSavePath

  objTextDefaultSavepath.Text = sDefaultSavePath
  objTextDefaultMetaPath.Text = sDefaultMetaPath
    
  If bolHalt Then
    objCheckHalt.Value = vbChecked
  Else
    objCheckHalt.Value = vbUnchecked
  End If
    
  If bolPb2kLayoutFix Then
    objCheckPb2kLayoutFix.Value = vbChecked
  Else
    objCheckPb2kLayoutFix.Value = vbUnchecked
  End If
    
  If bolFixPar Then
    objCheckFixPar.Value = vbChecked
  Else
    objCheckFixPar.Value = vbUnchecked
  End If
  
  If bolRebuildLinkStructure Then
    objCheckMangleLinks.Value = vbChecked
  Else
    objCheckMangleLinks.Value = vbUnchecked
  End If
  
  If bolDisableBrokenXhtmlLinks Then
    objCheckDisableXhtmlLinks.Value = vbChecked
  Else
    objCheckDisableXhtmlLinks.Value = vbUnchecked
  End If
    
  If bolEstimateBrokenXhtmlLinks Then
    objCheckEstimateXhtmlLinks.Value = vbChecked
  Else
    objCheckEstimateXhtmlLinks.Value = vbUnchecked
  End If
    
  objSliderMergeIfLower.Value = lClipLessThan
  objTextMergeIfLower.Text = lClipLessThan
  objSliderMergeIfShorter2.Value = lFirstClipLessThan
  objTextMergeIfShorter2.Text = lFirstClipLessThan
  objSliderMergeAndNextIsShorter.Value = lNextClipLessThan
  objTextMergeAndNextIsShorter.Text = lNextClipLessThan
  objSliderClipEndBeginSpan.Value = lClipSpan
  objTextClipEndBeginSpan.Text = lClipSpan

  If bolMergeShortPhrases Then
    objCheckMergeShort.Value = vbChecked
  Else
    objCheckMergeShort.Value = vbUnchecked
  End If
  
  If bolUseValidator Then
    objCheckValidateJob.Value = vbChecked
  Else
    objCheckValidateJob.Value = vbUnchecked
  End If
  
  If bolIncludeNCErrors Then
    objCheckIncludeNCErrors.Value = vbChecked
  Else
    objCheckIncludeNCErrors.Value = vbUnchecked
  End If
  
  If bolIncludeWarnings Then
    objCheckIncludeWarnings.Value = vbChecked
  Else
    objCheckIncludeWarnings.Value = vbUnchecked
  End If
  
  If bolIncludeADVADTD Then
    objCheckIncludeAdvancedADTD.Value = vbChecked
  Else
    objCheckIncludeAdvancedADTD.Value = vbUnchecked
  End If
  
  If bolCreateStandalone Then
    objCheckCreateStandalone.Value = vbChecked
  Else
    objCheckCreateStandalone.Value = vbUnchecked
  End If
  
  objTextStandalonePath.Text = sStandalonePath
  
  If bolSyncWValidator Then
    objCheckSyncWValidator.Value = vbChecked
  Else
    objCheckSyncWValidator.Value = vbUnchecked
  End If
  
  If bolValidatorLightMode Then
    objCheckValLightMode.Value = vbChecked
  Else
    objCheckValLightMode.Value = vbUnchecked
  End If
  
  objSliderTimeFluctuation.Value = lTimeFluctuation
  objTextTimeFluctuation.Text = CStr(lTimeFluctuation)

  objTextExtPath.Text = sExtPath
  objTextVTMPath.Text = sVTMPath
  objTextTempPath.Text = sTempPath

  
  If bolDoVerboseLog Then
    objCheckVerboseLog.Value = vbChecked
  Else
    objCheckVerboseLog.Value = vbUnchecked
  End If
  
  If bolAddCss Then
    objCheckAddCss.Value = vbChecked
  Else
    objCheckAddCss.Value = vbUnchecked
  End If
  
  If bolMakeTrueNccOnly Then
    objCheckMakeTrueNccOnly.Value = vbChecked
  Else
    objCheckMakeTrueNccOnly.Value = vbUnchecked
  End If
  
'  If bolPointTargetsToPar Then
'    objCheckPointTargetsToPar.Value = vbChecked
'  Else
'    objCheckPointTargetsToPar.Value = vbUnchecked
'  End If zzz
          
  If lSmilTarget = SMILTARGET_NOCHANGE Then
    objRadioSmilTargetNoChange.Value = True
  ElseIf lSmilTarget = SMILTARGET_PAR Then
    objRadioSmilTargetPar.Value = True
  Else 'lSmilTarget = SMILTARGET_TEXT
    objRadioSmilTargetText.Value = True
  End If
            
  bolOwnChange = False
  
' Disable/Enable parts of the interface

  If bolRegenerating Then bolJobTabEnabled = False
    
' *** General interface
    
  objCmdAdd.Enabled = Not bolRegenerating
  objCmdAddJoblist.Enabled = Not bolRegenerating
  objCmdStop.Enabled = bolRegenerating
  
  objCmdRemove.Enabled = bolJobTabEnabled
  objCmdRemoveAll.Enabled = bolJobTabEnabled
  objCmdRun.Enabled = bolJobTabEnabled

' *** Jobs tab

  objJobList.Enabled = True

' *** Job properties tab

  objCmdFolderBrws.Enabled = bolJobTabEnabled
  objCmdMetaBrws.Enabled = bolJobTabEnabled
  objCmdRestore.Enabled = bolJobTabEnabled
  objTextFoldername.Enabled = bolJobTabEnabled
  objTextMetaFile.Enabled = bolJobTabEnabled
  objTextPrefix.Enabled = bolJobTabEnabled
  objComboCharset.Enabled = bolJobTabEnabled
  objComboDTBType.Enabled = bolJobTabEnabled
  objComboIanaCS.Enabled = bolJobTabEnabled
  objCheckNumeric.Enabled = bolJobTabEnabled
  objCheckSeqRen.Enabled = bolJobTabEnabled
  objCheckMoveBook.Enabled = bolJobTabEnabled
  objRadioMetaImp.Enabled = bolJobTabEnabled
  objRadioMetaPres.Enabled = bolJobTabEnabled
  objRadioNewFolder.Enabled = bolJobTabEnabled
  objRadioSameFolder.Enabled = bolJobTabEnabled
  
  objCmdSetAll.Enabled = bolJobTabEnabled
  
  objLabelDTBType.Enabled = bolJobTabEnabled
  objLabelCharset.Enabled = bolJobTabEnabled
  objLabelPrefix.Enabled = bolJobTabEnabled
    
  If (Not bolRegenerating) And (bolJobTabEnabled) Then
    
    If objComboCharset.Text = "other" Then
      objComboIanaCS.Enabled = True
    Else
      objComboIanaCS.Enabled = False
    End If
    
    objTextMetaFile.Enabled = objRadioMetaImp.Value
    objCmdMetaBrws.Enabled = objTextMetaFile.Enabled
  
    If (objCheckSeqRen.Value = vbUnchecked) Then
      objCheckNumeric.Enabled = False
      objTextPrefix.Enabled = False
    Else
      objCheckNumeric.Enabled = True
'      If (objCheckNumeric.Value = vbUnchecked) Then _
        objTextPrefix.Enabled = True Else objTextPrefix.Enabled = False
      If (objCheckNumeric.Value = vbUnchecked) Then
        objTextPrefix.Enabled = True
        objLabelPrefix.Enabled = True
      Else
        objTextPrefix.Enabled = False
        objLabelPrefix.Enabled = False
      End If
    End If
  
    objCheckMoveBook.Enabled = objRadioSameFolder.Value
    'mg20030221
    'objTextFoldername.Enabled = fncCheck2Bol(objCheckMoveBook)
    objTextFoldername.Enabled = fncCheck2Bol(objCheckMoveBook) Or objRadioNewFolder.Value
    objCmdFolderBrws.Enabled = objTextFoldername.Enabled
  End If 'If (Not bolRegenerating) And (bolJobTabEnabled)
  
' *** Validation settings tab

  objCheckValidateJob.Enabled = Not bolRegenerating
  
  objCheckIncludeNCErrors.Enabled = bolUseValidator And (Not bolRegenerating)
  objCheckIncludeWarnings.Enabled = bolUseValidator And (Not bolRegenerating)
  objCheckIncludeAdvancedADTD.Enabled = bolUseValidator And (Not bolRegenerating)
  objCheckCreateStandalone.Enabled = bolUseValidator And (Not bolRegenerating)
  objTextStandalonePath.Enabled = bolUseValidator And bolCreateStandalone And (Not bolRegenerating)
  objCmdValReportFolderBrws.Enabled = bolUseValidator And bolCreateStandalone And (Not bolRegenerating)
  
  objCheckSyncWValidator.Enabled = bolUseValidator And (Not bolRegenerating)
  objCheckValLightMode.Enabled = bolUseValidator And (Not bolRegenerating)
  objSliderTimeFluctuation.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objTextTimeFluctuation.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  
  objCmdValExtFolderBrws.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objTextExtPath.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objCmdValVtmFolderBrws.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objTextVTMPath.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objCmdValTmpFolderBrws.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objTextTempPath.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  
  objLabelStandalonePath.Enabled = bolUseValidator And (Not bolRegenerating)
  objLabelVTMPath.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objLabelTempPath.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objLabelExtPath.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objLabelTimeFluct.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objLabelTimeFLuctMax.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  objLabelTimeFluctMin.Enabled = bolUseValidator And (Not bolSyncWValidator) And (Not bolRegenerating)
  
' *** Advanced settings tab
  objCheckPb2kLayoutFix.Enabled = Not bolRegenerating
  objCheckAddCss.Enabled = Not bolRegenerating
  objCheckHalt.Enabled = Not bolRegenerating
  
  objCheckFixPar.Enabled = Not bolRegenerating
  objCheckMangleLinks.Enabled = Not bolRegenerating
  objCheckDisableXhtmlLinks.Enabled = Not bolRegenerating
  objCheckEstimateXhtmlLinks.Enabled = bolDisableBrokenXhtmlLinks And (Not bolRegenerating)
  objCheckMakeTrueNccOnly.Enabled = Not bolRegenerating
'  objCheckPointTargetsToPar.Enabled = Not bolRegenerating
  objRadioSmilTargetNoChange.Enabled = Not bolRegenerating
  objRadioSmilTargetPar.Enabled = Not bolRegenerating
  objRadioSmilTargetText.Enabled = Not bolRegenerating
    
  objCheckMergeShort.Enabled = Not bolRegenerating
    
  objSliderMergeIfLower.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objSliderMergeIfShorter2.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objSliderMergeAndNextIsShorter.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objSliderClipEndBeginSpan.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objTextMergeIfLower.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objTextMergeIfShorter2.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objTextMergeAndNextIsShorter.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objTextClipEndBeginSpan.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  
  objLabelMergeAndNextIsShorter.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objLabelMergeIfLower.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objLabelMergeIfShorter2.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  objLabelClipEndBeginSpan.Enabled = bolMergeShortPhrases And (Not bolRegenerating)
  
  objTextSaveLogPath.Text = sLogPath
  
' Update info from the current job item if any
  If bolFrameJobs Then
    objTextLog.Text = objJobsLog.sLog
    If (lCurrentJob <= lJobCount And lCurrentJob > 0) Then
      objTextLog.Text = objTextLog.Text & aJobItems(lCurrentJob).objLog.sLog
'      If (Not bolRegenerating) Then fncAddReport aJobItems(lCurrentJob).objReport
    End If
  End If
End Sub

' This function adds a report object to the log textbox
Public Function fncAddReport(objReport As Object) As Boolean
  Dim lCounter As Long, objReportItem As Object
  
  If objReport Is Nothing Then Exit Function
  
  For lCounter = 0 To objReport.lFailedTestCount - 1
    objReport.fncRetrieveFailedTestItem lCounter, objReportItem
    
  ' This IF statement determins if we're going to show this error in the UI or not
    If (Not (objReportItem.sFailType = "warning" And _
        bolIncludeWarnings = False)) And _
       (Not (objReportItem.sFailType = "error" And _
        objReportItem.sFailClass = "non-critical" And _
        bolIncludeNCErrors = False)) Then
    
      objTextLog.Text = objTextLog.Text & _
        objReportItem.sFailType & " @ " & objReportItem.sAbsPath & " [" & _
        objReportItem.lLine & ":" & objReportItem.lColumn & "]" & " " & _
        objReportItem.sShortDesc & ", " & objReportItem.sComment & vbCrLf & vbCrLf
    End If
  Next lCounter
End Function

' This function updates the joblist
Public Function fncUpdateJobList()
  Dim lCounter As Long, objListItem As ListItem, lIcon As Long
  Dim sRegResult As String, sValResult As String
  
  Do Until objJobList.ListItems.Count = 0
    objJobList.ListItems.Remove (1)
  Loop
  
  For lCounter = 1 To lJobCount
' Select an icon depending on what status the job has
    If aJobItems(lCounter).bolRegRun And _
      (((Not aJobItems(lCounter).sErrorType = "") Or _
      (Not aJobItems(lCounter).bolRendered)) And _
      (Not aJobItems(lCounter).bolValResult)) Then
      If aJobItems(lCounter).bolRegResult And aJobItems(lCounter).bolValResult Then
        lIcon = 2
      Else
        If (Not aJobItems(lCounter).bolRegResult) Or _
          (Not aJobItems(lCounter).bolRendered) Then
          lIcon = 1
        Else
          If aJobItems(lCounter).sErrorType = "warning" Then lIcon = 4 Else lIcon = 1
        End If
      End If
    Else
      lIcon = 3
    End If

' Add the listitem and show the path in the first column
    Set objListItem = objJobList.ListItems.Add(, , aJobItems(lCounter).sPath, _
      lIcon, lIcon)

' Add some info about the progress of this job in the other two columns
    If aJobItems(lCounter).bolRegRun Then
      If aJobItems(lCounter).bolRegResult Then sRegResult = "Pass" Else sRegResult = "Fail"
      If Not aJobItems(lCounter).bolValResult Then
        If Not aJobItems(lCounter).bolRendered Then sValResult = "not rendered"
        Select Case aJobItems(lCounter).sErrorType
          Case "error"
            sValResult = aJobItems(lCounter).sErrorClass & " " & _
              aJobItems(lCounter).sErrorType
          Case "warning"
            sValResult = aJobItems(lCounter).sErrorType
        End Select
      Else
        sValResult = "Pass"
      End If
    Else
      sRegResult = "Not run"
      sValResult = "Not run"
    End If
    
    objListItem.ListSubItems.Add , , sRegResult
    objListItem.ListSubItems.Add , , sValResult
  Next lCounter

' Select the current job to update it's screen information
  If (Not lJobCount = 0) And (lCurrentJob <= lJobCount) Then
    objJobList.ListItems.Item(lCurrentJob).Selected = True
    objJobList_Click
  End If
  
  objLabelBatch.Caption = "Batch list (" & lJobCount & " jobs)"
End Function

' This is a common function for using the common dialog control to open file
Private Function fncOpenFile(sMask As String, sFileName As String, _
  bolMustExist, sOutPut As String) As Boolean
  
  On Error GoTo ErrorH
  
  With objCommonDialog
    .CancelError = True
    .Filter = sMask
    .FilterIndex = 1
    .FileName = sFileName
    If bolMustExist Then .Flags = cdlOFNFileMustExist
    .Flags = cdlOFNNoChangeDir
    .ShowOpen
    sOutPut = .FileName
  End With
   
  fncOpenFile = True
ErrorH:
End Function

Public Property Let bolBusy(ibolBusy As Boolean)
  bolprivBusy = ibolBusy
  If bolprivBusy Then MousePointer = 11 Else MousePointer = 0
  DoEvents
End Property

Public Property Get bolBusy() As Boolean
  bolBusy = bolprivBusy
End Property

'************************************************
'******************* M E N U ********************
'************************************************

Private Sub mnuRunBatch_Click()
  objCmdRun_Click
End Sub

Private Sub mnuStopBatch_Click()
  objCmdStop_Click
End Sub

Private Sub mnuAddJob_Click()
  objCmdAdd_Click
End Sub

Private Sub mnuAddJobList_Click()
  objCmdAddJoblist_Click
End Sub

Private Sub mnuRemoveJob_Click()
  objCmdRemove_Click
End Sub

Private Sub mnuRemoveAll_Click()
  objCmdRemoveAll_Click
End Sub

Private Sub mnuJobsTabFocus_Click()
  objTabstrip.Tabs(1).Selected = True
End Sub

Private Sub mnuJobPropTabFocus_Click()
  objTabstrip.Tabs(2).Selected = True
End Sub

Private Sub mnuValSettingsTabFocus_Click()
  objTabstrip.Tabs(3).Selected = True
End Sub

Private Sub mnuAdvancedSettingsTabFocus_Click()
  objTabstrip.Tabs(4).Selected = True
End Sub

Private Sub LogSettingsTabFocus_Click()
  objTabstrip.Tabs(5).Selected = True
End Sub

Private Sub mnuJobWindowFocus_Click()
  objTabstrip.Tabs(1).Selected = True
  objJobList.SetFocus
End Sub

Private Sub mnuLogWinFocus_Click()
  objTabstrip.Tabs(1).Selected = True
  objTextLog.SetFocus
End Sub
