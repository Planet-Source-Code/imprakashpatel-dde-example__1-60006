VERSION 5.00
Begin VB.Form DDE 
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get Data From Excel"
      Height          =   510
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Manuall"
      Top             =   765
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   465
      HideSelection   =   0   'False
      Left            =   1365
      LinkItem        =   "R1C1"
      LinkTopic       =   "Excel|Sheet1"
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   $"DDE.frx":0000
      Top             =   1590
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "'open one excel file and Type Something in first cell and then click button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   540
      TabIndex        =   4
      Top             =   90
      Width           =   3660
   End
   Begin VB.Label Label1 
      Caption         =   "A1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   285
      TabIndex        =   2
      Top             =   750
      Width           =   945
   End
End
Attribute VB_Name = "DDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Copyright (c) 2005 by Prakah Patel
'*
'* This software is the proprietary information of Pd Systems.
'* Use is subject to license terms.
'*
'* @author  Prakash Patel
'* @version 1.0
'* @date    31 March 2004
'*
'***************************************************************************


'***************************************************************************
'open one excel file and Type Something in first cell
'***************************************************************************

Private Sub Command1_Click()
   Dim CurRow As String
   Static Row   ' Worksheet row number.
   Row = Row + 1   ' Increment Row.
   If Row = 1 Then   ' First time only.
      ' Make sure the link isn't active.
      Text1.LinkMode = 0
      ' Set the application name and topic name.
      Text1.LinkTopic = "Excel|Sheet1"
      Text1.LinkItem = "R1C1"   ' Set LinkItem.
      Text1.LinkMode = 1   ' Set LinkMode to Automatic.
   Else
      ' Update the row in the data item.
      CurRow = "R" & Row & "C1"
      Text1.LinkItem = CurRow   ' Set LinkItem.
   End If
End Sub

