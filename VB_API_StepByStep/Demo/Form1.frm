VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   3570
   ClientTop       =   3960
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6585
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3375
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'initialize listview header
Call init_listView
'insert default records
Call insert_DefaultRec
End Sub

Private Sub init_listView()
With lv
  Call .ColumnHeaders.Add(, , "ID", 500)      '0
  Call .ColumnHeaders.Add(, , "Name", 1000)    '1
  Call .ColumnHeaders.Add(, , "Address", 3000) '2
End With
End Sub

Private Sub insert_DefaultRec()
Dim itemRetn As ListItem

With lv
  Set itemRetn = .ListItems.Add(, "Emp" & 1, "Emp" & 1)
      itemRetn.SubItems(1) = "Brandon"
      itemRetn.SubItems(2) = "SS24/19 Bangsar"
  Set itemRetn = Nothing

  Set itemRetn = .ListItems.Add(, "Emp" & 2, "Emp" & 2)
      itemRetn.SubItems(1) = "Danny Bunny"
      itemRetn.SubItems(2) = "Singapore"
  Set itemRetn = Nothing
  
  Set itemRetn = .ListItems.Add(, "Emp" & 3, "Emp" & 3)
      itemRetn.SubItems(1) = "Marie"
      itemRetn.SubItems(2) = "2033 Fordburg Gauteng"
  Set itemRetn = Nothing
  
  Set itemRetn = .ListItems.Add(, "Emp" & 4, "Emp" & 4)
      itemRetn.SubItems(1) = "Denise"
      itemRetn.SubItems(2) = "Jalan Cantoment, 10350 Penang"
  Set itemRetn = Nothing
  
End With

End Sub


Private Sub lv_DblClick()
Dim curItemIndex As Long
Dim curSubItemIndex As Long
Dim lvi As LVITEM
Dim i As Integer

curItemIndex = lv.SelectedItem.Index - 1   'since it is zero-based, thus have to decrement by one
curSubItemIndex = 2   'we would like to modify the address column

lvi.mask = LVIF_TEXT
lvi.iItem = curItemIndex
lvi.iSubItem = curSubItemIndex
lvi.pszText = Text1.Text

If ListView_SetItem(lv.hWnd, lvi) Then MsgBox "Successful!", vbInformation

End Sub

