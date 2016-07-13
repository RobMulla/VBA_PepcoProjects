Option Compare Database

Public WR_NO As String
Public SiteStreetNo As String
Public SiteStreetName As String
Public SiteStreetType As String
Public SiteCity As String
Public SiteState As String
Public SiteZip As String

Public CustName As String

Public BillAmmount As Currency
Public BillAddress As String
Public BillCity As String
Public BillState As String
Public BillZip As String
Public BillName As String

Public WRtoQuery As String

Public SAName As String
Public SADept As String
Public SAPhone As String
Public SASignaturePath As String

Public Path As String

Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

   DoCmd.MoveSize 200, 200

   DoCmd.OpenForm "frmLogoutTimer", , , , , acHidden
   'DoCmd.OpenForm "frmLogoutTimer", , , , , acWindowNormal
   
Exit_Here:
   Exit Sub
Err_Handler:
   MsgBox Err.Description, vbExclamation, "Error"
   Resume Exit_Here
End Sub


Private Sub CostLetterV2_Click()

Call InputWR '' Calls function for input box for WR#

''make sure they put a wr #
If WRtoQuery = "" Then
    MsgBox "You Must enter a Work Request Number"
    Exit Sub
End If

''make sure service associate is selected
If IsNull(SASelector.Value) Then
    MsgBox "Please select a Service Associate Name"
    Exit Sub
End If

Call PullSAInfo '' Pull info on SAName
Call PullWRData ''Calls function that prompts for WR# and pulls values from WMIS

Path = "Z:\PDelivery\WAS\Engineering_Design\4405_Engineering_Pepco\ServiceAssociateToolBox\Letters\CostLetter.docx"

Call fillCostLetter ''Calls function that opens word and fills out cost letter

End Sub

Private Sub CostLetterWord_Click()

Call InputWR '' Calls function for input box for WR#

''make sure they put a wr #
If WRtoQuery = "" Then
    MsgBox "You Must enter a Work Request Number"
    Exit Sub
End If

''make sure service associate is selected
If IsNull(SASelector.Value) Then
    MsgBox "Please select a Service Associate Name"
    Exit Sub
End If

Call PullSAInfo '' Pull info on SAName
Call PullWRData ''Calls function that prompts for WR# and pulls values from WMIS

Path = "Z:\PDelivery\WAS\Engineering_Design\4405_Engineering_Pepco\ServiceAssociateToolBox\Letters\CostLetter2.docx"

Call fillCostLetter ''Calls function that opens word and fills out cost letter

End Sub

Function PullSAInfo()


SAName = SASelector.Value
SADept = DLookup("[SADepartment]", "SAContactInfo", "SAName='" & SASelector.Value & "'")
SAPhone = DLookup("[SAPhone]", "SAContactInfo", "SAName='" & SASelector.Value & "'")
SASignaturePath = DLookup("[SignaturePath]", "SAContactInfo", "SAName='" & SASelector.Value & "'")

End Function


Private Sub HighwayLetter_Click()
Call InputWR '' Calls function for input box for WR#

''make sure they put a wr #
If WRtoQuery = "" Then
    MsgBox "You Must enter a Work Request Number"
    Exit Sub
End If

''make sure service associate is selected
If IsNull(SASelector.Value) Then
    MsgBox "Please select a Service Associate Name"
    Exit Sub
End If


Call PullSAInfo '' Pull info on SAName
Call PullWRData ''Calls function that prompts for WR# and pulls values from WMIS

Path = "Z:\PDelivery\WAS\Engineering_Design\4405_Engineering_Pepco\ServiceAssociateToolBox\Letters\HighwayCostLetter.docx"


Call fillCostLetter ''Calls function that opens word and fills out cost letter
End Sub

Private Sub RazeLetterWord_Click()

Call InputWR '' Calls function for input box for WR#

''make sure they put a wr #
If WRtoQuery = "" Then
    MsgBox "You Must enter a Work Request Number"
    Exit Sub
End If

''make sure service associate is selected
If IsNull(SASelector.Value) Then
    MsgBox "Please select a Service Associate Name"
    Exit Sub
End If

Call PullSAInfo '' Pull info on SAName
Call PullWRData ''Calls function that prompts for WR# and pulls values from WMIS

Path = "Z:\PDelivery\WAS\Engineering_Design\4405_Engineering_Pepco\ServiceAssociateToolBox\Letters\DemoLetter2.docx"

Call fillRazeLetter ''Calls function that opens word and fills out raze letter

End Sub



Function fillRazeLetter()
Dim appword As Word.Application
Dim doc As Word.Document
TodayDate = Format(Now(), "mmmm dd, yyyy")
On Error Resume Next
Error.Clear

Set appword = GetObject(, "word.application")

If Err.Number <> 0 Then
    Set appword = New Word.Application
    appword.Visible = True
End If



Set doc = appword.Documents.Open(Path, , True)

With doc
    .FormFields("Date").Result = TodayDate
    .FormFields("BillName").Result = BillName
    .FormFields("BillAmmount").Result = BillAmmount
    .FormFields("BillAddress").Result = BillAddress
    .FormFields("BillAmmount").Result = BillAmmount
    .FormFields("BillCity").Result = BillCity
    .FormFields("BillState").Result = BillState
    .FormFields("BillZip").Result = BillZip
    
    
    .FormFields("SiteZip").Result = SiteZip
    .FormFields("SiteState").Result = SiteState
    .FormFields("SiteCity").Result = SiteCity
    .FormFields("SiteStreetType").Result = SiteStreetType
    .FormFields("SiteStreetName").Result = SiteStreetName
    .FormFields("SiteStreetNo").Result = SiteStreetNo
    .FormFields("BillName2").Result = BillName
    .FormFields("WorkRequest").Result = WR_NO
    
    .FormFields("CustName").Result = CustName

    .FormFields("SAName").Result = SAName
    .FormFields("SADeptartment").Result = SADept
    .FormFields("SAPhone").Result = SAPhone
    
        Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    doc.Bookmarks("Signature").Range.InlineShapes.AddPicture FileName:= _
        SASignaturePath, _
        LinkToFile:=False, SaveWithDocument:=True
     
        '' end signature stuff
    '' Title the document
    doc.SaveAs2 FileName:=("WR") + WR_NO + ("-") + SiteStreetNo + ("_") + SiteStreetName + ("_") + SiteStreetType + ("-") + ("Demo"), FileFormat:=17

    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=1 ''go to first line


End With

Set doc = Nothing
Set appword = Nothing

End Function
Function fillCostLetter()
Dim appword As Word.Application
Dim doc As Word.Document
TodayDate = Format(Now(), "mmmm dd, yyyy")
On Error Resume Next
Error.Clear

Set appword = GetObject(, "word.application")

If Err.Number <> 0 Then
    Set appword = New Word.Application
    appword.Visible = True
End If


Set doc = appword.Documents.Open(Path, , True)


With doc
    .FormFields("Date").Result = TodayDate
    .FormFields("BillName").Result = BillName
    .FormFields("BillAmmount").Result = BillAmmount
    .FormFields("BillAddress").Result = BillAddress
    .FormFields("BillAmmount").Result = BillAmmount
    .FormFields("BillCity").Result = BillCity
    .FormFields("BillState").Result = BillState
    .FormFields("BillZip").Result = BillZip
    
    
    .FormFields("SiteZip").Result = SiteZip
    .FormFields("SiteState").Result = SiteState
    .FormFields("SiteCity").Result = SiteCity
    .FormFields("SiteStreetType").Result = SiteStreetType
    .FormFields("SiteStreetName").Result = SiteStreetName
    .FormFields("SiteStreetNo").Result = SiteStreetNo
    .FormFields("BillName2").Result = BillName
    .FormFields("WorkRequest").Result = WR_NO
    
    .FormFields("CustName").Result = CustName

    .FormFields("SAName").Result = SAName
    .FormFields("SADeptartment").Result = SADept
    .FormFields("SAPhone").Result = SAPhone

    .FormFields("WorkRequest2").Result = WR_NO
    .FormFields("WorkRequest3").Result = WR_NO
    
    .FormFields("SiteZip2").Result = SiteZip
    .FormFields("SiteState2").Result = SiteState
    .FormFields("SiteCity2").Result = SiteCity
    .FormFields("SiteStreetType2").Result = SiteStreetType
    .FormFields("SiteStreetName2").Result = SiteStreetName
    .FormFields("SiteStreetNo2").Result = SiteStreetNo


'' Add signature
   '' appword.ActiveDocument.Bookmarks("Signature").Select

   '' Selection.Goto What:=wdGoToBookmark, Name:="Signature"
   '' objWork.ActiveDocument.Bookmarks("Signature").Range.Select
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    doc.Bookmarks("Signature").Range.InlineShapes.AddPicture FileName:= _
        SASignaturePath, _
        LinkToFile:=False, SaveWithDocument:=True
        
        '' end signature stuff

    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=1 ''go to first line

    '' Title the document
    doc.SaveAs2 FileName:=("WR") + WR_NO + ("-") + SiteStreetNo + ("_") + SiteStreetName + ("_") + SiteStreetType + ("-") + ("CostLetter")


End With


appword.Visible = True
appword.Activate


'' clear stuff

Set doc = Nothing
Set appword = Nothing
    

    
End Function

Function InputWR()
WRtoQuery = InputBox(Prompt:="Enter Work Request Number", Title:="WR Number", Default:=WRtoQuery)
End Function


Function PullWRData()

WR_NO = Nz(DLookup("[WR_NO]", "WRINFO", "WR_NO=" & WRtoQuery), "WR NO NOT IN WMIS")
SiteStreetNo = Nz(DLookup("[SiteStreetNo]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
SiteStreetName = Nz(DLookup("[SiteStreetName]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
SiteStreetType = Nz(DLookup("[SiteStreetType]", "WRINFO", "WR_NO=" & WRtoQuery), " ")
SiteCity = Nz(DLookup("[SiteCity]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
SiteState = Nz(DLookup("[SiteState]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
SiteZip = Nz(DLookup("[SiteZip]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")

CustName = Nz(DLookup("[CustName]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")

BillAmmount = Nz(DLookup("[BillAmmount]", "WRINFO", "WR_NO=" & WRtoQuery), "999999")
BillAddress = Nz(DLookup("[BillAddress]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
BillCity = Nz(DLookup("[BillCity]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
BillState = Nz(DLookup("[BillState]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
BillZip = Nz(DLookup("[BillZip]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")
BillName = Nz(DLookup("[BillName]", "WRINFO", "WR_NO=" & WRtoQuery), "NOT IN WMIS")

End Function

Private Sub RazeLetterWordPG_Click()

Call InputWR '' Calls function for input box for WR#

''make sure they put a wr #
If WRtoQuery = "" Then
    MsgBox "You Must enter a Work Request Number"
    Exit Sub
End If

''make sure service associate is selected
If IsNull(SASelector.Value) Then
    MsgBox "Please select a Service Associate Name"
    Exit Sub
End If

Call PullSAInfo '' Pull info on SAName
Call PullWRData ''Calls function that prompts for WR# and pulls values from WMIS

Path = "Z:\PDelivery\WAS\Engineering_Design\4405_Engineering_Pepco\ServiceAssociateToolBox\Letters\DemoLetterPG.docx"

Call fillRazeLetter ''Calls function that opens word and fills out raze letter

End Sub

Private Sub StreetLightLetter_Click()

Call InputWR '' Calls function for input box for WR#

''make sure they put a wr #
If WRtoQuery = "" Then
    MsgBox "You Must enter a Work Request Number"
    Exit Sub
End If

''make sure service associate is selected
If IsNull(SASelector.Value) Then
    MsgBox "Please select a Service Associate Name"
    Exit Sub
End If

Call PullSAInfo '' Pull info on SAName
Call PullWRData ''Calls function that prompts for WR# and pulls values from WMIS

Path = "Z:\PDelivery\WAS\Engineering_Design\4405_Engineering_Pepco\ServiceAssociateToolBox\Letters\StreeLightCostLetter.docx"

Call fillCostLetter ''Calls function that opens word and fills out cost letter

End Sub
