Private Sub ClearAll_Click()
  Dim x As Control
  For Each x In QueryDialogBox.Controls
     If TypeOf x Is MSForms.CheckBox Then x.Value = False
     If TypeOf x Is MSForms.TextBox Then x.Value = ""
     If TypeOf x Is MSForms.ComboBox Then x.Value = ""
  Next



'If QueryDialogBox.ClearAll.Value = True Then
'
'    QueryDialogBox.ClearAll.Value = False
'    QueryDialogBox.SelectAllCheckbox.Value = False
'    QueryDialogBox.SelectBasicCheckbox.Value = False
'
'    QueryDialogBox.ShowWRType.Value = False
'    QueryDialogBox.ShowWRStatus.Value = False
'    QueryDialogBox.ShowWRName.Value = False
'    QueryDialogBox.ShowWRAddress.Value = False
'    QueryDialogBox.ShowWROwner.Value = False
'    QueryDialogBox.ShowWROwnerInits.Value = False
'    QueryDialogBox.Show1150.Value = False
'    QueryDialogBox.Show1700.Value = False
'    QueryDialogBox.Show1900.Value = False
'    QueryDialogBox.Show1925.Value = False
'    QueryDialogBox.Show2000.Value = False
'    QueryDialogBox.Show2050.Value = False
'    QueryDialogBox.Show2100.Value = False
'    QueryDialogBox.Show2150.Value = False
'    QueryDialogBox.Show2200.Value = False
'    QueryDialogBox.Show2450.Value = False
'
'    QueryDialogBox.ShowCustReady.Value = False
'    QueryDialogBox.ShowMeterSet.Value = False
'    QueryDialogBox.ShowConstComplete.Value = False
'
'End If

End Sub


Private Sub DeleteQueryCommand_Click()

If Not (QueryDialogBox.LoadQuery) = "" Then
    DeleteQueryRow
End If

End Sub



Private Sub LoadQueryCommand_Click()

PullUserform

End Sub

Private Sub SaveQueryCommand_Click()

SaveUserform

End Sub

Private Sub SelectBasicCheckbox_Click()

If QueryDialogBox.SelectBasicCheckbox.Value = True Then

    QueryDialogBox.SelectAllCheckbox.Value = False
   

    QueryDialogBox.ShowWRType.Value = True
    QueryDialogBox.ShowWRStatus.Value = True
    QueryDialogBox.ShowWRName.Value = True
    QueryDialogBox.ShowWRAddress.Value = True
    QueryDialogBox.ShowWROwner.Value = False
    QueryDialogBox.ShowWROwnerInits.Value = True
    QueryDialogBox.Show1150.Value = False
    QueryDialogBox.Show1700.Value = False
    QueryDialogBox.Show1900.Value = False
    QueryDialogBox.Show1925.Value = False
    QueryDialogBox.Show2000.Value = False
    QueryDialogBox.Show2050.Value = False
    QueryDialogBox.Show2100.Value = False
    QueryDialogBox.Show2150.Value = False
    QueryDialogBox.Show2200.Value = False
    QueryDialogBox.Show2450.Value = False
    
    QueryDialogBox.ShowCustReady.Value = True
    QueryDialogBox.ShowMeterSet.Value = False
    QueryDialogBox.ShowConstComplete.Value = False
    
End If

End Sub

Private Sub UserForm_Initialize()

With QueryDialogBox.FilterStatus
    .AddItem "INIT"
    .AddItem "DESGN"
    .AddItem "CONST"
    .AddItem "PSCHE"
    .AddItem "SCHED"
    .AddItem "DSFNL"
    .AddItem "FINAL"
    .AddItem "ARCH"
End With

With QueryDialogBox.FilterDistrictCode
    .AddItem "600"
    .AddItem "601"
    .AddItem "611"
    .AddItem "612"
    .AddItem "613"
    .AddItem "614"
    .AddItem "616"
    .AddItem "617"
    .AddItem "619"
    .AddItem "620"
    .AddItem "622T"
    .AddItem "625D"
    .AddItem "626"
    .AddItem "627"
End With

With QueryDialogBox.LocalDistrictCombo
    .AddItem "MCEO"
    .AddItem "MCEU"
    .AddItem "MCWO"
    .AddItem "MCEO"
    .AddItem "PGNO"
    .AddItem "PGNU"
    .AddItem "PGSO"
    .AddItem "PGSU"
    .AddItem "DCCO"
    .AddItem "DCCU"
    .AddItem "DCEO"
    .AddItem "DCEU"
    .AddItem "DCWO"
    .AddItem "DCWU"
End With


'' Code below pulls saved queries

  Dim rng As Range

  'To fill based on range
  For Each rng In Worksheets("SavedQueries").Range("A1:A500")
  If Not IsEmpty(rng) Then
    QueryDialogBox.LoadQuery.AddItem rng.Value
    End If
  Next



End Sub

Public Sub RunQuery_Click()


If (QueryDialogBox.WhereWRNo.Value = False) And (QueryDialogBox.WhereDistrictCode.Value = False) And _
(QueryDialogBox.WhereWROwner.Value = False) And (QueryDialogBox.WhereWRType.Value = False) And _
(QueryDialogBox.WhereLocalDistrict.Value = False) And (QueryDialogBox.WhereCreatedBefore.Value = False) And _
(QueryDialogBox.WhereCreatedAfter.Value = False) And (QueryDialogBox.WhereStatus.Value = False) Then

    MsgBox "You must select a filter"
    Exit Sub
End If

Dim strSQL As String 'SQL Query


strSQL = CreateSQL()


'' Establish connection with WMIS

Import_from_WMIS strSQL 'Pulls wmis data for SQL string

Me.Hide ''Close the dialog

End Sub


Public Function CreateSQL() As String

Dim strSQL As String
Dim selectcmd As String
Dim fromcmd As String
Dim wherecmd As String

Dim WRNOs As String
Dim FilterDistrictCode As String
Dim FilterWROwner As String
Dim FilterWRType As String

WRNOs = QueryDialogBox.WRNum.Value
FilterDistrictCode = QueryDialogBox.FilterDistrictCode.Value
FilterWROwner = QueryDialogBox.FilterWROwner.Value
FilterWRType = QueryDialogBox.FilterWRType.Value

fromcmd = ""

wherecmd = "" 'initialize where command

If QueryDialogBox.WhereWRNo.Value = True Then
    wherecmd = wherecmd & " AND wr.WR_NO in (" & WRNum & ")"
End If

''adding new commands
If QueryDialogBox.WhereStatus.Value = True Then
    wherecmd = wherecmd & " AND wr.WR_STATUS_CODE = '" & FilterStatus & "'"
End If

If QueryDialogBox.WhereState.Value = True Then
    wherecmd = wherecmd & " AND wr.STATE = '" & FilterState & "'"
End If

'' end new commands


If QueryDialogBox.WhereDistrictCode.Value = True Then
    wherecmd = wherecmd & " AND wr.PLANNING_DISTRICT_CODE like " & FilterDistrictCode
End If

If QueryDialogBox.WhereWROwner.Value = True Then
    wherecmd = wherecmd & " AND ownername.PERSON_INITIALS = '" & FilterWROwner & "'"
End If

If QueryDialogBox.WhereLocalDistrict.Value = True Then
    If Not IsEmpty(QueryDialogBox.LocalDistrictCombo.Value) Then
        wherecmd = wherecmd & " AND wr.TAX_DISTRICT_CODE = '" & LocalDistrictCombo & "'"
    End If
End If

If QueryDialogBox.WhereWRType.Value = True Then
    wherecmd = wherecmd & " AND wr.WR_TYPE_CODE LIKE '" & FilterWRType & "'"
End If

If QueryDialogBox.WhereCreatedAfter.Value = True Then
    wherecmd = wherecmd & " AND wr.ENTRY_DATE >= '" & FilterCreatedAfter & "'"
End If

If QueryDialogBox.WhereCreatedBefore.Value = True Then
    wherecmd = wherecmd & " AND wr.ENTRY_DATE <= '" & FilterCreatedBefore & "'"
End If

If QueryDialogBox.IgnoreCancelled.Value = True Then
    wherecmd = wherecmd & " AND wr.WR_CANCEL_DATE is NULL"
End If

selectcmd = "" 'initialize select command

If QueryDialogBox.ShowWRType.Value = True Then
    selectcmd = selectcmd & ", wr.WR_TYPE_CODE"
End If

If QueryDialogBox.ShowWRStatus.Value = True Then
    selectcmd = selectcmd & ", wr.WR_STATUS_CODE"
End If

If QueryDialogBox.ShowWRName.Value = True Then
    selectcmd = selectcmd & ", wr.WR_NAME"
End If

If QueryDialogBox.ShowWRAddress.Value = True Then
    selectcmd = selectcmd & ", wr.ADDRESS_1"
End If

If QueryDialogBox.ShowCustReady.Value = True Then
    selectcmd = selectcmd & ", wr.CUSTOMER_READY_DATE"
End If

If QueryDialogBox.ShowConstComplete.Value = True Then
    selectcmd = selectcmd & ", wr.CONSTRUCTION_COMPLETE_DATE"
End If

If QueryDialogBox.ShowWRAddress.Value = True Then
    selectcmd = selectcmd & ", wr.METER_SET_DATE"
End If

If QueryDialogBox.ShowWROwner.Value = True Then
    selectcmd = selectcmd & ", ownername.NAME AS ""OWNER NAME"""
End If

If QueryDialogBox.ShowWROwnerInits.Value = True Then
    selectcmd = selectcmd & ", ownername.PERSON_INITIALS AS ""OWNER"""
End If

If QueryDialogBox.Show1150.Value = True Then
    selectcmd = selectcmd & ", tsk1150.COMMENTS AS ""1150 Comments"", tsk1150.TASK_STATUS_CODE AS ""1150 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk1150 ON wr.WR_NO = tsk1150.WR_NO AND tsk1150.WR_TASK_NO = 1150"
End If

If QueryDialogBox.Show1700.Value = True Then
    selectcmd = selectcmd & ", tsk1700.COMMENTS AS ""1700 Comments"", tsk1700.TASK_STATUS_CODE AS ""1700 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk1700 ON wr.WR_NO = tsk1700.WR_NO AND tsk1700.WR_TASK_NO = 1700"
End If

If QueryDialogBox.Show1900.Value = True Then
    selectcmd = selectcmd & ", tsk1900.COMMENTS AS ""1900 Comments"", tsk1900.TASK_STATUS_CODE AS ""1900 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk1900 ON wr.WR_NO = tsk1900.WR_NO AND tsk1900.WR_TASK_NO = 1900"
End If

If QueryDialogBox.Show1925.Value = True Then
    selectcmd = selectcmd & ", tsk1925.COMMENTS AS ""1925 Comments"", tsk1925.TASK_STATUS_CODE AS ""1925 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk1925 ON wr.WR_NO = tsk1925.WR_NO AND tsk1925.WR_TASK_NO = 1925"
End If

If QueryDialogBox.Show2000.Value = True Then
    selectcmd = selectcmd & ", tsk2000.COMMENTS AS ""2000 Comments"", tsk2000.TASK_STATUS_CODE AS ""2000 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk2000 ON wr.WR_NO = tsk2000.WR_NO AND tsk2000.WR_TASK_NO = 2000"
End If

If QueryDialogBox.Show2050.Value = True Then
    selectcmd = selectcmd & ", tsk2050.COMMENTS AS ""2050 Comments"", tsk2050.TASK_STATUS_CODE AS ""2050 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk2050 ON wr.WR_NO = tsk2050.WR_NO AND tsk2050.WR_TASK_NO = 2050"
End If

If QueryDialogBox.Show2100.Value = True Then
    selectcmd = selectcmd & ", tsk2100.COMMENTS AS ""2100 Comments"", tsk2100.TASK_STATUS_CODE AS ""2100 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk2100 ON wr.WR_NO = tsk2100.WR_NO AND tsk2100.WR_TASK_NO = 2100"
End If

If QueryDialogBox.Show2150.Value = True Then
    selectcmd = selectcmd & ", tsk2150.COMMENTS AS ""2150 Comments"", tsk2150.TASK_STATUS_CODE AS ""2150 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk2150 ON wr.WR_NO = tsk2150.WR_NO AND tsk2150.WR_TASK_NO = 2150"
End If

If QueryDialogBox.Show2200.Value = True Then
    selectcmd = selectcmd & ", tsk2200.COMMENTS AS ""2200 Comments"", tsk2200.TASK_STATUS_CODE AS ""2200 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk2200 ON wr.WR_NO = tsk2200.WR_NO AND tsk2200.WR_TASK_NO = 2200"
End If

If QueryDialogBox.Show2450.Value = True Then
    selectcmd = selectcmd & ", tsk2450.COMMENTS AS ""2450 Comments"", tsk2450.TASK_STATUS_CODE AS ""2450 Status"""
    fromcmd = fromcmd & " LEFT JOIN WR_TASK tsk2450 ON wr.WR_NO = tsk2450.WR_NO AND tsk2450.WR_TASK_NO = 2450"
End If
    

'CreateSQL = "SELECT wr.WR_NO" & selectcmd & " from WORK_REQUEST wr WHERE wr.WR_NO In (" & WRNOs & ")"

CreateSQL = "SELECT wr.WR_NO" & selectcmd & " FROM WORK_REQUEST wr" _
& fromcmd _
& " LEFT JOIN ALL_PEOPLE ownername ON ownername.PERSON_NO = wr.WR_OWNER_PERSON_NO" _
& " LEFT JOIN WR_CONTACT contact ON contact.WR_NO = wr.WR_NO" _
& " LEFT JOIN WR_CONTACT_PHONE phone ON phone.CONTACT_ID = contact.CONTACT_ID AND phone.WR_NO = wr.WR_NO AND phone.PHONE_ID = 1" _
& " WHERE wr.WR_NO IS NOT NULL AND wr.COMPANY_CODE = '7000' " & wherecmd



' wr.PLANNING_DISTRICT_CODE = '611' AND wr.FINAL_CLOSE_DATE is NULL" And wr.WR_CANCEL_DATE Is Null AND
End Function

Private Sub SelectAllCheckbox_Click()

If QueryDialogBox.SelectAllCheckbox.Value = True Then

    QueryDialogBox.SelectBasicCheckbox.Value = False

    QueryDialogBox.ShowWRType.Value = True
    QueryDialogBox.ShowWRStatus.Value = True
    QueryDialogBox.ShowWRName.Value = True
    QueryDialogBox.ShowWRAddress.Value = True
    QueryDialogBox.ShowWROwner.Value = True
    QueryDialogBox.ShowWROwnerInits.Value = True
    QueryDialogBox.Show1150.Value = True
    QueryDialogBox.Show1700.Value = True
    QueryDialogBox.Show1900.Value = True
    QueryDialogBox.Show1925.Value = True
    QueryDialogBox.Show2000.Value = True
    QueryDialogBox.Show2050.Value = True
    QueryDialogBox.Show2100.Value = True
    QueryDialogBox.Show2150.Value = True
    QueryDialogBox.Show2200.Value = True
    QueryDialogBox.Show2450.Value = True
    
    QueryDialogBox.ShowCustReady.Value = True
    QueryDialogBox.ShowMeterSet.Value = True
    QueryDialogBox.ShowConstComplete.Value = True

End If

End Sub

Private Sub WhereCreatedAfter_Click()
    
    QueryDialogBox.FilterCreatedAfter.Enabled = WhereCreatedAfter.Value

End Sub

Private Sub WhereCreatedBefore_Click()

    QueryDialogBox.FilterCreatedBefore.Enabled = WhereCreatedBefore.Value

End Sub

Private Sub WhereDistrictCode_Click()

    QueryDialogBox.FilterDistrictCode.Enabled = WhereDistrictCode.Value


End Sub

Private Sub WhereLocalDistrict_Click()

    QueryDialogBox.LocalDistrictCombo.Enabled = WhereLocalDistrict.Value

End Sub

Private Sub WhereState_Click()

    QueryDialogBox.FilterState.Enabled = WhereState.Value

End Sub

Private Sub WhereStatus_Click()

    QueryDialogBox.FilterStatus.Enabled = WhereStatus.Value


End Sub

Private Sub WhereWRNo_Click()

    QueryDialogBox.WRNum.Enabled = WhereWRNo.Value


End Sub

Private Sub WhereWROwner_Click()

    QueryDialogBox.FilterWROwner.Enabled = WhereWROwner.Value

End Sub

Private Sub WhereWRType_Click()
    
    QueryDialogBox.FilterWRType.Enabled = WhereWRType.Value
    
End Sub

