Attribute VB_Name = "mdl_0ffice_events"
Public Sub SE_ACT_OPEN(control As IRibbonControl)
    Select Case control.ID
        Case "se_open_dft"
            frmBatchOpen.Show
        Case "se_open_revmgr"
            Call openDummy
        Case "se_open_pdf"
            Call openSelectedPDF
    End Select
End Sub
Public Sub SE_ACT_PDF(control As IRibbonControl)
    Select Case control.ID
        Case "se_print_pdf"
            Call printPDF
        Case "se_print_alldft"
            Call printAll_PDF
        Case "se_print_all_DXF"
            Call PrintAll_DXF
        Case "se_print_all_DWG"
            Call PrintAll_DWG
    End Select
End Sub

Public Sub SE_ACT_TB(control As IRibbonControl)
    If seApp Is Nothing Then Call Conn2se
    If seApp.Documents.Count = 0 Then Exit Sub
    If Not seApp.ActiveDocument.Type = igDraftDocument Then
        MsgBox "Can't continue, active document is not a draft!", vbOKOnly + vbInformation, "Document Format Error"
    Else
        frmSE_tbm.Show
    End If
End Sub
Public Sub SE_ACT_CLOSE(control As IRibbonControl)
    Select Case control.ID
        Case "bt_st_close1"
            Call closeallDFT(True)
        Case "bt_st_close2"
            Call closeallDFT(False)
    End Select
End Sub
Public Sub SE_ACT_Print(control As IRibbonControl)
    frmPrintCenter.Show
'    Select Case control.ID
'        Case "bt_st_print1"
'            Call printAll_Paper
'        Case "bt_st_print2"
'            Call PrintSelected_Paper
'        Case "bt_st_print3"
'            Call PrintSelectedPDF_Paper
'    End Select
End Sub


Public Sub SE_ACT_SPEC(control As IRibbonControl)

    Select Case control.ID
        Case "bt_se_spec"
            If seApp Is Nothing Then Call Conn2se
            If seApp.Documents.Count = 0 Then Exit Sub
            If seApp.ActiveDocument.Type = igDraftDocument Then
                frmSpec.Show
            End If
    End Select
End Sub
Public Sub SE_ACT_Balloon(control As IRibbonControl)
    Select Case control.ID
        Case "bt_balloon_delete"
            Call deleteBB
        Case "bt_balloon_edit"
            Call editBB
    End Select
End Sub
Public Sub SE_ACT_CONFIG(control As IRibbonControl)
frmConfig.Show
End Sub


Public Sub DOMI_BOM_ACT(control As IRibbonControl)
    Select Case control.ID
        Case "bt_bom_switchcell"
            Call switchCell
        Case "bt_bom_VerticalMerge"
            Call verticalMerge
        Case "bt_bom_addraw"
            Call addRawMaterial
        Case "bt_bom_CheckUsage"
            Call bom_CheckUsage
        Case "bt_bom_Format_K3bom"
            Call initExportedBom
        Case "bt_bom_Get_cutsize"
            Call bom_Get_cutsize
        Case "bt_bom_Trans_2_SpotWeld"
            Call Trans2SpotWeld
    End Select
End Sub


Public Sub OPEN_TEMPLATE_ACT(control As IRibbonControl)
    Call OpenTableTemplate(control.ID)
End Sub

