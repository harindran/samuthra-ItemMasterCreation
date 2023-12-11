Public Class clsItemMaster
    Public Const frmType As String = "150"
    Dim TypeSelected As Boolean = False
    Dim objForm As SAPbouiCOM.Form
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim strSQL As String = ""
    Dim objRS As SAPbobsCOM.Recordset
    Dim objItemForm As SAPbouiCOM.Form
    Dim typecode, typedesc, mgcode, mgdesc, sgcode, sgdesc, ssgcode, ssgdesc As String
    Public Sub ItemEvent(FormUID As String, pval As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pval.BeforeAction = True Then
            Select Case pval.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    'If pval.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Not ValidateCode(FormUID) Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    'objItemForm = objAddOn.objApplication.Forms.Item(FormUID)
               
            End Select
        Else
            Select Case pval.EventType
                'Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                 
                'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                '    If pval.ItemUID = "U_Type" Then
                '        LoadMG(FormUID)
                '    ElseIf pval.ItemUID = "U_MG" Then
                '        LoadSG(FormUID)
                '    ElseIf pval.ItemUID = "U_SG" Then
                '        LoadSSG(FormUID)
                '    ElseIf pval.ItemUID = "U_SSG" Then
                '        CodeCreation(FormUID)
                '    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pval.ItemUID = "U_Type" Then
                        selectParameters(FormUID, pval)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pval.ItemUID = "U_SSGDesc" Then
                        CodeCreation(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    If pval.ItemUID = "" And TypeSelected Then
                        AssignParameters(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                    ResizeFields(FormUID)
            End Select
        End If
    End Sub
    Private Sub AssignParameters(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objItemForm.Items.Item("U_Type").Specific.String = typecode
        objItemForm.Items.Item("U_TypeDesc").Specific.String = typedesc
        objItemForm.Items.Item("U_MG").Specific.String = mgcode
        objItemForm.Items.Item("U_MGDesc").Specific.String = mgdesc
        objForm.Items.Item("U_SG").Specific.String = sgcode
        objForm.Items.Item("U_SGDesc").Specific.String = sgdesc
        objForm.Items.Item("U_SSG").Specific.String = ssgcode
        objForm.Items.Item("U_SSGDesc").Specific.String = ssgdesc
        TypeSelected = False


    End Sub
    Private Sub ResizeFields(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objForm.Items.Item("LType").Left = 774
            objForm.Items.Item("LType").Top = 63
            objForm.Items.Item("U_Type").Left = 856
            objForm.Items.Item("U_Type").Top = 63
            objForm.Items.Item("U_TypeDesc").Left = 901
            objForm.Items.Item("U_TypeDesc").Top = 63
            objForm.Items.Item("LMG").Left = 774
            objForm.Items.Item("LMG").Top = 78
            objForm.Items.Item("U_MG").Left = 856
            objForm.Items.Item("U_MG").Top = 78
            objForm.Items.Item("U_MGDesc").Left = 901
            objForm.Items.Item("U_MGDesc").Top = 78
            objForm.Items.Item("LSG").Left = 774
            objForm.Items.Item("LSG").Top = 93
            objForm.Items.Item("U_SG").Left = 856
            objForm.Items.Item("U_SG").Top = 93
            objForm.Items.Item("U_SGDesc").Left = 901
            objForm.Items.Item("U_SGDesc").Top = 93
            objForm.Items.Item("LSSG").Left = 774
            objForm.Items.Item("LSSG").Top = 108
            objForm.Items.Item("U_SSG").Left = 856
            objForm.Items.Item("U_SSG").Top = 108
            objForm.Items.Item("U_SSGDesc").Left = 901
            objForm.Items.Item("U_SSGDesc").Top = 108
        Catch ex As Exception
        End Try
    End Sub
    Private Sub selectParameters(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvents As SAPbouiCOM.ChooseFromListEvent
        objItemForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCFLEvents = pval
        Dim objTable As SAPbouiCOM.DataTable
        objTable = objCFLEvents.SelectedObjects

        If objTable.Rows.Count > 0 Then
            Try

                typecode = objTable.GetValue("U_TypeCode", 0)
                typedesc = objTable.GetValue("U_TypeDesc", 0)
                mgcode = objTable.GetValue("U_MGCode", 0)
                mgdesc = objTable.GetValue("U_MGDesc", 0)
                sgcode = objTable.GetValue("U_SGCode", 0)
                sgdesc = objTable.GetValue("U_SGDesc", 0)
                ssgcode = objTable.GetValue("U_SSGCode", 0)
                ssgdesc = objTable.GetValue("U_SSGDesc", 0)
                TypeSelected = True

                ' objItemForm.Items.Item("U_Type").Specific.String = typecode


                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_Type", 0, objTable.GetValue("U_TypeCode", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_TypeDesc", 0, objTable.GetValue("U_TypeDesc", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_MG", 0, objTable.GetValue("U_MGCode", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_MGDesc", 0, objTable.GetValue("U_MGDesc", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_SG", 0, objTable.GetValue("U_SGCode", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_SGDesc", 0, objTable.GetValue("U_SGDesc", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_SSG", 0, objTable.GetValue("U_SSGCode", 0))
                'objForm.DataSources.DBDataSources.Item("OITM").SetValue("U_SSGDesc", 0, objTable.GetValue("U_SSGDesc", 0))
            Catch ex As Exception
                ' objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short)
            End Try

            'Try
            '    objItemForm.Items.Item("U_TypeDesc").Specific.String = typedesc
            'Catch ex As Exception

            'End Try
            'Try
            '    objItemForm.Items.Item("U_MG").Specific.String = mgcode
            'Catch ex As Exception

            'End Try
            'Try
            '    objItemForm.Items.Item("U_MGDesc").Specific.String = mgdesc
            'Catch ex As Exception

            'End Try
            'Try
            '    objForm.Items.Item("U_SG").Specific.String = sgcode
            'Catch ex As Exception

            'End Try
            'Try
            '    objForm.Items.Item("U_SGDesc").Specific.String = sgdesc
            'Catch ex As Exception

            'End Try
            'Try
            '    objForm.Items.Item("U_SSG").Specific.String = ssgcode
            'Catch ex As Exception

            'End Try
            'Try
            '    objForm.Items.Item("U_SSGDesc").Specific.String = ssgdesc
            'Catch ex As Exception

            'End Try
            ''

        End If


    End Sub
    Public Sub LoadForm(ByVal FormUID As String)
        CreateObjects(FormUID)
        'LoadType(FormUID)
    End Sub
    Private Sub CreateObjects(ByVal FormUID As String)
        Dim oItem As SAPbouiCOM.Item
        Dim oLabel As SAPbouiCOM.StaticText
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button
        '    Dim oEditText As SAPbouiCOM.ComboBox
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        'Production No.
        Try
            Dim OCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLParams = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLParams.ObjectType = "MICODE"
            oCFLParams.UniqueID = "CFLCode"
            oCFLParams.MultiSelection = False
            OCFL = objForm.ChooseFromLists.Add(oCFLParams)

            objForm.Freeze(True)
            '    oEditText = objForm.Items.Item("U_MG").Specific

            'Dim left As Integer = objForm.Items.Item("5").Left + (objForm.Items.Item("5").Width * 3) + 5
            'Dim Top As Integer = objForm.Items.Item("5").Top
            'Dim Height As Integer = 14
            'Dim ValueWidth As Integer = 30
            'Dim Width As Integer = 120
            'Dim topdiff As Integer = 15
            'Dim leftdiff As Integer = 100

            'oItem = objForm.Items.Add("U_Type", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left
            'oItem.Width = ValueWidth
            'oItem.Height = Height
            'oItem.Top = Top
            'oItem.Visible = True
            'oItem.DisplayDesc = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_Type")
            '' oEditText.ChooseFromListUID = OCFL.UniqueID
            '' oEditText.ChooseFromListAlias = "U_TypeCode"
            'oItem = objForm.Items.Add("U_TypeDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left + ValueWidth
            'oItem.Width = Width
            'oItem.Height = Height
            'oItem.Top = Top
            'oItem.Visible = True
            'oItem.DisplayDesc = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_TypeDesc")
            'oItem = objForm.Items.Add("LType", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Left = left - leftdiff
            'oItem.Top = Top
            'oItem.Height = Height
            'oItem.Width = Width - ValueWidth
            'oItem.LinkTo = "U_Type"
            'oItem.Visible = True
            'oLabel = oItem.Specific
            'oLabel.Caption = "Type "


            'oItem = objForm.Items.Add("U_MG", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left
            'oItem.Width = ValueWidth
            'oItem.Height = Height
            'oItem.Top = Top + topdiff
            'oItem.Visible = True
            'oItem.DisplayDesc = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_MG")
            'oItem = objForm.Items.Add("U_MGDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left + ValueWidth
            'oItem.Width = Width
            'oItem.Height = Height
            'oItem.Top = Top + topdiff
            'oItem.Visible = True
            'oItem.DisplayDesc = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_MGDesc")
            'oItem = objForm.Items.Add("LMG", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Left = left - leftdiff
            'oItem.Top = Top + topdiff
            'oItem.Height = Height
            'oItem.Width = Width - ValueWidth
            'oItem.LinkTo = "U_MG"
            'oItem.Visible = True
            'oLabel = oItem.Specific
            'oLabel.Caption = "Main Group"


            'oItem = objForm.Items.Add("U_SG", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left
            'oItem.Width = ValueWidth
            'oItem.Height = Height
            'oItem.Top = Top + (2 * topdiff)
            'oItem.DisplayDesc = True
            'oItem.Visible = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_SG")
            'oItem = objForm.Items.Add("U_SGDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left + ValueWidth
            'oItem.Width = Width
            'oItem.Height = Height
            'oItem.Top = Top + (2 * topdiff)
            'oItem.DisplayDesc = True
            'oItem.Visible = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_SGDesc")
            'oItem = objForm.Items.Add("LSG", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Left = left - leftdiff
            'oItem.Top = Top + (2 * topdiff)
            'oItem.Height = Height
            'oItem.Width = Width - ValueWidth
            'oItem.LinkTo = "U_SG"
            'oItem.Visible = True
            'oLabel = oItem.Specific
            'oLabel.Caption = "Sub Group "


            'oItem = objForm.Items.Add("U_SSG", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left
            'oItem.Width = ValueWidth
            'oItem.Height = Height
            'oItem.Top = Top + (3 * topdiff)
            'oItem.DisplayDesc = True
            'oItem.Visible = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_SSG")
            'oItem = objForm.Items.Add("U_SSGDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = left + ValueWidth
            'oItem.Width = Width
            'oItem.Height = Height
            'oItem.Top = Top + (3 * topdiff)
            'oItem.DisplayDesc = True
            'oItem.Visible = True
            'oEditText = oItem.Specific
            'oEditText.DataBind.SetBound(True, "OITM", "U_SSGDesc")
            'oItem = objForm.Items.Add("LSSG", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Left = left - leftdiff
            'oItem.Top = Top + (3 * topdiff)
            'oItem.Height = Height
            'oItem.Width = Width - ValueWidth
            'oItem.LinkTo = "U_SSG"
            'oItem.Visible = True
            'oLabel = oItem.Specific
            'oLabel.Caption = "SubSubGroup "


            oItem = objForm.Items.Add("U_Type", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 856
            oItem.Width = 44
            oItem.Height = 14
            oItem.Top = 63
            oItem.Visible = True
            oItem.DisplayDesc = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_Type")
            oItem = objForm.Items.Add("U_TypeDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 901
            oItem.Top = 63
            oItem.Height = 14
            oItem.Width = 120
            oItem.Visible = True
            oItem.DisplayDesc = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_TypeDesc")
            oItem = objForm.Items.Add("LType", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 774 'objForm.Items.Item("214").Left + objForm.Items.Item("214").Width + 421 
            oItem.Width = 80
            oItem.Height = 14
            oItem.Top = 63
            oItem.LinkTo = "U_Type"
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "Type "


            oItem = objForm.Items.Add("U_MG", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 856
            oItem.Width = 44
            oItem.Height = 14
            oItem.Top = 78
            oItem.Visible = True
            oItem.DisplayDesc = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_MG")
            oItem = objForm.Items.Add("U_MGDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 901
            oItem.Top = 78
            oItem.Height = 14
            oItem.Width = 120
            oItem.Visible = True
            oItem.DisplayDesc = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_MGDesc")
            oItem = objForm.Items.Add("LMG", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 774
            oItem.Width = 80
            oItem.Height = 14
            oItem.Top = 78
            oItem.LinkTo = "U_MG"
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "Main Group"


            oItem = objForm.Items.Add("U_SG", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 856
            oItem.Width = 44
            oItem.Height = 14
            oItem.Top = 93
            oItem.DisplayDesc = True
            oItem.Visible = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_SG")
            oItem = objForm.Items.Add("U_SGDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 901
            oItem.Top = 93
            oItem.Height = 14
            oItem.Width = 120
            oItem.DisplayDesc = True
            oItem.Visible = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_SGDesc")
            oItem = objForm.Items.Add("LSG", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 774
            oItem.Width = 80
            oItem.Height = 14
            oItem.Top = 93
            oItem.LinkTo = "U_SG"
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "Sub Group "


            oItem = objForm.Items.Add("U_SSG", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 856
            oItem.Width = 44
            oItem.Height = 14
            oItem.Top = 108
            oItem.DisplayDesc = True
            oItem.Visible = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_SSG")
            oItem = objForm.Items.Add("U_SSGDesc", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 901
            oItem.Top = 108
            oItem.Height = 14
            oItem.Width = 120
            oItem.DisplayDesc = True
            oItem.Visible = True
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OITM", "U_SSGDesc")
            oItem = objForm.Items.Add("LSSG", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 774
            oItem.Width = 80
            oItem.Height = 14
            oItem.Top = 108
            oItem.LinkTo = "U_SSG"
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "SubSubGroup "

        Catch ex As Exception
            'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Finally
            objForm.Freeze(False)

        End Try

    End Sub
    Private Sub LoadMG(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objCmbType As SAPbouiCOM.ComboBox
        objCmbType = objForm.Items.Item("U_Type").Specific
        objCombo = objForm.Items.Item("U_MG").Specific
        If objAddOn.HANA Then
            strSQL = " Select  Distinct ""U_MGCode"", ""U_MGDesc"" from ""@MICODE"" where ""U_TypeCode""='" & objCmbType.Selected.Value & "';"
        Else

            strSQL = "Select  Distinct U_MGCode, U_MGDesc from [@MICODE] where U_TypeCode='" & objCmbType.Selected.Value & "'"
        End If
        If objCombo.ValidValues.Count > 0 Then
            Dim i As Integer = 0
            While i <= objCombo.ValidValues.Count - 1
                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            End While
        End If
        '   objCombo.ValidValues.Add("-", "-")
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        While Not objRS.EoF
            objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
            objRS.MoveNext()
        End While
        objRS = Nothing
    End Sub

    Private Sub LoadSG(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objCmbType As SAPbouiCOM.ComboBox
        objCmbType = objForm.Items.Item("U_MG").Specific
        objCombo = objForm.Items.Item("U_SG").Specific
        If objAddOn.HANA Then
            strSQL = " Select  Distinct ""U_SGCode"", ""U_SGDesc"" from ""@MICODE"" where ""U_MGCode""='" & objCmbType.Selected.Value & "';"
        Else

            strSQL = "Select  Distinct U_SGCode, U_SGDesc from [@MICODE] where U_MGCode='" & objCmbType.Selected.Value & "'"
        End If

        If objCombo.ValidValues.Count > 0 Then
            Dim i As Integer = 0
            While i <= objCombo.ValidValues.Count - 1
                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            End While
        End If

        'objCombo.ValidValues.Add("-", "-")
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        While Not objRS.EoF
            objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
            objRS.MoveNext()
        End While
        objRS = Nothing
    End Sub
    Private Sub LoadSSG(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objCmbType As SAPbouiCOM.ComboBox
        objCmbType = objForm.Items.Item("U_SG").Specific
        objCombo = objForm.Items.Item("U_SSG").Specific
        If objAddOn.HANA Then
            strSQL = " Select  Distinct ""U_SSGCode"", ""U_SSGDesc"" from ""@MICODE"" where ""U_SGCode""='" & objCmbType.Selected.Value & "';"
        Else
            strSQL = "Select  Distinct U_SSGCode, U_SSGDesc from [@MICODE] where U_SGCode='" & objCmbType.Selected.Value & "'"
        End If

        If objCombo.ValidValues.Count > 0 Then
            Dim i As Integer = 0
            While i <= objCombo.ValidValues.Count - 1
                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            End While
        End If
        'objCombo.ValidValues.Add("-", "-")
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        While Not objRS.EoF
            objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
            objRS.MoveNext()
        End While
        objRS = Nothing
    End Sub
    Private Sub LoadType(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try


            objCombo = objForm.Items.Item("U_Type").Specific
            If objAddOn.HANA Then
                strSQL = " Select  Distinct ""U_TypeCode"", ""U_TypeDesc"" from ""@MICODE"""
            Else

                strSQL = "Select  Distinct U_TypeCode, U_TypeDesc from [@MICODE] "
            End If

            If objCombo.ValidValues.Count > 0 Then
                 Dim i As Integer = 0
                While i <= objCombo.ValidValues.Count - 1
                    objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                End While
            End If
            ' objCombo.ValidValues.Add("-", "-")
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strSQL)
            While Not objRS.EoF
                objCombo.ValidValues.Add(objRS.Fields.Item(0).Value, objRS.Fields.Item(1).Value)
                objRS.MoveNext()
            End While
            objRS = Nothing
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub CodeCreation(ByVal FormUID As String)
        Dim StrSQL As String
        Dim objEdit As SAPbouiCOM.EditText


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try



            objEdit = objForm.Items.Item("U_MG").Specific
            Dim ExpItemCode As String = Trim(objEdit.Value)
            objEdit = objForm.Items.Item("U_SG").Specific
            ExpItemCode = ExpItemCode & Trim(objEdit.Value)
            objEdit = objForm.Items.Item("U_SSG").Specific
            ExpItemCode = ExpItemCode & Trim(objEdit.Value)

            If objAddOn.HANA Then
                StrSQL = "Select " & ExpItemCode & "||right(cast(ifnull((max(cast(ifnull(Right(""ItemCode"",3),'0')as int))),0)+1001 as varchar),3)  from OITM" & _
       " where ""ItemCode"" like '" & ExpItemCode & "%';"
            Else
                StrSQL = "Select Concat('" & ExpItemCode & "', Right(convert(varchar, isnull(max(convert(int, isnull(Right(ItemCode,3),'0'))),0) + 1001),3))  from OITM" & _
   " where ItemCode like '" & ExpItemCode & "%'"
            End If
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(StrSQL)
            If Not objRS.EoF Then
                objForm.Items.Item("5").Specific.String = Trim(CStr(objRS.Fields.Item(0).Value))

            End If
        Catch ex As Exception
        End Try
        '     StrSQL = "select $[$U_MG.0.0]||''||$[$U_SG.0.0]||''||$[$U_SSG.0.0]||''||right(cast(ifnull((max(cast(ifnull(Right(""ItemCode"",3),'0')as int))),0)+1001 as varchar),3)  from OITM" & _
        '" where ""ItemCode"" like $[$U_MG.0.0]||''||$[$U_SG.0.0]||''||$[$U_SSG.0.0]||'''%' ;"


        ' StrSQL ="select (select $[$U_MG.0.0] from dummy)||(select $[$U_SB.0.0] from dummy)||(select $[$U_SSB.0.0] from dummy)|| (select right(cast(ifnull((max(cast(ifnull(Right("ItemCode",3),'0')as int))),0)+1001 as varchar),3) from OITM  where  LENGTH("ItemCode") = 7 AND "ItemCode" like (select $[$U_MG.0.0] from dummy)||(select $[$U_SB.0.0] from dummy)||(select $[$U_SSB.0.0]  from dummy)||'%') from Dummy;

    End Sub
    Public Function ValidateCode(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try

        
        Dim ItemCode As String = Trim(objForm.Items.Item("5").Specific.String)
        objCombo = objForm.Items.Item("U_MG").Specific
            Dim ExpItemCode As String = Trim(objCombo.Value)
        objCombo = objForm.Items.Item("U_SG").Specific
            ExpItemCode = ExpItemCode & Trim(objCombo.Value)
        objCombo = objForm.Items.Item("U_SSG").Specific
            ExpItemCode = ExpItemCode & Trim(objCombo.Value)

        If objAddOn.HANA Then
                strSQL = "Select " & ExpItemCode & "||right(cast(ifnull((max(cast(ifnull(Right(""ItemCode"",3),'0')as int))),0)+1001 as varchar),3)  from OITM" & _
       " where ""ItemCode"" like '" & ExpItemCode & "%';"
        Else
                strSQL = "Select Concat('" & ExpItemCode & "', Right(convert(varchar, isnull(max(convert(int, isnull(Right(ItemCode,3),'0'))),0) + 1001),3))  from OITM" & _
   " where ItemCode like '" & ExpItemCode & "%'"
        End If
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        If Not objRS.EoF Then
            If ItemCode = Trim(objRS.Fields.Item(0).Value) Then
                Return True
            End If
            End If
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
            Return False
        End Try
        objAddOn.objApplication.SetStatusBarMessage("Please check itemcode", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Return False
    End Function

End Class
