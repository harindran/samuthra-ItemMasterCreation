Imports System.IO
Imports SAPbouiCOM.Framework

Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public SOMenuID As String = "0"
   
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0
    Public objItemMaster As clsItemMaster
  
    Public HANA As Boolean = True
    'Public HANA As Boolean = False

    Public HWKEY() As String = New String() {"R1574408489", "L1653539483", "A1836445156", "M0394249985", "V0913316776", "F0123559701", "L1552968038", "M0090876837", "H0922924113"}
    Private Sub CheckLicense()

    End Sub
    'Function isValidLicense() As Boolean
    '    Try
    '        If objApplication.Forms.ActiveForm.TypeCount > 1 Then
    '            For i As Integer = 0 To objApplication.Forms.ActiveForm.TypeCount - 1
    '                objApplication.Forms.ActiveForm.Close()
    '            Next
    '        End If
    '        objApplication.Menus.Item("257").Activate()
    '        Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
    '        objApplication.Forms.ActiveForm.Close()

    '        For i As Integer = 0 To HWKEY.Length - 1
    '            If HWKEY(i).Trim = CrrHWKEY.Trim Then
    '                Return True
    '            End If
    '        Next
    '        MsgBox("Add-on installation failed due to license mismatch", MsgBoxStyle.OkOnly, "License Management")
    '        Return False
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    '    Return True
    'End Function

    Function isValidLicense() As Boolean
        Try
            If objApplication.Forms.ActiveForm.TypeCount > 0 Then
                For i As Integer = 0 To objApplication.Forms.ActiveForm.TypeCount - 1
                    objApplication.Forms.ActiveForm.Close()
                Next
            End If
            objapplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objapplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objapplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            'MsgBox(ex.ToString)
            objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
        Return True
    End Function
    Public Sub Intialize(ByVal args() As String)
        Try
            Dim oapplication As Application
            If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
            objapplication = Application.SBO_Application
            If isValidLicense() Then
                objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objcompany = Application.SBO_Application.Company.GetDICompany()
                Try
                    createObjects()
                    createTables()
                    createUDOs()
                    loadMenu()
                Catch ex As Exception
                    objAddOn.objApplication.MessageBox(ex.ToString)
                    End
                End Try
                objApplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                oapplication.Run()
            Else
                objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'System.Windows.Forms.Application.Run()
        Catch ex As Exception
            objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Public Sub Intialize()
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            'incomingPayment()
            'OutgoingPayment()
            createTables()
            createObjects()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString)
            End
        End Try
        If isValidLicense() Then
            objApplication.SetStatusBarMessage("Addon connected successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Else
            objApplication.SetStatusBarMessage("Failed To Connect, Please Check The License Configuration", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objCompany.Disconnect()
            objApplication = Nothing
            objCompany = Nothing
            End
        End If
    End Sub
    Private Sub createUDOs()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        Dim ct1(1) As String
        ct1(0) = ""
        objUDFEngine.createUDO("MICODE", "MICODE", "MICODE", ct1, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, False)
    End Sub
    Private Sub createObjects()
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)
        objItemMaster = New clsItemMaster
     
    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case clsItemMaster.frmType
                    objItemMaster.ItemEvent(FormUID, pVal, BubbleEvent)

            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        'If BusinessObjectInfo.BeforeAction Then
        '    Try

        '        If BusinessObjectInfo.FormTypeEx = "150" And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
        '            If objItemMaster.ValidateCode(BusinessObjectInfo.FormUID) Then
        '            Else
        '                BubbleEvent = False
        '            End If
        '        End If

        '    Catch ex As Exception
        '        objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '    End Try
        'End If
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
            Select Case pVal.MenuUID
               
            End Select
        Else
            Try
                Select Case pVal.MenuUID
                    
                    Case "1282" ', "1290", "1288", "1289", "1291"
                       
                    Case "ditem"
                    Case "3073"
                        objItemMaster.LoadForm(objApplication.Forms.ActiveForm.UniqueID)


                End Select
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub
    Private Sub loadMenu()
        'If objApplication.Menus.Item("43520").SubMenus.Exists("QC") Then Return
        'MenuCount = objApplication.Menus.Item("43520").SubMenus.Count

        'CreateMenu(Application.StartupPath + "\jc1.png", MenuCount + 1, "QC Management", SAPbouiCOM.BoMenuType.mt_POPUP, "QC", objApplication.Menus.Item("43520"))
        'CreateMenu("", 1, "Quality Check", SAPbouiCOM.BoMenuType.mt_STRING, clsQC.Formtype, objApplication.Menus.Item("QC"))

    End Sub
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    Private Sub createTables()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        ' WriteSMSLog("0")
       

        objUDFEngine.AddAlphaField("OITM", "Type", "Type", 5)
        objUDFEngine.AddAlphaField("OITM", "MG", "Main Group", 5)
        objUDFEngine.AddAlphaField("OITM", "SG", "SubGroup", 5)
        objUDFEngine.AddAlphaField("OITM", "SSG", "SubSubGroup", 5)

        objUDFEngine.AddAlphaField("OITM", "TypeDesc", "Type", 50)
        objUDFEngine.AddAlphaField("OITM", "MGDesc", "Main Group", 50)
        objUDFEngine.AddAlphaField("OITM", "SGDesc", "SubGroup", 50)
        objUDFEngine.AddAlphaField("OITM", "SSGDesc", "SubSubGroup", 50) '' 

        objUDFEngine.CreateTable("MICODE", "Codification", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objUDFEngine.AddAlphaField("MICODE", "TypeCode", "TypeCode", 5)
        objUDFEngine.AddAlphaField("MICODE", "TypeDesc", "TypeDesc", 50)
        objUDFEngine.AddAlphaField("MICODE", "MGCode", "MGCode", 5)
        objUDFEngine.AddAlphaField("MICODE", "MGDesc", "MGDesc", 50)
        objUDFEngine.AddAlphaField("MICODE", "SGCode", "SGCode", 5)
        objUDFEngine.AddAlphaField("MICODE", "SGDesc", "SGDesc", 50)
        objUDFEngine.AddAlphaField("MICODE", "SSGCode", "SSGCode", 5)
        objUDFEngine.AddAlphaField("MICODE", "SSGDesc", "SSGDesc", 50)


        '*******************  Table ******************* START********************************* END
    End Sub
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        If eventInfo.BeforeAction Then
        Else
            If eventInfo.FormUID.Contains("MIREJDET") And (eventInfo.ItemUID = "13") And eventInfo.Row > 0 Then

                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try

                    If objAddOn.objApplication.Menus.Exists("ditem") Then
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    End If
                Catch ex As Exception

                End Try
                Try

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280").SubMenus.Item("ditem")
                    ZB_row = eventInfo.Row
                Catch ex As Exception
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "ditem"
                    oCreationPackage.String = "Delete Row"
                    oCreationPackage.Enabled = True

                    oMenuItem = objAddOn.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                    ZB_row = eventInfo.Row
                End Try
                If eventInfo.ItemUID <> "13" Then
                    '   Dim oMenuItem As SAPbouiCOM.MenuItem
                    '  Dim oMenus As SAPbouiCOM.Menus
                    Try
                        objAddOn.objApplication.Menus.RemoveEx("ditem")
                    Catch ex As Exception
                        ' MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
            End If
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)


    End Sub
    Public Sub WriteSMSLog(ByVal Str As String)
        Dim fs As FileStream
        Dim chatlog As String = System.Windows.Forms.Application.StartupPath & "\Log_" & Today.ToString("yyyyMMdd") & ".txt"
        If File.Exists(chatlog) Then
        Else
            fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
            fs.Close()
        End If
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub
    Private Sub addJobCardReporttype()
        'Dim rptTypeService As SAPbobsCOM.ReportTypesService
        'Dim newType As SAPbobsCOM.ReportType
        'Dim newtypeParam As SAPbobsCOM.ReportTypeParams
        'Dim newReportParam As SAPbobsCOM.ReportLayoutParams
        'Dim ReportExists As Boolean = False
        'Try


        '    Dim newtypesParam As SAPbobsCOM.ReportTypesParams
        '    rptTypeService = objAddOn.objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '    newtypesParam = rptTypeService.GetReportTypeList

        '    Dim i As Integer
        '    For i = 0 To newtypesParam.Count - 1
        '        If newtypesParam.Item(i).TypeName = clsJobCard.FormType And newtypesParam.Item(i).MenuID = clsJobCard.FormType Then
        '            ReportExists = True
        '            Exit For
        '        End If
        '    Next i

        '    If Not ReportExists Then
        '        rptTypeService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
        '        newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)


        '        newType.TypeName = clsJobCard.FormType
        '        newType.AddonName = "JC2Addon"
        '        newType.AddonFormType = clsJobCard.FormType
        '        newType.MenuID = clsJobCard.FormType
        '        newtypeParam = rptTypeService.AddReportType(newType)

        '        Dim rptService As SAPbobsCOM.ReportLayoutsService
        '        Dim newReport As SAPbobsCOM.ReportLayout
        '        rptService = objCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
        '        newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
        '        newReport.Author = objCompany.UserName
        '        newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
        '        newReport.Name = clsJobCard.FormType
        '        newReport.TypeCode = newtypeParam.TypeCode

        '        newReportParam = rptService.AddReportLayout(newReport)

        '        newType = rptTypeService.GetReportType(newtypeParam)
        '        newType.DefaultReportLayout = newReportParam.LayoutCode
        '        rptTypeService.UpdateReportType(newType)

        '        Dim oBlobParams As SAPbobsCOM.BlobParams
        '        oBlobParams = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
        '        oBlobParams.Table = "RDOC"
        '        oBlobParams.Field = "Template"
        '        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
        '        oKeySegment = oBlobParams.BlobTableKeySegments.Add
        '        oKeySegment.Name = "DocCode"
        '        oKeySegment.Value = newReportParam.LayoutCode

        '        Dim oFile As FileStream
        '        oFile = New FileStream(Application.StartupPath + "\JobCard.rpt", FileMode.Open)
        '        Dim fileSize As Integer
        '        fileSize = oFile.Length
        '        Dim buf(fileSize) As Byte
        '        oFile.Read(buf, 0, fileSize)
        '        oFile.Dispose()

        '        Dim oBlob As SAPbobsCOM.Blob
        '        oBlob = objCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
        '        oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
        '        objCompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
        '    End If
        'Catch ex As Exception
        '    objApplication.MessageBox(ex.ToString)
        'End Try

    End Sub

    Private Sub objApplication_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objApplication.LayoutKeyEvent

        ''BubbleEvent = True
        'If eventInfo.BeforeAction = True Then
        '    If eventInfo.FormUID.Contains(clsJobCard.FormType) Then
        '        objJobCard.LayoutKeyEvent(eventInfo, BubbleEvent)
        '    End If
        'End If
    End Sub
End Class


