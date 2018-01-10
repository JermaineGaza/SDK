Option Strict Off
Option Explicit On

Imports SAPbouiCOM
Imports System.Threading
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.IO
Imports System.Globalization
Imports System.Collections.Generic

Friend Class MegaM

    'Global Variables to be used in ApplicationDim testOP As Boolean
    Dim comBox As SAPbouiCOM.ComboBox ' global combo for the audit pack
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Dim oForm As SAPbouiCOM.Form
    Dim oInvoiceDoc As SAPbobsCOM.Documents
    Dim oMatrix, aMatrix As SAPbouiCOM.Matrix
    Dim oColumns As SAPbouiCOM.Columns
    Dim oColumn As SAPbouiCOM.Column
    Dim oCell As SAPbouiCOM.Cell
    Dim oButton As SAPbouiCOM.Button
    Dim oCB As SAPbouiCOM.ComboBox
    Dim oStaticText As SAPbouiCOM.StaticText
    Dim oEditText, anEditText As SAPbouiCOM.EditText
    Dim creationPackage As SAPbouiCOM.FormCreationParams
    Dim oItem As SAPbouiCOM.Item
    Dim oLink As SAPbouiCOM.LinkedButton
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCons As SAPbouiCOM.Conditions
    Dim oCompany As SAPbobsCOM.Company
    Dim oCon As SAPbouiCOM.Condition
    Dim LRes, oRow As Integer
    Dim lRetCode, lErrCode, lErrCode2, lRetCode2, retVal As Long
    Dim dictionary As New Collections.Generic.Dictionary(Of String, String)
    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDataTable1, oDataTable2, oDataTable3, oDataTable4 As SAPbouiCOM.DataTable
    Dim count, rows As Integer
    Dim isDuplicate As Boolean
    Dim FileName, sErrMsg, sErrMsg2, AttachmentPath, SaveName As String
    Public Sub New()
        'Create new instance of MegaM and Setting Up of Application
        MyBase.New()
        SetApplication()
        AddMenuItem()

    End Sub
    Private Sub SetApplication()
        'Application and connecting to Business One
        Dim oSboGuiApi As New SAPbouiCOM.SboGuiApi
        Dim sConnStr, sCookie As String
        Try
            sConnStr = Environment.GetCommandLineArgs.GetValue(1)
            oSboGuiApi.Connect(sConnStr)
            SBO_Application = oSboGuiApi.GetApplication()
            oCompany = New SAPbobsCOM.Company()
            sCookie = oCompany.GetContextCookie
            sConnStr = SBO_Application.Company.GetConnectionContext(sCookie)
            retVal = oCompany.SetSboLoginContext(sConnStr)
            If (retVal <> 0) Then
                oCompany.GetLastError(retVal, sErrMsg)
                SBO_Application.StatusBar.SetText("Error SBO Login :" & retVal & ", Error: " & sErrMsg)
            End If
            retVal = oCompany.Connect()
            If (retVal <> 0) Then
                oCompany.GetLastError(retVal, sErrMsg)
                SBO_Application.StatusBar.SetText("Error Connecting to Company :" & retVal & ", Error: " & sErrMsg)
            Else
            End If
            count = 0
            isDuplicate = False
        Catch ex As Exception
            MsgBox("Exception : " & ex.Message)

        Finally
            oSboGuiApi = Nothing
        End Try
    End Sub
    Public Sub AddMenuItem()
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

        Try
            'Item(3072) is Inventory Module
            oMenuItem = SBO_Application.Menus.Item("3072")
            oMenus = oMenuItem.SubMenus
            oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            'If oMenus.Exists("Logistics") Then
            '    oMenuItem = SBO_Application.Menus.Item("Logistics")
            '    oMenus = oMenuItem.SubMenus
            'Else
            '    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            '    oCreationPackage.UniqueID = "Logistics"
            '    oCreationPackage.String = "Logistics Department"

            '    oCreationPackage.Position = 1

            '    Try
            '        oMenuItem = oMenus.AddEx(oCreationPackage)
            '        oMenus = oMenuItem.SubMenus
            '    Catch ex1 As Exception
            '        'SBO_Application.MessageBox("Error :" & ex1.Message)
            '    End Try
            'End If

            If oMenus.Exists("Login") Then
            Else
                Try

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "Login"
                    oCreationPackage.String = "Login"
                    oCreationPackage.Position = 1

                    oMenus.AddEx(oCreationPackage)

                Catch ex2 As Exception
                    'SBO_Application.MessageBox("Error :" & ex2.Message)
                End Try
            End If

        Catch ex As Exception
            'SBO_Application.MessageBox("Error :" & ex.Message)
        End Try

    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        'All SBO_Application menu events are handled here
        Dim UnID As String
        Dim Exists As Boolean

        If (pVal.MenuUID = "3072") And (pVal.BeforeAction = True) Then
            'This part is catching the duplicate event
            oForm = SBO_Application.Forms.ActiveForm()
            'Checks if record exists
            Exists = UDO_Exist(oForm)
            UnID = oForm.UniqueID
            If UnID.StartsWith("UDO") Then
                'Enters here if it is a UDO form
                oForm.Mode = BoFormMode.fm_OK_MODE
            End If
        
        ElseIf (pVal.MenuUID = "Login") And (pVal.BeforeAction = True) Then
            SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "Login", "")
        End If

    End Sub


    'definition forInit_UDO_Form() 

    Sub Init_UDO_Form(ByVal UDOformUID As String)
        'Start Init
        Dim obcm As SAPbouiCOM.ButtonCombo
        Dim oitem As SAPbouiCOM.Item

        'SBO_Application.MessageBox(UDOformUID)
        ' Adding the form
        oForm = SBO_Application.Forms.Item(UDOformUID)
        oForm.Freeze(True)
        oForm.EnableMenu("3072", True)
        Try
            oitem = oForm.Items.Item("Item_0")
            obcm = oitem.Specific
        Catch ex As Exception
        End Try

        Try
            ''to do remove this line and make more specific
            'obcm.ValidValues.Add("Sales Quotation", "")

            ' add udo specific copy to in forms 
            If UDOformUID.Contains("Log") Then
                obcm.ValidValues.Add("Login", "")
            End If

            '' add udo specific copy to in forms 
            'If UDOformUID.Contains("QAF60") Then
            '    obcm.ValidValues.Add("QAF/60F-U", "")
            'End If

            '' add udo specific copy to in forms 
            'If UDOformUID.Contains("MC14") Then
            '    obcm.ValidValues.Add("QAF1", "")
            'End If

            '' add udo specific copy to in forms 
            'If UDOformUID.Contains("MC1") Then
            '    obcm.ValidValues.Add("MC6", "")
            '    obcm.ValidValues.Add("MC7", "")
            'End If

        Catch ex As Exception
        End Try

        'Try

        '    oDataTable1 = oForm.DataSources.DataTables.Add("OADP")
        '    oDataTable3 = oForm.DataSources.DataTables.Add("@MC14HD")
        '    oDataTable1.ExecuteQuery("Select TOP 1 AttachPath from OADP")
        '    AttachmentPath = oDataTable1.GetValue(0, 0)

        'Catch ex As Exception

        'End Try
        oForm.Freeze(False)
    End Sub


    Private Sub SBO_Application_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent


        'This section catches all SBO_Application item events
        Dim EventEnum As SAPbouiCOM.BoEventTypes
        EventEnum = pVal.EventType

        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        EventEnum = pVal.EventType
        ' Check if this is a form of the UDO and Initialises the form
        If (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False) Then
            If FormUID.StartsWith("UDO") Then
                ' SBO_Application.MessageBox("Hello World!")
                Init_UDO_Form(FormUID)
            End If
        End If

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.Before_Action = False Then

            If pVal.ItemUID = "DriverName" Then
                oCFLEvento = pVal


                Dim aForm As SAPbouiCOM.Form
                Dim oED1 As SAPbouiCOM.EditText
                Dim oItem1 As SAPbouiCOM.Item
                Dim sCFL_ID As String
                Dim oDataTable As SAPbouiCOM.DataTable
                Dim val As String

                Try
                    val = ""
                    oDataTable = oCFLEvento.SelectedObjects
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    aForm = SBO_Application.Forms.Item(FormUID)
                    oItem1 = aForm.Items.Item("23_U_E")

                    oED1 = oItem1.Specific
                    Try

                        val = oDataTable.GetValue(1, 0)
                        oED1.Value = val
                    Catch ex As Exception
                    End Try

                Catch ex3 As Exception
                    SBO_Application.MessageBox(ex3.Message)
                End Try
            End If
        End If


        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.Before_Action = False Then

            If pVal.ItemUID = "DeliveryName" Then
                oCFLEvento = pVal


                Dim aForm As SAPbouiCOM.Form
                Dim oED1 As SAPbouiCOM.EditText
                Dim oItem1 As SAPbouiCOM.Item
                Dim sCFL_ID As String
                Dim oDataTable As SAPbouiCOM.DataTable
                Dim val As String

                Try
                    val = ""
                    oDataTable = oCFLEvento.SelectedObjects
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    aForm = SBO_Application.Forms.Item(FormUID)
                    oItem1 = aForm.Items.Item("24_U_E")

                    oED1 = oItem1.Specific
                    Try

                        val = oDataTable.GetValue(1, 0)
                        oED1.Value = val
                    Catch ex As Exception
                    End Try

                Catch ex3 As Exception
                    SBO_Application.MessageBox(ex3.Message)
                End Try
            End If
        End If

        If EventEnum = BoEventTypes.et_CHOOSE_FROM_LIST Then
            oCFLEvento = pVal

            If (pVal.ItemUID = "DriverName") And (pVal.Before_Action = False) Then
                Dim aForm As SAPbouiCOM.Form
                Dim oED1 As SAPbouiCOM.EditText
                Dim oItem1 As SAPbouiCOM.Item
                Dim sCFL_ID As String
                Dim oDataTable As SAPbouiCOM.DataTable
                Dim val As String

                Try

                    val = ""
                    oDataTable = oCFLEvento.SelectedObjects
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    aForm = SBO_Application.Forms.Item(FormUID)
                    oItem1 = aForm.Items.Item("EditTxt")
                    oED1 = oItem1.Specific

                    Try
                        val = oDataTable.GetValue(0, 0)
                    Catch ex As Exception

                    End Try


                    anEditText.Value = val

                Catch ex3 As Exception
                    SBO_Application.MessageBox(ex3.Message)

                    '    If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                    '    Else
                    '        oForm.Mode = BoFormMode.fm_UPDATE_MODE
                    '    End If
                    '    End if
                    'Catch ex As Exception


                End Try

            End If


        End If



    End Sub
    Function UDO_Exist(ByRef UDOForm As SAPbouiCOM.Form) As Boolean
        Dim itemDocEntry As SAPbouiCOM.Item
        Dim DocEntText As SAPbouiCOM.EditText
        Dim DocEnt, DocType As String
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData

        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService

        itemDocEntry = UDOForm.Items.Item("0_U_E")
        DocEntText = itemDocEntry.Specific
        DocEnt = DocEntText.Value
        DocType = UDOForm.BusinessObject.Type.ToString
        oCompanyService = oCompany.GetCompanyService
        Try
            oGeneralService = oCompanyService.GetGeneralService(DocType)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralParams.SetProperty("DocEntry", DocEnt)
        Catch
        End Try

        Try
            oGeneralData = oGeneralService.GetByParams(oGeneralParams)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Private Sub WriteToMatrix(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent)

        Dim Matrix2 As SAPbouiCOM.Matrix
        Dim Edit As SAPbouiCOM.EditText
        Dim anEditText1, anEditText2, anEditText3 As SAPbouiCOM.EditText


        Dim oDataTable As SAPbouiCOM.DataTable
        Dim DocType As String = ""
        Dim DocNum As String = ""
        Dim DocDate As String = ""
        Dim sCFL_ID As String
        Dim EventEnum As SAPbouiCOM.BoEventTypes
        Dim myCFLEvento As SAPbouiCOM.IChooseFromListEvent
        Dim acels As SAPbouiCOM.Cells
        Dim acel As SAPbouiCOM.Cell


        EventEnum = pVal.EventType
        Matrix2 = oForm.Items.Item("0_U_G").Specific
        Edit = oForm.Items.Item("Item_5").Specific
        myCFLEvento = pVal
        Dim myDataTable As SAPbouiCOM.DataTable


        Dim SystemOrUdo As String
        Dim edCheck As SAPbouiCOM.ComboBox = oForm.Items.Item("Item_4").Specific
        SystemOrUdo = edCheck.Selected.Description

        Try



            myDataTable = myCFLEvento.SelectedObjects
            sCFL_ID = myCFLEvento.ChooseFromListUID

            ' If SystemOrUdo.Contains("MC") Or SystemOrUdo.Contains("QAF") Then
            If SystemOrUdo.Contains("Login") Then

                DocType = myDataTable.GetValue("Object", 0)
                DocNum = myDataTable.GetValue("DocEntry", 0)
                DocDate = myDataTable.GetValue("CreateDate", 0)
            Else
                DocType = oForm.Items.Item("Item_4").Specific.Selected.Description
                DocNum = myDataTable.GetValue("DocEntry", 0)
                If SystemOrUdo.Contains("Licence1") Or SystemOrUdo.Contains("MEMORANDUM") Or SystemOrUdo.Contains("Cert") Or SystemOrUdo.Contains("Inspection Report") Or SystemOrUdo.Contains("PM/MRE/19") Or SystemOrUdo.Contains("PMMRE18") Or SystemOrUdo.Contains("POLE INSPECTION") Or SystemOrUdo.Contains("Sawn Timber Aridged") Then
                    DocDate = myDataTable.GetValue("CreateDate", 0)
                Else
                    DocDate = myDataTable.GetValue("TaxDate", 0)
                End If
            End If


            'Try
            '    Dim oEdit1 As SAPbouiCOM.EditText
            '    oEdit1 = Matrix2.Columns.Item("Col0").Cells.Item(Matrix2.RowCount).Specific
            '    oEdit1.Value = DocType.ToString()
            'Catch ex As Exception

            'End Try


            'Try
            '    Dim oEdit2 As SAPbouiCOM.EditText
            '    oEdit2 = Matrix2.Columns.Item("Col1").Cells.Item(Matrix2.RowCount).Specific
            '    oEdit2.Value = DocNum.ToString()
            'Catch ex As Exception

            'End Try

            'Try
            '    Dim oEdit3 As SAPbouiCOM.EditText
            '    oEdit3 = Matrix2.Columns.Item("Col2").Cells.Item(Matrix2.RowCount).Specific
            '    oEdit3.Value = DocDate.ToString()
            'Catch ex As Exception
            'End Try

            Matrix2.AddRow(1)


        Catch ex As Exception
            SBO_Application.MessageBox("niggerson" & ex.Message)
        End Try




    End Sub
    Private Sub SetProjectDocs(UDOformUID As String)
        Try
            oForm.DataSources.UserDataSources.Add("LinkedDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        Catch ex As Exception
        End Try

        Try
            oCFLs = oForm.ChooseFromLists
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "Login"
            oCFLCreationParams.UniqueID = "Login"
            oCFL = oCFLs.Add(oCFLCreationParams)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Valid_Values_For_Choose_List()


        Try
            Dim aValidValue As String

            Dim oED1 As SAPbouiCOM.EditText
            Try
                'oForm = SBO_Application.Forms.Item(UDOformUID)
                aValidValue = comBox.Selected.Description
            Catch ex As Exception
                SBO_Application.MessageBox(ex.Message)
            End Try

            Dim oED As SAPbouiCOM.EditText
            'oED = oForm.Items.Item("Item_5").Specific
            oED1 = oForm.Items.Item("Item_5").Specific
            oED1.DataBind.SetBound(True, "", "LinkedDS")

            'Adding the custom Choose From Lists
            If aValidValue = "MC14" Then
                oED1.ChooseFromListUID = "CFL2"
                oED1.ChooseFromListAlias = "DocNum"

                ' anEditText = oED


            ElseIf aValidValue = "MC5" Then
                oED1.ChooseFromListUID = "CFL3"
                oED1.ChooseFromListAlias = "DocNum"

                anEditText = oED
            ElseIf aValidValue = "MC6" Then
                oED1.ChooseFromListUID = "CFL4"
                oED1.ChooseFromListAlias = "DocNum"
            End If
        Catch e As Exception
            SBO_Application.MessageBox(e.Message)
        End Try
    End Sub

End Class
