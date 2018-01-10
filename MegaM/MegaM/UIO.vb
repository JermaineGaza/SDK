Public Class UIO
    'Dim oForm As SAPbouiCOM.Form
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

    Public Function EditTextText(oForm As SAPbouiCOM.Form, name As String) As String
        Dim oitem As SAPbouiCOM.Item
        Dim oEd As SAPbouiCOM.EditText
        oitem = oForm.Items.Item(name)
        oEd = oitem.Specific
        Return oEd.Value.ToString()
    End Function

    Public Function ComboBoxText(oForm As SAPbouiCOM.Form, cboxid As String) As String
        Dim oitem As SAPbouiCOM.Item
        Dim oComBox As SAPbouiCOM.ComboBox
        oitem = oForm.Items.Item(cboxid)
        oComBox = oitem.Specific
        Return oComBox.Selected.Value
    End Function

    Public Function ButtonCaption(oForm As SAPbouiCOM.Form, btnid As String) As String
        Dim oitem As SAPbouiCOM.Item
        Dim oBtn As SAPbouiCOM.Button
        oitem = oForm.Items.Item(btnid)
        oBtn = oitem.Specific
        Return oBtn.Caption.ToString()
    End Function

    Public Function EnableButtons(oForm As SAPbouiCOM.Form, names() As String) As Boolean 'E to enable D to disable
        Dim oitem As SAPbouiCOM.Item
        Dim oBtn As SAPbouiCOM.Button
        For Each name As String In names
            oitem = oForm.Items.Item(name)
            oBtn = oitem.Specific
            oBtn.Item.Enabled = True
        Next
        Return True
    End Function

    Public Function DisableButtons(oForm As SAPbouiCOM.Form, names() As String) As Boolean 'E to enable D to disable
        Dim oitem As SAPbouiCOM.Item
        Dim oBtn As SAPbouiCOM.Button
        For Each name As String In names
            oitem = oForm.Items.Item(name)
            oBtn = oitem.Specific
            oBtn.Item.Enabled = False
        Next
        Return True
    End Function

    Public Function EnableEditTexts(oForm As SAPbouiCOM.Form, names() As String) As Boolean 'E to enable D to disable
        Dim oitem As SAPbouiCOM.Item
        Dim oEd As SAPbouiCOM.EditText

        For Each name As String In names
            oitem = oForm.Items.Item(name)
            oEd = oitem.Specific
            oEd.Item.Enabled = True
        Next
        Return True
    End Function

    Public Function DisableEditTexts(oForm As SAPbouiCOM.Form, names() As String) As Boolean 'E to enable D to disable
        Dim oitem As SAPbouiCOM.Item
        Dim oEd As SAPbouiCOM.EditText

        For Each name As String In names
            oitem = oForm.Items.Item(name)
            oEd = oitem.Specific
            oEd.Item.Enabled = False
        Next
        Return True
    End Function


    Public Function EnableComboBoxes(oForm As SAPbouiCOM.Form, names() As String) As Boolean  'E to enable D to disable
        Dim oitem As SAPbouiCOM.Item
        'Dim oEd As SAPbouiCOM.EditText
        Dim oCb As SAPbouiCOM.ComboBox

        For Each name As String In names
            oitem = oForm.Items.Item(name)
            oCb = oitem.Specific
            oCb.Item.Enabled = True
        Next
        Return True
    End Function

    Public Function DisableComboBoxes(oForm As SAPbouiCOM.Form, names() As String) As Boolean  'E to enable D to disable
        Dim oitem As SAPbouiCOM.Item
        'Dim oEd As SAPbouiCOM.EditText
        Dim oCb As SAPbouiCOM.ComboBox

        For Each name As String In names
            oitem = oForm.Items.Item(name)
            oCb = oitem.Specific
            oCb.Item.Enabled = False
        Next
        Return True
    End Function

    Function WriteEditText(oForm As SAPbouiCOM.Form, name As String, value As String) As Boolean
        Dim oitem As SAPbouiCOM.Item
        Dim oEd As SAPbouiCOM.EditText
        oitem = oForm.Items.Item(name)
        oEd = oitem.Specific
        oEd.String = value
        Return True
    End Function

    Function WriteButton(oForm As SAPbouiCOM.Form, name As String, value As String) As Boolean
        Dim oitem As SAPbouiCOM.Item
        Dim oEd As SAPbouiCOM.Button
        oitem = oForm.Items.Item(name)
        oEd = oitem.Specific
        oEd.Caption = value
        Return True
    End Function

    Function AddMatrixRow(oForm As SAPbouiCOM.Form, matrixname As String) As Boolean
        Dim aMartix As SAPbouiCOM.Matrix
        Dim oitem As SAPbouiCOM.Item
        oitem = oForm.Items.Item(matrixname)
        aMartix = oitem.Specific
        aMartix.AddRow()
        Return True
    End Function

    Function RemoveMatrixRow(oForm As SAPbouiCOM.Form, matrixname As String) As Boolean
        Dim aMartix As SAPbouiCOM.Matrix
        Dim oitem As SAPbouiCOM.Item
        oitem = oForm.Items.Item(matrixname)
        aMartix = oitem.Specific
        For ii As Integer = 1 To aMartix.RowCount
            If aMartix.IsRowSelected(ii) Then
                aMartix.DeleteRow(ii)
            End If
        Next
        Return True
    End Function


End Class
