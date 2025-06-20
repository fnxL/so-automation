Option Explicit
Private ws As Worksheet

' Global SAP Objects
Private SapGuiAuto As Object
Private Application1 As Object
Private Connection As Object
Private session As Object

' SAP Key Constants
Private Const KEY_ENTER As Integer = 0
Private Const KEY_F3 As Integer = 3 ' F3 = Back
Private Const KEY_CTRL_S As Integer = 3 ' Ctrl+S = Save


' Constants
Private Const CONFIG_ORDER_TYPE As String = "AO4"
Private Const CONFIG_SALES_ORG As String = "AO5"
Private Const CONFIG_DIST_CHANNEL As String = "AO6"
Private Const SAP_DIVISION As String = "00"
Private Const CONFIG_ATTACH_PATH As String = "AJ1"
Private Const CONFIG_STAKE_HOLDER As String = "AH1"

' Column Constants
Private Const COL_PO_NUMBER As Integer = 1
Private Const COL_SO_NUMBER As Integer = 2
Private Const COL_SOLD_TO As Integer = 3
Private Const COL_SHIP_TO As Integer = 4
Private Const COL_PAYMENT_TERM As Integer = 5
Private Const COL_INCOTERM As Integer = 6
Private Const COL_INCOTERM_2 As Integer = 7
Private Const COL_END_CUSTOMER As Integer = 9
Private Const COL_CHANNEL_TYPE As Integer = 10
Private Const COL_SUB_CHANNEL_TYPE As Integer = 11
Private Const COL_ORDER_TYPE As Integer = 12
Private Const COL_SHIP_START As Integer = 13
Private Const COL_SHIP_CANCEL As Integer = 14
Private Const COL_PORT_OF_SHIPMENT As Integer = 15
Private Const COL_FINAL_DESTINATION As Integer = 16
Private Const COL_COUNTRY_DESTINATION As Integer = 17
Private Const COL_PORT_DISCHARGE As Integer = 18
Private Const COL_MATERIAL As Integer = 19
Private Const COL_QUANTITY As Integer = 20
' RUGS ONLY
Private Const COL_RUGS_SORT_NO As Integer = 22
Private Const COL_SHADE_NO_PD As Integer = 23
Private Const COL_PRINTING_SHADE_NO As Integer = 24
Private Const COL_PRODUCT_PACKING_TYPE As Integer = 26
Private Const COL_SET_NO As Integer = 28
Private Const COL_SHADE_NO_YD As Integer = 29
Private Const COL_DESTINATION As Integer = 30
Private Const COL_PLANT As Integer = 31
Private Const COL_CUSTOMER_MATERIAL As Integer = 32
Private Const COL_PIS As Integer = 33
Private Const COL_NOTIFY As Integer = 35
Private Const COL_PO_FILENAME As Integer = 36
Private Const COL_PO_FORMAT As Integer = 37

' Main Procedure
Public Sub RugsSalesOrders()
    Dim lastRow As Long
    Dim currentRow As Long
    Dim headerRow As Long

    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, COL_PO_NUMBER).End(xlUp).Row

    If Not InitSAPSession() Then
        MsgBox "Failed to connect to SAP." & vbCrLf & "Please ensure SAP GUI is running and you are logged in.", vbCritical, "SAP Connection Error"
        Goto CleanExit
    End If


    For currentRow = 7 To lastRow
        If ws.Cells(currentRow, COL_PO_NUMBER).Value <> "" And ws.Cells(currentRow, COL_SO_NUMBER).Value = 0 Then
            headerRow = currentRow
            CreateNewSaleOrder(headerRow)
        End If
    Next currentRow

    ' Processing completed
    CleanExit :
    Application.ScreenUpdating = True
    Set session = Nothing
    Set Connection = Nothing
    Set Application1 = Nothing
    Set SapGuiAuto = Nothing
    Exit Sub

    GlobalErrorHandler :
    MsgBox "An unexpected error occurred:" & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Description:  " & Err.Description & vbCrLf & vbCrLf & _
        "The macro will now stop.", vbCritical, "Critical Error"
    Goto CleanExit
End Sub

Private Sub CreateNewSaleOrder(ByVal headerRow as Long)
    ' Create SO (va01)
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva01"
    Call FillInitalOrgData()

    ' Overview Screen
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = ws.Cells(headerRow, COL_PO_NUMBER).Value
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = ws.Cells(headerRow, COL_SOLD_TO).Value
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = ws.Cells(headerRow, COL_SHIP_TO).Value
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").Text = Format(Date, "dd.mm.yyyy")
    session.findById("wnd[0]").sendVKey KEY_ENTER
    ' Incoterms 2
    If ws.Cells(headerRow, COL_INCOTERM_2).Value <> 0 Then
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/txtVBKD-INCO2").Text = ws.Cells(headerRow, COL_INCOTERM_2).Value
        session.findById("wnd[0]").sendVKey KEY_ENTER
        session.findById("wnd[0]").sendVKey KEY_ENTER
    End If

    Call FillHeaderData(headerRow)

    ' Overview Screen > Fast Data Entry (Fill Line Items)
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/cmbRV45A-MUEBS").Key = "ZWILCHAR"

    Dim lineItemRow As Long
    Dim lineItemCounter As Long
    Dim tempLineItemCounter as Long
    Dim linesInCurrentOrder As Long
    Dim gridRow As Long
    Dim savedSONumber As String

    lineItemRow = headerRow
    lineItemCounter = 0
    While ws.Cells(lineItemRow, COL_MATERIAL).Value <> 0
        If tempLineItemCounter > 0 And tempLineItemCounter Mod 9 = 0 Then
            session.findById("wnd[0]").sendVKey KEY_ENTER
            tempLineItemCounter = tempLineItemCounter + 1
        End If
        gridRow = tempLineItemCounter Mod 9
        ' Fill data for one line item
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MABNR[2," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_MATERIAL).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/txtRV45A-KWMENG[3," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_QUANTITY).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT01[6," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_RUGS_SORT_NO).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT02[7," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_SHADE_NO_PD).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT03[8," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_PRINTING_SHADE_NO).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT05[10," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_PRODUCT_PACKING_TYPE).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT07[12," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_SET_NO).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT08[13," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_SHADE_NO_YD).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtRV45A-MWERT09[14," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_DESTINATION).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtVBAP-WERKS[19," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_PLANT).Value

        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/tblSAPMV45ATCTRL_U_MILL_SE_KONFIG/ctxtVBAP-KDMAT[29," & gridRow & "]").Text = ws.Cells(lineItemRow, COL_CUSTOMER_MATERIAL).Value


        lineItemCounter = lineItemCounter + 1
        tempLineItemCounter = tempLineItemCounter + 1
        lineItemRow = lineItemRow + 1

    Wend
    linesInCurrentOrder = lineItemCounter
    session.findById("wnd[0]").sendVKey KEY_ENTER

    Call AttachPIS(headerRow)
    Call HitEnter(linesInCurrentOrder)

    Dim statusBarMessage As String
    statusBarMessage = session.findById("wnd[0]/sbar").Text
    savedSONumber = Split(statusBarMessage, " ")(3)

    ' Update Excel
    If IsNumeric(savedSONumber) Then
        For lineItemRow = headerRow To headerRow + linesInCurrentOrder - 1
            ws.Cells(lineItemRow, COL_SO_NUMBER).Value = savedSONumber
        Next lineItemRow
    End If

    ' Attach PO Copy
    If ws.Cells(headerRow, COL_PO_FILENAME).Value <> "" And ws.Cells(headerRow, COL_PO_FORMAT).Value <> "" Then
        AttachFile savedSONumber, _
            CStr(ws.Range(CONFIG_ATTACH_PATH).Value), _
            CStr(ws.Cells(headerRow, COL_PO_FILENAME).Value), _
            CStr(ws.Cells(headerRow, COL_PO_FORMAT).Value)
    End If
End Sub

Private Sub AttachPIS(ByVal headerRow As Long)
    ' Just save if no PIS
    If Trim(ws.Cells(headerRow, COL_PIS).Value) = "" Then
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        Exit Sub
    End If
    ' Select All Button
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:7901/subSUBSCREEN_TC:SAPMV45A:7905/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKAL").press
    ' Extras Menu
    session.findById("wnd[0]/mbar/menu[3]/menu[10]").Select
    session.findById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,0]").text = "PIS"
    session.findById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKNR[1,0]").text = ws.Cells(headerRow, COL_PIS).Value
    session.findById("wnd[1]").sendVKey KEY_ENTER
End Sub

Private Sub HitEnter(ByVal linesInCurrentOrder As Long)
    Dim i As Long
    For i = 0 To linesInCurrentOrder - 1
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    Next i
End Sub

Private Sub AttachFile(ByVal SONumber As String, ByVal filePath As String, ByVal fileName As String, ByVal fileExtension As String)
    ' Open SO Document
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
    session.findById("wnd[0]").sendVKey KEY_ENTER
    Application.Wait(Now + TimeValue("00:00:03"))

    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = SONumber
    session.findById("wnd[0]").sendVKey KEY_ENTER

    session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
    session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = filePath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName & "." & fileExtension
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press

End Sub

Private Sub FillHeaderData(ByVal headerRow As Long)
    ' Open Header Details
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press

    ' Header > Partners > Set End Customer
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").Key = "ZE"
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").Text = ws.Cells(headerRow, COL_END_CUSTOMER).Value
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").SetFocus
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,4]").caretPosition = 6
    session.findById("wnd[0]").sendVKey KEY_ENTER

    ' Header > Texts Tab > Notify/Remitter
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "Z041", "Column1"
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "Z041", "Column1"
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "Z041", "Column1"
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = ws.Cells(headerRow, COL_NOTIFY).Value

    ' Header > Additional Data A Tab
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").Select
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4309/ctxtVBAK-ZECOM").Text = ws.Cells(headerRow, COL_CHANNEL_TYPE).Value
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4309/ctxtVBAK-ZSCTYP").Text = ws.Cells(headerRow, COL_SUB_CHANNEL_TYPE).Value
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-ZORD_TYPE").Key = ws.Cells(headerRow, COL_ORDER_TYPE).Value

    ' Header > Additional Data B Tab
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13").Select

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_STARTDT").Text = Format(ws.Cells(headerRow, COL_SHIP_START).Value, "dd.mm.yyyy")

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_CANDT").Text = Format(ws.Cells(headerRow, COL_SHIP_CANCEL).Value, "dd.mm.yyyy")

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_EXFACDT").Text = Format(ws.Cells(headerRow, COL_SHIP_START).Value - 1, "dd.mm.yyyy")

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_HANODT").Text = Format(ws.Cells(headerRow, COL_SHIP_START).Value + 1, "dd.mm.yyyy")

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_POS").Text = ws.Cells(headerRow, COL_PORT_OF_SHIPMENT).Value

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_FDEST").Text = ws.Cells(headerRow, COL_FINAL_DESTINATION).Value
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtVBAK-ZZSD_CFDEST").Text = ws.Cells(headerRow, COL_COUNTRY_DESTINATION).Value
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSD_POD").Text = ws.Cells(headerRow, COL_PORT_DISCHARGE).Value

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZSTAKE_HOLDER").Text = ws.Range(CONFIG_STAKE_HOLDER).Value

    ' Go back to overview screen
    session.findById("wnd[0]").sendVKey KEY_F3
End Sub


Private Sub FillInitalOrgData()
    session.findById("wnd[0]").sendVKey KEY_ENTER
    session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = ws.Range(CONFIG_ORDER_TYPE).Value
    session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = ws.Range(CONFIG_SALES_ORG).Value
    session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = ws.Range(CONFIG_DIST_CHANNEL).Value
    session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = SAP_DIVISION
    session.findById("wnd[0]").sendVKey KEY_ENTER
End Sub


Private Function InitSAPSession() As Boolean
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application1 = SapGuiAuto.GetScriptingEngine
    Set Connection = Application1.Children(0)
    Set session = Connection.Children(0)
    session.findById("wnd[0]").maximize

    If session Is Nothing Then
        Err.Raise vbObjectError + 1000, , "SAP connection failed. Please ensure SAP GUI is running and you are logged in."
        InitSAPSession = False
    End If
    InitSAPSession = True
End Function

