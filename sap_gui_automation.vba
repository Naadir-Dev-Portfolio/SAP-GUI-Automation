
Sub Process_[COMPANY_REPORT_ID]()

        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        
    'User input
    Dim FileDate As String
    Dim FileName As String
        FileDate = InputBox("EnterDate(dd/mm/yyyy) ", "Enter the same date entered in BW (for the day required)")
        FileDate = Format(CDate(FileDate), "yyyymmdd")
        FileName = "Rant_and_Rave_Data_" & FileDate & ".csv"

    'Sap Variables
    Dim SapGui
    Dim Applic
    Dim connection
    Dim session
    Dim WSHShell
    Dim fso As Variant
    
    'Non Sap Variables
    Dim wb, wb2 As Workbook
    Dim Template, Start_Date, End_Date As String
    
        Template = "Z:\13. Activity\01. Daily\[COMPANY_REPORT_ID] - PRMI374 Rant and Rave Data (ER&R)\[COMPANY_REPORT_ID] - PRMI374 Rant and Rave Data (ER&R).xlsm"
        Set wb = Workbooks.Open(FileName:=Template)
            Start_Date = wb.Sheets("Config").Range("SD")
            End_Date = wb.Sheets("Config").Range("ED")
        '-------------------------------------------------
        Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
        Set WSHShell = CreateObject("WScript.Shell")
            Do Until WSHShell.AppActivate("SAP Logon ")
                Application.Wait Now + TimeValue("0:00:01")
            Loop
        Set WSHShell = Nothing

    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.OpenConnection("PR0 ECC", True) 'Change the text inbetween the quotes to look at your specific
    'connection name,
    Set session = connection.Children(0)
        'This will Log in to SAP
        session.findById("wnd[0]").maximize
        
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "Enter your SAP username here TOM" 'UsernameVar
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Guess what goes here" 'Password_Var
        session.findById("wnd[0]").sendVKey 0
        
        'Log in to Transaction Code ZMMD_IW65

        session.findById("wnd[0]/tbar[0]/okcd").Text = "IW49N"
        session.findById("wnd[0]").sendVKey 0
        
            session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select

            session.findById("wnd[1]/usr/txtV-LOW").Text = "[COMPANY_REPORT_ID]_SBTM"
            session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
            session.findById("wnd[1]/tbar[0]/btn[8]").press


            session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2").Select
            session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERDAT-LOW").Text = Start_Date
            session.findById("wnd[0]/usr/tabsTABSTRIP_TABBLOCK1/tabpS_TAB2/ssub%_SUBSCREEN_TABBLOCK1:RI_ORDER_OPERATION_LIST:1200/ctxtS_ERDAT-HIGH").Text = End_Date

            session.findById("wnd[0]/tbar[1]/btn[8]").press

            'Export Data
                session.findById("wnd[0]/mbar/menu[0]/menu[6]").Select
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/tbar[0]/btn[0]").press

                'Find Exported Data
                For Each w In Workbooks
                    If w.Name Like "Worksheet in Basis (1)*" Then w.Activate
                Next w
                
                Set wb2 = ActiveWorkbook
                    wb2.SaveAs FileName:="Z:\13. Activity\01. Daily\[COMPANY_REPORT_ID] - PRMI374 Rant and Rave Data (ER&R)\Source Data\SAP\Worksheet in Basis (1).xlsx"
                    wb2.Close False
                    
                    End_SAP
                    
                wb.Activate
                Application.Calculate
                
                'Refresh Lookup Queries
                wb.refreshall
                
                wb.Worksheets("Output").ListObjects("Output").QueryTable.Refresh BackgroundQuery:=False
                DoEvents
                'Now to export the data
            Dim wb3 As Workbook
            Set wb3 = Workbooks.Add(xlWBATWorksheet) '<-- new workbook with one sheet
            wb.Sheets("Output").Copy Before:=wb3.Sheets(1) '<-- put sheet into new workbook making it the first sheet
            wb3.Sheets("Output").Select
            wb3.Sheets("Sheet1").Delete '<-- delete the original sheet that we haven't used
            wb3.Sheets("Output").Range("A1").Select
            wb3.SaveAs FileName:=wb.Path & "\Export\" & FileName, FileFormat:=xlCSV
            wb3.SaveAs FileName:="https://cadentgasltd.sharepoint.com/sites/BusChg/MITeam/Reporting/" & _
                "Documents/PRMI%20Daily%20Reports/PRMI374%20-%20Daily%20Rant%20and%20Rave/" & _
                "PRMI374%20-%20Data/" & FileName, FileFormat:=xlCSV
            wb3.Close False
            wb.Close False
            
            Email[COMPANY_REPORT_ID]
            Send[COMPANY_REPORT_ID]CSVEmail
            
            LoadFZ
            
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub