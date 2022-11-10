Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dtSidebarMenu, dt_UserLogin, dt_File1, dt_NPP, dt_Periode, dt_Bulan
Dim DownloadPath, iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0011-010 - Validasi Field report PHR pada searching portal.xlsx", "SCD0011")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin, "Periode : " & dt_Bulan &" "& dt_Periode, "Nama File : " & dt_File1))
REM ------- Digisales
iteration = Environment.Value("ActionIteration")
Call DA_Login()
Call FR_GoTo_SidebarMenu(dtSidebarMenu)

If iteration = 1 Then
	Call GenerateProductHolding()
	Call DA_Logout("0")
Else
	Call GenerateProductHolding()
	Call ListProductHolding()
	Call ExportExcel()
	Call DA_Logout("0")
	REM ------- Open File EXCEL
	Call OpenLastDownloadFile(dt_File1, "EXCEL")
End If



Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDigisales, LibReport, LibRepo, objSysInfo
	Dim tempDigisalesPath, tempDigisalesPath2, PathDigisales
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDigisalesPath 	= Environment.Value("TestDir")
	tempDigisalesPath2 	= InStrRev(tempDigisalesPath, "\")
	PathDigisales 		= Left(tempDigisalesPath, tempDigisalesPath2)
	
	LibPathDigisales	= PathDigisales & "Lib_Repo_Excel\LibDigisales\"
	LibReport			= PathDigisales & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDigisales & "Lib_Repo_Excel\Repo\"
	DownloadPath		= "C:\Users\" & objSysInfo.UserName & "\Downloads"
	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_report.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Home_Digisales_Web.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Report.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")

End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin				= DataTable.Value("USER",dtLocalSheet)
	dt_File1					= DataTable.Value("FILE1",dtLocalSheet)
	dt_NPP						= DataTable.Value("TEXT1",dtLocalSheet)
	dt_Periode					= DataTable.Value("TEXT2",dtLocalSheet)
	dt_Bulan					= DataTable.Value("TEXT3",dtLocalSheet)
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
'	REM ---------- Menu
	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU" ,dtLocalSheet)
End Sub
