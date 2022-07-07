Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dtSidebarMenu, dtSidebar_SubMenu, dtSidebar_Submenu_Submenu, dt_Hostname
Dim  noRek, noJurnal, trxDate, Openprogram

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0012_Upload File ke Storage.xlsx", "SCD0012")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Hostname : " & dt_Hostname))

REM ------- Filezilla
Call Login_Filezilla()
Call Create_SiteManager()
Call UploadStorage()
'call ExecuteSQL()
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

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_Filezilla.qfl")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Heidi.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Hostname					= DataTable.Value("HOSTNAME",dtLocalSheet)
'	dt_Database1				= DataTable.Value("DATABASE_1",dtLocalSheet)
'	dt_Database2				= DataTable.Value("DATABASE_2",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
'	
'	REM ---------- Menu
'	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU" ,dtLocalSheet)
'	dtSidebar_toSubmenu			= DataTable.Value("SIDEBAR_SUBMENU" ,dtLocalSheet)
'	dtSidebar_Submenu_Submenu	= DataTable.Value("SIDEBAR_SUBMENU_SUBMENU", dtLocalSheet)
'	dtMenu_Merchant_Pembelian	= DataTable.Value("MENU_MERCHANT_PEMBELIAN" ,dtLocalSheet)
'
'	REM ---- Transaksi
'	jenisPembelian				= DataTable.Value("JENIS_PEMBELIAN_PLN" ,dtLocalSheet)
'	idPelanggan					= DataTable.Value("ID_PELANGGAN" ,dtLocalSheet)
'	NoReff						= DataTable.Value("NO_REFF" ,dtLocalSheet)
'	Nominal						= DataTable.Value("NOMINAL" ,dtLocalSheet)
'	PINTransaksi				= DataTable.Value("PIN_TRX" ,dtLocalSheet)
'	
End Sub
