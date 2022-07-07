Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dtSidebarMenu, dt_UserLogin, dt_File1, dt_Text1
Dim UploadPathDigisales

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DigisalesLib_Report.xlsx", "SCD0013_Upload List File.xlsx", "SCD0013")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_UserLogin, "Isi Field Keterangan : " & dt_Text1, "Nama File : " & dt_File1))
REM ------- Digisales
Call DA_Login()
call FR_GoTo_SidebarMenu(dtSidebarMenu)
call UploadFileDistribution()
Call DA_Logout("0")

Call OpenFile( UploadPathDigisales , dt_File1, "CHROME" )

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
	UploadPathDigisales = PathDigisales & "File_Upload\"
	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- Digisales lib
	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDigisales & "Digisales_FileDistribution.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Home_Digisales_Web.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Upload_File_List.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	
'	REM --- Transaction Library
'	LoadFunctionLibrary (LibPathDigisales & "Digisales46_Pembelian.qfl")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Merchant_Pembelian.tsr")
'	
'	REM --- Laporan Transaksi
'	LoadFunctionLibrary (LibPathDigisales & "Digisales46_LaporanTransaksi.qfl")
'	Call RepositoriesCollection.Add(LibRepo & "RP_LaporanTransaksi.tsr")
'	
'	REM --- Verifications Library
'	LoadFunctionLibrary (LibPathDigisales & "DigisalesLib_Verifikasi.qfl")
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_UserLogin				= DataTable.Value("USER",dtLocalSheet)
	dt_File1					= DataTable.Value("FILE1",dtLocalSheet)
	dt_Text1					= DataTable.Value("TEXT1",dtLocalSheet)
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
'	
'	REM ---------- Menu
	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU" ,dtLocalSheet)
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
'	REM ------  Verifications
'	noJurnal 					= DataTable.Value("OUT_NO_JURNAL", dtLocalSheet)
'	trxDate  					= DataTable.Value("OUT_TRX_DATE", dtLocalSheet)
'	noRek 						= DataTable.Value("NO_REKENING", dtlocalsheet)
End Sub
