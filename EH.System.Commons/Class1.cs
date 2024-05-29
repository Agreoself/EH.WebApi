// REM HR AD sync console program, by Changtu Wang in Xiamen, Aug. 2014
using Microsoft.VisualBasic;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices; // import 微软自带的AD操作相关的类
using System.Net; // 网络操作函数
using System.Net.Mail; // 邮件调用相关
using System.Text;
using System.Text.RegularExpressions; // text下的子模块

public static class Module1
{
    private static string forXM_US = ""; // "US" "XM"
    private static bool forLeaveOnly = false; // for LOA set expired / remove expired only, don't send others alert email
    private static bool forTest = true; // for ehiAD call need set True
    // private ITDBname As String = "ITauto"
    private static string ITDBname = "ITtest";
    private static var sftpIncomingDirectory = "/workday/"; // "/incoming/"
    private static string HRsftpFileName = "ehealth3.csv"; // '"ehealth.txt" ' "ehealth2.csv"
    private static bool APIorManualUpload = true;
    private static bool needDownloadHRfile = true; // for ehiAD call need set False
    private static var HRfileForTestName = "WD0509.csv"; // "ehealth3.csv"
    private static bool needSaveEEinfo2DB = true;
    private static bool END_afterSaveEEinfo2DB = false;

    private static bool needStop = true; // Need to Stop after going live with Workday+Okta Automation
    private static bool needQueryAD = true; // for ehiAD call need set True
    private static var allUsersForTestName = "ADallusers20220509_1055.csv";
    private static bool needSendAlertEmail = true; // for ehiAD call need set False
    private static bool needRecordInDB = true; // for ehiAD call need set False
    private static bool needSendADfileToHR = false; // for ehiAD call need set False, RFS-136351, Vishal email sent 04/07/2020: send to Gaja
    private static bool needCheckNewUsersNotLogon = false;
    private static bool needShowNoHRrecord = false;
    private static bool needUploadS3 = true;
    private static bool needOomnitzaReport = false;
    private static bool needCoupaReport = false;

    private static var appNameVersion = "WorkdayAD2"; // System.Reflection.Assembly.GetEntryAssembly.GetName.Name   Process.GetCurrentProcess.MainModule.FileName
    private static var appNameAddUser = "addUser0.9b.exe";
    private static string rootFullName = "Francis Soistman";
    private static string CEO = "Fran Soistman";
    private static var serviceAccount = "prisminst";
    private static string entOpsEmail = "EntOps@ehealthinsurance.com";
    private static string CSVfilenameOomnitza = "OomnitzaAD.csv";
    private static string CSVfilenameCoupa = "CoupaActiveDirectoryUsers.csv";
    private static int DaysCanRemoveManager = 3;
    private static var OnLeaveCheckDays = 3;
    private static var OnLeaveIgnoreTimesAllowed = 2;
    private static string inactiveYesStr = "Yes";
    private static int daysInactiveLimit = 30;
    private static int DaysInactiveLimitNewHire = 7;
    private static int timeDifferenceUT = 1;
    private static int timeDifferenceMA = 3;
    private static int timeDifferenceXM = 16; // PDT: 15, PST: 16 'if isPDT(), timeDifferenceXM -= 1
    private static int outOfDate_ThresholdHours = 12; // 2
    private static bool needSetLOApwdNeverExpires = false;
    private static int LOApwdNeverExpires_ThresholdDays = 7;
    private static int noShow_ThresholdDays = 28;
    private static bool needDisableAtEndOfLastWorkingDay = true; // CTPR-9 Penny approved
    private static bool needCompareAccountName = false;
    // private CPuseWDID As Boolean = True

    private static bool needSetDisableForInactive = true; // ENGR-286826 PCI-Strengthen existing process to remove/disable inactive AD user accounts (8.1.4) 
    private static string noNeedAccountValueHR = "N";

    private static string jobTitleUpdateDelayLocationStr = ".XM.";
    private static int jobTitleUpdateDelayDays = 19; // ENGR-288803 Delay Title Change for AD sync
    private static string jobTitleUpdateDelayEndDayStr = "06/28/2015";
    private static bool needUpdateJobTitleDelay = false;

    private static bool needUpdatePhoneNotes = false; // RFS-56571 for copy EEID to show on Outlook properties tab
    private static bool needUpdaeTerminatedProperties = false;
    private static bool needUpdateDepartmentCode = true; // RFS-202979
    private static bool needUpdateDivisionCode = true; // RFS-202979

    private static string hrMIS = "Workday"; // "WFN" '"HRB"
    private static var ftpDownloadFolder = @"C:\ehiAD\hrb\ftp_dnload";
    private static var ftpEEdataFolder = @"C:\ehiAD\hrb\ee_data";
    private static var sftpBatch = "hrSFTP.txt";
    private static string ftpFileName = hrMIS; // "ehealth_standard_empoyee" ' "eHealth_StdEmp_Conn"
    private static var strWorkFolder = @"C:\ehiAD";
    private static var strScriptFolder = "script";
    private static var strEXEfolder = "exe";
    private static var strJIRAfolder = "JIRA";
    private static var strReportSubFolder = "report";
    private static var strLogSubFolder = "log";
    private static var strDataSubfolder = "data";
    private static var strHRsubFolder = "hrb";
    private static var strFTP_DNLOADsubFolder = "ftp_dnload";
    private static var strEE_DATAsubFolder = "ee_data";
    private static var strPrimeSubfolder = "Prime";
    private static var strTempfolder = "temp";
    private static var NoNeedADfileName = strWorkFolder + @"\" + strDataSubfolder + @"\NoNeedAD.csv";
    private static string minor = "minor";
    private static string major = "major";
    private static var strScope = "subtree";
    private static var defaultOU = ""; // "OU=EHI,"
    private static string SeasonalGroupName = ""; // "sdl-okta-jobvite-seasonal"

    private static string reportTalbeName = "ADuserReport";
    private static string onLeaveTalbeName = "ADonLeave";
    private static string contactITtalbeName = "ADcontactIT";
    private static string disabledTalbeName = "ADdisabled";
    private static string logFilename, errStr;

    // Dim jobChangeTalbeName As String = "ADjobChange"
    private static bool SQLDBOK = false;
    private static bool hasRights = false;
    private static string generalReprotTodayFilename = strWorkFolder + @"\" + strTempfolder + @"\" + "ReportToday.csv";
    private static string runnerName = "";
    private static DataTable tableTMP = new DataTable();

    private static bool needClearEmployeeNumber = false;
    private static bool needUpdateEmployeeID = false; // normally don't update employeeID, but may need update it for new hire with requisitionID
    private static bool needUpdateAccountControl = true;
    private static bool needUpdateMOCaccount = false;
    private static bool needUpdateMailbox = false;

    private static bool needUpdateLegalFirstName = true; // RFS-140202
    private static bool needUpdateFirstName = true;
    private static bool needUpdateMiddleName = true;
    private static bool needUpdateLastName = false;
    private static bool needUpdateDisplayName = true;

    private static bool needUpdateJobTitle = true;
    private static bool needUpdateHomeDeptment = true;
    private static bool needUpdateUserSegment = true;
    private static bool needRemoveManagerForTerminated = true;
    private static bool needHideEmailForTerminated = false; // for hiding email address to show on Outlook properties tab
    private static bool needUpdateManager = true;
    private static bool needUpdateCompany = true;
    private static bool needAutoSetReqID = true;

    private static bool needUseHRstreetAddress = false;
    private static bool needUpdateAddress = true;
    private static bool needUpdateDescription = false;
    private static bool needUpdateOffice = true;
    private static bool needUpdatePhone = false;
    private static bool needUpdateMobile = true; // CTEWS-4653
    private static bool needUpdateFax = false;
    private static bool needUpdateScriptPath = false;
    private static bool needUpdateHomeDirectory = false;
    // Dim needUpdateCostCenter As Boolean = False
    private static string needUpdateCostcenterLocationStr = ".XM.";
    private static bool needUpdateOthers = false;
    private static bool needCreateADaccount = false;
    private static bool needEmailMismatchReport = false;
    private static bool needEmailNamingConventionMismatchReport = false;
    private static bool needPrimaryEmailMismatchReport = false; // set True if need primary email address mismatch report (including email domain mismatch) 'Vishal email Sent: June 11, 2019
    private static string needDiscrepancyEmailDomainLocationStr = "";
    private static string NoNeedDiscrepancyNewHireEmailLocationStr = ".RMT.";
    private static bool needListFTEnonFTEgroupMemberShip = false;
    private static string ignoreOnLeaveEEIDs = "";
    private static string ignoreUpdateEEIDs = "";
    private static string actualIgnoredEEIDs = "";
    private static string alreadyExpiredForLeave = "";
    private static string alreadyExpiredForInactivate = "";
    private static string hasPreferredLastnameEEIDs = "";
    public static string svctestSamIDs = ""; // for ehiAD call
    public static string NONsvctestSamIDs = ""; // for ehiAD call
    public static string noShowSamIDs = ""; // for ehiAD call
    public static string SOXsamIDs = ""; // for ehiAD call
    public static string ignoreRehireSamIDs = "";

    private static var strTitle = "employeeID,previousID,samAccountName,DistinguishedName,LegalFirstName_AD,FirstName_AD,LastName_AD,MiddleName_AD,Status_AD,JobTitle_AD,HomeDepartment_AD,WorkEmail_AD,WhenCreated_AD,WhenChanged_AD,LastLogonTimeStamp_AD,employeeType_AD,City_AD,Manager_AD,managerID_AD,managerID,ManagerDN,ManagerWorkEmail,UserSegment,PrimaryWorkMobile,FirstName,MiddleName,LastName,NickName,Jobtitle,JobEffDate,BusinessUnit,HomeDepartment,Status,statusEffDate,EmpClass,OriginalHireDate,HireDate,Address,City,State,PostalCode,DivisionCode,LocationAux1,Office,CostCenter,NeedAccount,Name_mismatch,FTE_mismatch,Title_mismatch,Department_mismatch,Manager_mismatch,eeID_mismatch,samID_mismatch";
    private static string Name_mismatch, FTE_mismatch, Title_mismatch, Department_mismatch, Manager_mismatch, noteStrEmployeeID, noteStrSamAccountName, noteStr;
    private static string EEIDincorrect = ".";
    private static int maxAllowedRecipient = 20;
    private static int maxAllowedAttachment = 10;

    private static string[] SendTo = new string[21], ccTo = new string[21], bccTo = new string[21];
    // Dim FileAttach(maxAllowedAttachment) As String
    private static string EntOps = "";

    private static int ADS_PROPERTY_CLEAR = 1;
    private static int ADS_PROPERTY_UPDATE = 2;
    private static int ADS_PROPERTY_APPEND = 3;
    private static int ADS_PROPERTY_DELETE = 4;
    private static int TS_SESSION_DISCONNECT = 0;
    private static int TS_SESSION_END = 1;
    private static int TS_SESSION_ANY_CLIENT = 0;
    private static int TS_SESSION_ORIGINATING = 1;
    private static int ADS_UF_ACCOUNTDISABLE = 0x2;
    private static int ADS_UF_PASSWD_CANT_CHANGE = 0x40;
    private static int ADS_UF_DONT_EXPIRE_PASSWD = 0x10000;
    private static int ADS_UF_PASSWORD_EXPIRED = 0x800000;
    private static int ADS_UF_ENCRYPTED_TEXT_PASSWD = 0x80;
    private static int ADS_UF_SMARTCARD_REQUIRED = 0x40000;
    private static int ADS_UF_ACCOUNT_TRUSTED = 0x80000;
    private static int ADS_UF_ACCOUNT_SENSITIVE = 0x100000;
    private static int ADS_UF_DES_ENCRYPTION = 0x200000;
    private static int ADS_UF_KERBEROS_PREAUTH = 0x400000;

    private static int ADS_GROUP_TYPE_GLOBAL_GROUP = 0x2;
    private static int ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = 0x4;
    private static int ADS_GROUP_TYPE_LOCAL_GROUP = 0x4;
    private static int ADS_GROUP_TYPE_UNIVERSAL_GROUP = 0x8;
    private static uint ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000;

    private static int HRrowCount, intRow, intRowCount, intOKcount, intErrorCount, intNeedDisabledCount, intNotFoundCount, intUAC, index;
    private static string lineStr, tmpStr, arg1, arg2, arg3, arg4, arg5, arg6, arg7;

    private static object objRootDSE, fileAllADusers, connQueryAllUsers, cmdQueryAllUsers, rsQueryAllUsers;
    private static object connQueryManager, cmdQueryManager, rsQueryManager;
    private static object objField, objConnection, objCommand, objRecordSet;
    private static object objOU, objGroup, objUser, objManager;

    private static string strRoot, strDomain, strLDAP, strfilter;
    private static string[] OU = new string[6], CN = new string[6];
    private static string newCN, strCNnet;
    private static int intOUcount, intCNcount, correctCount;

    private static string Jobtitle, HomeDepartment, DepartmentAux1, OriginalHireDate, HireDate, EmpClass, AllUsersCSVFileName, JobEffDate, BusinessUnit, UserSegment;
    private static string Workemail, DistinguishedName, CNname, CNOUkey, samAccountName, employeeID, FirstName, LastName, MiddleName, NickName, DisplayName;
    private static string wkdID, preID, reqID, ADusersHRfilenameCSV, ADusersHRfilenameXLS, ADusersHRactiveCSV;
    private static string Status, LocationAux1, ManagerID, ManagerDN, ManagerWorkEmail;
    private static string strOUnet, FTEgroup, nonFTEgroup, newHireEmail, statusEffDateStr, leaveReturningEffDateStr;
    private static bool needDisabledForInactive, needDisabledForNotLogon;
    private static string Description, PhysicalDeliveryOfficeName, ipPhone, ProfilePath, ScriptPath, HomeDirectory, HomeDrive, HomePhone, CostCenter;
    private static string Workphone, Mobile, Fax, Company, Address, PObox, PostalCode, City, State, Country, TPA_location_vendor;
    private static string PhoneNotes, wWWHomePage, otherTelephone, URL, AccountExpirationDate, memberOf;
    private static string defaultGroup, MOC_enabled, onLeaveSamIDs, HR_AccountName;

    private static string samID2, eeID2, status2, DistinguishedNameFound, samAccountNameFound, samAccountNameFound2, employeeIDfound;
    private static string NeedAccount, DivisionCode, strExportFileNewUser, strReportfileDiscrepancy, strReportfileEmailMismatch;
    private static string firstNameAD, lastNameAD, middleNameAD, cityAD, PhysicalDeliveryOfficeNameAD, jobtitleAD, homeDepartmentAD, employeeidAD, statusAD, WorkemailAD, whenCreated, whenChanged, LastLogonTimeStamp, managerAD;
    private static string LegalFirstNameAD, reqidAD, proxyAddressesAD, employeeTypeAD;
    private static string strExportFileEntire, strExportFileFound, strExportFileNotFound, strExportFileManager, strExportFileNeedDisabled, strReportfileName;
    private static string strExportFileGroupAction, strReportfileNameAll, strReportPrime, strReportPrimeFailed, strReportCoupa, strReportOomnitza;
    private static StreamWriter reportfileEntire, reportfileFound, reportfileNotFound, reportfileManagerEmployeeID, reportfileNeedDisabled;
    private static StreamWriter reportfileGroupAction, reportfileDiscrepancy, reportfileEmailMismatch;
    private static int discrepancyCount, discrepancyCountXM, emailMismatchCount, emailMismatchCountXM;
    private static int HRneedTerminatedCount = 0;

    private static ArrayList LeaveReturnEEID = new ArrayList(), DB_OnLeaveEEID = new ArrayList();
    private static ArrayList DB_OnLeaveSamID = new ArrayList(), DB_DisabledSamIDs = new ArrayList();
    private static ArrayList DB_OnLeaveEffDate = new ArrayList(), DB_LeaveReturningSamID = new ArrayList(), DB_LeaveReturningEffDate = new ArrayList();

    private static ArrayList HR_eeID = new ArrayList(), HR_firstName = new ArrayList(), HR_lastName = new ArrayList(), HR_middleName = new ArrayList(), HR_nickName = new ArrayList(), HR_workEmail = new ArrayList(), HR_workEmailKey = new ArrayList(), HR_status = new ArrayList(), HR_managerWorkEmail = new ArrayList(), HR_locationAux1 = new ArrayList(), HR_managerID = new ArrayList();
    private static ArrayList WDID = new ArrayList(), previousID = new ArrayList(), duplicatedWDID = new ArrayList();
    private static ArrayList HR_LocationCode = new ArrayList(), HR_DivisionCode = new ArrayList(), HR_CompanyCode = new ArrayList(), HR_ManagementLevel = new ArrayList();
    private static ArrayList HR_needAccount = new ArrayList(), emailCClist = new ArrayList(), HR_statusEffDate = new ArrayList(), HR_needDisabled = new ArrayList(), HR_UserSegment = new ArrayList();
    private static ArrayList HR_AddressStreet2 = new ArrayList(), HR_AddressStreet3 = new ArrayList(), HR_TPA_location_vendor = new ArrayList();
    private static ArrayList HR_eeID0 = new ArrayList(), HR_FTE0 = new ArrayList(); // for locate managerID when updating (HR_eeID will be reset when doing property update)
    private static int colHR_eeID, colHR_firstName, colHR_lastName, colHR_middleName, colHR_nickName, colHR_workEmail, colHR_status, colHR_managerWorkEmail, colHR_locationAux1, colHR_managerID;
    private static int colHR_NeedAccount, colHR_CostCenter, colHR_statusEffDate, colHR_needDisabled;
    private static int colHR_location, colHR_AddressStreet2, colHR_AddressStreet3, colHR_TPA_location_vendor;
    private static string TPA_vendor, TPA_location;
    private static ArrayList HR_DistinguishedName = new ArrayList(), HR_samAccountName = new ArrayList(), AD_status = new ArrayList(), HR_Jobtitle = new ArrayList(), HR_HomeDepartment = new ArrayList(), HR_managerDN = new ArrayList(), HR_Address = new ArrayList(), HR_PostalCode = new ArrayList(), HR_City = new ArrayList(), HR_State = new ArrayList(), HR_Country = new ArrayList();
    private static ArrayList HR_Workphone = new ArrayList(), HR_mobile = new ArrayList(), HR_Fax = new ArrayList(), HR_office = new ArrayList(), HR_ProfilePath = new ArrayList(), HR_FTE = new ArrayList(), HR_CostCenter = new ArrayList(), HR_OriginalHireDate = new ArrayList(), HR_HireDate = new ArrayList(), HR_JobEffDate = new ArrayList(), HR_BusinessUnit = new ArrayList();
    private static ArrayList LastLogonTimeStampList = new ArrayList();
    private static int colHR_DistinguishedName, colHR_samAccountName, colAD_status, colHR_Jobtitle, colHR_HomeDepartment, colHR_managerDN, colHR_Address, colHR_PostalCode, colHR_City, colHR_State, colHR_Country;
    private static int colHR_DepartmentAux1, colHR_Workphone, colHR_Fax, colHR_office, colHR_ProfilePath, colHR_FTE, colHR_OriginalHireDate, colHR_HireDate, colHR_JobEffDate;
    private static int colHR_RehireDate, colHR_BusinessUnit;

    private static ArrayList eeIDlist = new ArrayList(), samIDlist = new ArrayList(), firstNameList = new ArrayList(), lastNameList = new ArrayList(), middleNameList = new ArrayList(), userDNlist = new ArrayList(), statusList = new ArrayList(), samIDs = new ArrayList(), eeIDs = new ArrayList(), statusS = new ArrayList(), indexList = new ArrayList();
    private static ArrayList HR_reqID = new ArrayList(), reqIDlist = new ArrayList(), inactiveDaysList = new ArrayList(), inactiveList = new ArrayList(), legalFirstNameList = new ArrayList();
    private static bool needDisableInactive, needDisableNotLogonByStartdate;
    private static ArrayList jobtitleList = new ArrayList(), HomeDepartmentList = new ArrayList(), WorkemailList = new ArrayList(), WorkemailListKey = new ArrayList(), whenCreatedList = new ArrayList(), whenChangedList = new ArrayList(), PhysicalDeliveryOfficeNameList = new ArrayList(), cityList = new ArrayList(), managerList = new ArrayList();
    private static ArrayList proxyAddressesList = new ArrayList(), employeeTypeList = new ArrayList();

    private static bool hasError, hasWriteEntireError, samIDprovided, accountFound, samID2hasActive;
    private static bool needUpdateProperties, partialNeedUpdated, hasUpdated, isXM;

    private static bool hasWorkEmail = false;
    private static bool hasManagerID = false;
    private static bool hasManagerWorkEmail = false;
    private static bool hasJobTitle = false;
    private static bool hasHomeDepartment = false;
    private static bool emailChanged = false;
    private static bool hrFileReady = false;

    private static StreamWriter swriterGeneralToday, swReport, swLog;
    private static string strReport;
    private static string strGeneral = null;
    private static ArrayList NoNeedAD_EEIDs = new ArrayList();

    private static ArrayList FTEnonFTEgroupNamesList = new ArrayList(), memberOfList = new ArrayList();
    private static SearchResultCollection ehiDCs;

    private static string waitNoteStr = "Press a key to exit the program...";
    private static string newHireNotLogonStr = "new hire/rehire not logon over ";
    private static ArrayList HRactiveEEIDlist = new ArrayList(), HRactiveREQIDlist = new ArrayList(), HRterminatedEEIDlist = new ArrayList(), HRterminatedREQIDlist = new ArrayList();

    // Main 函数为包含主要的程序逻辑

    public static void Main(string[] args)
    {
        hasWriteEntireError = false;
        arg1 = "";
        arg2 = "";
        arg3 = "";
        if (args.Length > 0)
        {
            arg1 = Strings.LCase(Strings.Trim(args[0]));
            if (args.Length > 1)
            {
                arg2 = Strings.LCase(Strings.Trim(args[1]));
                if (args.Length > 2)
                    arg3 = Strings.LCase(Strings.Trim(args[2]));
            }
        }

        if ((Strings.InStr(arg1, "?") > 0))
        {
            Console.WriteLine("This program need run in Command Prompt mode.");
            // Console.WriteLine("If need create account, please input 1 parameters: create")
            return;
        }

        try
        {
            object objADSystemInfo = Interaction.CreateObject("ADSystemInfo");
            objUser = Interaction.GetObject("LDAP://" + objADSystemInfo.UserName);
            runnerName = objUser.samAccountName;
        }
        // runnerName = CreateObject("WScript.Network").UserName
        catch (Exception ex)
        {
            runnerName = "unknow";
            // whoami   echo %username%
            try
            {
                object ReadComputerName;
                ReadComputerName = Interaction.CreateObject("WScript.Shell");
                string ComputerName, RegPath;
                RegPath = @"HKLM\System\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName";
                ComputerName = ReadComputerName.RegRead(RegPath);
                if (ComputerName == "XMCORP-B672YY2")
                    runnerName = "changtu";
                else if (ComputerName == "AWSENTMSUTIL01")
                    runnerName = "prisminst";
            }
            catch (Exception ex2)
            {
            }
        }

        objRootDSE = Interaction.GetObject("LDAP://RootDSE");
        strRoot = objRootDSE.GET("DefaultNamingContext");

        strDomain = strRoot;
        strLDAP = "LDAP://" + strDomain;

        checkSQLDB(); // must run after strDomain set value

        getHRactiveEEIDlist();

        // If Not copyCSVtoSQLDB("C:\ehiAD\changeReport20190624_1839.csv") Then
        // Console.WriteLine("error")
        // End If
        // Console.WriteLine("Save to DB OK.")
        // Return

        logFilename = strWorkFolder + @"\" + strLogSubFolder + @"\" + "wdAD" + DateTime.Now.ToString("yyyyMMddHHmm") + ".txt";
        if (!createLogFile(logFilename, false))
        {
            Console.WriteLine("Cannot open the file for writing, please close it first: " + logFilename);
            // Console.WriteLine(waitNoteStr)
            // Console.ReadKey()
            return;
        }

        hasRights = false;
        if (Strings.LCase(runnerName) == "changtu" || Strings.LCase(runnerName) == "prisminst")
        {
            hasRights = true;
            swLog.WriteLine("Welcome " + runnerName + " to use this AD sync tool.");
            Console.WriteLine("log file: " + logFilename);
            Console.WriteLine(DateTime.Now + " start sync, please wait...");
        }
        else
        {
            swLog.WriteLine(runnerName + ", you have no rights to use this tool. Please contact Changtu.Wang@ehealth.com.");
            if (!forTest)
            {
                swLog.Close();
                return;
            }
        }

        FileSystem.ChDrive("C:");
        if (!System.IO.Directory.Exists(strWorkFolder))
        {
            try
            {
                FileSystem.MkDir(strWorkFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder);
                swLog.Close();
                return;
            }
        }
        FileSystem.ChDir(strWorkFolder);

        if (!System.IO.Directory.Exists(strScriptFolder))
        {
            try
            {
                FileSystem.MkDir(strScriptFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strScriptFolder);
                swLog.Close();
                return;
            }
        }
        if (!System.IO.Directory.Exists(strEXEfolder))
        {
            try
            {
                FileSystem.MkDir(strEXEfolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strEXEfolder);
                swLog.Close();
                return;
            }
        }
        if (!System.IO.Directory.Exists(strJIRAfolder))
        {
            try
            {
                FileSystem.MkDir(strJIRAfolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strJIRAfolder);
                swLog.Close();
                return;
            }
        }
        if (!System.IO.Directory.Exists(strReportSubFolder))
        {
            try
            {
                FileSystem.MkDir(strReportSubFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strReportSubFolder);
                swLog.Close();
                return;
            }
        }

        if (!System.IO.Directory.Exists(strLogSubFolder))
        {
            try
            {
                FileSystem.MkDir(strLogSubFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strLogSubFolder);
                swLog.Close();
                return;
            }
        }

        if (!System.IO.Directory.Exists(strDataSubfolder))
        {
            try
            {
                FileSystem.MkDir(strDataSubfolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strDataSubfolder);
                swLog.Close();
                return;
            }
        }

        if (!System.IO.Directory.Exists(strHRsubFolder))
        {
            try
            {
                FileSystem.MkDir(strHRsubFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strHRsubFolder);
                swLog.Close();
                return;
            }
        }

        if (!Directory.Exists(strPrimeSubfolder))
        {
            try
            {
                FileSystem.MkDir(strPrimeSubfolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strPrimeSubfolder);
                swLog.Close();
                return;
            }
        }

        if (!System.IO.Directory.Exists(strTempfolder))
        {
            try
            {
                FileSystem.MkDir(strTempfolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strTempfolder);
                swLog.Close();
            }
        }

        FileSystem.ChDir(strHRsubFolder);
        if (!System.IO.Directory.Exists(strFTP_DNLOADsubFolder))
        {
            try
            {
                FileSystem.MkDir(strFTP_DNLOADsubFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strHRsubFolder + @"\" + strFTP_DNLOADsubFolder);
                swLog.Close();
                return;
            }
        }

        if (!System.IO.Directory.Exists(strEE_DATAsubFolder))
        {
            try
            {
                FileSystem.MkDir(strEE_DATAsubFolder);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Cannot create folder " + strWorkFolder + @"\" + strHRsubFolder + @"\" + strEE_DATAsubFolder);
                swLog.Close();
                return;
            }
        }

        swLog.WriteLine(appNameVersion + ".exe begin running at " + DateTime.Now);
        setSyncSettings();

        if (needStop)
        {
            // need set after reading settings
            needUpdateFirstName = false;
            needUpdateLegalFirstName = false;
            needUpdateMiddleName = false;
            needUpdateDisplayName = false;

            needUpdateJobTitle = false;
            needUpdateHomeDeptment = false;
            needUpdateDescription = false;

            needUpdateManager = false;
            needUpdateOffice = false;
            needUpdateAddress = false;

            needUpdateMobile = false;
            needUpdateCostcenterLocationStr = "";
            needUpdateDepartmentCode = false;
            needUpdateDivisionCode = false;
            needUpdateCompany = false;
            needAutoSetReqID = false;

            needUpdateScriptPath = false;
            needUpdateHomeDirectory = false;
            needUpdatePhone = false;
            needUpdateFax = false;

            needUpdateUserSegment = false;

            if (forLeaveOnly)
            {
                needRemoveManagerForTerminated = false;
                needHideEmailForTerminated = false;
            }
        }


        int tryTimes = 0;
        string HRfileTime = "";
        string YYYYMMDD_hhtt = getYYYYMMDD_hhtt();
        var SFTPbatchFilename = strWorkFolder + @"\" + sftpBatch;
        string ccAddress = "";
        if (forTest)
            ccAddress = "";

        eeIDlist.Clear();
        reqIDlist.Clear();
        samIDlist.Clear();
        legalFirstNameList.Clear();
        firstNameList.Clear();
        lastNameList.Clear();
        middleNameList.Clear();
        userDNlist.Clear();
        statusList.Clear();
        inactiveDaysList.Clear();
        inactiveList.Clear();

        jobtitleList.Clear();
        HomeDepartmentList.Clear();
        WorkemailList.Clear();
        WorkemailListKey.Clear();
        proxyAddressesList.Clear();
        whenCreatedList.Clear();
        whenChangedList.Clear();
        PhysicalDeliveryOfficeNameList.Clear();
        employeeTypeList.Clear();
        cityList.Clear();
        managerList.Clear();
        LastLogonTimeStampList.Clear();

        ADusersHRfilenameCSV = "";
        ADusersHRfilenameXLS = "";
        if (needQueryAD)
        {
            // listAllUsersCSV() will set values for eeIDlist, reqIDlist...
            private int sleepTimeAgain = 30000;
    tryTimes = 0;
            while ((tryTimes< 2 && !listAllUsersCSV()))
            {
                tryTimes += 1;
                swLog.WriteLine(DateTime.Now + " tried time(s): " + tryTimes);
                System.Threading.Thread.Sleep(sleepTimeAgain* tryTimes);
            }

if (hasError)
{
    swLog.WriteLine("listAllUsersCSV() failed, exit the program " + DateTime.Now);
    swLog.Close();
    return;
}
        }
        else
{
    AllUsersCSVFileName = strWorkFolder + @"\" + strReportSubFolder + @"\" + allUsersForTestName;
    if (Strings.InStr(AllUsersCSVFileName, ".csv") < 1)
        AllUsersCSVFileName += ".csv";
    if (!System.IO.File.Exists(AllUsersCSVFileName))
    {
        swLog.WriteLine("file does not exist: " + AllUsersCSVFileName);
        swLog.Close();
        return;
    }
    try
    {
        readAllCSV2ArrayList();
    }
    catch (Exception ex)
    {
        swLog.WriteLine(ex.ToString() + " exit the program " + DateTime.Now);
        swLog.Close();
        return;
    }
}

swLog.Close();
if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

string filename = strWorkFolder + @"\psftp.exe";
if (!System.IO.File.Exists(filename))
{
    swLog.WriteLine("file does not exist: " + filename);
    swLog.Close();
    return;
}

private string hrTimeFileName = "hrTime.txt";
try
{
    swriterGeneralToday = new System.IO.StreamWriter(strWorkFolder + @"\" + hrTimeFileName, false); // always create new file
    swriterGeneralToday.WriteLine("ls " + sftpIncomingDirectory + HRsftpFileName + "*");
    swriterGeneralToday.Close();
}
catch (Exception ex)
{
    swLog.WriteLine("Cannot open the file for writing: " + strWorkFolder + @"\" + hrTimeFileName);
    swLog.Close();
    return;
}

var getFileTimeCMDStr = strWorkFolder + @"\psftp -b " + strWorkFolder + @"\" + hrTimeFileName + " -P 22 -p";
getFileTimeCMDStr += "w " + Strings.Chr(97) + Strings.Chr(68) + Strings.Chr(112) + Strings.Chr(65) + Strings.Chr(97) + Strings.Chr(50) + Strings.Chr(53);
getFileTimeCMDStr += Strings.Chr(89) + " adp" + Strings.Chr(64) + "prsftp01.ehealthinsurance.com"; // 10.8.5.101

hrFileReady = false;
if (!needDownloadHRfile)
    filename = ftpEEdataFolder + @"\" + HRfileForTestName;
else
{
    swLog.WriteLine("Reading " + hrMIS + " file received time from SFTP server directory " + sftpIncomingDirectory + " ...");
    FileSystem.ChDrive("C:");
    FileSystem.ChDir(strWorkFolder);
    int sleepTime = 300000; // 5 minutes
    int tryTimesCap = 2;
    if (forTest)
        sleepTime = 3000;
    tryTimes = 0;
    do
    {
        tryTimes += 1;
        HRfileTime = getFileTimeCMD(getFileTimeCMDStr);
        if (HRfileTime != "")
        {
            swLog.WriteLine("  SFTP file: " + HRfileTime);
            if (hrFileReady)
            {
                swLog.WriteLine("Latest " + hrMIS + " connection file is ready.");
                break;
            }
            else if (tryTimes < tryTimesCap)
            {
                swLog.WriteLine(DateTime.Now + ": Latest " + hrMIS + " file is not ready, wait and try again...");
                System.Threading.Thread.Sleep(sleepTime);
            }
        }
    }
    while ((tryTimes < tryTimesCap && !hrFileReady));

    if (!hrFileReady)
    {
        ccAddress = ""; // "Nirmal.Mehta@ehealth.com"  'HR的相关负责人，如果workday上的数据没有及时上传至对应的sftp中。
        string contentStr;
        if (HRfileTime == "")
        {
            tmpStr = DateTime.Now + ": reading " + hrMIS + " file information from SFTP server failed!";
            swLog.WriteLine(tmpStr);
            contentStr = tmpStr + " ( may be network problem )";
        }
        else
        {
            tmpStr = DateTime.Now + ": Latest " + hrMIS + " file still not received on SFTP server!";
            swLog.WriteLine(tmpStr);
            swLog.WriteLine("The existing file is " + HRfileTime);
            contentStr = tmpStr + " ( existing file is " + HRfileTime + ")";
        }
        swLog.WriteLine("Sending alert email for " + hrMIS + " file still not received on SFTP server...");
        sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, "AD sync: Latest " + hrMIS + " file problem", contentStr, "");
        // should not try download old HR connection file, it may be too old that has many discrepancies with AD

        if (!forTest)
        {
            // If Not forTest, exit the program here, do not try download any more
            sendADusersHRfile(ADusersHRfilenameCSV);
            swLog.Close();
            return;
        }
    }
    swLog.WriteLine("");

    // always download from SFTP first, set SFTP file name:
    filename = ftpEEdataFolder + @"\" + ftpFileName + "_" + YYYYMMDD_hhtt + ".csv";
    try
    {
        swriterGeneralToday = new System.IO.StreamWriter(SFTPbatchFilename, false); // always create new file
        swriterGeneralToday.WriteLine("get " + sftpIncomingDirectory + HRsftpFileName + " " + filename);
        swriterGeneralToday.Close();
    }
    catch (Exception ex)
    {
        swLog.WriteLine("Cannot open the file for writing, please close it first: " + SFTPbatchFilename);
        swLog.Close();
        return;
    }
}

if (needDownloadHRfile)
{
    // If Not forTest, exit the program already, do not run into this section
    FileSystem.ChDrive("C:");
    FileSystem.ChDir(strWorkFolder);
    if (!APIorManualUpload)
        HRfileTime = getFileTimeCMD(getFileTimeCMDStr);
    swLog.WriteLine("Downloading......" + HRfileTime);

    int sleepTime = 4000;
    int tryTimesCap = 20;
    if (forTest)
        tryTimesCap = 2;
    bool sftpOK = false;
    string downloadCMDStr = strWorkFolder + @"\psftp -b " + SFTPbatchFilename + " -P 22 -p";
    downloadCMDStr += "w " + Strings.Chr(97) + Strings.Chr(68) + Strings.Chr(112) + Strings.Chr(65) + Strings.Chr(97) + Strings.Chr(50) + Strings.Chr(53);
    downloadCMDStr += Strings.Chr(89) + " adp" + Strings.Chr(64) + "prsftp01.ehealthinsurance.com";

    try
    {
        Process p = new Process();
        p.StartInfo.FileName = @"C:\Windows\system32\cmd.exe";
        p.StartInfo.Arguments = "/c " + downloadCMDStr;
        p.StartInfo.UseShellExecute = false;
        p.StartInfo.RedirectStandardInput = true;
        p.StartInfo.RedirectStandardOutput = true;
        p.StartInfo.RedirectStandardError = true;
        p.StartInfo.CreateNoWindow = false;
        p.StartInfo.WorkingDirectory = strWorkFolder;
        p.Start();
        p.StandardInput.WriteLine("Exit");
        // Dim output As String = p.StandardOutput.ReadToEnd()
        p.Close();
        // swLog.WriteLine(output)

        System.Threading.Thread.Sleep(3000);
        sftpOK = true;
        tryTimes = 0;
        while (!System.IO.File.Exists(filename))
        {
            System.Threading.Thread.Sleep(sleepTime);
            tryTimes += 1;
            if (tryTimes > tryTimesCap)
            {
                sftpOK = false;
                tmpStr = "AD sync: Download " + hrMIS + " file from SFTP time out!";
                swLog.WriteLine(tmpStr);
                sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, tmpStr, tmpStr, "");
                break;
            }
        }
        if (sftpOK)
        {
            if (!APIorManualUpload)
                HRfileTime = getFileTimeCMD(getFileTimeCMDStr);
            Console.WriteLine("SFTP file downloaded is: " + HRfileTime);
            swLog.WriteLine("Downloaded OK: " + filename);
        }
    }
    catch (Exception ex)
    {
        tmpStr = "Download " + hrMIS + " file from SFTP failed!" + privateants.vbCrLf + ex.ToString();
        swLog.WriteLine(tmpStr);
        swLog.WriteLine("Sending alert email for SFTP download failed...");
        sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, "AD sync: Download " + hrMIS + " file from SFTP failed!", tmpStr, "");
    }

    if (!sftpOK)
    {
        // a fail-safe here: the HR connection file in FTP server only has few records, no longer try to download from old FTP server
        return;

        ftpFileName = "eHealth_StdEmp_Conn";
        filename = ftpEEdataFolder + @"\" + ftpFileName + "_" + YYYYMMDD_hhtt + ".csv";
        try
        {
        }
        // DownloadHRB(filename)
        catch (Exception ex)
        {
            tmpStr = "Download " + hrMIS + " file from FTP failed!" + privateants.vbCrLf + ex.ToString();
            swLog.WriteLine(tmpStr);
            swLog.WriteLine("Sending alert email for FTP download failed...");
            sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, "AD sync: Download " + hrMIS + " file from FTP failed!", tmpStr, "");
            swLog.Close();
            return;
        }
    }
    swLog.WriteLine("");
}

if (!System.IO.File.Exists(filename))
{
    swLog.WriteLine("file does not exist: " + filename);
    sendADusersHRfile(ADusersHRfilenameCSV);
    swLog.Close();
    return;
}

// need openCSVFileForChecking before saveDB
if (!openCSVFileForChecking(filename))
{
    swLog.Close();
    return;
}

if (HRrowCount < 1)
{
    swLog.WriteLine(hrMIS + " file has no record.");
    // should also pull a report from AD
    swLog.Close();
    return;
}

// saveDB3(False)
saveDB4();

// swLog.Close()
// Return

if (needSaveEEinfo2DB)
{
    if (END_afterSaveEEinfo2DB)
    {
        // swLog.WriteLine("Done testing, end here.")
        swLog.Close();
        return; // if need download only, end here
    }
    else
    {
        if (!forLeaveOnly)
        {
        }
        swLog.Close();
    }
}
else
    swLog.Close();

if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

// openCSVFileForChecking() will look at On Leave record in DB, if DB has problem, exist the program
if (!getOnLeaveEffDateFromDB())
{
    sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: " + appNameVersion + " running failed", "please check DB. Detailed please see log file " + logFilename, "");
    swLog.WriteLine("getOnLeaveEffDateFromDB() failed, exit the program " + DateTime.Now);
    swLog.Close();
    return;
}

if (!getDisabledSamidFromDB())
{
    sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: " + appNameVersion + " running failed", "please check DB", "");
    swLog.WriteLine("getDisabledSamidFromDB() failed, exit the program " + DateTime.Now);
    swLog.Close();
    return;
}

generalReprotTodayFilename = strWorkFolder + @"\" + strTempfolder + @"\" + "Report" + getYYYYMMDD_hhtt() + ".csv";
try
{
    swriterGeneralToday = new System.IO.StreamWriter(generalReprotTodayFilename, false); // always create new file
}
catch (Exception ex)
{
    swLog.WriteLine("Cannot open the file for writing, please close it first: " + generalReprotTodayFilename);
    swLog.Close();
    return;
}
swriterGeneralToday.WriteLine("AutomateTime,Tool,employeeID,samAccountName,firstName,lastName,status,department,jobTitle,note1,note2,note3,note4,note5,note6,note7,note8,note9,note10,note11,note12");


// read existing dont need AD employee records
if (!openCSVFileForNoNeedAD(NoNeedADfileName))
    swLog.WriteLine("Open CSV file problem: " + NoNeedADfileName);

strReportfileNameAll = strWorkFolder + @"\" + strReportSubFolder + @"\" + "updateReport" + getYYYYMMDD_hhtt() + ".txt";
strReportfileName = strReportfileNameAll;
try
{
    swReport = new System.IO.StreamWriter(strReportfileName, false); // create new file for 1st time
}
catch (Exception ex)
{
    swLog.WriteLine("Cannot open the file for writing, please close it first: " + strReportfileName);
    swLog.Close();
    return;
}

swLog.Close();
if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

// readEntOpsMembers
DirectoryEntry entry = new DirectoryEntry("LDAP://" + strDomain);
DirectorySearcher Searcher = new DirectorySearcher(entry);
Searcher.Filter = "(&(objectCategory=group)(mail=" + entOpsEmail + "))";
Searcher.PropertiesToLoad.Clear();
Searcher.PropertiesToLoad.Add("mail");
Searcher.PropertiesToLoad.Add("member");
SearchResult result = Searcher.FindOne;
if (!result == null)
{
    if (result.Properties("member").Count > 0)
    {
        // loop members
        for (int j = 0; j <= result.Properties("member").Count - 1; j++)
        {
            DistinguishedNameFound = result.Properties("member")(j).ToString;
            object objMember = Interaction.GetObject("LDAP://" + DistinguishedNameFound);
            try
            {
                EntOps += LCase(objMember.mail) + ";";
            }
            catch (Exception ex)
            {
            }
        }
    }
}
if (needCheckNewUsersNotLogon)
    ehiDCs = getDCs();

if (!checkAD())
{
    swReport.Close();
    swLog.Close();

    return;
}
swReport.Close();

swLog.Close();
if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

string strSubject;
swLog.WriteLine("");
// swLog.WriteLine("Sending alert email for AD not found employee...")
ccAddress = "";

if (!forLeaveOnly)
{
    if (!prepareADHRactive())
    {
        Console.WriteLine(DateTime.Now + " prepare AD users HR active report failed!");
        swLog.WriteLine(DateTime.Now + " prepare AD users HR active report failed!");
        ADusersHRactiveCSV = "";
        sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: prepare AD users HR active report failed!", "Today's AD users HR active report failed.", "");
    }
    else if (ADusersHRactiveCSV != "")
    {
        if (needUploadS3 && !forTest)
        {
            S3base s3 = new S3base();
            tmpStr = strWorkFolder + @"\" + strReportSubFolder + @"\";
            s3.S3key = ADusersHRactiveCSV.Replace(tmpStr, ""); // "ADusers_20200701_1515.csv"
            if (s3.upload)
            {
                Console.WriteLine(DateTime.Now + " " + s3.resultStr);
                swLog.WriteLine(DateTime.Now + " " + s3.resultStr);
            }
            else
            {
                Console.WriteLine(DateTime.Now + " upload to S3 failed: " + s3.errStr);
                swLog.WriteLine(DateTime.Now + " upload to S3 failed: " + s3.errStr);
                sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: upload AD users HR active report to S3 failed!", "Sending " + ADusersHRactiveCSV + " failed!", "");
            }
        }
    }
}
swLog.Close();
if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

if (discrepancyCount > 0 && !forLeaveOnly)
{
    string toAddress = entOpsEmail + ";Mithlesh.Prasad@ehealth.com";
    ccAddress = "Nirmal.Mehta@ehealth.com;";
    if (discrepancyCountXM > 0)
        ccAddress += "Stacey.Chen@ehealth.com";

    strSubject = "Daily comparison between " + hrMIS + " and AD: " + discrepancyCount + " record";
    if (discrepancyCount > 1)
        strSubject += "s";
    strSubject += " failed to match";
    string strMessage = "Please see attached file.";
    sendAlertEmail(toAddress, ccAddress, strSubject, strMessage, strReportfileDiscrepancy);
}
if (emailMismatchCount > 0 && !forLeaveOnly)
{
    string toAddress = entOpsEmail;
    if (emailMismatchCountXM > 0)
    {
    }
    ccAddress = "Changtu.Wang@ehealth.com";
    // ccAddress = "Nirmal.Mehta@ehealth.com;"
    // ccAddress &= checkEntOps("Changtu.Wang@ehealth.com")

    strSubject = emailMismatchCount + " primary email address mismatch found between " + hrMIS + " and AD";
    string strMessage = "Please see attached file.";
}
if (HRneedTerminatedCount > 0 && !forLeaveOnly)
{
    strSubject = "AD sync: " + HRneedTerminatedCount + " employee may need terminated in " + hrMIS;
    string strMessage = "Please check needDisabled column in " + strExportFileFound;
    sendAlertEmail("Changtu.Wang@ehealth.com", "", strSubject, strMessage, "");
}

if (needListFTEnonFTEgroupMemberShip && !forLeaveOnly)
    listFTEnonFTEgroupMemberShip();

if (intNotFoundCount > 0)
{
    strSubject = "AD sync not found: " + intNotFoundCount + " employee in " + hrMIS + " cannot be found in AD";
    // sendAlertEmail("Changtu.Wang@ehealth.com", "", strSubject, "All not found attached.", strExportFileNotFound & ";" & strExportFileFound)

    bool needPickUpNewUser = false;
    if (needPickUpNewUser)
    {
        System.Threading.Thread.Sleep(1000);
        readCSV2table(strExportFileNotFound, tableTMP);
        if (hasError)
        {
            System.Threading.Thread.Sleep(3000);
            readCSV2table(strExportFileNotFound, tableTMP);
        }
        intNotFoundCount = 0; // set the flag to 0 before set value again by calling pickNewUser()
        if (!hasError)
        {
            try
            {
                pickNewUser();
            }
            catch (Exception ex)
            {
                sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync pickNewUser failed", ex.ToString(), "");
            }
        }

        // after pickNewUser, intNotFoundCount may be changed
        if (intNotFoundCount > 0)
        {
            strSubject = "AD sync new users: " + intNotFoundCount + " employee has no AD account";
            string strMessage = "Hello IT Team Members,<br/><br/>";
            strMessage += "Discrepancy detected between " + hrMIS + " and AD.<br/>";
            strMessage += "The " + hrMIS + " AD sync tool cannot locate AD account for the employee.<br/>";
            strMessage += "Please see the attachment for details.<br/>";
            strMessage += "After confirming with New Hire JIRA issue, you may use the data for batch creating new AD users.";

            // new hire information is from JobVite now, no longer need send hrMIS based newhires.csv to IT guys
            ccAddress = "";
            sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, strSubject, strMessage, strExportFileNewUser);
        }
    } // needPickUpNewUser
}

// If Not forLeaveOnly AndAlso Not preparePrime() Then
// Console.WriteLine(Now & " prepare Prime report failed!")
// swLog.WriteLine(Now & " prepare Prime report failed!")
// strReportPrime = ""
// strReportPrimeFailed = ""
// sendAlertEmail("Anthony.King@ehealthinsurance.com;Changtu.Wang@ehealth.com", "Frances.SantiagoCruz@ehealthinsurance.com;Colleen.Pulkowski@ehealthinsurance.com", "Captivate Prime report from " & hrMIS & " AD sync failed", "prepare Prime report failed today!", "")
// End If

if (needCoupaReport && !forLeaveOnly && !prepareCoupa())
{
    Console.WriteLine(DateTime.Now + " prepare Coupa report failed!");
    swLog.WriteLine(DateTime.Now + " prepare Coupa report failed!");
    strReportCoupa = "";
    // strReportCoupaFailed = ""
    sendAlertEmail("Leah.Nolan@ehealth.com;monika.rayabarapu@ehealth.com", "Changtu.Wang@ehealth.com", "CoupaActiveDirectoryUsers report from " + hrMIS + " AD sync failed", "prepare Coupa report failed today!", "");
}

if (needOomnitzaReport)
{
    if (!forLeaveOnly && !prepareOomnitza())
    {
        Console.WriteLine(DateTime.Now + " prepare Oomnitza report failed!");
        swLog.WriteLine(DateTime.Now + " prepare Oomnitza report failed!");
        strReportOomnitza = "";
        sendAlertEmail("Changtu.Wang@ehealth.com;Windy.Chen@ehealth.com", "", "AD sync: prepare Oomnitza report failed!", "Today's Oomnitza report failed.", "");
    }
    if (strReportOomnitza != "" && !forLeaveOnly)
    {
        if (!putFileToSharedOomnitza())
        {
            swLog.WriteLine(DateTime.Now + " put Oomnitza report to shared failed!");
            sendAlertEmail("Changtu.Wang@ehealth.com;Windy.Chen@ehealth.com", "", "AD sync: put Oomnitza report to shared folder failed!", @"\\xmprspapp01\\CSV_Files\\" + CSVfilenameOomnitza + " may be being used by another process?", strReportOomnitza);
        }
    }
}

swLog.Close();
if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

swLog.WriteLine("");
if (!openCSVFileForUpdating(strExportFileNeedDisabled))
{
    swLog.Close();
    return;
}
bool needSave2SQLDB = false;
if (HRrowCount > 0)
{
    strReportfileName = strWorkFolder + @"\" + strReportSubFolder + @"\" + "setExpiredReport" + getYYYYMMDD_hhtt() + ".txt";
    try
    {
        swReport = new System.IO.StreamWriter(strReportfileName, false); // create new file for 1st time
    }
    catch (Exception ex)
    {
        swLog.WriteLine("Cannot open the file for writing, please close it first: " + strReportfileName);
        swLog.Close();
        return;
    }

    if (!System.IO.File.Exists(strExportFileFound))
    {
        swLog.WriteLine(strExportFileFound + " file does not exist, please check.");
        swLog.Close();
        return;
    }

    if (!UpdateFoundUserAfterChecking(true))
    {
        swReport.Close();
        swLog.Close();
        return;
    }
    swReport.Close();

    if (intOKcount > 0)
    {
        if (needRecordInDB)
            needSave2SQLDB = true;
        swLog.WriteLine("");
        swLog.WriteLine(DateTime.Now + " Sending alert email for AD already expired...");
        string subjectStr = intOKcount + " AD user";
        if (intOKcount > 1)
            subjectStr += "s";
        subjectStr += " set expired by the " + hrMIS + " into AD sync";

        string toAddress = entOpsEmail;
        ccAddress = "";
        ccAddress += checkEntOps("Nirmal.Mehta@ehealth.com");
        ccAddress += checkEntOps("Changtu.Wang@ehealth.com");

        string strMessage = "Hello IT Team Members,<br/><br/>";
        strMessage += "Attached report is about AD users set expired by the AD sync tool.<br/>";
        strMessage += "<br/>";
        strMessage += "The reason for expiring is the status in " + hrMIS + " is Leave or the AD user is not logon over " + daysInactiveLimit + " days.<br/>";
        strMessage += "<br/>";
        strMessage += "Please see the attachment for details.<br/>";
        strMessage += "<br/>";
        strMessage += "This is just notification and no action is required.";

        sendAlertEmail(toAddress, ccAddress, subjectStr, strMessage, strReportfileName); // strExportFileNeedDisabled)
    }
}

if (!openCSVFileForUpdating(strExportFileFound))
{
    swLog.Close();
    return;
}
if (HRrowCount > 0)
{
    strReportfileName = strReportfileNameAll;
    if (!System.IO.File.Exists(strExportFileFound))
    {
        swLog.WriteLine(strExportFileFound + " file does not exist, please check.");
        swLog.Close();
        return;
    }

    try
    {
        swReport = new System.IO.StreamWriter(strReportfileName, true); // append for 2nd time
    }
    catch (Exception ex)
    {
        swLog.WriteLine("Cannot open the file for writing, please close it first: " + strReportfileName);
        swLog.Close();
        return;
    }

    if (!UpdateFoundUserAfterChecking(false))
    {
        swReport.Close();
        swLog.Close();
        return;
    }
    swReport.Close();

    swriterGeneralToday.Close();

    if ((correctCount + intOKcount) > 0 || ((alreadyExpiredForLeave != "" || alreadyExpiredForInactivate != "") && !forLeaveOnly))
    {
        if (needRecordInDB)
            needSave2SQLDB = true;

        swLog.WriteLine("");
        swLog.WriteLine(DateTime.Now + " Sending alert email for AD user updated...");
        string subjectStr = (correctCount + intOKcount) + " AD user";
        if ((correctCount + intOKcount) > 1)
            subjectStr += "s";
        subjectStr += " updated by the " + hrMIS + " into AD sync";

        string toAddress = "";
        ccAddress = "";
        if ((correctCount + intOKcount) > 0)
        {
            toAddress = entOpsEmail;
            ccAddress += checkEntOps("Nirmal.Mehta@ehealth.com");
            // Huan's email 03 November 2016 ask added Bernette and Stacey
            ccAddress += checkEntOps("Stacey.Chen@ehealth.com");
        }
        else
            toAddress = "Changtu.Wang@ehealth.com";

        string strMessage = "Hello IT Team Members,<br/><br/>";
        strMessage += subjectStr + ".<br/>";
        strMessage += "Please see the attachment for details.<br/>";
        strMessage += "Please double check.";

        sendAlertEmail(toAddress, ccAddress, subjectStr, strMessage, strReportfileName);
    }
    else if (forLeaveOnly && (correctCount + intOKcount) < 1)
    {
        try
        {
            File.Delete(strReportfileName);
        }
        catch (Exception ex)
        {
        }
    }
}

// not in HR system list also need check and hide email address to show on Outlook properties tab, already hiding with CTPR-13
// some new hires AD already created and set disabled, but HR system has already deleted records
if (!forLeaveOnly)
    DoGroupMemberAction();

swriterGeneralToday.Close();

// delete non-OnLeave records, after updating On Leave users
deleteNonOnLeave();
deleteDisabledDaysAgo();

swLog.Close();
if (!createLogFile(logFilename, true))
{
    Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
    return;
}

if (needSave2SQLDB)
{
    try
    {
        if (copyCSVtoSQLDB(generalReprotTodayFilename))
        {
            if (intOKcount > 0)
            {
                tmpStr = intOKcount + " row(s) recorded in SQL DB " + ITDBname + " OK.";
                swLog.WriteLine(tmpStr);
            }
        }
        else
            sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: " + appNameVersion + " copy " + generalReprotTodayFilename + " to " + ITDBname + " has error", errStr, generalReprotTodayFilename);
    }
    catch (Exception ex)
    {
        sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: " + appNameVersion + " copy " + generalReprotTodayFilename + " to " + ITDBname + " failed", ex.ToString(), generalReprotTodayFilename);
    }
}

if (needCreateADaccount && intNotFoundCount > 0)
{
    tmpStr = strWorkFolder + @"\" + appNameAddUser;
    if (!System.IO.File.Exists(tmpStr))
    {
        swLog.WriteLine("");
        swLog.WriteLine(tmpStr + " does not exist. Cannot add new user.");
    }
    else
    {
        if (!openCSVFileForUpdating(strExportFileNewUser))
        {
            swLog.Close();
            return;
        }

        if (HRrowCount > 0)
        {
            hasError = false;
            var strCreateReportfileName = strWorkFolder + @"\" + strReportSubFolder + @"\" + "createReport" + getYYYYMMDD_hhtt() + ".txt";
            tmpStr = strWorkFolder + @"\" + appNameAddUser + " " + "\"" + strExportFileNewUser + "\" " + strCreateReportfileName;
            try
            {
                Process p = new Process();
                p.StartInfo.FileName = @"C:\Windows\system32\cmd.exe";
                p.StartInfo.Arguments = "/c " + tmpStr;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = false;
                p.StartInfo.WorkingDirectory = strWorkFolder;
                p.Start();
                p.StandardInput.WriteLine("Exit");
                string output = p.StandardOutput.ReadToEnd();
                p.Close();
                swLog.WriteLine(output);
                System.Threading.Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                swLog.WriteLine("Call " + tmpStr + " failed.");
            }
        }
    }
}

if (!forLeaveOnly)
{
    // if not needQueryAD, will no HR file generte and send
    if (ADusersHRfilenameXLS == "")
        sendADusersHRfile(ADusersHRfilenameCSV);
    else
    {
        // Need waiting the new generated Excel file is read to use
        hasError = false;
        sendADusersHRfile(ADusersHRfilenameXLS);
        if (!hasError)
        {
            try
            {
                System.IO.File.Delete(ADusersHRfilenameCSV);
            }
            catch (Exception ex)
            {
            }
        }
    }
}

strReportPrime = "";
strReportPrimeFailed = "";
if (strReportPrime != "" && !forLeaveOnly)
{
    var XMUS = "";
    if (forXM_US != "")
        XMUS = " (for " + forXM_US + "). ";

    string attachFileNames = strReportPrime;
    if (strReportPrimeFailed != "")
        attachFileNames += ";" + strReportPrimeFailed;
}

if (needCoupaReport && strReportCoupa != "" && !forLeaveOnly)
{
    if (putFileToSharedCoupa())
        swLog.WriteLine(DateTime.Now + " put CoupaActiveDirectoryUsers.csv to shared folder successfully");
    else
    {
        swLog.WriteLine(DateTime.Now + " put CoupaActiveDirectoryUsers report to shared folder failed!");
        sendAlertEmail("Leah.Nolan@ehealth.com;monika.rayabarapu@ehealth.com", "Changtu.Wang@ehealth.com", "AD sync: put CoupaActiveDirectoryUsers report to shared folder failed!", "Please process attached flie manually.", strReportCoupa);
    }
}

if (forLeaveOnly)
{
    try
    {
        if (!forTest)
            File.Move(AllUsersCSVFileName, AllUsersCSVFileName.Replace(strReportSubFolder, strTempfolder));
        File.Move(strExportFileEntire, strExportFileEntire.Replace(strReportSubFolder, strTempfolder));
        File.Move(strExportFileFound, strExportFileFound.Replace(strReportSubFolder, strTempfolder));
        File.Move(strExportFileNotFound, strExportFileNotFound.Replace(strReportSubFolder, strTempfolder));
        File.Move(strExportFileNeedDisabled, strExportFileNeedDisabled.Replace(strReportSubFolder, strTempfolder));
        File.Move(strExportFileGroupAction, strExportFileGroupAction.Replace(strReportSubFolder, strTempfolder));
        File.Move(strReportfileDiscrepancy, strReportfileDiscrepancy.Replace(strReportSubFolder, strTempfolder));
        swLog.WriteLine("Some report files generated this time has been moved to folder " + strTempfolder);
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.ToString());
    }
}

Console.WriteLine(appNameVersion + " complete OK at " + DateTime.Now);
swLog.WriteLine(appNameVersion + " complete OK at " + DateTime.Now);
swLog.Close();

while ((!swLog == null))
    swLog.Close();
try
{
    sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync log", " please see attached.", logFilename);
}
catch (Exception ex)
{
}
    }

    public static bool openCSVFileForChecking(string aFileName)
{
    duplicatedWDID.Clear();
    WDID.Clear();
    previousID.Clear();
    HR_eeID.Clear();
    HR_reqID.Clear();
    HR_firstName.Clear();
    HR_lastName.Clear();
    HR_middleName.Clear();
    HR_nickName.Clear();
    HR_workEmail.Clear();
    HR_workEmailKey.Clear();
    HR_status.Clear();
    HR_statusEffDate.Clear();
    HR_managerWorkEmail.Clear();
    HR_locationAux1.Clear();
    HR_managerID.Clear();

    HR_Jobtitle.Clear();
    HR_JobEffDate.Clear();
    HR_BusinessUnit.Clear();
    HR_HomeDepartment.Clear();
    HR_samAccountName.Clear();
    HR_UserSegment.Clear();

    HR_Address.Clear();
    HR_AddressStreet2.Clear();
    HR_AddressStreet3.Clear();
    HR_TPA_location_vendor.Clear();
    HR_City.Clear();
    HR_State.Clear();
    HR_PostalCode.Clear();
    HR_Country.Clear();

    HR_Workphone.Clear();
    HR_mobile.Clear();
    HR_Fax.Clear();
    HR_CostCenter.Clear();
    HR_office.Clear();
    HR_ProfilePath.Clear();
    HR_FTE.Clear();
    HR_needAccount.Clear();

    HR_OriginalHireDate.Clear();
    HR_HireDate.Clear();
    HR_LocationCode.Clear();
    HR_DivisionCode.Clear();
    HR_CompanyCode.Clear();
    HR_ManagementLevel.Clear();


    if (aFileName == null)
        return false;
    if (!File.Exists(aFileName))
        return false;
    object input = null;
    int tryTimes = 0;
    while (input == null)
    {
        try
        {
            tryTimes += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(aFileName);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTimes > 90)
        {
            tmpStr = "AD sync: Open file time out!" + aFileName;
            if (swLog != null)
                swLog.WriteLine(tmpStr);
            sendAlertEmail("Changtu.Wang@ehealth.com", "", tmpStr, tmpStr, "");
            break;
        }
    }
    if (input == null)
        return false;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    hasManagerID = false;
    colHR_eeID = -1;
    int colHR_previousID = -1;
    int colHR_reqID = -1;
    colHR_firstName = -1;
    colHR_lastName = -1;
    colHR_middleName = -1;
    colHR_nickName = -1;
    colHR_workEmail = -1;
    colHR_status = -1;
    colHR_statusEffDate = -1;
    colHR_managerWorkEmail = -1;
    colHR_location = -1;
    colHR_locationAux1 = -1;
    colHR_HomeDepartment = -1;
    colHR_DepartmentAux1 = -1;
    colHR_Jobtitle = -1;
    colHR_JobEffDate = -1;
    colHR_BusinessUnit = -1;
    colHR_samAccountName = -1;
    int colHR_UserSegment = -1;
    colHR_Address = -1;
    colHR_AddressStreet2 = -1;
    colHR_AddressStreet3 = -1;
    colHR_TPA_location_vendor = -1;
    int colHR_TPA_vendor = -1;
    colHR_City = -1;
    colHR_State = -1;
    colHR_PostalCode = -1;
    colHR_Country = -1;

    colHR_Workphone = -1;
    int colHR_mobile = -1;
    colHR_Fax = -1;
    colHR_CostCenter = -1;
    colHR_office = -1;
    colHR_ProfilePath = -1;
    colHR_FTE = -1;
    colHR_OriginalHireDate = -1;
    colHR_HireDate = -1;
    colHR_RehireDate = -1;
    colHR_NeedAccount = -1;
    int colHR_locationCode = -1;
    int colHR_DivisionCode = -1;
    int colHR_CompanyCode = -1;
    int colHR_ManagementLevel = -1;

    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        // aTitle = aTitle.Replace("#", "")
        if (aTitle == "lastname")
            colHR_lastName = HRcolCount;
        else if (aTitle == "firstname")
            colHR_firstName = HRcolCount;
        else if (aTitle == "middlename")
            colHR_middleName = HRcolCount;
        else if (aTitle == "preferredname" || aTitle == "nickname")
            colHR_nickName = HRcolCount;
        else if (aTitle == "employeeid")
            colHR_eeID = HRcolCount;
        else if (aTitle == "previousid")
            colHR_previousID = HRcolCount;
        else if (aTitle == "requisitionid")
            colHR_reqID = HRcolCount;
        else if (aTitle == "workcontact:workemail" || aTitle == "workemail")
        {
            colHR_workEmail = HRcolCount;
            hasWorkEmail = true;
        }
        else if (aTitle == "positionstatus" || aTitle == "status")
            colHR_status = HRcolCount;
        else if (aTitle == "statuseffectivedate")
            colHR_statusEffDate = HRcolCount;
        else if (aTitle.Contains("reportstoemail") || aTitle == "managerworkemail")
        {
            colHR_managerWorkEmail = HRcolCount;
            hasManagerWorkEmail = true;
        }
        else if (aTitle == "businesstitle" || aTitle == "jobtitle")
        {
            colHR_Jobtitle = HRcolCount;
            hasJobTitle = true;
        }
        else if (aTitle == "jobeffdate")
            colHR_JobEffDate = HRcolCount;
        else if (aTitle == "businessunit")
            colHR_BusinessUnit = HRcolCount;
        else if (aTitle == "homedepartment" || aTitle == "homedepartmentdescription")
        {
            colHR_HomeDepartment = HRcolCount;
            hasHomeDepartment = true;
        }
        else if (aTitle == "departmentaux1")
            colHR_DepartmentAux1 = HRcolCount;
        else if (aTitle == "locationdescription" || aTitle == "location")
            colHR_location = HRcolCount;
        else if (aTitle == "locationauxiliary1" || aTitle == "locationaux1")
            colHR_locationAux1 = HRcolCount;
        else if (aTitle == "reportstoassociateid" || aTitle == "managereeid")
        {
            colHR_managerID = HRcolCount;
            hasManagerID = true;
        }
        else if (aTitle == "samaccountname")
            colHR_samAccountName = HRcolCount;
        else if (aTitle == "usersegment")
            colHR_UserSegment = HRcolCount;
        else if (aTitle == "locationstreet1" || aTitle == "address")
            colHR_Address = HRcolCount;
        else if (aTitle == "locationstreet2")
            colHR_AddressStreet2 = HRcolCount;
        else if (aTitle == "locationstreet3")
            colHR_AddressStreet3 = HRcolCount;
        else if (aTitle.Contains("tpa_location_vendor"))
            colHR_TPA_location_vendor = HRcolCount;
        else if (aTitle.Contains("supplier"))
            colHR_TPA_vendor = HRcolCount;
        else if (aTitle == "locationcity" || aTitle == "city")
            colHR_City = HRcolCount;
        else if (aTitle == "locationstate" || aTitle == "state")
            colHR_State = HRcolCount;
        else if (aTitle == "locationzip" || aTitle == "postalcode")
            colHR_PostalCode = HRcolCount;
        else if (aTitle == "locationcountry" || aTitle == "country")
            colHR_Country = HRcolCount;
        else if (aTitle == "workphone")
            colHR_Workphone = HRcolCount;
        else if (aTitle == "primaryworkmobile")
            colHR_mobile = HRcolCount;
        else if (aTitle == "workfax")
            colHR_Fax = HRcolCount;
        else if (aTitle == "costcenter" || aTitle == "pager")
            colHR_CostCenter = HRcolCount;
        else if (aTitle == "office")
            colHR_office = HRcolCount;
        else if (aTitle == "profilepath")
            colHR_ProfilePath = HRcolCount;
        else if (aTitle == "workercategorydescription" || aTitle == "empclass")
            colHR_FTE = HRcolCount;
        else if (aTitle == "originalhiredate")
            colHR_OriginalHireDate = HRcolCount;
        else if (aTitle == "hiredate")
            colHR_HireDate = HRcolCount;
        else if (aTitle == "rehiredate")
            colHR_RehireDate = HRcolCount;
        else if (aTitle == "needsadaccount")
            colHR_NeedAccount = HRcolCount;
        else if (aTitle == "locationcode")
            colHR_locationCode = HRcolCount;
        else if (aTitle == "divisioncode")
            colHR_DivisionCode = HRcolCount;
        else if (aTitle == "companycode")
            colHR_CompanyCode = HRcolCount;
        else if (aTitle == "managementlevel")
            colHR_ManagementLevel = HRcolCount;
        HRcolCount = HRcolCount + 1;
    }

    HRrowCount = 0;
    if (HRcolCount < 1)
    {
        swLog.WriteLine(aFileName + ": " + " has no column.");
        return false;
    }
    if (colHR_middleName < 0)
        needUpdateMiddleName = false;
    // If colHR_TPA_vendor < 0 Then
    // needUpdateAddress = False
    // End If

    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);
            if (colHR_previousID < 0)
                previousID.Add("");
            else
                previousID.Add(rows.Item(HRrowCount)(colHR_previousID).replace(",", "_"));

            if (colHR_eeID < 0)
                WDID.Add("");
            else
            {
                // deal with strange EEID that may be contains ","
                tmpStr = rows.Item(HRrowCount)(colHR_eeID).replace(",", "_");
                while ((tmpStr.Length < 6))
                    tmpStr = "0" + tmpStr;
                if (WDID.IndexOf(tmpStr) >= 0)
                    duplicatedWDID.Add(tmpStr);
                WDID.Add(tmpStr);
            }

            // if has previousID, use it as employeeID
            // always use WDID since 01/09/2020
            tmpStr = WDID(HRrowCount); // previousID(HRrowCount)
            if (tmpStr == "")
                HR_eeID.Add(WDID(HRrowCount));
            else
            {
                while ((tmpStr.Length < 6))
                    tmpStr = "0" + tmpStr;
                HR_eeID.Add(tmpStr);
            }
            HR_eeID0.Add(HR_eeID(HRrowCount));
            if (colHR_reqID < 0)
                HR_reqID.Add("");
            else
            {
                // deal with strange reqID that may be contains "," or ";"
                // get last reqID
                int pos = -1;
                string theReqID = Trim(rows.Item(HRrowCount)(colHR_reqID));
                theReqID = theReqID.Replace(" ", "");
                if (theReqID.Contains(","))
                {
                    pos = theReqID.LastIndexOf(",");
                    if (pos > 0)
                        theReqID = theReqID.Substring(pos + 1, theReqID.Length - pos - 1);
                }
                if (theReqID.Contains(";"))
                {
                    pos = theReqID.LastIndexOf(";");
                    if (pos > 0)
                        theReqID = theReqID.Substring(pos + 1, theReqID.Length - pos - 1);
                }
                HR_reqID.Add(theReqID);
            }
            if (colHR_firstName < 0)
                HR_firstName.Add("");
            else
                HR_firstName.Add(rows.Item(HRrowCount)(colHR_firstName));
            if (colHR_lastName < 0)
                HR_lastName.Add("");
            else
                HR_lastName.Add(rows.Item(HRrowCount)(colHR_lastName));
            if (colHR_middleName < 0)
                HR_middleName.Add("");
            else
                HR_middleName.Add(rows.Item(HRrowCount)(colHR_middleName));
            if (colHR_nickName < 0)
                HR_nickName.Add("");
            else
                HR_nickName.Add(rows.Item(HRrowCount)(colHR_nickName));

            if (colHR_workEmail < 0)
            {
                HR_workEmail.Add("");
                HR_workEmailKey.Add("");
            }
            else
            {
                tmpStr = rows.Item(HRrowCount)(colHR_workEmail);
                if (Strings.InStr(tmpStr, "@") > 0)
                {
                    HR_workEmail.Add(tmpStr);
                    HR_workEmailKey.Add(Microsoft.VisualBasic.Left(Strings.LCase(tmpStr), Strings.InStr(tmpStr, "@"))); // only email key need LCase
                }
                else
                {
                    HR_workEmail.Add("");
                    HR_workEmailKey.Add("");
                }
            }
            if (colHR_status < 0)
                HR_status.Add("");
            else
                HR_status.Add(rows.Item(HRrowCount)(colHR_status));
            if (colHR_statusEffDate < 0)
                HR_statusEffDate.Add("");
            else
                HR_statusEffDate.Add(rows.Item(HRrowCount)(colHR_statusEffDate));

            if (InStr(HR_status.Item(HR_status.Count - 1), "Leave") > 0)
            {
                if (HR_statusEffDate(HR_statusEffDate.Count - 1) == "")
                {
                    // look at On Leave record in DB
                    int OnLeaveIndex = DB_OnLeaveEEID.IndexOf(HR_eeID(HR_eeID.Count - 1));
                    if (OnLeaveIndex >= 0)
                        HR_statusEffDate(HR_statusEffDate.Count - 1) = DB_OnLeaveEffDate(OnLeaveIndex);
                }
            }

            if (colHR_managerWorkEmail < 0)
                HR_managerWorkEmail.Add("");
            else
                HR_managerWorkEmail.Add(LCase(rows.Item(HRrowCount)(colHR_managerWorkEmail)));

            // try add locationAux1 first, then location
            // should map location to locationAux1
            if (colHR_locationAux1 < 0)
            {
                if (colHR_location < 0)
                    tmpStr = "";
                else
                    tmpStr = Trim(rows.Item(HRrowCount)(colHR_location));
            }
            else
                tmpStr = rows.Item(HRrowCount)(colHR_locationAux1);

            if (tmpStr == "")
                HR_locationAux1.Add("");
            else
            {
                tmpStr = mapLocation_Aux1(Strings.UCase(tmpStr));
                HR_locationAux1.Add(tmpStr);
            }

            if (hasManagerID)
            {
                tmpStr = rows.Item(HRrowCount)(colHR_managerID);
                if (tmpStr != "")
                {
                    while ((tmpStr.Length < 6))
                        tmpStr = "0" + tmpStr;
                }
                HR_managerID.Add(tmpStr);
            }
            else
                HR_managerID.Add("");
            if (hasJobTitle)
                HR_Jobtitle.Add(rows.Item(HRrowCount)(colHR_Jobtitle));
            else
                HR_Jobtitle.Add("");
            if (colHR_JobEffDate < 0)
                HR_JobEffDate.Add("");
            else
                HR_JobEffDate.Add(rows.Item(HRrowCount)(colHR_JobEffDate));
            if (colHR_BusinessUnit < 0)
                HR_BusinessUnit.Add("");
            else
                HR_BusinessUnit.Add(rows.Item(HRrowCount)(colHR_BusinessUnit));
            tmpStr = "";
            if (hasHomeDepartment)
                tmpStr = rows.Item(HRrowCount)(colHR_HomeDepartment);
            HR_HomeDepartment.Add(tmpStr);

            if (colHR_samAccountName < 0)
                HR_samAccountName.Add("");
            else
                HR_samAccountName.Add(rows.Item(HRrowCount)(colHR_samAccountName));
            if (colHR_UserSegment < 0)
                HR_UserSegment.Add("");
            else
                HR_UserSegment.Add(rows.Item(HRrowCount)(colHR_UserSegment));
            if (colHR_Address < 0)
                HR_Address.Add("");
            else
                HR_Address.Add(rows.Item(HRrowCount)(colHR_Address));
            if (colHR_AddressStreet2 < 0)
                HR_AddressStreet2.Add("");
            else
                HR_AddressStreet2.Add(rows.Item(HRrowCount)(colHR_AddressStreet2));
            if (colHR_AddressStreet3 < 0)
                HR_AddressStreet3.Add("");
            else
                HR_AddressStreet3.Add(rows.Item(HRrowCount)(colHR_AddressStreet3));
            if (colHR_TPA_location_vendor < 0)
            {
                if (colHR_TPA_vendor < 0)
                    HR_TPA_location_vendor.Add("");
                else
                {
                    string location = HR_locationAux1(HRrowCount); // .replace("TPA: ", "")
                    string vendor = Trim(rows.Item(HRrowCount)(colHR_TPA_vendor));
                    if (location == "")
                        HR_TPA_location_vendor.Add(vendor);
                    else if (vendor == "")
                    {
                        if (location.StartsWith("TPA:"))
                            HR_TPA_location_vendor.Add(location);
                        else
                            HR_TPA_location_vendor.Add("");
                    }
                    else
                        HR_TPA_location_vendor.Add(location + "; " + vendor);
                }
            }
            else
                HR_TPA_location_vendor.Add(rows.Item(HRrowCount)(colHR_TPA_location_vendor));
            if (colHR_City < 0)
                HR_City.Add("");
            else
                HR_City.Add(rows.Item(HRrowCount)(colHR_City));
            if (colHR_State < 0)
                HR_State.Add("");
            else
                HR_State.Add(rows.Item(HRrowCount)(colHR_State));
            if (colHR_PostalCode < 0)
                HR_PostalCode.Add("");
            else
                HR_PostalCode.Add(rows.Item(HRrowCount)(colHR_PostalCode));
            if (colHR_Country < 0)
                HR_Country.Add("");
            else
                HR_Country.Add(rows.Item(HRrowCount)(colHR_Country));

            if (colHR_Workphone < 0)
                HR_Workphone.Add("");
            else
                HR_Workphone.Add(rows.Item(HRrowCount)(colHR_Workphone));
            if (colHR_mobile < 0)
                HR_mobile.Add("");
            else
                HR_mobile.Add(rows.Item(HRrowCount)(colHR_mobile));
            if (colHR_Fax < 0)
                HR_Fax.Add("");
            else
                HR_Fax.Add(rows.Item(HRrowCount)(colHR_Fax));
            // always use standard field name "CostCenter"
            if (colHR_CostCenter < 0)
                HR_CostCenter.Add("");
            else
                HR_CostCenter.Add(rows.Item(HRrowCount)(colHR_CostCenter));
            if (colHR_office < 0)
                HR_office.Add("");
            else
                HR_office.Add(rows.Item(HRrowCount)(colHR_office));
            if (colHR_ProfilePath < 0)
                HR_ProfilePath.Add("");
            else
                HR_ProfilePath.Add(rows.Item(HRrowCount)(colHR_ProfilePath));
            if (colHR_FTE < 0)
                HR_FTE.Add("");
            else
                HR_FTE.Add(rows.Item(HRrowCount)(colHR_FTE));
            HR_FTE0.Add(HR_FTE(HRrowCount));

            // after HR system upgrade to WorkForce Now, “Hire Date” is now “OriginalHireDate”. And if they are a re-hire, then the “Rehire Date” field will be populated. 
            // so if “Rehire Date” is not blank, “Hire Date” should be set the “Rehire Date”
            // Workday should use OriginalHireDate and hireDate as column name, if OriginalHireDate <> hireDate then is rehire date
            if (colHR_RehireDate < 0)
            {
                if (colHR_OriginalHireDate < 0)
                    HR_OriginalHireDate.Add("");
                else
                    HR_OriginalHireDate.Add(rows.Item(HRrowCount)(colHR_OriginalHireDate));
                if (colHR_HireDate < 0)
                    HR_HireDate.Add("");
                else
                    HR_HireDate.Add(rows.Item(HRrowCount)(colHR_HireDate));
            }
            else
            {
                if (colHR_HireDate < 0)
                    HR_OriginalHireDate.Add("");
                else
                    HR_OriginalHireDate.Add(rows.Item(HRrowCount)(colHR_HireDate));

                if (Trim(rows.Item(HRrowCount)(colHR_RehireDate)) == "")
                {
                    if (colHR_HireDate < 0)
                        HR_HireDate.Add("");
                    else
                        HR_HireDate.Add(rows.Item(HRrowCount)(colHR_HireDate));
                }
                else
                    HR_HireDate.Add(rows.Item(HRrowCount)(colHR_RehireDate));
            }

            if (colHR_NeedAccount < 0)
                HR_needAccount.Add("");
            else
                HR_needAccount.Add(rows.Item(HRrowCount)(colHR_NeedAccount));

            if (colHR_locationCode < 0)
                HR_LocationCode.Add("");
            else
                HR_LocationCode.Add(rows.Item(HRrowCount)(colHR_locationCode));
            if (colHR_DivisionCode < 0)
                HR_DivisionCode.Add("");
            else
            {
                tmpStr = rows.Item(HRrowCount)(colHR_DivisionCode);
                if ((tmpStr.Length < 2))
                    tmpStr = "0" + tmpStr;
                HR_DivisionCode.Add(tmpStr);
            }
            if (colHR_CompanyCode < 0)
                HR_CompanyCode.Add("");
            else
                HR_CompanyCode.Add(rows.Item(HRrowCount)(colHR_CompanyCode));
            if (colHR_ManagementLevel < 0)
                HR_ManagementLevel.Add("");
            else
                HR_ManagementLevel.Add(rows(HRrowCount)(colHR_ManagementLevel));

            // swLog.WriteLine(HR_eeID.Item(HRrowCount) & " " & HR_firstName.Item(HRrowCount) & " " & HR_lastName.Item(HRrowCount))
            HRrowCount += 1;
        }
    }
    catch (Exception ex)
    {
        input.Close();
        string ccAddress = "Steve.Arnoldus@ehealth.com;Nirmal.Mehta@ehealth.com;";

        tmpStr = "Latest " + hrMIS + " connection file downloaded " + aFileName + " has problem (row# " + HRrowCount + 1 + " ), please check with SFTP server and ADP!";
        tmpStr += " Please upload the connection file manually to SFTP server adp@prsftp01.ehealthinsurance.com, overwrite the existing bad file ";
        tmpStr += sftpIncomingDirectory + HRsftpFileName + ",";
        tmpStr += "then ask IT guys run the " + hrMIS + @" AD sync programe once on server SJentMSutil01:     C:\ehiAD\hrbAD.bat";
        sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, "AD sync: The " + hrMIS + " connection file downloaded on server SJentMSutil01 has problem!", tmpStr, "");
        swLog.WriteLine(tmpStr);
        return false;
    }
    input.Close();

    if (false && hasManagerID)
    {
        // replace managerID with previousID
        string pID;
        for (int i = 0; i <= HR_managerID.Count - 1; i++)
        {
            if (HR_managerID(i) == "")
                continue;
            index = WDID.IndexOf(HR_managerID(i));
            if (index >= 0)
            {
                pID = previousID(index);
                if (pID != "")
                    HR_managerID(i) = pID;
            }
        }
    }

    if (duplicatedWDID.Count > 0)
    {
        for (int i = 0; i <= duplicatedWDID.Count - 1; i++)
        {
            int delIndex = WDID.IndexOf(duplicatedWDID(i));
            while (delIndex >= 0 && delIndex < (WDID.Count - 1))
            {
                if (WDID(delIndex + 1) != duplicatedWDID(i))
                    // keep last one
                    break;
                if (HR_status(delIndex) != "Terminated")
                {
                    if (HR_status(delIndex + 1) == "Terminated")
                        // delete next row
                        delIndex += 1;
                }
                WDID.RemoveAt(delIndex);
                previousID.RemoveAt(delIndex);
                HR_eeID.RemoveAt(delIndex);
                HR_reqID.RemoveAt(delIndex);
                HR_firstName.RemoveAt(delIndex);
                HR_lastName.RemoveAt(delIndex);
                HR_middleName.RemoveAt(delIndex);
                HR_nickName.RemoveAt(delIndex);
                HR_workEmail.RemoveAt(delIndex);
                HR_workEmailKey.RemoveAt(delIndex);
                HR_status.RemoveAt(delIndex);
                HR_statusEffDate.RemoveAt(delIndex);
                HR_managerWorkEmail.RemoveAt(delIndex);
                HR_locationAux1.RemoveAt(delIndex);
                HR_managerID.RemoveAt(delIndex);

                HR_Jobtitle.RemoveAt(delIndex);
                HR_JobEffDate.RemoveAt(delIndex);
                HR_BusinessUnit.RemoveAt(delIndex);
                HR_HomeDepartment.RemoveAt(delIndex);
                HR_samAccountName.RemoveAt(delIndex);
                HR_UserSegment.RemoveAt(delIndex);

                HR_Address.RemoveAt(delIndex);
                HR_AddressStreet2.RemoveAt(delIndex);
                HR_AddressStreet3.RemoveAt(delIndex);
                HR_TPA_location_vendor.RemoveAt(delIndex);
                HR_City.RemoveAt(delIndex);
                HR_State.RemoveAt(delIndex);
                HR_PostalCode.RemoveAt(delIndex);
                HR_Country.RemoveAt(delIndex);

                HR_Workphone.RemoveAt(delIndex);
                HR_mobile.RemoveAt(delIndex);
                HR_Fax.RemoveAt(delIndex);
                HR_CostCenter.RemoveAt(delIndex);
                HR_office.RemoveAt(delIndex);
                HR_ProfilePath.RemoveAt(delIndex);
                HR_FTE.RemoveAt(delIndex);
                HR_needAccount.RemoveAt(delIndex);

                HR_OriginalHireDate.RemoveAt(delIndex);
                HR_HireDate.RemoveAt(delIndex);
                HR_LocationCode.RemoveAt(delIndex);
                HR_DivisionCode.RemoveAt(delIndex);
                HR_CompanyCode.RemoveAt(delIndex);
                HR_ManagementLevel.RemoveAt(delIndex);

                HR_eeID0.RemoveAt(delIndex);
                HR_FTE0.RemoveAt(delIndex);

                delIndex = WDID.IndexOf(duplicatedWDID(i));
            }
        }
    }
    HRrowCount = WDID.Count;
    Console.WriteLine(aFileName + ": " + HRrowCount + " records");
    if (swLog != null)
        swLog.WriteLine(aFileName + ": " + HRrowCount + " records");
    return true;
}

private static bool openCSVFileForUpdating(string aFileName)
{
    HR_eeID.Clear();
    HR_reqID.Clear(); // when open for updating, HR_reqID is used for storing a mark that need update EEID for a new hire that only has reqID but no EEID
    HR_firstName.Clear();
    HR_lastName.Clear();
    HR_middleName.Clear();
    HR_nickName.Clear();
    HR_workEmail.Clear(); // AD found file only has workemail_ad, here HR_workEmail means actual AD email address
    HR_workEmailKey.Clear();
    HR_status.Clear();
    HR_statusEffDate.Clear();
    HR_managerWorkEmail.Clear();
    HR_locationAux1.Clear();
    HR_managerID.Clear();
    HR_managerDN.Clear();

    HR_Jobtitle.Clear();
    HR_JobEffDate.Clear();
    HR_BusinessUnit.Clear();
    HR_HomeDepartment.Clear();
    HR_DistinguishedName.Clear();
    HR_samAccountName.Clear();
    HR_UserSegment.Clear();
    AD_status.Clear();

    HR_Address.Clear();
    HR_City.Clear();
    HR_State.Clear();
    HR_PostalCode.Clear();
    HR_Country.Clear();

    HR_Workphone.Clear();
    HR_mobile.Clear();
    HR_Fax.Clear();
    HR_CostCenter.Clear();
    HR_office.Clear();
    HR_ProfilePath.Clear();
    HR_FTE.Clear();
    HR_needAccount.Clear();
    HR_needDisabled.Clear();
    HR_DivisionCode.Clear();


    if (aFileName == null)
        return false;
    object input = null;
    int tryTimes = 0;
    while (input == null)
    {
        try
        {
            tryTimes += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(aFileName);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(1000);
        }
        if (tryTimes > 100)
            break;
    }
    if (input == null)
        return false;

    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    hasManagerID = false;
    colHR_eeID = -1;

    colHR_firstName = -1;
    colHR_lastName = -1;
    colHR_middleName = -1;
    colHR_nickName = -1;
    colHR_workEmail = -1;
    colHR_status = -1;
    colHR_statusEffDate = -1;
    colHR_managerWorkEmail = -1;
    colHR_locationAux1 = -1;
    colHR_HomeDepartment = -1;
    colHR_DepartmentAux1 = -1;
    colHR_managerID = -1;
    colHR_managerDN = -1;
    colHR_Jobtitle = -1;
    colHR_JobEffDate = -1;
    colHR_BusinessUnit = -1;

    colHR_DistinguishedName = -1;
    colHR_samAccountName = -1;
    int colHR_UserSegment = -1;
    colAD_status = -1;
    colHR_Address = -1;
    colHR_City = -1;
    colHR_State = -1;
    colHR_PostalCode = -1;
    colHR_Country = -1;

    colHR_Workphone = -1;
    int colHR_mobile = -1;
    colHR_Fax = -1;
    // always use standard field name "CostCenter" in generated CSV file
    colHR_CostCenter = -1;
    colHR_office = -1;
    colHR_ProfilePath = -1;
    colHR_FTE = -1;
    colHR_OriginalHireDate = -1;
    colHR_HireDate = -1;
    colHR_NeedAccount = -1;
    colHR_needDisabled = -1;
    int colHR_DivisionCode = -1;

    int colName_mismatch = -1;
    int colSamID_mismatch = -1;


    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        if (aTitle == "lastname")
            colHR_lastName = HRcolCount;
        else if (aTitle == "firstname")
            colHR_firstName = HRcolCount;
        else if (aTitle == "middlename")
            colHR_middleName = HRcolCount;
        else if (aTitle == "nickname" || aTitle == "preferredname")
            colHR_nickName = HRcolCount;
        else if (aTitle == "employeeid" || aTitle == "workdayid")
            colHR_eeID = HRcolCount;
        else if (aTitle == "workemail_ad")
        {
            colHR_workEmail = HRcolCount;
            hasWorkEmail = true;
        }
        else if (aTitle == "status")
            colHR_status = HRcolCount;
        else if (aTitle == "statuseffectivedate" || aTitle == "statuseffdate")
            colHR_statusEffDate = HRcolCount;
        else if (aTitle == "status_ad")
            colAD_status = HRcolCount;
        else if (aTitle == "managerworkemail")
        {
            colHR_managerWorkEmail = HRcolCount;
            hasManagerWorkEmail = true;
        }
        else if (aTitle == "jobtitle")
        {
            colHR_Jobtitle = HRcolCount;
            hasJobTitle = true;
        }
        else if (aTitle == "jobeffdate")
            colHR_JobEffDate = HRcolCount;
        else if (aTitle == "businessunit")
            colHR_BusinessUnit = HRcolCount;
        else if (aTitle == "homedepartment")
        {
            colHR_HomeDepartment = HRcolCount;
            hasHomeDepartment = true;
        }
        else if (aTitle == "departmentaux1")
            colHR_DepartmentAux1 = HRcolCount;
        else if (aTitle == "locationaux1")
            colHR_locationAux1 = HRcolCount;
        else if (aTitle == "managerid" || aTitle == "managereeid")
        {
            colHR_managerID = HRcolCount;
            hasManagerID = true;
        }
        else if (aTitle == "managerdn")
            colHR_managerDN = HRcolCount;
        else if (aTitle == "distinguishedname")
            colHR_DistinguishedName = HRcolCount;
        else if (aTitle == "samaccountname")
            colHR_samAccountName = HRcolCount;
        else if (aTitle == "usersegment")
            colHR_UserSegment = HRcolCount;
        else if (aTitle == "address")
            colHR_Address = HRcolCount;
        else if (aTitle == "city")
            colHR_City = HRcolCount;
        else if (aTitle == "state")
            colHR_State = HRcolCount;
        else if (aTitle == "postalcode")
            colHR_PostalCode = HRcolCount;
        else if (aTitle == "country")
            colHR_Country = HRcolCount;
        else if (aTitle == "workphone")
            colHR_Workphone = HRcolCount;
        else if (aTitle == "primaryworkmobile")
            colHR_mobile = HRcolCount;
        else if (aTitle == "workfax")
            colHR_Fax = HRcolCount;
        else if (aTitle == "costcenter" || aTitle == "pager")
            colHR_CostCenter = HRcolCount;
        else if (aTitle == "office")
            colHR_office = HRcolCount;
        else if (aTitle == "profilepath")
            colHR_ProfilePath = HRcolCount;
        else if (aTitle == "empclass")
            colHR_FTE = HRcolCount;
        else if (aTitle == "originalhiredate")
            colHR_OriginalHireDate = HRcolCount;
        else if (aTitle == "hiredate")
            colHR_HireDate = HRcolCount;
        else if (aTitle == "needaccount")
            colHR_NeedAccount = HRcolCount;
        else if (aTitle == "divisioncode")
            colHR_DivisionCode = HRcolCount;
        else if (aTitle == "needdisabled")
            colHR_needDisabled = HRcolCount;
        else if (aTitle == "name_mismatch")
            colName_mismatch = HRcolCount;
        else if (aTitle == "samid_mismatch")
            colSamID_mismatch = HRcolCount;
        HRcolCount = HRcolCount + 1;
    }

    HRrowCount = 0;
    if (HRcolCount < 1)
    {
        swLog.WriteLine(aFileName + ": " + " has no column.");
        return false;
    }
    while ((!input.endofdata))
    {
        rows.Add(input.readfields);
        if (colHR_eeID < 0)
            HR_eeID.Add("");
        else
            // deal with strange EEID that may be contains ","
            HR_eeID.Add(rows.Item(HRrowCount)(colHR_eeID).replace(",", "_"));
        if (colHR_firstName < 0)
            HR_firstName.Add("");
        else
            HR_firstName.Add(rows.Item(HRrowCount)(colHR_firstName));
        if (colHR_lastName < 0)
            HR_lastName.Add("");
        else
            HR_lastName.Add(rows.Item(HRrowCount)(colHR_lastName));
        if (colHR_middleName < 0)
            HR_middleName.Add("");
        else
            HR_middleName.Add(rows.Item(HRrowCount)(colHR_middleName));
        if (colHR_nickName < 0)
            HR_nickName.Add("");
        else
            HR_nickName.Add(rows.Item(HRrowCount)(colHR_nickName));
        if (colHR_workEmail < 0)
        {
            HR_workEmail.Add("");
            HR_workEmailKey.Add("");
        }
        else
        {
            tmpStr = rows.Item(HRrowCount)(colHR_workEmail);
            if (Strings.InStr(tmpStr, "@") > 0)
            {
                HR_workEmail.Add(tmpStr);
                HR_workEmailKey.Add(Microsoft.VisualBasic.Left(Strings.LCase(tmpStr), Strings.InStr(tmpStr, "@"))); // only email key need LCase
            }
            else
            {
                HR_workEmail.Add("");
                HR_workEmailKey.Add("");
            }
        }

        if (colHR_status < 0)
            HR_status.Add("");
        else
            HR_status.Add(rows.Item(HRrowCount)(colHR_status));
        if (colHR_statusEffDate < 0)
            HR_statusEffDate.Add("");
        else
            HR_statusEffDate.Add(rows.Item(HRrowCount)(colHR_statusEffDate));

        if (colHR_managerWorkEmail < 0)
            HR_managerWorkEmail.Add("");
        else
            HR_managerWorkEmail.Add(LCase(rows.Item(HRrowCount)(colHR_managerWorkEmail)));
        if (colHR_locationAux1 < 0)
            HR_locationAux1.Add("");
        else
            HR_locationAux1.Add(rows.Item(HRrowCount)(colHR_locationAux1));
        if (colHR_managerID < 0)
            HR_managerID.Add("");
        else
            HR_managerID.Add(rows.Item(HRrowCount)(colHR_managerID));
        if (colHR_managerDN < 0)
            HR_managerDN.Add("");
        else
            HR_managerDN.Add(rows.Item(HRrowCount)(colHR_managerDN));

        if (hasJobTitle)
            HR_Jobtitle.Add(rows.Item(HRrowCount)(colHR_Jobtitle));
        else
            HR_Jobtitle.Add("");
        if (colHR_JobEffDate < 0)
            HR_JobEffDate.Add("");
        else
            HR_JobEffDate.Add(rows.Item(HRrowCount)(colHR_JobEffDate));
        if (colHR_BusinessUnit < 0)
            HR_BusinessUnit.Add("");
        else
            HR_BusinessUnit.Add(rows.Item(HRrowCount)(colHR_BusinessUnit));
        tmpStr = "";
        if (hasHomeDepartment)
            tmpStr = rows.Item(HRrowCount)(colHR_HomeDepartment);
        HR_HomeDepartment.Add(tmpStr);

        if (colHR_DistinguishedName < 0)
            HR_DistinguishedName.Add("");
        else
            HR_DistinguishedName.Add(rows.Item(HRrowCount)(colHR_DistinguishedName));
        if (colHR_samAccountName < 0)
            HR_samAccountName.Add("");
        else
            HR_samAccountName.Add(rows.Item(HRrowCount)(colHR_samAccountName));
        if (colHR_UserSegment < 0)
            HR_UserSegment.Add("");
        else
            HR_UserSegment.Add(rows.Item(HRrowCount)(colHR_UserSegment));
        if (colAD_status < 0)
            AD_status.Add("");
        else
            AD_status.Add(rows.Item(HRrowCount)(colAD_status));
        if (colHR_Address < 0)
            HR_Address.Add("");
        else
            HR_Address.Add(rows.Item(HRrowCount)(colHR_Address));
        if (colHR_City < 0)
            HR_City.Add("");
        else
            HR_City.Add(rows.Item(HRrowCount)(colHR_City));
        if (colHR_State < 0)
            HR_State.Add("");
        else
            HR_State.Add(rows.Item(HRrowCount)(colHR_State));
        if (colHR_PostalCode < 0)
            HR_PostalCode.Add("");
        else
            HR_PostalCode.Add(rows.Item(HRrowCount)(colHR_PostalCode));
        if (colHR_Country < 0)
            HR_Country.Add("");
        else
            HR_Country.Add(rows.Item(HRrowCount)(colHR_Country));

        if (colHR_Workphone < 0)
            HR_Workphone.Add("");
        else
            HR_Workphone.Add(rows.Item(HRrowCount)(colHR_Workphone));
        if (colHR_mobile < 0)
            HR_mobile.Add("");
        else
            HR_mobile.Add(rows.Item(HRrowCount)(colHR_mobile));
        if (colHR_Fax < 0)
            HR_Fax.Add("");
        else
            HR_Fax.Add(rows.Item(HRrowCount)(colHR_Fax));
        if (colHR_CostCenter < 0)
            HR_CostCenter.Add("");
        else
            HR_CostCenter.Add(rows.Item(HRrowCount)(colHR_CostCenter));
        if (colHR_office < 0)
            HR_office.Add("");
        else
            HR_office.Add(rows.Item(HRrowCount)(colHR_office));
        if (colHR_ProfilePath < 0)
            HR_ProfilePath.Add("");
        else
            HR_ProfilePath.Add(rows.Item(HRrowCount)(colHR_ProfilePath));
        if (colHR_FTE < 0)
            HR_FTE.Add("");
        else
            HR_FTE.Add(rows.Item(HRrowCount)(colHR_FTE));
        if (colHR_OriginalHireDate < 0)
            HR_OriginalHireDate.Add("");
        else
            HR_OriginalHireDate.Add(rows.Item(HRrowCount)(colHR_OriginalHireDate));
        if (colHR_HireDate < 0)
            HR_HireDate.Add("");
        else
            HR_HireDate.Add(rows.Item(HRrowCount)(colHR_HireDate));
        if (colHR_NeedAccount < 0)
            HR_needAccount.Add("");
        else
            HR_needAccount.Add(rows.Item(HRrowCount)(colHR_NeedAccount));
        if (colHR_DivisionCode < 0)
            HR_DivisionCode.Add("");
        else
            HR_DivisionCode.Add(rows.Item(HRrowCount)(colHR_DivisionCode));
        if (colHR_needDisabled < 0)
            HR_needDisabled.Add("");
        else
            HR_needDisabled.Add(rows.Item(HRrowCount)(colHR_needDisabled));
        if (colName_mismatch < 0)
            HR_reqID.Add("");
        else
        {
            string nameMismatchStr = rows.Item(HRrowCount)(colName_mismatch);
            if (nameMismatchStr.Contains("need set EEID for reqID"))
                // HR_reqID is used for storing a mark that need update EEID for a new hire that only has reqID but no EEID
                HR_reqID.Add("need set EEID for reqID");
            else if (nameMismatchStr != "")
            {
                // HR_reqID also used for storing a mark that cannot update AD properties if EEID is blank and has name mismatch
                if (colSamID_mismatch < 0)
                    HR_reqID.Add("Name mismatch: " + nameMismatchStr);
                else
                {
                    string samIDmismatchStr = rows.Item(HRrowCount)(colSamID_mismatch);
                    if (samIDmismatchStr != "")
                        HR_reqID.Add("Name mismatch: " + nameMismatchStr);
                    else
                        HR_reqID.Add("");
                }
            }
            else
                HR_reqID.Add("");
        }
        // swLog.WriteLine(HR_eeID.Item(HRrowCount) & " " & HR_firstName.Item(HRrowCount) & " " & HR_lastName.Item(HRrowCount))
        HRrowCount = HRrowCount + 1;
    }
    input.Close();

    if (HRrowCount > 0)
    {
        Console.WriteLine(aFileName + ": " + HRrowCount + " records, please wait...");
        swLog.WriteLine(aFileName + ": " + HRrowCount + " records, please wait...");
    }
    else
        swLog.WriteLine(aFileName + ": " + HRrowCount + " record, no need process.");
    return true;
}

public static void getYYYYMMDD_hhtt()
{
    string YYYY, MM, DD, hh, tt;
    tmpStr = DateTime.Now();
    YYYY = DateTime.Year(tmpStr);
    MM = DateTime.Month(tmpStr);
    if (Strings.Len(MM) < 2)
        MM = "0" + MM;
    DD = DateTime.Day(tmpStr);
    if (Strings.Len(DD) < 2)
        DD = "0" + DD;
    hh = DateTime.Hour(tmpStr);
    if (Strings.Len(hh) < 2)
        hh = "0" + hh;
    tt = Strings.Trim(DateTime.Minute(tmpStr));
    if (Strings.Len(tt) < 2)
        tt = "0" + tt;
    getYYYYMMDD_hhtt = YYYY + MM + DD + "_" + hh + tt;
}

private static void DownloadHRB(string aFilename)
{
    return;

    swLog.WriteLine("Downloading " + hrMIS + " file from FTP server...");
    getFileFromFTP(ftpDownloadFolder + @"\" + ftpFileName + ".csv", ftpFileName + ".csv", "ftp.eease.com", "ehealth", "1yAwTZxu");
    tmpStr = @"C:\gnupg\gpg --batch -o " + aFilename + " -d " + ftpDownloadFolder + @"\" + ftpFileName + ".csv";

    Process p = new Process();
    p.StartInfo.FileName = @"C:\Windows\system32\cmd.exe";
    p.StartInfo.Arguments = "/c " + tmpStr;
    p.StartInfo.UseShellExecute = false;
    p.StartInfo.RedirectStandardInput = true;
    p.StartInfo.RedirectStandardOutput = true;
    p.StartInfo.RedirectStandardError = true;
    p.StartInfo.CreateNoWindow = false;
    p.StartInfo.WorkingDirectory = strWorkFolder;
    p.Start();
    p.StandardInput.WriteLine("Exit");
    string output = p.StandardOutput.ReadToEnd();
    p.Close();
    if (Strings.Trim(output) != "")
    {
        swLog.WriteLine("gpg error: " + output);
        return;
    }
    swLog.WriteLine("Downloaded OK: " + aFilename);
}

private static void getFileFromFTP(string localFile, string remoteFile, string host, string username, string password)
{
    string URI = "FTP://" + host + "/" + remoteFile;
    FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create(URI);

    ftp.Credentials = new NetworkCredential(username, password);
    ftp.KeepAlive = false;
    ftp.UseBinary = false;
    // Define the action required (in this case, download a file)
    ftp.Method = WebRequestMethods.Ftp.DownloadFile;

    using (FtpWebResponse response = (FtpWebResponse)ftp.GetResponse())
    {
        using (System.IO.Stream responseStream = response.GetResponseStream())
        {
            // loop to read & write to file
            using (System.IO.FileStream fs = new System.IO.FileStream(localFile, System.IO.FileMode.Create))
            {
                byte[] buffer = new byte[2048];
                int read = 0;
                do
                {
                    read = responseStream.Read(buffer, 0, buffer.Length);
                    fs.Write(buffer, 0, read);
                }
                while (!read == 0); // see Note(1)
                responseStream.Close();
                fs.Flush();
                fs.Close();
            }
            responseStream.Close();
        }
        response.Close();
    }
}

private static DateTime GetUserLastLogonTime(string TargetUsername)
{
    AccountExpirationDate = null;
    DateTime LastLogonDate = "01/01/1900";
    if (TargetUsername != "")
    {
        using (DirectorySearcher Searcher = new DirectorySearcher(new DirectoryEntry()))
        {
            Searcher.Filter = "(&(sAMAccountName=" + TargetUsername + "))";
            Searcher.PropertiesToLoad.Add("lastLogonTimestamp");
            Searcher.PropertiesToLoad.Add("AccountExpires"); // don't use AccountExpirationDate
            SearchResult UserAccount = Searcher.FindOne;

            Int64 RawExpires = 0;
            DateTime ExpirationDate;
            try
            {
                RawExpires = System.Convert.ToInt64(UserAccount.Properties("accountExpires")(0));
                if (RawExpires > 0)
                {
                    ExpirationDate = DateTime.FromFileTime(RawExpires);
                    AccountExpirationDate = ExpirationDate.ToString();
                }
            }
            catch (Exception ex)
            {
            }

            Int64 RawLastLogon = 0;
            try
            {
                RawLastLogon = System.Convert.ToInt64(UserAccount.Properties("lastLogonTimestamp")(0)); // error: CN=sprivett,
                LastLogonDate = DateTime.FromFileTime(RawLastLogon);
            }
            catch (Exception ex)
            {
            }
        }
    }
    return LastLogonDate;
}

private static void getLocationITcontactEmail()
{
    string ITcontact1 = "";
    string ITcontact2 = "";
    string ITcontact3 = "";
    string ITcontact4 = "";
    string DNOU = "";
    if (LocationAux1 == "" && DistinguishedName != "")
    {
        // use DN to decide the location approximately
        DNOU = Strings.UCase(DistinguishedName).Replace(",OU=EHI," + Strings.UCase(strRoot), "");
        DNOU = DNOU.Substring(DNOU.LastIndexOf(",") + 1);
    }

    ArrayList ITemailList = new ArrayList();
    if (LocationAux1 != "")
        ITemailList = getAssociatedITemailFromDB(LocationAux1);
    else if (DNOU != "")
    {
        // RMT usually in OU=GR, OU=SL
        string mapLocation = "SCS";
        switch (DNOU)
        {
            case "OU=SCS":
                {
                    mapLocation = "SCS";
                    break;
                }

            case "OU=MV":
                {
                    mapLocation = "SCS"; // "MTNV" 'DC,SF
                    break;
                }

            case "OU=GR":
                {
                    mapLocation = "GDRV"; // RMT
                    break;
                }

            case "OU=SL":
                {
                    mapLocation = "UT"; // RMT
                    break;
                }

            case "OU=XM":
                {
                    mapLocation = "XM";
                    break;
                }

            case "OU=MAYNARD":
                {
                    mapLocation = "MA";
                    break;
                }
        }
        ITemailList = getAssociatedITemailFromDB(mapLocation);
    }

    if (ITemailList.Count > 0)
    {
        ITcontact1 = ITemailList.Item(0);
        if (ITemailList.Count > 1)
        {
            ITcontact2 = ITemailList.Item(1);
            if (ITemailList.Count > 2)
            {
                ITcontact3 = ITemailList.Item(2);
                if (ITemailList.Count > 3)
                    ITcontact4 = ITemailList.Item(3);
            }
        }
    }

    if (ITcontact1 != "" && emailCClist.IndexOf(ITcontact1) < 0)
        emailCClist.Add(ITcontact1);
    if (ITcontact2 != "" && emailCClist.IndexOf(ITcontact2) < 0)
        emailCClist.Add(ITcontact2);
    if (ITcontact3 != "" && emailCClist.IndexOf(ITcontact3) < 0)
        emailCClist.Add(ITcontact3);
    if (ITcontact4 != "" && emailCClist.IndexOf(ITcontact4) < 0)
        emailCClist.Add(ITcontact4);
}

private static void mapOU_Address()
{
    // From: Huan Nguyen  Sent: 16 April, 2014 02:09
    // So there needs to be a distinction between the location that they are charged to (via home department) 
    // and the actual physical location they work at (based on location code).  
    if (LocationAux1 == "")
    {
        PhysicalDeliveryOfficeName = "";
        return;
    }
    // USA Remote   TPA: Remote    TPA: Hobart, IN
    if (Strings.UCase(LocationAux1) == "RMT" || Strings.InStr(Strings.UCase(LocationAux1), "REMOTE") > 0 || Strings.UCase(LocationAux1).StartsWith("TPA:"))
    {
        strOUnet = "OU=Users,OU=MV,OU=EHI,";
        PhysicalDeliveryOfficeName = "Remote";
        Address = ""; // USA Remote	123 Main Street

        City = "Remote";
        // TODO: Workday should use Supplier
        if (TPA_location_vendor != "")
        {
            string[] lv;
            try
            {
                lv = TPA_location_vendor.Split(";");
            }
            catch (Exception ex)
            {
            }
            if (lv.Length > 1)
                TPA_vendor = Strings.Trim(lv[1]);
            else if (lv.Length > 0)
            {
                TPA_location = lv[0];
                TPA_vendor = TPA_location;
            }
            if (TPA_vendor != "")
                City = "Remote - " + TPA_vendor;
        }
        // State = ""
        PostalCode = "";
        Country = "US";
        if (Workphone == "")
            Workphone = "650.584.2700";
        if (Fax == "")
            Fax = "650.961.2110";
    }
    else if (Strings.UCase(LocationAux1) == "SCS" || Strings.UCase(LocationAux1).StartsWith("SANTA CLARA"))
    {
        strOUnet = "OU=Users,OU=MV,OU=EHI,";
        PhysicalDeliveryOfficeName = "Santa Clara Square";
        Address = "2625 Augustine Drive, Second Floor";
        City = "Santa Clara";
        State = "CA";
        PostalCode = "95054";
        Country = "US";
        Workphone = "650.584.2700";
        Fax = "650.961.2110";
    }
    else if (Strings.UCase(LocationAux1) == "SF" || Strings.InStr(Strings.UCase(LocationAux1), "SAN FRANCISCO") > 0)
    {
        // SF(): "128 King Street, 4th floor" "San Francisco" "CA" "94107" "US"
        strOUnet = "OU=Users,OU=MV,OU=EHI,"; // OU=Multi,OU=EHI,?
        PhysicalDeliveryOfficeName = "San Francisco";
        if (Address == "")
            Address = "128 King Street, 4th floor";
        if (City == "")
            City = "San Francisco";
        if (State == "")
            State = "CA";
        if (PostalCode == "")
            PostalCode = "94107";
        Country = "US";
        if (Workphone == "")
            Workphone = "650.584.2700";
        if (Fax == "")
            Fax = "650.961.2110";
    }
    else if (Strings.UCase(LocationAux1) == "GDRV" || Strings.InStr(Strings.UCase(LocationAux1), "GOLD RIVER") > 0)
    {
        strOUnet = "OU=Users,OU=GR,OU=EHI,";
        PhysicalDeliveryOfficeName = "Gold River";
        if (Address == "")
            Address = "11919 Foundation Place, Suite 100";
        if (City == "")
            City = "Gold River";
        if (State == "")
            State = "CA";
        if (PostalCode == "")
            PostalCode = "95670";
        Country = "US";
        if (Workphone == "")
            Workphone = "800.977.8860";
        if (Fax == "")
            Fax = "916.608.6189";
    }
    else if (Strings.UCase(LocationAux1) == "UT" || Strings.InStr(Strings.UCase(LocationAux1), "SALT LAKE CITY") > 0)
    {
        strOUnet = "OU=Users,OU=SL,OU=EHI,";
        PhysicalDeliveryOfficeName = "Salt Lake City";
        if (Address == "")
            Address = "2875 South Decker Lake Drive, Suite 400";
        if (City == "")
            City = "Salt Lake City";
        if (State == "")
            State = "UT";
        if (PostalCode == "")
            PostalCode = "84119";
        Country = "US";
        if (Workphone == "")
            Workphone = "650.584.2700";
        if (Fax == "")
            Fax = "650.961.2110";
    }
    else if (Strings.UCase(LocationAux1) == "XM" || Strings.InStr(Strings.UCase(LocationAux1), "XIAMEN") > 0)
    {
        strOUnet = "OU=Users,OU=XM,OU=EHI,"; // CN=Sophie Chen,OU=Users,OU=XM,OU=EHI,
                                             // if XMFinance:  CN=schen,OU=XMFinance,OU=Users,OU=GR,OU=EHI,
        if (HomeDepartment.Contains("(Ubao)") || Strings.LCase(Workemail).Contains("@ubao.com"))
            strOUnet = "OU=Ubao," + strOUnet;
        PhysicalDeliveryOfficeName = "Xiamen";
        if (Address == "")
            Address = "9F, Chuangxin Building, Software Park";
        if (City == "")
            City = "Xiamen";
        if (State == "")
            State = "Fujian";
        if (PostalCode == "")
            PostalCode = "361005";
        Country = "CN";
        if (Workphone == "")
            Workphone = "011-86-592-2517000";
        if (Fax == "")
            Fax = "011-86-592-2517111";
    }
    else if (Strings.UCase(LocationAux1) == "DC" || Strings.InStr(Strings.UCase(LocationAux1), "WASHINGTON") > 0)
    {
        strOUnet = "OU=Users,OU=MV,OU=EHI,";
        PhysicalDeliveryOfficeName = "Washington D.C.";
        if (Address == "")
            Address = "1615 L Street, NW Suite 400";// 1100"
        if (City == "")
            City = "Washington";// D.C.
        if (State == "")
            State = "DC";
        if (PostalCode == "")
            PostalCode = "20036";
        Country = "US";
        if (Workphone == "")
            Workphone = "202.506.1096";
        if (Fax == "")
            Fax = "650.961.2110";
    }
    else if (Strings.UCase(LocationAux1).Contains("AUSTIN"))
    {
        strOUnet = "OU=Users,OU=AUS,OU=EHI,";
        PhysicalDeliveryOfficeName = "Austin";
        if (Address == "")
            Address = "13620 Ranch Road 620, Suite A250";// "505 E Palm Valley Blvd Ste 240, Round Rock"
        if (City == "")
            City = "Austin";
        if (State == "")
            State = "TX";
        if (PostalCode == "")
            PostalCode = "78717";// "78664"
        Country = "US";
        if (Workphone == "")
            Workphone = "(866) 894-3258";
        if (Fax == "")
        {
        }
    }
    else if (Strings.UCase(LocationAux1).Contains("IN") || Strings.InStr(Strings.UCase(LocationAux1), "INDIANAPOLIS") > 0)
    {
        strOUnet = "OU=Users,OU=IN,OU=EHI,";
        PhysicalDeliveryOfficeName = "Indianapolis";
        if (Address == "")
            Address = "9190 Priority Way West Drive Suite 300";// "135 N Pennsylvania St, Ste 1700"
        if (City == "")
            City = "Indianapolis";
        if (State == "")
            State = "IN";
        if (PostalCode == "")
            PostalCode = "46240";
        Country = "US";
        if (Workphone == "")
        {
        }
        if (Fax == "")
        {
        }
    }
    else if (Strings.UCase(LocationAux1) == "MA" || Strings.InStr(Strings.UCase(LocationAux1), "WESTFORD") > 0)
    {
        strOUnet = "OU=Users,OU=Maynard,OU=EHI,";
        PhysicalDeliveryOfficeName = "Westford";
        if (Address == "")
            Address = "2 Technology Park Drive";
        if (City == "")
            City = "Westford";
        if (State == "")
            State = "MA";
        if (PostalCode == "")
            PostalCode = "01886";
        Country = "US";
        if (Workphone == "")
            Workphone = "650.584.2700";
        if (Fax == "")
            Fax = "650.961.2110";
    }
    else
    {
        // default OU?
        strOUnet = "OU=unknown,OU=EHI,";
        PhysicalDeliveryOfficeName = "";
        Workphone = "";
        Fax = "";
    }
}

// always set default value before read actual value
private static void set_ScriptPath_HomeDirectory(object aSamID)
{
    // if aSamID = "", will provide root string only
    ScriptPath = "";
    ProfilePath = "";
    HomeDrive = "";
    HomeDirectory = "";
    if (LocationAux1 == "")
        return;
    HomeDrive = "U:";
    // USA Remote   TPA: Remote    TPA: Hobart, IN
    if (Strings.UCase(LocationAux1) == "RMT" || Strings.InStr(Strings.UCase(LocationAux1), "REMOTE") > 0)
    {
        HomeDrive = "";
        ScriptPath = ""; // "user.bat"
        if (aSamID != "")
            HomeDirectory = "";// "\\sjdfs01\" & aSamID & "$"
    }
    else if (Strings.UCase(LocationAux1) == "SCS" || Strings.UCase(LocationAux1).StartsWith("SANTA CLARA"))
    {
        ScriptPath = "user.bat";
        if (aSamID != "")
            HomeDirectory = @"\\sjdfs01\" + aSamID + "$";
    }
    else if (Strings.UCase(LocationAux1) == "SF" || Strings.InStr(Strings.UCase(LocationAux1), "SAN FRANCISCO") > 0)
    {
        ScriptPath = "user.bat";
        if (aSamID != "")
            HomeDirectory = @"\\sjdfs01\" + aSamID + "$";
    }
    else if (Strings.UCase(LocationAux1) == "GDRV" || Strings.InStr(Strings.UCase(LocationAux1), "GOLD RIVER") > 0)
    {
        ScriptPath = "ntlogin.bat";
        // if memberof GR_Finance or RRS then
        if (Strings.InStr(memberOf, "GR_Finance;") > 0)
            ScriptPath = "finance.bat";
        if (aSamID != "")
        {
            ProfilePath = @"\\grfs01\profiles$\" + aSamID;
            HomeDirectory = @"\\grfs01\" + aSamID + "$";
        }
    }
    else if (Strings.UCase(LocationAux1) == "UT" || Strings.InStr(Strings.UCase(LocationAux1), "SALT LAKE CITY") > 0)
    {
        ScriptPath = "sllogin.bat";
        if (aSamID != "")
        {
            ProfilePath = @"\\sldfs01\profile$\" + aSamID;
            HomeDirectory = @"\\sldfs01\" + aSamID + "$";
        }
    }
    else if (Strings.UCase(LocationAux1) == "XM" || Strings.InStr(Strings.UCase(LocationAux1), "XIAMEN") > 0)
    {
        if (HomeDepartment.Contains("Revenue Accounting"))
            ScriptPath = "xmfinance.vbs";
        else
            ScriptPath = "testxmlogon.vbs";// xmlogon.bat testxmlogon.bat
        if (aSamID != "")
            HomeDirectory = @"\\xmfilesrv01\HOME$\" + aSamID;
    }
    else if (Strings.UCase(LocationAux1) == "DC" || Strings.InStr(Strings.UCase(LocationAux1), "WASHINGTON") > 0)
    {
        ScriptPath = "user.bat";
        if (aSamID != "")
            HomeDirectory = @"\\wolverine\" + aSamID + "$";
    }
    else if (Strings.UCase(LocationAux1).Contains("AUSTIN"))
    {
    }
    else if (Strings.UCase(LocationAux1) == "IN" || Strings.InStr(Strings.UCase(LocationAux1), "INDIANAPOLIS") > 0)
    {
    }
    else if (Strings.UCase(LocationAux1) == "MA" || Strings.InStr(Strings.UCase(LocationAux1), "WESTFORD") > 0)
    {
        ScriptPath = "malogin.bat";
        if (aSamID != "")
        {
            // ProfilePath = "\\mamsutil01\Profiles$\" & aSamID
            ProfilePath = @"\\wfpritfs01\Profiles$\" + aSamID;
            // HomeDirectory = "\\mamsutil01\" & aSamID & "$"
            HomeDirectory = @"\\wfpritfs01\Home$" + aSamID;
        }
    }
}

public static bool checkAD()
{
    if (HRrowCount < 1)
    {
        swLog.WriteLine(hrMIS + " file has been opened, but no data to process.");
        return false;
    }

    int minRow = 1;
    int maxRow = HRrowCount;

    DateTime startTime = DateTime.Now;
    Console.WriteLine(startTime + " checking AD, please wait...");
    swLog.WriteLine(startTime + " checking AD, please wait...");

    // set all group names for use in writeAccountFound()
    FTEnonFTEgroupNamesList.Clear();
    FTEnonFTEgroupNamesList.Add("RemoteUsers_FTE");
    FTEnonFTEgroupNamesList.Add("RemoteUsers_NonFTE");
    FTEnonFTEgroupNamesList.Add("SantaClara_FTE");
    FTEnonFTEgroupNamesList.Add("SantaClara_nonFTE");
    FTEnonFTEgroupNamesList.Add("SanFrancisco_FTE");
    FTEnonFTEgroupNamesList.Add("SanFrancisco_NonFTE");
    FTEnonFTEgroupNamesList.Add("GoldRiver_FTE");
    FTEnonFTEgroupNamesList.Add("GoldRiver_NonFTE");
    FTEnonFTEgroupNamesList.Add("SaltLake_FTE");
    FTEnonFTEgroupNamesList.Add("SaltLake_NonFTE");
    // FTEnonFTEgroupNamesList.Add("Westford_FTE")
    // FTEnonFTEgroupNamesList.Add("Westford_NonFTE")
    FTEnonFTEgroupNamesList.Add("Xiamen_FTE");
    FTEnonFTEgroupNamesList.Add("Xiamen_NonFTE");
    FTEnonFTEgroupNamesList.Add("Austin_FTE");
    FTEnonFTEgroupNamesList.Add("Austin_nonFTE");
    FTEnonFTEgroupNamesList.Add("Indianapolis_FTE");
    FTEnonFTEgroupNamesList.Add("Indianapolis_NonFTE");

    tmpStr = getYYYYMMDD_hhtt();
    strExportFileEntire = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADuserCheck" + tmpStr + ".csv";
    strExportFileFound = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADfound" + tmpStr + ".csv";
    strExportFileNotFound = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADnotFound" + tmpStr + ".csv";
    strExportFileNewUser = strWorkFolder + @"\" + strReportSubFolder + @"\" + "NewHire" + tmpStr + ".csv";
    strExportFileManager = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ManagerEmployeeID" + tmpStr + ".csv";
    strExportFileNeedDisabled = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADneedDisabled" + tmpStr + ".csv";
    strReportfileDiscrepancy = strWorkFolder + @"\" + strReportSubFolder + @"\" + hrMIS + "_ADdiscrepancy" + tmpStr + ".csv";
    strReportfileEmailMismatch = strWorkFolder + @"\" + strReportSubFolder + @"\" + hrMIS + "_ADemailMismatch" + tmpStr + ".csv";
    strExportFileGroupAction = strWorkFolder + @"\" + strReportSubFolder + @"\" + "GroupAction" + tmpStr + ".csv";
    try
    {
        reportfileGroupAction = new StreamWriter(strExportFileGroupAction, false);
        reportfileDiscrepancy = new StreamWriter(strReportfileDiscrepancy, false);
        // reportfileEmailMismatch = New StreamWriter(strReportfileEmailMismatch, False)
        reportfileEntire = new StreamWriter(strExportFileEntire, false);
        reportfileFound = new StreamWriter(strExportFileFound, false);
        reportfileNotFound = new StreamWriter(strExportFileNotFound, false);
        reportfileNeedDisabled = new StreamWriter(strExportFileNeedDisabled, false);
    }
    catch (Exception ex)
    {
        swLog.WriteLine("file writing failed: " + ex.ToString());
        return false;
    }
    reportfileEntire.WriteLine(strTitle + ",defaultGroup,memberOf,Notes,needDisabled");
    reportfileFound.WriteLine(strTitle + ",defaultGroup,memberOf,Notes,needDisabled");
    reportfileNotFound.WriteLine(strTitle + ",defaultGroup,memberOf,Notes,container,profilePath,scriptPath,homeDirectory,homeDrive,MOC_enabled,workEmail");
    reportfileNeedDisabled.WriteLine(strTitle + ",defaultGroup,memberOf,Notes,needDisabled");
    reportfileGroupAction.writeline("groupDN,action,memberDN");
    reportfileDiscrepancy.writeline("WorkdayID,previousID,FirstName,LastName,NickName,Location,status,description,detail");
    // reportfileEmailMismatch.writeline("employeeID,FirstName,LastName,NickName,Location,status,description,detail")

    intRowCount = 0;
    intNotFoundCount = 0;
    discrepancyCount = 0;
    discrepancyCountXM = 0;
    emailMismatchCount = 0;
    emailMismatchCountXM = 0;

    hasError = false;
    onLeaveSamIDs = "";
    emailCClist.Clear();
    LeaveReturnEEID.Clear();
    for (int i = 0; i <= HRrowCount - 1; i++)
    {
        if ((i < minRow - 1))
            continue;
        if ((i >= maxRow))
            break;
        intRowCount = intRowCount + 1;

        FirstName = Trim(HR_firstName.Item(i));
        LastName = Trim(HR_lastName.Item(i));
        MiddleName = Trim(HR_middleName.Item(i));

        NickName = Trim(HR_nickName.Item(i));
        NickName = NickName.Replace("(", "");
        NickName = NickName.Replace(")", "");
        wkdID = Trim(WDID(i));
        employeeID = Trim(HR_eeID.Item(i));
        preID = Trim(previousID(i));
        reqID = Trim(HR_reqID.Item(i));
        LocationAux1 = UCase(Trim(HR_locationAux1.Item(i)));
        if (LocationAux1 == "")
            continue;
        else if (forXM_US != "")
        {
            if (forXM_US == "US")
            {
                if (LocationAux1 == "XM" || LocationAux1.Contains("XIAMEN"))
                    continue;
            }
            else if (LocationAux1 != "XM" && !LocationAux1.Contains("XIAMEN"))
                continue;
        }

        isXM = false;
        if (LocationAux1 == "XM" || LocationAux1.Contains("XIAMEN"))
            isXM = true;

        Jobtitle = Trim(HR_Jobtitle.Item(i)); // HR connection file has no JobTitle and Manager value for new hires before hire date, so Jobtitle may be blank
        JobEffDate = Trim(HR_JobEffDate.Item(i));
        BusinessUnit = Trim(HR_BusinessUnit(i));
        HomeDepartment = Trim(HR_HomeDepartment.Item(i)); // HR_HomeDepartment already prepared by openCSVFileForChecking()
        UserSegment = Trim(HR_UserSegment(i));
        if (FirstName == "" || LastName == "")
        {
            swLog.WriteLine("row #" + i + " FirstName or LastName is blank, please check!");
            continue;
        }

        Workemail = Trim(HR_workEmail.Item(i));

        Status = Trim(HR_status.Item(i));
        if (Status == "Leave" || Status == "L")
            Status = "On Leave";
        else if (Status == "Deceased" || Status == "T")
            Status = "Terminated";
        else if (Status == "A")
            Status = "Active";
        statusEffDateStr = Trim(HR_statusEffDate.Item(i));

        if (Status == "Terminated")
        {
            if (forLeaveOnly)
                continue;
        }
        else if (Status == "Active")
        {
            int OnLeaveIndex = DB_OnLeaveEEID.IndexOf(employeeID);
            if (OnLeaveIndex < 0)
            {
                if (forLeaveOnly)
                    // Dim leaveReturningIndex As Integer = DB_LeaveReturningEEID.IndexOf(employeeID)
                    continue;
            }
            else
                LeaveReturnEEID.Add(employeeID);
        }

        NeedAccount = Trim(HR_needAccount.Item(i));
        DivisionCode = Trim(HR_DivisionCode(i));
        EmpClass = Trim(HR_FTE.Item(i));
        OriginalHireDate = Trim(HR_OriginalHireDate.Item(i));
        HireDate = Trim(HR_HireDate.Item(i));
        ManagerID = Trim(HR_managerID.Item(i));
        if (ManagerID != "")
        {
            while ((ManagerID.Length < 6))
                ManagerID = "0" + ManagerID;
        }
        ManagerDN = ""; // Trim(HR_managerDN.Item(i))
        ManagerWorkEmail = Trim(HR_managerWorkEmail.Item(i)); // HR manager work email may be not the same as AD

        Workphone = Trim(HR_Workphone.Item(i));
        Mobile = Trim(HR_mobile.Item(i));
        Fax = Trim(HR_Fax.Item(i));
        CostCenter = Trim(HR_CostCenter.Item(i));
        if (CostCenter != "")
        {
            if (isXM)
                CostCenter = "1" + CostCenter;
            else
                CostCenter = "0" + CostCenter;
        }

        Address = "";
        City = "";
        State = "";
        PostalCode = "";
        Country = "";
        TPA_location_vendor = Trim(HR_TPA_location_vendor.Item(i));
        if (needUseHRstreetAddress)
        {
            Address = Trim(HR_Address.Item(i));
            if (TPA_location_vendor != "")
                Address = TPA_location;// TPA_location_vendor
            if (Trim(HR_AddressStreet2.Item(i)) != "")
                Address += ", " + Trim(HR_AddressStreet2.Item(i));
            if (Trim(HR_AddressStreet3.Item(i)) != "")
                Address += ", " + Trim(HR_AddressStreet3.Item(i));
            City = HR_City.Item(i);
            State = HR_State.Item(i);
            PostalCode = HR_PostalCode.Item(i);
            if (PostalCode != "")
            {
                while ((PostalCode.Length < 5))
                    PostalCode = "0" + PostalCode;
            }
            Country = HR_Country.Item(i);
        }

        PObox = "";
        mapOU_Address();

        HR_AccountName = LCase(HR_samAccountName.Item(i));
        samAccountName = HR_AccountName;
        if (NeedAccount == noNeedAccountValueHR && samAccountName != "")
        {
            if (Status != "Terminated")
            {
                if (LCase(HR_FTE(i)) != "non-employee")
                {
                    reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",no need AD account,but " + hrMIS + " input " + samAccountName);
                    addDiscrepancyCount(LocationAux1);
                }
            }
            samAccountName = "";
        }
        samIDprovided = true;
        if (samAccountName == "")
        {
            samIDprovided = false;
            // If employeeID = "" Then 'only guess samID if eeID is blank
            if (isXM)
            {
                if (NickName == "")
                    samAccountName = Strings.LCase(FirstName.Replace(" ", "")) + Strings.LCase(Microsoft.VisualBasic.Left(Strings.LTrim(LastName), 1));
                else
                    samAccountName = Strings.LCase(NickName.Replace(" ", "")) + Strings.LCase(Microsoft.VisualBasic.Left(LastName, 1));
            }
            else if (NickName == "")
                samAccountName = Strings.LCase(Microsoft.VisualBasic.Left(FirstName, 1)) + Strings.LCase(LastName);
            else
                samAccountName = Strings.LCase(Microsoft.VisualBasic.Left(NickName, 1)) + Strings.LCase(LastName);
            // Xiamen also has "'" included in name
            samAccountName = samAccountName.Replace("'", ""); // James O'Brien 101227
            samAccountName = samAccountName.Replace(" ", ""); // lafraid of hawk
                                                              // From: Philip Zinn Sent: Tuesday, August 26, 2014 7:44 AM  RE: non-standard AD account and email address
                                                              // “when an AD user account has a hyphen in it, Back Office accounts fail during creation”
                                                              // so we should always remove "-"
            samAccountName = samAccountName.Replace("-", "");
            samAccountName = samAccountName.Replace(",", ""); // 300661 Albert (Disandro, Jr)
        }
        hasError = false;
        queryADforHRrow(i);
        if (hasError)
        {
            hasWriteEntireError = true;
            swLog.WriteLine("====>errors occured in queryADforHRrow(" + i + ")!.");
            swLog.WriteLine("");
        }
    }

    reportfileNotFound.Close();
    reportfileEntire.Close();
    reportfileFound.Close();

    string theLine = "";
    int noHRrecord = crossCheckHR();
    if (noHRrecord > 0)
    {
        theLine = noHRrecord + " AD user";
        if (noHRrecord > 1)
            theLine += "s";
        theLine += " has no corresponding " + hrMIS + " record that may need set disabled!";
        swLog.WriteLine(theLine);
    }

    if (needCheckNewUsersNotLogon)
    {
        int NewHireInactiveCount = checkNewUsersNotLogon();
        theLine = NewHireInactiveCount + " new AD user";
        if (NewHireInactiveCount > 1)
            theLine += "s";
        theLine += " has not logon over " + DaysInactiveLimitNewHire + " days that may need set disabled!";
        swLog.WriteLine(theLine);
    }

    reportfileGroupAction.close();
    reportfileNeedDisabled.close();
    reportfileDiscrepancy.Close();
    // reportfileEmailMismatch.Close()

    Console.WriteLine(DateTime.Now + " checked OK. Spent " + DateTime.DateDiff("s", startTime, DateTime.Now) + " seconds.");
    swLog.WriteLine(DateTime.Now + " checked OK. Spent " + DateTime.DateDiff("s", startTime, DateTime.Now) + " seconds.");

    swLog.WriteLine("");
    swLog.WriteLine("Checking " + strRoot + ", " + intRowCount + " users processed.");
    swLog.WriteLine("All current AD users is list in " + AllUsersCSVFileName + ".");
    swLog.WriteLine("Checking result for all is in " + strExportFileEntire + ".");
    swLog.WriteLine("Found users is also in " + strExportFileFound + ".");
    swLog.WriteLine("Not found is also in " + strExportFileNotFound + ".");
    swLog.WriteLine("Need disabled is also in " + strExportFileNeedDisabled + ".");
    return true;
}

public static void queryADforHRrow(object iRowNum)
{
    Name_mismatch = "";
    FTE_mismatch = "";
    Title_mismatch = "";
    Department_mismatch = "";
    Manager_mismatch = "";
    noteStrEmployeeID = "";
    noteStrSamAccountName = "";
    noteStr = "";
    lineStr = "";
    accountFound = false;
    DistinguishedNameFound = "";
    samAccountNameFound = "";
    samAccountNameFound2 = "";
    employeeIDfound = "";
    employeeidAD = "";
    LegalFirstNameAD = "";
    firstNameAD = "";
    lastNameAD = "";
    middleNameAD = "";
    statusAD = "";
    managerAD = "";
    samIDs.Clear();
    eeIDs.Clear();
    statusS.Clear();
    samID2 = "";
    samID2hasActive = false;
    set_defaultGroup(); // both not found and found users need checking  _FTE or _NonFTE DL group mismatch
    memberOf = "";
    AccountExpirationDate = "";
    // full checking will also check if eeID = "" 
    if (employeeID != "" && employeeID != "0")
    {
        // always check employeeID
        // findByEmployeeID(employeeID, "User")
        findByEmployeeIDlistMoreValue(employeeID); // quickly
        if (DistinguishedNameFound == "")
            noteStrEmployeeID = noteStrEmployeeID + "No AD user has EEID " + employeeID + ". ";
        else
        {
            if (samAccountNameFound2 != "")
            {
                noteStrSamAccountName = noteStrSamAccountName + employeeID + "= " + samAccountNameFound2 + ". "; // has duplicate AD accounts
                                                                                                                 // If Status = "Active" Then
                reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",duplicate AD accounts," + employeeID + "= " + samAccountNameFound2);
                addDiscrepancyCount(LocationAux1);
            }
            if (Strings.LCase(samAccountName) != Strings.LCase(samAccountNameFound))
            {
                if (samIDprovided)
                {
                    noteStrSamAccountName = noteStrSamAccountName + "(AD)" + samAccountNameFound + "<===>" + samAccountName + "(" + hrMIS + "). ";
                    if (needCompareAccountName)
                    {
                        if (Status != "Terminated")
                        {
                            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",account name mismatch," + "(AD)" + samAccountNameFound + "<===>" + samAccountName + "(" + hrMIS + ")");
                            addDiscrepancyCount(LocationAux1);
                        }
                    }
                }
                else
                    noteStrSamAccountName = noteStrSamAccountName + "not standard (" + samAccountNameFound + " : " + samAccountName + "). ";

                // since account already found by employeeID, disregard any name/email mismatch, don't need further judging
                samAccountName = samAccountNameFound;
            }
            accountFound = true;
            checkMismatch();
            // lineStr = (iRowNum + 1) & ","
            lineStr = wkdID + ",";
            writeAccountFound();
        }
    }

    // if not found, then should findByReqID
    if (!accountFound && (reqID != "" || employeeID != ""))
    {
        string usedID = employeeID;
        if (reqID != "")
            usedID = reqID;
        DistinguishedNameFound = "";
        if (!findByReqID_indexWay(usedID))
        {
            if (reqID != "")
                noteStrEmployeeID = noteStrEmployeeID + "No AD user has reqID " + reqID + ". ";
        }
        else
        {
            if (Strings.LCase(samAccountName) != Strings.LCase(samAccountNameFound))
            {
                if (samIDprovided)
                {
                    noteStrSamAccountName = noteStrSamAccountName + "(AD)" + samAccountNameFound + "<===>" + samAccountName + "(" + hrMIS + "). ";
                    if (needCompareAccountName)
                    {
                        if (Status != "Terminated")
                        {
                            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + " (reqID:" + usedID + "),\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",account name mismatch," + "(AD)" + samAccountNameFound + "<===>" + samAccountName + "(" + hrMIS + ")");
                            addDiscrepancyCount(LocationAux1);
                        }
                    }
                }
                else
                    noteStrSamAccountName = noteStrSamAccountName + "not standard (" + samAccountNameFound + " : " + samAccountName + "). ";
                // since account already found by reqID, disregard any name/email mismatch, don't need further judging
                samAccountName = samAccountNameFound;
            }
            accountFound = true;
            checkMismatch();


            if (employeeidAD != "" && employeeidAD == wkdID)
            {
            }
            else if (employeeidAD == "" || WDID.IndexOf(employeeidAD) < 0)
            {
                if (Name_mismatch == "")
                    // if AD name is the same as HR system name, we trust the AD account is located accurate, we can mark it in the foundUsers report, for later updating EEID
                    Name_mismatch = "\"name OK, but need set EEID for reqID:" + usedID + "\"";
                else if (Name_mismatch.Contains(minor))
                    Name_mismatch = Name_mismatch.Replace(minor, minor + " but need set EEID for reqID ");
                else
                {
                }
            }

            lineStr = wkdID + ",";
            writeAccountFound();
        }
    }

    // do search by email key
    if (!accountFound && Workemail != "")
    {
        string emailKey = Strings.LCase(Microsoft.VisualBasic.Left(Workemail, Strings.InStr(Workemail, "@")));
        // Find by email only if HR emailKey is unique
        if (HR_workEmailKey.IndexOf(emailKey) == HR_workEmailKey.LastIndexOf(emailKey) || (employeeID != "" && ignoreUpdateEEIDs.Contains("." + employeeID + ".")))
        {
            if (findByEmailListKeyMoreValue(emailKey))
            {
                samAccountName = samAccountNameFound;
                accountFound = true;
                checkMismatch();
                lineStr = wkdID + ",";
                writeAccountFound();
                if (employeeIDfound != employeeID && Conversion.Val(employeeIDfound) != Conversion.Val(employeeID))
                {
                    DateTime createDate = new DateTime();
                    try
                    {
                        createDate = whenCreated;
                    }
                    catch (Exception ex)
                    {
                    }
                    if (employeeIDfound == "")
                    {
                        noteStrEmployeeID += samAccountName + ": (AD EEID is blank)<===>(" + hrMIS + ")" + employeeID + ". ";
                        if (Status != "Terminated" && reqidAD == "")
                        {
                            // if already set reqID, if no name mismatch, the EEID will be updated automatically, so don't report this discrepany
                            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",AD EEID is blank,AD user \"" + samAccountName + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need set EEID");
                            addDiscrepancyCount(LocationAux1);
                        }
                    }
                    else if (employeeIDfound != reqID)
                    {
                        if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                        {
                            noteStrEmployeeID += samAccountName + ": " + employeeIDfound + "(AD)<===>(" + hrMIS + ")" + employeeID + ". ";
                            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + " (AD:" + employeeIDfound + "),\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",AD EEID incorrect,AD user \"" + samAccountName + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need correct EEID");
                            addDiscrepancyCount(LocationAux1);

                            EEIDincorrect += employeeIDfound + ".";
                        }
                    }
                }
            }
        }
        else
        {
        }
    }

    // do search by samAccountName
    string samIDforNewUser = ""; // default should set blank, if samAccountName not in use, then use it
    if (!accountFound && samAccountName != "")
    {
        DistinguishedNameFound = "";
        // findBySamAccountName(samAccountName, "User")
        findBySamIDlistMoreValue(samAccountName); // quickly
        if (DistinguishedNameFound == "")
        {
            noteStr = noteStr + "Invalid: " + samAccountName + " ? ";
            samIDforNewUser = samAccountName;
            samAccountName = "";
        }
        else if (samIDprovided || employeeID != "")
        {
            if (employeeIDfound == employeeID || (employeeIDfound == "" && samIDprovided))
            {
                // If employeeIDfound = employeeID Then 'No employeeID account will list in not sure file
                accountFound = true;
                checkMismatch();
                lineStr = wkdID + ",";
                writeAccountFound();
            }
            else if (employeeIDfound != "")
            {
                noteStrEmployeeID = noteStrEmployeeID + samAccountName + ": " + employeeIDfound + "(AD)<===>(" + hrMIS + ")" + employeeID + ". ";
                samAccountName = "";
            }
        }// If samIDprovided Or employeeID <> ""
    }

    // do search by name
    if (!accountFound)
    {
        DistinguishedNameFound = "";
        if (NeedAccount != noNeedAccountValueHR)
        {
            if (NickName == "")
                SearchByNameMoreValue(FirstName);
            else
                SearchByNameMoreValue(NickName);
        }
        if (DistinguishedNameFound != "")
        {
            // should avoid duplicated mismatch notes
            if (employeeID == "" && !samIDprovided)
                checkMismatch();
            if (employeeIDfound == "")
            {
                accountFound = true;
                lineStr = wkdID + ",";
                writeAccountFound();
            }
            else
            {
                // if AD already has employeID set, must be employeeIDfound <> employeeID here
                noteStrEmployeeID += samAccountName + " is " + employeeIDfound + ". ";
                samAccountName = "";
            }
        }
    }

    if (!accountFound)
    {
        if (employeeID == "")
        {
            if (samAccountName != "")
                noteStr = noteStr + "Not sure: " + samAccountName + " ? ";
            noteStr = noteStr + "No employeeID. No user account found in AD.";
        }
        set_ScriptPath_HomeDirectory(samIDforNewUser);
        samAccountName = "";
        if (employeeID != "" && NoNeedAD_EEIDs.IndexOf(employeeID) >= 0)
        {
            if (NeedAccount == "")
                NeedAccount = noNeedAccountValueHR;
        }
        else if (NeedAccount != noNeedAccountValueHR)
        {
            if (Strings.LCase(Status) == "active")
            {
            }
        }

        // should also find manage EEid and DN for creating new account use
        lineStr = wkdID + ",\"" + preID + "\",\"" + samAccountName + "\",,,,,,,,,,,,,,,,,";

        if (ManagerID == "" && Strings.InStr(ManagerWorkEmail, "@") > 0)
        {
            // chinese last name and first name may be the same, e.g. chen feng 200842, 200109, and AD only store nickname as firstname, so need compare email 
            // try find the manager employee ID by HR manager email. but HR email may not match AD email.
            // just use the key part before "@" to compare, may be @ehealthinsurance.com or @ehealth.com
            // index = WorkemailList.IndexOf(ManagerWorkEmail)
            index = WorkemailListKey.IndexOf(Microsoft.VisualBasic.Left(Strings.LCase(ManagerWorkEmail), Strings.InStr(ManagerWorkEmail, "@")));
            if (index >= 0)
                ManagerID = eeIDlist.Item(index);
            else
            {
                index = HR_workEmailKey.IndexOf(Microsoft.VisualBasic.Left(Strings.LCase(ManagerWorkEmail), Strings.InStr(ManagerWorkEmail, "@")));
                if (index >= 0)
                    ManagerID = HR_eeID.Item(index);
                noteStr = noteStr + " manager email: " + ManagerWorkEmail + "? ";
            }
            if (ManagerID != "")
            {
                while ((ManagerID.Length < 6))
                    ManagerID = "0" + ManagerID;
            }
        }
        if (ManagerID != "" && ManagerDN == "")
        {
            index = eeIDlist.IndexOf(ManagerID);
            if (index >= 0)
                ManagerDN = userDNlist.Item(index);
        }

        // don't compare AD not found
        lineStr += ManagerID + "," + "\"" + ManagerDN + "\"" + "," + ManagerWorkEmail + ",\"" + UserSegment + "\",";
        lineStr += Mobile + ",";
        lineStr = lineStr + "\"" + FirstName + "\",";
        lineStr = lineStr + "\"" + MiddleName + "\",";
        lineStr = lineStr + "\"" + LastName + "\",";
        lineStr = lineStr + "\"" + NickName + "\",";
        lineStr = lineStr + "\"" + Jobtitle + "\",";
        lineStr += "\"" + JobEffDate + "\",";
        lineStr += "\"" + BusinessUnit + "\",";
        lineStr = lineStr + "\"" + HomeDepartment + "\",";
        lineStr = lineStr + "\"" + Status + "\",";
        lineStr = lineStr + "\"" + statusEffDateStr + "\",";
        lineStr = lineStr + "\"" + EmpClass + "\",";
        lineStr += "\"" + OriginalHireDate + "\",";
        lineStr = lineStr + "\"" + HireDate + "\",";
        lineStr = lineStr + "\"" + Address + "\",";
        lineStr = lineStr + "\"" + City + "\",";
        lineStr = lineStr + "\"" + State + "\",";
        lineStr = lineStr + "\"" + PostalCode + "\",";
        lineStr = lineStr + "\"" + DivisionCode + "\",";
        lineStr = lineStr + "\"" + LocationAux1 + "\",";
        lineStr = lineStr + "\"" + PhysicalDeliveryOfficeName + "\",";
        lineStr = lineStr + "\"" + CostCenter + "\",";
        lineStr = lineStr + "\"" + NeedAccount + "\",";

        if (NeedAccount != noNeedAccountValueHR)
            intNotFoundCount += 1;

        generateNewHireEmail();
        // 'USA Remote   TPA: Remote    TPA: Hobart, IN
        // remote user don't need mailbox. other new hire should always show this email naming convention discrepancy
        if (needEmailNamingConventionMismatchReport && emailChanged && Strings.UCase(NeedAccount) != noNeedAccountValueHR)
        {
            if (!NoNeedDiscrepancyNewHireEmailLocationStr.Contains("." + LocationAux1 + "."))
            {
                // noteStr = noteStr & " " & hrMIS & " is " & Trim(HR_workEmail.Item(iRowNum)) & "?"
                if (Status != "Terminated")
                {
                    tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",email naming convention," + hrMIS + " is " + Workemail + ". Should be " + newHireEmail + "?";
                    // reportfileEmailMismatch.WriteLine(tmpStr)
                    // addEmailMismatchCount(LocationAux1)
                    reportfileDiscrepancy.WriteLine(tmpStr);
                    addDiscrepancyCount(LocationAux1);
                }
            }
        }

        lineStr = lineStr + ",,,,," + noteStrEmployeeID + "," + noteStrSamAccountName;
        // not AD account found, only report discrepancy if need account
        // If HR_needAccount(iRowNum) <> noNeedAccountValueHR AndAlso noteStrEmployeeID <> "No AD user has EEID " & employeeID & ". " Then
        // reportfileDiscrepancy.WriteLine(wkdID & "," & preID & ",""" & FirstName & """,""" & LastName & """,""" & NickName & """,""" & LocationAux1 & """," & Status & ",EEID mismatch," & noteStrEmployeeID & " " & noteStrSamAccountName)
        // addDiscrepancyCount(LocationAux1)
        // End If
        try
        {
            reportfileNotFound.WriteLine(lineStr + "," + defaultGroup + ",," + noteStr + ",\"" + strOUnet + strDomain + "\"," + ProfilePath + "," + ScriptPath + "," + HomeDirectory + "," + HomeDrive + "," + MOC_enabled + "," + newHireEmail);
            reportfileEntire.WriteLine(lineStr + "," + defaultGroup + ",," + noteStr + ","); // has needDisabled
        }
        catch (Exception ex)
        {
            hasError = true;
            hasWriteEntireError = true;
            swLog.WriteLine("Write report file failed, please check if there is invalid text character: " + privateants.vbCrLf + wkdID + " " + samAccountName);
        }
    }
}

private static void generateNewHireEmail()
{
    // SOP_eHealth Naming Standards_080317_V0.02.docx: default email domain is all "@ehealth.com" now
    // set default email domain
    string emailDomain = "@ehealth.com"; // "@ehealthinsurance.com"

    string neatFirstName = FirstName.Replace(" ", "");
    if (NickName != "")
        neatFirstName = NickName.Replace(" ", "");
    // From: Charles Escoto  Sent: Thursday, July 30, 2015 9:27 AM: let’s keep the apostrophe out of the main email address to keep the email creation as simple as possible. The IT team can just include the apostrophe in the alias email address.
    neatFirstName = neatFirstName.Replace("'", "");
    neatFirstName = neatFirstName.Replace("-", "");
    string neatLastName = LastName.Replace(" ", "");
    neatLastName = neatLastName.Replace("'", "");
    neatLastName = neatLastName.Replace("-", "");

    string emailKey = neatFirstName + "." + neatLastName;
    // for existing AD users, must use HR system emailKey for checking mismatch
    // but for new hires, need create a standard new email address
    // need check workemail first, if already in use, try to create a new one:
    // steve.smith@ehealth.com (primary choice)
    // s.smith@ ehealth.com
    // stevie.smith@ehealth.com (user choice)
    // If WorkemailListKey.IndexOf(LCase(emailKey) & "@") > 0 Then
    // If emailKey <> neatFirstName & "." & neatLastName Then
    // emailKey = neatFirstName & "." & neatLastName
    // End If
    // End If
    if (WorkemailListKey.IndexOf(Strings.LCase(emailKey) + "@") > 0)
        emailKey = Microsoft.VisualBasic.Left(neatFirstName, 1) + "." + neatLastName;

    int tryTime = 1;
    while (WorkemailListKey.IndexOf(Strings.LCase(emailKey) + "@") > 0)
    {
        // the workemail already in use
        tryTime += 1;
        emailKey = neatFirstName + tryTime + "." + neatLastName;
    }
    newHireEmail = emailKey + emailDomain;

    emailChanged = false;
    // ignore TPA emails
    if (Strings.LCase(Workemail).Contains("@ehealth") && Strings.LCase(newHireEmail) != Strings.LCase(Workemail))
        emailChanged = true;
}

private static void set_defaultGroup()
{
    MOC_enabled = "";
    defaultGroup = ""; // "PGP Users;"
    FTEgroup = "RemoteUsers_FTE";
    nonFTEgroup = "RemoteUsers_NonFTE";
    if (LocationAux1 == "")
        return;
    // USA Remote   TPA: Remote    TPA: Hobart, IN
    if (Strings.UCase(LocationAux1) == "RMT" || Strings.InStr(Strings.UCase(LocationAux1), "REMOTE") > 0)
    {
        FTEgroup = "RemoteUsers_FTE";
        nonFTEgroup = "RemoteUsers_NonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "SCS" || Strings.UCase(LocationAux1).StartsWith("SANTA CLARA"))
    {
        FTEgroup = "SantaClara_FTE";
        nonFTEgroup = "SantaClara_nonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "SF" || Strings.InStr(Strings.UCase(LocationAux1), "SAN FRANCISCO") > 0)
    {
        FTEgroup = "SanFrancisco_FTE";
        nonFTEgroup = "SanFrancisco_NonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "GDRV" || Strings.InStr(Strings.UCase(LocationAux1), "GOLD RIVER") > 0)
    {
        FTEgroup = "GoldRiver_FTE";
        nonFTEgroup = "GoldRiver_NonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "UT" || Strings.InStr(Strings.UCase(LocationAux1), "SALT LAKE CITY") > 0)
    {
        FTEgroup = "SaltLake_FTE";
        nonFTEgroup = "SaltLake_NonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "XM" || Strings.InStr(Strings.UCase(LocationAux1), "XIAMEN") > 0)
    {
        FTEgroup = "Xiamen_FTE";
        nonFTEgroup = "Xiamen_NonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "DC" || Strings.InStr(Strings.UCase(LocationAux1), "WASHINGTON") > 0)
    {
        // Charles email 11/13/2014
        FTEgroup = "RemoteUsers_FTE";
        nonFTEgroup = "RemoteUsers_NonFTE";
    }
    else if (Strings.UCase(LocationAux1).Contains("AUSTIN"))
    {
        FTEgroup = "Austin_FTE";
        nonFTEgroup = "Austin_NonFTE";
    }
    else if (Strings.UCase(LocationAux1) == "IN" || Strings.InStr(Strings.UCase(LocationAux1), "INDIANAPOLIS") > 0)
    {
        FTEgroup = "Indianapolis_FTE";
        nonFTEgroup = "Indianapolis_NonFTE";
    }
    // if use old format HR connection file, there is no EmpClass column, so if EmpClass is blank, should not update to nonFTEgroup wrongly
    if (EmpClass != "")
    {
        // So I can just see if "Regular" included in Worker Category Description to distinguish FTE or NonFTE employee
        // Huan 2016/09/20 15:14 
        // yes, i just asked Valerie for clarification
        // don't really need to distinguish between FT and PT
        // if "Regular" then = FTE
        // and non-FTE is all others 
        if (Strings.LCase(EmpClass) == "non-employee")
        {
            defaultGroup = "";
            FTEgroup = "";
        }
        else
            // Huan email sent on 10/17/2020: Yes…Seasonal workers are classified under the “Employee” category
            if (EmpClass.Contains("Regular"))
            defaultGroup += FTEgroup + ";";
        else
            defaultGroup += nonFTEgroup + ";";
    }
}

public static void checkMismatch()
{
    // Even if first name and last name match, still need report " has duplicated account name" if any
    if (samIDprovided && Status != "Terminated")
    {
        if (accountFound)
        {
            if (samAccountNameFound != "" && Strings.LCase(samAccountNameFound) != HR_AccountName)
            {
                string otherEEIDname = "";
                int index = samIDlist.IndexOf(HR_AccountName);
                if (index >= 0)
                    otherEEIDname = " (" + HR_AccountName + " is for " + eeIDlist.Item(index) + " " + firstNameList.Item(index) + " " + lastNameList.Item(index) + ")";
                if (needCompareAccountName)
                {
                    if (!noteStrSamAccountName.Contains("<===>"))
                    {
                        if (samAccountNameFound.Contains(".") || samAccountNameFound.Contains(" "))
                            tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + "," + "AD account name nonstandard," + samAccountNameFound + " need correct to " + HR_AccountName + otherEEIDname + "?";
                        else
                            tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",account name mismatch," + "(AD)" + samAccountNameFound + "<===>" + HR_AccountName + "(" + hrMIS + ")" + otherEEIDname;
                        reportfileDiscrepancy.WriteLine(tmpStr);
                        addDiscrepancyCount(LocationAux1);
                    }
                }
            }
        }
        else if (employeeIDfound != "" && employeeIDfound != employeeID)
        {
            // check really has duplicated account name in HR system
            int index1, index2;
            index1 = HR_samAccountName.IndexOf(samAccountName);
            index2 = HR_samAccountName.LastIndexOf(samAccountName);
            if (index1 != index2)
            {
                tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + "," + hrMIS + " has duplicated account name," + samAccountName + " is for " + employeeIDfound + " " + firstNameAD + " " + lastNameAD;
                if (Status != "Terminated")
                {
                    // TODO: if rehire input as new hire in WFN, we should not suggest new account name, but should figure out the existing accout
                    if (newSAMID(FirstName, LastName, isXM) == "")
                        tmpStr += " (a rehire user " + objUser.samAccountName + "?)";
                    else
                        tmpStr += " (suggested new account name: " + newSAMID(FirstName, LastName, isXM) + ")";
                }
                reportfileDiscrepancy.WriteLine(tmpStr);
                addDiscrepancyCount(LocationAux1);
            }
        }
    }

    // Chinese english name will stored in NickName in HR system
    // if name contains "-" or " ", AD may use part of the firstname or lastname
    // If (LCase(firstNameAD) <> LCase(NickName) And LCase(firstNameAD) <> LCase(FirstName)) Or LCase(lastNameAD) <> LCase(LastName) Or (middleNameAD <> "" And LCase(middleNameAD) <> LCase(MiddleName)) Then
    // ignore middle name compare
    if (NickName == "")
        tmpStr = Strings.LCase(FirstName);
    else
        tmpStr = Strings.LCase(NickName);
    if ((Strings.LCase(firstNameAD) != tmpStr) || (Strings.LCase(lastNameAD) != Strings.LCase(LastName) && !hasPreferredLastnameEEIDs.Contains("." + employeeID + ".")))
    {
        string showMinor = checkMinorMismatch(firstNameAD, lastNameAD, FirstName, LastName, NickName);
        if (NickName != "")
            Name_mismatch = Name_mismatch + "\"" + showMinor + "(AD)" + firstNameAD + "," + middleNameAD + " ," + lastNameAD + " <===> " + FirstName + "(" + NickName + ")," + MiddleName + " ," + LastName + "(" + hrMIS + ").\"";
        else
            Name_mismatch = Name_mismatch + "\"" + showMinor + "(AD)" + firstNameAD + "," + middleNameAD + " ," + lastNameAD + " <===> " + FirstName + "," + MiddleName + " ," + LastName + "(" + hrMIS + ").\"";
        if (Status != "Terminated")
        {
            // if just hyphen mismatch, don't report it
            if ((showMinor != minor) || (Strings.LCase(lastNameAD.Replace("-", " ")) != Strings.LCase(LastName.Replace("-", " ")) && !hasPreferredLastnameEEIDs.Contains("." + employeeID + ".")))
            {
                // CTPR-9 Ken F. approved the fix:
                // WFN may be input samAccountName incorrectly, check again by workemail
                tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",AD user \"" + samAccountName + "\" fullname mismatch," + Name_mismatch.Replace(major, "");
                reportfileDiscrepancy.WriteLine(tmpStr);
                addDiscrepancyCount(LocationAux1);
            }
        }
    }

    // If City <> "" Then
    // If City <> cityAD Then
    // FTE_mismatch = FTE_mismatch & """" & "(AD)" & cityAD & " <===> " & City & "(" & hrMIS & ")." & """"
    // End If
    // End If

    // don't compare terminated
    if (Status != "Terminated")
    {
        if (Jobtitle != "")
        {
            if (Jobtitle != jobtitleAD)
            {
                Title_mismatch = Title_mismatch + "\"" + "(AD)" + jobtitleAD + " <===> " + Jobtitle + "(" + hrMIS + ")." + "\"";
                // after Workday + Okta integration, AD sync no longer update title, so need report title mismatch
                reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",title mismatch," + Title_mismatch);
                addDiscrepancyCount(LocationAux1);
            }
        }
        if (HomeDepartment != "")
        {
            if (HomeDepartment != homeDepartmentAD)
            {
                Department_mismatch = Department_mismatch + "\"" + "(AD)" + homeDepartmentAD + " <===> " + HomeDepartment + "(" + hrMIS + ")." + "\"";
                // after Workday + Okta integration, AD sync no longer update department, so need report department mismatch
                reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",department mismatch," + Department_mismatch);
                addDiscrepancyCount(LocationAux1);
            }
        }

        string keyWorkemailAD = "";
        string keyWorkemail = "";
        if (Strings.InStr(WorkemailAD, "@") > 0)
            keyWorkemailAD = Microsoft.VisualBasic.Left(WorkemailAD, Strings.InStr(WorkemailAD, "@"));
        if (Strings.InStr(Workemail, "@") > 0)
            keyWorkemail = Microsoft.VisualBasic.Left(Workemail, Strings.InStr(Workemail, "@"));

        if (Strings.LCase(keyWorkemailAD) != Strings.LCase(keyWorkemail))
        {
            noteStr = noteStr + "email: (AD)" + WorkemailAD + " <===> " + Workemail + "(" + hrMIS + ").";
            if (accountFound)
            {
                if ((NickName != "" && Strings.LCase(firstNameAD) != Strings.LCase(NickName)) && Strings.LCase(lastNameAD) == Strings.LCase(LastName))
                {
                    // if HR system already record the same email as AD, ignore reporting this discrepancy: AD not use preferred name as firstname
                    tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",AD not use preferred name as firstname,AD firstname should use '" + NickName + "'";
                    reportfileDiscrepancy.WriteLine(tmpStr);
                    addDiscrepancyCount(LocationAux1);
                }
                if (needEmailMismatchReport && WorkemailAD != "")
                {
                    // email mismatch here only compare key part, ignore domain part. So should always show this discrepancy
                    // Penny Bayer's email Sent: 28 October, 2016 02:29 agree this improvement: 
                    // If the email address in WFN already in AD email address alias list, we should not report the email mismatch, even the primary email address is not the same. 
                    // Harpal Harika email Sent: Saturday, August 25, 2018 2:52 AM: 
                    // looks like there are multiple emails generated for the same user accounts. Is it possible to consolidate all actions take for the users accounts in 1 email.
                    if (Strings.LCase(proxyAddressesAD).Contains(Strings.LCase(Workemail)))
                    {
                        if (needPrimaryEmailMismatchReport)
                        {
                            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",email primary,(AD)" + WorkemailAD + " <===> " + Workemail + "(" + hrMIS + ")");
                            addDiscrepancyCount(LocationAux1);
                        }
                    }
                    else
                    {
                        // reportfileEmailMismatch.WriteLine(employeeID & ",""" & FirstName & """,""" & LastName & """,""" & NickName & """,""" & LocationAux1 & """," & Status & ",email mismatch,(AD)" & WorkemailAD & " <===> " & Workemail & "(" & hrMIS & ")")
                        // addEmailMismatchCount(LocationAux1)
                        reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",email mismatch,(AD)" + WorkemailAD + " <===> " + Workemail + "(" + hrMIS + ")");
                        addDiscrepancyCount(LocationAux1);
                    }
                }
            }
        }
        else if (needEmailMismatchReport)
        {
            if (needDiscrepancyEmailDomainLocationStr.Contains("." + LocationAux1 + "."))
            {
                if (WorkemailAD != "" && Strings.LCase(WorkemailAD) != Strings.LCase(Workemail))
                {
                    // Penny Bayer's email Sent: 28 October, 2016 02:29 agree this improvement: 
                    // If the email address in WFN already in AD email address alias list, we should not report the email mismatch, even the primary email address is not the same. 
                    // Harpal Harika email Sent: Saturday, August 25, 2018 2:52 AM: 
                    // looks like there are multiple emails generated for the same user accounts. Is it possible to consolidate all actions take for the users accounts in 1 email.

                    // Vishal email Sent: Tuesday, June 11, 2019 3:17 AM
                    // “Primary email address must match between AD and WFN.”
                    if (!Strings.LCase(proxyAddressesAD).Contains(Strings.LCase(Workemail)))
                    {
                        // reportfileEmailMismatch.WriteLine(employeeID & ",""" & FirstName & """,""" & LastName & """,""" & NickName & """,""" & LocationAux1 & """," & Status & ",email domain,(AD)" & WorkemailAD & " <===> " & Workemail & "(" & hrMIS & ")")
                        // addEmailMismatchCount(LocationAux1)
                        reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",email domain,(AD)" + WorkemailAD + " <===> " + Workemail + "(" + hrMIS + ")");
                        addDiscrepancyCount(LocationAux1);
                    }
                    else if (needPrimaryEmailMismatchReport)
                    {
                        // primary email address
                        reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",email primary,(AD)" + WorkemailAD + " <===> " + Workemail + "(" + hrMIS + ")");
                        addDiscrepancyCount(LocationAux1);
                    }
                }
            }
        }
    }
}

public static void writeAccountFound()
{
    string FTEnonFTEname;
    GetUserMemberOf(samAccountName);
    if (Strings.LCase(Status) != "terminated" && Strings.LCase(statusAD) != "terminated")
    {
        // correct _FTE / _NonFTE group as defaultGroup, also add into sdl-okta-jobvite-seasonal for new hire (RFS-160604)
        // If employeeID <> "101297" Then 'CTPR-85 Do not update email distribution list for Holly Makris
        if (SeasonalGroupName != "")
        {
            if (EmpClass.StartsWith("Seasonal"))
            {
                if (!Strings.UCase(memberOf).Contains(Strings.UCase(SeasonalGroupName)))
                {
                    FTE_mismatch += "add " + SeasonalGroupName + ". ";
                    if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                        reportfileGroupAction.WriteLine("\"" + getGroupDN(SeasonalGroupName) + "\",add,\"" + DistinguishedNameFound + "\"");
                }
            }
        }
        if (defaultGroup.Contains(FTEgroup))
        {
            // need remove nonFTEgroup, also need remove other location group. so loop all FTEnonFTEgroup except want one
            for (int t = 0; t <= FTEnonFTEgroupNamesList.Count - 1; t++)
            {
                FTEnonFTEname = FTEnonFTEgroupNamesList.Item(t);
                if (Strings.UCase(FTEnonFTEname) != Strings.UCase(FTEgroup))
                {
                    if (Strings.UCase(memberOf).Contains(Strings.UCase(FTEnonFTEname)))
                    {
                        FTE_mismatch += "remove " + FTEnonFTEname + ". ";
                        if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                            reportfileGroupAction.WriteLine("\"" + getGroupDN(FTEnonFTEname) + "\",remove,\"" + DistinguishedNameFound + "\"");
                    }
                }
            }
            if (!Strings.UCase(memberOf).Contains(Strings.UCase(FTEgroup)))
            {
                // need add FTEgroup
                FTE_mismatch += "add " + FTEgroup + ". ";
                if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                    reportfileGroupAction.WriteLine("\"" + getGroupDN(FTEgroup) + "\",add,\"" + DistinguishedNameFound + "\"");
            }
        }
        if (defaultGroup.Contains(nonFTEgroup))
        {
            for (int t = 0; t <= FTEnonFTEgroupNamesList.Count - 1; t++)
            {
                FTEnonFTEname = FTEnonFTEgroupNamesList.Item(t);
                if (Strings.UCase(FTEnonFTEname) != Strings.UCase(nonFTEgroup))
                {
                    if (Strings.UCase(memberOf).Contains(Strings.UCase(FTEnonFTEname)))
                    {
                        FTE_mismatch += "remove " + FTEnonFTEname + ". ";
                        if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                            reportfileGroupAction.WriteLine("\"" + getGroupDN(FTEnonFTEname) + "\",remove,\"" + DistinguishedNameFound + "\"");
                    }
                }
            }
            if (!Strings.UCase(memberOf).Contains(Strings.UCase(nonFTEgroup)))
            {
                // need add nonFTEgroup
                FTE_mismatch += "add " + nonFTEgroup + ". ";
                if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                    reportfileGroupAction.WriteLine("\"" + getGroupDN(nonFTEgroup) + "\",add,\"" + DistinguishedNameFound + "\"");
            }
        }
    }
    else if (Strings.LCase(Status) == "terminated" && isNeedDisableTerminateRightNow()
                && DB_DisabledSamIDs.IndexOf(samAccountName) < 0)
    {
        // only remove _FTE / _NonFTE group for HR terminated, ignore HR On Leave
        // Charles email 11/04/2014: is there a way to update DL’s through a script if there is a change to a location in HR system?   
        // Also, can a terminated employees report be run containing any distribution lists they are currently on?

        // should be loop memberOf, remove all groups for terminated?
        // Yes. CTEWS-2977 remove users from all Groups and clear organization reporting from AD upon termination
        // but, CTEWS-3367 update WFN AD sync tool to remove groups, manager, and hide from GAL 3 days after account disabling
        if (DaysCanRemoveManager > 0 && Strings.LCase(statusAD) == "terminated")
        {
            if (memberOf != "")
            {
                FTE_mismatch += "remove ";
                for (int i = 0; i <= memberOfList.Count - 1; i++)
                {
                    string gCN = memberOfList(i);
                    FTE_mismatch += gCN.Substring(0, gCN.IndexOf(","));
                    if (i < memberOfList.Count - 1)
                        FTE_mismatch += "; ";
                    else
                        FTE_mismatch += ".";
                    if (!ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                        reportfileGroupAction.WriteLine("\"" + memberOfList(i) + "\",remove,\"" + DistinguishedNameFound + "\"");
                }
            }
        }
    }
    // If FTE_mismatch <> "" Then
    // reportfileDiscrepancy.WriteLine(wkdID & "," & preID & ",""" & FirstName & """,""" & LastName & """,""" & NickName & """,""" & LocationAux1 & """," & Status & ",email distribution update, " & FTE_mismatch)
    // addDiscrepancyCount(LocationAux1)
    // End If
    // set_ScriptPath_HomeDirectory(samAccountName)

    lineStr = lineStr + "\"" + preID + "\","; // always display HR system employeeID
    lineStr = lineStr + "\"" + samAccountName + "\","; // samAccountNameFound
    lineStr = lineStr + "\"" + DistinguishedNameFound + "\",";
    lineStr += "\"" + LegalFirstNameAD + "\",";
    lineStr = lineStr + "\"" + firstNameAD + "\",";
    lineStr = lineStr + "\"" + lastNameAD + "\",";
    lineStr = lineStr + "\"" + middleNameAD + "\",";
    lineStr = lineStr + statusAD + ",";
    lineStr = lineStr + "\"" + jobtitleAD + "\",";
    lineStr = lineStr + "\"" + homeDepartmentAD + "\",";
    lineStr = lineStr + "\"" + WorkemailAD + "\",";
    lineStr = lineStr + "\"" + whenCreated + "\",";
    lineStr = lineStr + "\"" + whenChanged + "\",";
    lineStr += "\"" + LastLogonTimeStamp + "\",";

    lineStr = lineStr + "\"" + employeeTypeAD + "\",";
    lineStr = lineStr + "\"" + cityAD + "\",";
    lineStr = lineStr + "\"" + managerAD + "\",";

    string ManagerWorkEmailAD = "";
    if (Strings.Trim(managerAD) != "")
    {
        index = userDNlist.IndexOf(managerAD);
        if (index < 0)
            lineStr = lineStr + "\"Not_Found\",";
        else
        {
            lineStr = lineStr + "\"" + eeIDlist.Item(index) + "\",";
            ManagerWorkEmailAD = WorkemailList(index);
        }
    }
    else
        lineStr = lineStr + ",";// ",,,,"

    if (Strings.InStr(ManagerWorkEmail, "@") > 0 && ManagerID == "")
    {
        // chinese last name and first name may be the same, e.g. chen feng 200842, 200109, and AD only store nickname as firstname, so need compare email 
        // try find the manager employee ID by HR manager email. but HR email may not match AD email.
        // just use the key part before "@" to compare, may be @ehealthinsurance.com or @ehealth.com
        // index = WorkemailList.IndexOf(ManagerWorkEmail)
        index = WorkemailListKey.IndexOf(Microsoft.VisualBasic.Left(Strings.LCase(ManagerWorkEmail), Strings.InStr(ManagerWorkEmail, "@")));
        if (index >= 0)
            ManagerID = eeIDlist.Item(index);
        else
        {
            index = HR_workEmailKey.IndexOf(Microsoft.VisualBasic.Left(Strings.LCase(ManagerWorkEmail), Strings.InStr(ManagerWorkEmail, "@")));
            if (index >= 0)
                ManagerID = HR_eeID.Item(index);
            noteStr = noteStr + " manager email: " + ManagerWorkEmail + "? ";
        }
    }
    if (ManagerID != "" && ManagerDN == "")
    {
        index = eeIDlist.IndexOf(ManagerID);
        int index2 = eeIDlist.LastIndexOf(ManagerID);
        // if not found, or manager has 2+ accounts, should found again by workEmail
        if (index < 0 || index != index2)
        {
            if (ManagerWorkEmail != "")
                index = WorkemailListKey.IndexOf(Microsoft.VisualBasic.Left(Strings.LCase(ManagerWorkEmail), Strings.InStr(ManagerWorkEmail, "@")));
        }

        if (index < 0)
        {
            int mngIndex = HR_eeID.IndexOf(ManagerID);
            if (mngIndex < 0 || LCase(HR_FTE(mngIndex)) != "non-employee")
            {
                if (Status != "Terminated")
                {
                    reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",AD managerEEID not found," + Microsoft.VisualBasic.Left(Strings.LCase(ManagerWorkEmail), Strings.InStr(ManagerWorkEmail, "@")) + " need set EEID: " + ManagerID);
                    addDiscrepancyCount(LocationAux1);
                }
            }
        }
        else if (Status != "Terminated")
        {
            ManagerDN = userDNlist.Item(index);
            if (ManagerDN == DistinguishedNameFound && !ManagerDN.Contains(CEO))
            {
                Manager_mismatch += "\"" + hrMIS + " manager is set to the employee self" + "\"";
                reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",manager set to self," + Manager_mismatch);
                addDiscrepancyCount(LocationAux1);
            }
        }
    }

    if (Status == "Terminated")
    {
        if (managerAD != "")
        {
            string hireDateStr = "";
            try
            {
                object oUser = Interaction.GetObject("LDAP://" + DistinguishedNameFound);
                hireDateStr = oUser.extensionAttribute9;
            }
            catch (Exception ex)
            {
            }
            if (hireDateStr != "")
            {
                DateTime hireDate = "2020-08-01";
                try
                {
                    hireDate = Strings.Trim(hireDateStr);
                }
                catch (Exception ex)
                {
                }

                if (hireDate < DateTime.Today)
                    Manager_mismatch += "need REMOVE manager.";
            }
        }
    }
    else
        // don't compare terminated
        if (ManagerDN != "" && ManagerDN != managerAD)
    {
        string managerADcn = "";
        if (managerAD != "")
        {
            managerADcn = Microsoft.VisualBasic.Left(managerAD, Strings.InStr(managerAD, ",") - 1);
            managerADcn = managerADcn.Replace("CN=", "");
        }
        tmpStr = "";
        tmpStr = Microsoft.VisualBasic.Left(ManagerDN, Strings.InStr(ManagerDN, ",") - 1);
        tmpStr = tmpStr.Replace("CN=", "");
        Manager_mismatch += "\"" + "(AD)" + managerADcn + " <===> " + tmpStr;

        bool isNonEmployee = false;
        if (HR_eeID.IndexOf(ManagerID) < 0)
        {
        }
        else if (LCase(HR_FTE(HR_eeID.IndexOf(ManagerID))) == "non-employee")
            isNonEmployee = true;
        if (isNonEmployee)
            Manager_mismatch += "(" + hrMIS + " Non-Employee)." + "\"";
        else
        {
            Manager_mismatch += "(" + hrMIS + ")." + "\"";
            // manager mismatch is much important, always report this discrepancy if manager is normal employee
            // but normally this mismatch will be updated soon in the sync running, so only need report it if has name mismatch
            // after Workday + Okta integration, AD sync no longer update manager, so need report manager mismatch
            // Dim showMinor As String = checkMinorMismatch(firstNameAD, lastNameAD, FirstName, LastName, NickName)
            // If (LCase(firstNameAD.Replace("-", " ")) <> LCase(NickName.Replace("-", " ")) AndAlso LCase(firstNameAD.Replace("-", " ")) <> LCase(FirstName.Replace("-", " "))) _
            // If (showMinor <> minor) OrElse LCase(lastNameAD.Replace("-", " ")) <> LCase(LastName.Replace("-", " ")) Then
            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",manager mismatch," + Manager_mismatch);
            addDiscrepancyCount(LocationAux1);
        }
    }

    // lineStr = lineStr & "," ' getNewManagerEmployeeID & ","
    // HR system data will be listed as new data in report file:
    lineStr += ManagerID + "," + "\"" + ManagerDN + "\"" + "," + ManagerWorkEmailAD + ",\"" + UserSegment + "\",";

    lineStr += Mobile + ",";
    lineStr = lineStr + "\"" + FirstName + "\",";
    lineStr = lineStr + "\"" + MiddleName + "\",";
    lineStr = lineStr + "\"" + LastName + "\",";
    lineStr = lineStr + "\"" + NickName + "\",";
    lineStr = lineStr + "\"" + Jobtitle + "\",";
    lineStr += "\"" + JobEffDate + "\",";
    lineStr += "\"" + BusinessUnit + "\",";
    lineStr = lineStr + "\"" + HomeDepartment + "\",";
    lineStr = lineStr + "\"" + Status + "\",";
    lineStr = lineStr + "\"" + statusEffDateStr + "\",";
    lineStr = lineStr + "\"" + EmpClass + "\",";
    lineStr += "\"" + OriginalHireDate + "\",";
    lineStr = lineStr + "\"" + HireDate + "\",";
    lineStr = lineStr + "\"" + Address + "\",";
    lineStr = lineStr + "\"" + City + "\",";
    lineStr = lineStr + "\"" + State + "\",";
    lineStr = lineStr + "\"" + PostalCode + "\",";
    lineStr = lineStr + "\"" + DivisionCode + "\",";
    lineStr = lineStr + "\"" + LocationAux1 + "\",";
    lineStr = lineStr + "\"" + PhysicalDeliveryOfficeName + "\",";
    lineStr = lineStr + "\"" + CostCenter + "\",";
    lineStr = lineStr + "\"" + NeedAccount + "\",";

    noteStr = noteStr + ",";

    if (Status.Contains("Leave"))
        onLeaveSamIDs += "'" + samAccountName + "',";
    // if HR system is Active but No need account, should not set AD disabled, report discrepancy instead
    if (((Strings.LCase(Status) != "active" || Strings.UCase(NeedAccount) == noNeedAccountValueHR) && DistinguishedNameFound != "" && (Strings.LCase(statusAD) != "terminated" || samID2hasActive)) || (needDisableInactive && needSetDisableForInactive) || AccountExpirationDate != "")
    {
        bool needDisableOnLeave = false;
        if (Status.Contains("Leave"))
        {
            if (!ignoreOnLeaveEEIDs.Contains("." + employeeID + "."))
                // here only check if AD status is Active
                // needDisableOnLeave = isNeedDisableOnLeave(samAccountName) 'Active, Terminated need delete OnLeaveTable record
                needDisableOnLeave = isNeedDisableOnLeaveByEffDate();
        }
        DateTime theHireDate = new DateTime();
        theHireDate = HireDate;
        bool isRehire = false;
        if (OriginalHireDate != HireDate)
            isRehire = true;
        if (!ignoreRehireSamIDs.Contains("." + samAccountName + "."))
        {
            // if is rehire and before hire date( allow 3 days lag), don't set AD disabled for inactive over 90 days
            if (((!Status.Contains("Leave")) && (!isRehire || (isRehire && DateTime.Now > theHireDate.AddDays(3)))) || (needDisableOnLeave) || Strings.UCase(NeedAccount) == noNeedAccountValueHR)
            {
                if (Strings.LCase(Status) != "terminated" || (Strings.LCase(Status) == "terminated" && isNeedDisableTerminateRightNow()))
                {
                    bool disableOrExpire = false;
                    if (Strings.LCase(Status) == "terminated")
                    {
                        if (Strings.LCase(statusAD) != "terminated")
                        {
                            if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                                noteStr += "\"ignore disable\"";
                            else
                            {
                                string hireDateStr = "";
                                try
                                {
                                    object oUser = Interaction.GetObject("LDAP://" + DistinguishedNameFound);
                                    hireDateStr = oUser.extensionAttribute9;
                                }
                                catch (Exception ex)
                                {
                                }
                                if (hireDateStr != "")
                                {
                                    DateTime hireDate = "2020-08-01";
                                    try
                                    {
                                        hireDate = Strings.Trim(hireDateStr);
                                    }
                                    catch (Exception ex)
                                    {
                                    }

                                    if (hireDate < DateTime.Today)
                                    {
                                        noteStr += "\"AD account need set disabled!\"";
                                        disableOrExpire = true;
                                    }
                                    else
                                        noteStr += "\"future hire\"";
                                }
                            }
                        }
                    }
                    else
                        // if already set expired, should not set again and again
                        if (AccountExpirationDate == "")
                    {
                        if (needDisableOnLeave)
                        {
                            noteStr += "\"AD account need set expired for Leave!\"";
                            disableOrExpire = true;
                        }
                        else if (needDisableInactive)
                        {
                            if (hasLeaveReturningRecord(samAccountName))
                                noteStr += "\"leave returning on " + leaveReturningEffDateStr + "\"";
                            else
                            {
                                string hireDateStr = "";
                                try
                                {
                                    object oUser = Interaction.GetObject("LDAP://" + DistinguishedNameFound);
                                    hireDateStr = oUser.extensionAttribute9;
                                }
                                catch (Exception ex)
                                {
                                }
                                if (hireDateStr != "")
                                {
                                    DateTime hireDate = "2020-08-01";
                                    try
                                    {
                                        hireDate = Strings.Trim(hireDateStr);
                                    }
                                    catch (Exception ex)
                                    {
                                    }

                                    int inactiveDays;
                                    inactiveDays = DateTime.DateDiff("d", hireDate, DateTime.Today);
                                    if (inactiveDays > daysInactiveLimit)
                                    {
                                        noteStr += "\"AD account need set expired for inactive over " + daysInactiveLimit + " days!\"";
                                        disableOrExpire = true;
                                    }
                                    else
                                        noteStr += "\"change hire date?\"";
                                }
                            }
                        }
                    }
                    else if (AccountExpirationDate.Contains("/209"))
                        noteStr += "\"need set Never expires\"";
                    else if (needDisableOnLeave)
                    {
                        noteStr += "\"already expired for Leave\"";
                        alreadyExpiredForLeave += employeeID + " " + samAccountName + "; ";
                    }
                    else
                    {
                        int leaveReturnIndex = LeaveReturnEEID.IndexOf(employeeID);
                        if (leaveReturnIndex < 0)
                        {
                            noteStr += "\"already expired for inactive\"";
                            alreadyExpiredForInactivate += employeeID + " " + samAccountName + "; ";
                        }
                        else
                            noteStr += "\"Need remove ExpirationDate for Leave return\"";
                    }
                    intNeedDisabledCount = intNeedDisabledCount + 1;
                    tmpStr = lineStr + Name_mismatch + ",\"" + FTE_mismatch + "\"," + Title_mismatch + "," + Department_mismatch + "," + Manager_mismatch + "," + noteStrEmployeeID + "," + noteStrSamAccountName + ",";
                    tmpStr += ",\"" + memberOf + "\"," + noteStr;
                    if (Strings.UCase(NeedAccount) == noNeedAccountValueHR)
                    {
                        tmpStr = wkdID + "," + preID + ",\"" + FirstName + "\",\"" + LastName + "\",\"" + NickName + "\",\"" + LocationAux1 + "\"," + Status + ",AD user \"" + samAccountName + "\" is enabled,but " + hrMIS + " Needs AD Account is No";
                        if (LCase(HR_FTE(HR_eeID.IndexOf(employeeID))) == "non-employee")
                            tmpStr += " (Non-employee)";
                        reportfileDiscrepancy.WriteLine(tmpStr);
                        addDiscrepancyCount(LocationAux1);
                    }
                    else if (disableOrExpire)
                    {
                        try
                        {
                            reportfileNeedDisabled.WriteLine(tmpStr);
                        }
                        catch (Exception ex)
                        {
                            hasError = true;
                            hasWriteEntireError = true;
                            swLog.WriteLine("Write report file failed, please check if there is invalid text character: " + privateants.vbCrLf + tmpStr);
                        }
                    }
                }
            }
        }
    }
    else if ((Strings.LCase(Status) == "active" && Strings.UCase(NeedAccount) != noNeedAccountValueHR) && DistinguishedNameFound != "" && (Strings.LCase(statusAD) != "active" && !samID2hasActive))
    {
        // if HireDate still in future, it must be an rehire, so ignore the discrepancy
        DateTime theHireDate = new DateTime();
        theHireDate = HireDate;
        if (theHireDate < DateTime.Now && Jobtitle != "")
        {
            // AD account may be disabled for inactive over 90 days, ignore the discrepancy
            DateTime LastLogonTime;
            int inactiveDays;
            LastLogonTime = LastLogonTimeStamp;  // GetUserLastLogonTime(samAccountName)
            inactiveDays = DateTime.DateDiff("d", LastLogonTime, DateTime.Now);

            if (inactiveDays < daysInactiveLimit)
            {
                DateTime ADchangedDate = new DateTime();
                try
                {
                    ADchangedDate = whenChanged;
                }
                catch (Exception ex)
                {
                }
                if (DateTime.Now > ADchangedDate.AddDays(2))
                {
                    noteStr = noteStr + "\"" + hrMIS + " need Terminated?\"";
                    HRneedTerminatedCount += 1;
                }
            }
        }
    }
    else if (Strings.LCase(Status) == "active")
    {
        if (hasLeaveReturningRecord(samAccountName))
        {
            DateTime lastDate = new DateTime(), returnDate = new DateTime();
            try
            {
                lastDate = LastLogonTimeStamp;
                returnDate = leaveReturningEffDateStr;
                if (lastDate > returnDate)
                    // remove the “Password never expires” 'CTEWS-5225
                    deleteLeaveReturning(samAccountName);
            }
            catch (Exception ex)
            {
            }
        }
    }
    lineStr = lineStr + Name_mismatch + ",\"" + FTE_mismatch + "\"," + Title_mismatch + "," + Department_mismatch + "," + Manager_mismatch + "," + noteStrEmployeeID + "," + noteStrSamAccountName + ",";
    lineStr += "\"" + defaultGroup + "\"," + "\"" + memberOf + "\",";
    lineStr += noteStr;
    try
    {
        reportfileEntire.WriteLine(lineStr);
        reportfileFound.WriteLine(lineStr);
    }
    catch (Exception ex)
    {
        hasWriteEntireError = true;
        // hasError = True
        swLog.WriteLine("");
        swLog.WriteLine("====>Write report file failed, please check if there is invalid text character: " + privateants.vbCrLf + lineStr);

        // TODO: should try to replace the invalid character in Hex mode or binary mode
        try
        {
            lineStr = lineStr.Replace(jobtitleAD, jobtitleAD.Substring(0, jobtitleAD.IndexOf(" ")) + " ??");
            reportfileEntire.WriteLine(lineStr);
            reportfileFound.WriteLine(lineStr);
        }
        catch (Exception ex2)
        {
        }
    }
}

private static void getGroupDN(string groupSamID)
{
    getGroupDN = "";
    if (groupSamID != "")
    {
        object connQueryGroup = Interaction.CreateObject("ADODB.Connection");
        object cmdQueryGroup = Interaction.CreateObject("ADODB.Command");
        connQueryGroup.Provider = "ADsDSOObject";
        connQueryGroup.Open("Active Directory Provider");
        // objConnection.Open "Provider=ADsDSOObject;"
        cmdQueryGroup.ActiveConnection = connQueryGroup;
        cmdQueryGroup.Properties("Page Size") = 100;
        strfilter = "(&(objectCategory=Group)(samAccountName=" + groupSamID + "))";
        cmdQueryGroup.commandtext = "<LDAP://" + strRoot + ">;" + strfilter + ";DistinguishedName;" + strScope;
        object rsQueryGroup = cmdQueryGroup.EXECUTE;
        if (rsQueryGroup.RecordCount > 0)
            getGroupDN = rsQueryGroup.Fields("DistinguishedName").value;
    }
}

public static void getManagerDN(string eeID)
{
    // if found once, should not find again and again, add to arrayList: ManagerID, ManagerDN
    ManagerDN = "";
    if (eeID != "")
    {
        connQueryManager = Interaction.CreateObject("ADODB.Connection");
        cmdQueryManager = Interaction.CreateObject("ADODB.Command");
        connQueryManager.Provider = "ADsDSOObject";
        connQueryManager.Open("Active Directory Provider");
        // objConnection.Open "Provider=ADsDSOObject;"
        cmdQueryManager.ActiveConnection = connQueryManager;

        cmdQueryManager.Properties("Page Size") = 100;
        strfilter = "(&(objectCategory=Person)(objectClass=User)(employeeID=" + eeID + "))";
        cmdQueryManager.commandtext = "<LDAP://" + strRoot + ">;" + strfilter + ";DistinguishedName;" + strScope;
        rsQueryManager = cmdQueryManager.EXECUTE;
        if (rsQueryManager.RecordCount > 0)
            // DistinguishedName never null:
            ManagerDN = rsQueryManager.Fields("DistinguishedName").value;
    }
}

private static bool findByReqID_indexWay(string findID)
{
    needDisableInactive = false;
    index = reqIDlist.IndexOf(findID);
    if (index < 0)
    {
        if (findID.StartsWith("0"))
        {
            // if AD and HR system reqID just the starting 0 different, consider it's the same:
            tmpStr = Microsoft.VisualBasic.Right(findID, Strings.Len(findID) - 1);
            index = eeIDlist.IndexOf(tmpStr);
            if (index < 0)
                return false;
        }
        else
            return false;
    }
    employeeidAD = eeIDlist(index);
    DistinguishedNameFound = userDNlist.Item(index);
    samAccountNameFound = samIDlist.Item(index);
    // getUserAttributes()
    LegalFirstNameAD = legalFirstNameList(index);
    firstNameAD = firstNameList.Item(index);
    lastNameAD = lastNameList.Item(index);
    middleNameAD = middleNameList.Item(index);
    statusAD = statusList.Item(index);
    if (inactiveList.Item(index) == inactiveYesStr)
    {
        int inactiveDays = 0;
        try
        {
            inactiveDays = inactiveDaysList.Item(index);
        }
        catch (Exception ex)
        {
        }
        if (inactiveDays >= daysInactiveLimit)
            needDisableInactive = true;
    }
    else if (inactiveList.Item(index).startswith("expired:"))
        AccountExpirationDate = inactiveList.Item(index).replace("expired:", "");

    jobtitleAD = jobtitleList.Item(index);
    homeDepartmentAD = HomeDepartmentList.Item(index);
    WorkemailAD = WorkemailList.Item(index);
    proxyAddressesAD = proxyAddressesList(index);
    whenCreated = whenCreatedList.Item(index);
    whenChanged = whenChangedList.Item(index);
    LastLogonTimeStamp = LastLogonTimeStampList(index);
    PhysicalDeliveryOfficeNameAD = PhysicalDeliveryOfficeNameList.Item(index);
    employeeTypeAD = employeeTypeList(index);
    cityAD = cityList.Item(index);
    managerAD = managerList.Item(index);
    return true;
}

private static void findByEmployeeIDlistMoreValue(string findID)
{
    needDisableInactive = false;
    index = eeIDlist.IndexOf(findID);
    if (index < 0)
    {
        if (findID.StartsWith("0"))
        {
            // if AD 30265,  HR system 030265 , consider it's the same:
            tmpStr = Microsoft.VisualBasic.Right(findID, Strings.Len(findID) - 1);
            while ((tmpStr.StartsWith("0")))
                tmpStr = Microsoft.VisualBasic.Right(tmpStr, Strings.Len(tmpStr) - 1);
            index = eeIDlist.IndexOf(tmpStr);
            if (index < 0)
                return;
        }
        else
        {
            // don't allow AD employeeID use lower case EEID, just return
            return;

            index = eeIDlist.IndexOf(Strings.LCase(findID));
            if (index < 0)
                return;
        }
    }
    indexList.Clear();
    indexList.Add(index); // including first found
                          // find more employee ID and put its index in indexList
    for (var j = index + 1; j <= eeIDlist.Count - 1; j++)
    {
        if (eeIDlist.Item(j) == findID)
            indexList.Add(j);
    }

    reqidAD = reqIDlist.Item(index);
    DistinguishedNameFound = userDNlist.Item(index);
    samAccountNameFound = samIDlist.Item(index);
    // getUserAttributes()
    LegalFirstNameAD = legalFirstNameList.Item(index);
    firstNameAD = firstNameList.Item(index);
    lastNameAD = lastNameList.Item(index);
    middleNameAD = middleNameList.Item(index);
    statusAD = statusList.Item(index);
    if (inactiveList.Item(index) == inactiveYesStr)
    {
        int inactiveDays = 0;
        try
        {
            inactiveDays = inactiveDaysList.Item(index);
        }
        catch (Exception ex)
        {
        }
        if (inactiveDays >= daysInactiveLimit)
            needDisableInactive = true;
    }
    else if (inactiveList.Item(index).startswith("expired:"))
        AccountExpirationDate = inactiveList.Item(index).replace("expired:", "");

    jobtitleAD = jobtitleList.Item(index);
    homeDepartmentAD = HomeDepartmentList.Item(index);
    WorkemailAD = WorkemailList.Item(index);
    proxyAddressesAD = proxyAddressesList(index);
    whenCreated = whenCreatedList.Item(index);
    whenChanged = whenChangedList.Item(index);
    LastLogonTimeStamp = LastLogonTimeStampList(index);
    PhysicalDeliveryOfficeNameAD = PhysicalDeliveryOfficeNameList.Item(index);
    employeeTypeAD = employeeTypeList(index);
    cityAD = cityList.Item(index);
    managerAD = managerList.Item(index);

    if (indexList.Count < 2)
        return;

    for (var j = 0; j <= indexList.Count - 1; j++)
    {
        samID2 = LCase(samIDlist.Item(indexList.Item(j)));
        status2 = "(" + statusList.Item(indexList.Item(j)) + ")";
        if (Strings.InStr(status2, "Active") > 0)
            samID2hasActive = true;
        samIDs.Add(samID2);
        statusS.Add(status2);
        if (samID2 == samAccountName || (statusAD == "Terminated" && status2 == "(Active)"))
        {
            tmpStr = samAccountNameFound;
            samAccountNameFound = samID2;
            samID2 = tmpStr;
            DistinguishedNameFound = userDNlist.Item(indexList.Item(j));
            // getUserAttributes()
            LegalFirstNameAD = legalFirstNameList.Item(indexList(j));
            firstNameAD = firstNameList.Item(indexList.Item(j));
            lastNameAD = lastNameList.Item(indexList.Item(j));
            middleNameAD = middleNameList.Item(indexList.Item(j));
            statusAD = statusList.Item(indexList.Item(j));
            needDisableInactive = false;
            if (inactiveList.Item(indexList.Item(j)) == inactiveYesStr)
            {
                int inactiveDays = 0;
                try
                {
                    inactiveDays = inactiveDaysList.Item(j);
                }
                catch (Exception ex)
                {
                }
                if (inactiveDays >= daysInactiveLimit)
                    needDisableInactive = true;
            }
            else if (inactiveList.Item(indexList.Item(j)).startswith("expired:"))
                AccountExpirationDate = inactiveList.Item(indexList.Item(j)).replace("expired:", "");

            jobtitleAD = jobtitleList.Item(indexList.Item(j));
            homeDepartmentAD = HomeDepartmentList.Item(indexList.Item(j));
            WorkemailAD = WorkemailList.Item(indexList.Item(j));
            proxyAddressesAD = proxyAddressesList(j);
            whenCreated = whenCreatedList.Item(indexList.Item(j));
            whenChanged = whenChangedList.Item(indexList.Item(j));
            LastLogonTimeStamp = LastLogonTimeStampList(indexList(j));
            PhysicalDeliveryOfficeNameAD = PhysicalDeliveryOfficeNameList.Item(indexList.Item(j));
            employeeTypeAD = employeeTypeList(indexList.Item(j));
            cityAD = cityList.Item(indexList.Item(j));
            managerAD = managerList.Item(indexList.Item(j));
        }
    }
    for (j = 0; j <= samIDs.Count - 1; j++)
        samAccountNameFound2 = samAccountNameFound2 + samIDs.Item(j) + statusS.Item(j) + "  ";
}

private static void findBySamIDlistMoreValue(string findID)
{
    needDisableInactive = false;
    index = samIDlist.IndexOf(findID);
    if (index < 0)
        return;
    DistinguishedNameFound = userDNlist.Item(index);
    employeeIDfound = eeIDlist.Item(index);
    employeeidAD = employeeIDfound;
    reqidAD = reqIDlist.Item(index);
    // getUserAttributes()
    LegalFirstNameAD = legalFirstNameList.Item(index);
    firstNameAD = firstNameList.Item(index);
    lastNameAD = lastNameList.Item(index);
    middleNameAD = middleNameList.Item(index);
    statusAD = statusList.Item(index);
    if (inactiveList.Item(index) == inactiveYesStr)
    {
        int inactiveDays = 0;
        try
        {
            inactiveDays = inactiveDaysList.Item(index);
        }
        catch (Exception ex)
        {
        }
        if (inactiveDays >= daysInactiveLimit)
            needDisableInactive = true;
    }
    else if (inactiveList.Item(index).startswith("expired:"))
        AccountExpirationDate = inactiveList.Item(index).replace("expired:", "");

    jobtitleAD = jobtitleList.Item(index);
    homeDepartmentAD = HomeDepartmentList.Item(index);
    WorkemailAD = WorkemailList.Item(index);
    proxyAddressesAD = proxyAddressesList(index);
    whenCreated = whenCreatedList.Item(index);
    whenChanged = whenChangedList.Item(index);
    LastLogonTimeStamp = LastLogonTimeStampList(index);
    PhysicalDeliveryOfficeNameAD = PhysicalDeliveryOfficeNameList.Item(index);
    employeeTypeAD = employeeTypeList(index);
    cityAD = cityList.Item(index);
    managerAD = managerList.Item(index);
}

private static bool findByEmailListKeyMoreValue(string findID)
{
    needDisableInactive = false;
    if (findID == "")
        return false;
    if (Jobtitle == "COBRA Special Dependents")
        return false;
    index = WorkemailListKey.IndexOf(findID);
    if (index < 0)
        return false;

    string nameKey = LCase(firstNameList.Item(index) + "." + lastNameList.Item(index));
    // fixing incorrect discrepancy report:  WFN is Dawn.DawdyMoore@ehealth.com. Should be D.DawdyMoore@ehealth.com?
    nameKey = nameKey.Replace(" ", "") + "@";
    if (nameKey != findID && nameKey.Replace("-", "") != findID)
        return false;
    DistinguishedNameFound = userDNlist.Item(index);
    samAccountNameFound = samIDlist.Item(index);
    employeeIDfound = eeIDlist.Item(index);
    employeeidAD = employeeIDfound;
    reqidAD = reqIDlist.Item(index);
    // getUserAttributes()
    LegalFirstNameAD = legalFirstNameList.Item(index);
    firstNameAD = firstNameList.Item(index);
    lastNameAD = lastNameList.Item(index);
    middleNameAD = middleNameList.Item(index);
    statusAD = statusList.Item(index);
    if (inactiveList.Item(index) == inactiveYesStr)
    {
        int inactiveDays = 0;
        try
        {
            inactiveDays = inactiveDaysList.Item(index);
        }
        catch (Exception ex)
        {
        }
        if (inactiveDays >= daysInactiveLimit)
            needDisableInactive = true;
    }
    else if (inactiveList.Item(index).startswith("expired:"))
        AccountExpirationDate = inactiveList.Item(index).replace("expired:", "");

    jobtitleAD = jobtitleList.Item(index);
    homeDepartmentAD = HomeDepartmentList.Item(index);
    WorkemailAD = WorkemailList.Item(index);
    proxyAddressesAD = proxyAddressesList(index);
    whenCreated = whenCreatedList.Item(index);
    whenChanged = whenChangedList.Item(index);
    LastLogonTimeStamp = LastLogonTimeStampList(index);
    PhysicalDeliveryOfficeNameAD = PhysicalDeliveryOfficeNameList.Item(index);
    employeeTypeAD = employeeTypeList(index);
    cityAD = cityList.Item(index);
    managerAD = managerList.Item(index);
    return true;
}

public static void findBySamAccountName(string findID, string ADtype)
{
    DistinguishedNameFound = "";
    employeeIDfound = "";
    // Search for a User Account in Active Directory
    objConnection = Interaction.CreateObject("ADODB.Connection");
    objConnection.Open("Provider=ADsDSOObject;");

    objCommand = Interaction.CreateObject("ADODB.Command");
    objCommand.ActiveConnection = objConnection;
    // strfilter = "(&(objectCategory=Person)(objectClass=Contact))" 'InetOrgPerson

    objCommand.CommandText = "<LDAP://" + defaultOU + strDomain + ">;(&(objectCategory=" + ADtype + ")" + "(samAccountName=" + findID + "));distinguishedName,employeeID;subtree";

    try
    {
        objRecordSet = objCommand.Execute;
    }
    catch (Exception ex)
    {
        // noteStr = findID & " in findBySamAccountName: " & ex.Message
        // reportfileNotFound.WriteLine(lineStr & Name_mismatch & "," & FTE_mismatch & "," & Title_mismatch & "," & Department_mismatch & "," & noteStrEmployeeID & "," & noteStrSamAccountName & "," & noteStr)
        swLog.WriteLine(findID + " in findBySamAccountName: " + ex.Message + privateants.vbCrLf + objCommand.CommandText);
        return;
    }
    if (objRecordSet.RecordCount > 0)
    {
        try
        {
            DistinguishedNameFound = objRecordSet.Fields("distinguishedName").value;
            if (!Information.IsDBNull(objRecordSet.Fields("employeeID").value))
            {
                employeeIDfound = objRecordSet.Fields("employeeID").value;
                employeeidAD = employeeIDfound;
            }
        }
        catch (Exception ex)
        {
            swLog.WriteLine("Get objRecordSet value error in function findBySamAccountName.");
        }
    }
    objRecordSet = null;
    objConnection.Close();
}

public static void findByEmployeeID(string findID, string ADtype)
{
    // Search for a User Account in Active Directory
    objConnection = Interaction.CreateObject("ADODB.Connection");
    objConnection.Open("Provider=ADsDSOObject;");

    objCommand = Interaction.CreateObject("ADODB.Command");
    objCommand.ActiveConnection = objConnection;
    // strfilter = "(&(objectCategory=Person)(objectClass=Contact))" 'InetOrgPerson
    objCommand.CommandText = "<LDAP://" + defaultOU + strDomain + ">;(&(objectCategory=" + ADtype + ")" + "(employeeID=" + findID + "));distinguishedName,samAccountName,userAccountControl;subtree";
    // objCommand.CommandText = "<LDAP://" & defaultOU & strDomain & ">;(&(objectCategory=" & ADtype & ")" & "(employeeID=" & findID & "));" & strAttributes & ";subtree"

    try
    {
        objRecordSet = objCommand.Execute;
    }
    catch (Exception ex)
    {
        swLog.WriteLine(findID + " in findByEmployeeID(findID: " + ex.Message + privateants.vbCrLf + objCommand.CommandText);
        return;
    }
    if (objRecordSet.RecordCount > 0)
    {
        try
        {
            DistinguishedNameFound = objRecordSet.Fields("distinguishedName").value;
            samAccountNameFound = objRecordSet.Fields("samAccountName").value;
        }
        catch (Exception ex)
        {
            swLog.WriteLine("Get objRecordSet value error in function findByEmployeeID.");
        }
        // getUserAttributesRS()
        statusAD = "";
        try
        {
            intUAC = objRecordSet.Fields("userAccountControl").value;
            if (intUAC & ADS_UF_ACCOUNTDISABLE)
                statusAD = "Terminated";
            else
                statusAD = "Active";
        }
        catch (Exception ex)
        {
        }

        // one employeeID may have 2 accounts, like Xiamen RA team
        if (objRecordSet.RecordCount > 1)
        {
            objRecordSet.MoveFirst();
            while (!(objRecordSet.EOF))
            {
                try
                {
                    samID2 = LCase(objRecordSet.Fields("samAccountName").value);
                    try
                    {
                        intUAC = objRecordSet.Fields("userAccountControl").value;
                        if (intUAC & ADS_UF_ACCOUNTDISABLE)
                            status2 = "(Terminated)";
                        else
                        {
                            status2 = "(Active)";
                            samID2hasActive = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        status2 = "";
                    }
                    // If Not samIDs.Contains(samID2) Then
                    samIDs.Add(samID2);
                    statusS.Add(status2);
                    // End If
                    if (samID2 == samAccountName || (statusAD == "Terminated" && status2 == "(Active)"))
                    {
                        tmpStr = samAccountNameFound;
                        samAccountNameFound = samID2;
                        samID2 = tmpStr;
                        DistinguishedNameFound = objRecordSet.Fields("distinguishedName").value;
                    }
                }
                catch (Exception ex)
                {
                    swLog.WriteLine("Get objRecordSet value error in function findByEmployeeID.");
                }
                objRecordSet.MoveNext();
            }
            for (var j = 0; j <= samIDs.Count - 1; j++)
                samAccountNameFound2 = samAccountNameFound2 + "[ " + samIDs.Item(j) + " ]" + statusS.Item(j) + "  ";
        }
    }
    objRecordSet = null;
    objConnection.Close();
}

private static void SearchByNameMoreValue(string findName)
{
    needDisableInactive = false;
    index = -1;
    for (int findindex = 0; findindex <= samIDlist.Count - 1; findindex++)
    {
        if (firstNameList.Item(findindex) == findName && lastNameList.Item(findindex) == LastName)
        {
            index = findindex;
            break;
        }
    }
    if (index < 0)
        return;

    indexList.Clear();
    indexList.Add(index); // including first found
                          // find more name match and put its index in indexList
    for (var j = index + 1; j <= samIDlist.Count - 1; j++)
    {
        if (firstNameList.Item(j) == findName && lastNameList.Item(j) == LastName)
            indexList.Add(j);
    }

    DistinguishedNameFound = userDNlist.Item(index);
    samAccountNameFound = samIDlist.Item(index);
    samAccountName = samAccountNameFound;
    employeeIDfound = eeIDlist.Item(index);
    employeeidAD = employeeIDfound;
    reqidAD = reqIDlist.Item(index);
    // getUserAttributes()
    LegalFirstNameAD = legalFirstNameList(index);
    firstNameAD = firstNameList.Item(index);
    lastNameAD = lastNameList.Item(index);
    middleNameAD = middleNameList.Item(index);
    statusAD = statusList.Item(index);
    if (inactiveList.Item(index) == inactiveYesStr)
    {
        int inactiveDays = 0;
        try
        {
            inactiveDays = inactiveDaysList.Item(index);
        }
        catch (Exception ex)
        {
        }
        if (inactiveDays >= daysInactiveLimit)
            needDisableInactive = true;
    }

    jobtitleAD = jobtitleList.Item(index);
    homeDepartmentAD = HomeDepartmentList.Item(index);
    WorkemailAD = WorkemailList.Item(index);
    proxyAddressesAD = proxyAddressesList(index);
    whenCreated = whenCreatedList.Item(index);
    whenChanged = whenChangedList.Item(index);
    LastLogonTimeStamp = LastLogonTimeStampList(index);
    PhysicalDeliveryOfficeNameAD = PhysicalDeliveryOfficeNameList.Item(index);
    employeeTypeAD = employeeTypeList(index);
    cityAD = cityList.Item(index);
    managerAD = managerList.Item(index);

    if (indexList.Count < 2)
    {
        eeID2 = "";
        return;
    }
    for (var j = 0; j <= indexList.Count - 1; j++)
    {
        samID2 = samIDlist.Item(indexList.Item(j));
        eeID2 = eeIDlist.Item(indexList.Item(j));
        status2 = "(" + statusList.Item(indexList.Item(j)) + ")";
        if (Strings.InStr(status2, "Active") > 0)
            samID2hasActive = true;
        samIDs.Add(samID2);
        eeIDs.Add(eeID2);
        statusS.Add(status2);
    }
    for (j = 0; j <= samIDs.Count - 1; j++)
        samAccountNameFound2 = samAccountNameFound2 + "[ " + samIDs.Item(j) + " ]" + statusS.Item(j) + "  ";
    for (j = 0; j <= eeIDs.Count - 1; j++)
        samAccountNameFound2 = samAccountNameFound2 + "[ " + eeIDs.Item(j) + " ]" + statusS.Item(j) + "  ";
    samAccountName = samAccountNameFound;
}

private static bool UpdateFoundUserAfterChecking(bool needUpdateTerminated)
{
    // both set AD account disabled and update other properties call this function
    intRowCount = 0;
    intOKcount = 0;
    intErrorCount = 0;
    emailCClist.Clear();
    if (HRrowCount < 1)
    {
        swLog.WriteLine(hrMIS + " file has been opened, but no data to process.");
        return true;
    }

    for (int i = 0; i <= HRrowCount - 1; i++)
    {
        intRowCount = intRowCount + 1;

        FirstName = Trim(HR_firstName.Item(i));
        LastName = Trim(HR_lastName.Item(i));
        MiddleName = Trim(HR_middleName.Item(i));

        NickName = Trim(HR_nickName.Item(i));
        NickName = NickName.Replace("(", "");
        NickName = NickName.Replace(")", "");
        employeeID = Trim(HR_eeID.Item(i));
        Status = Trim(HR_status.Item(i));
        if (Status == "Leave" || Status == "L")
            Status = "On Leave";
        else if (Status == "Deceased" || Status == "T")
            Status = "Terminated";
        else if (Status == "A")
            Status = "Active";

        statusEffDateStr = Trim(HR_statusEffDate.Item(i));

        NeedAccount = Trim(HR_needAccount.Item(i));
        DivisionCode = Trim(HR_DivisionCode(i));
        statusAD = Trim(AD_status.Item(i));
        // for inactive over 90 days, will be marked in this field in ADneedDisabled.CSV: needDisabled
        AccountExpirationDate = "";
        needDisabledForInactive = false;
        needDisabledForNotLogon = false;
        if (HR_needDisabled.Item(i).Contains("need set expired for Leave"))
            AccountExpirationDate = "Not set";
        else if (HR_needDisabled.Item(i).Contains("Need remove ExpirationDate for Leave"))
            AccountExpirationDate = "Need remove";
        else if (HR_needDisabled.Item(i).Contains("need set Never expires"))
            AccountExpirationDate = "need set Never expires";
        else if (HR_needDisabled.Item(i).Contains("AD account need set expired for inactive"))
            needDisabledForInactive = true;
        else if (needCheckNewUsersNotLogon)
        {
            if (HR_needDisabled.Item(i).Contains(newHireNotLogonStr + DaysInactiveLimitNewHire + " days"))
                needDisabledForNotLogon = true;
        }

        Workemail = Trim(HR_workEmail.Item(i)); // workemail should not be updated with HR value
        Jobtitle = Trim(HR_Jobtitle.Item(i)); // if not hasJobTitle, the list will be blank list, so Jobtitle is blank
        if (Jobtitle.EndsWith("."))
            Jobtitle = Jobtitle.Remove(Jobtitle.Length - 1);

        JobEffDate = Trim(HR_JobEffDate.Item(i));
        BusinessUnit = Trim(HR_BusinessUnit(i));
        HomeDepartment = Trim(HR_HomeDepartment.Item(i)); // HR_HomeDepartment already prepared by openCSVFileForChecking()
        CostCenter = Trim(HR_CostCenter.Item(i));
        // CostCenter.Length >= 4 in generated CSV file
        EmpClass = Trim(HR_FTE.Item(i));
        OriginalHireDate = Trim(HR_OriginalHireDate.Item(i));
        HireDate = Trim(HR_HireDate.Item(i));
        ManagerID = Trim(HR_managerID.Item(i));
        ManagerDN = Trim(HR_managerDN.Item(i));
        ManagerWorkEmail = Trim(HR_managerWorkEmail.Item(i));
        UserSegment = Trim(HR_UserSegment(i));
        Mobile = Trim(HR_mobile.Item(i));
        LocationAux1 = UCase(Trim(HR_locationAux1.Item(i)));

        DistinguishedName = LCase(Trim(HR_DistinguishedName.Item(i)));

        samAccountName = LCase(Trim(HR_samAccountName.Item(i)));
        GetUserMemberOf(samAccountName);
        mapOU_Address();
        set_ScriptPath_HomeDirectory(samAccountName);

        Address = Trim(HR_Address.Item(i));
        PObox = "";
        City = Trim(HR_City.Item(i));
        State = Trim(HR_State.Item(i));
        PostalCode = Trim(HR_PostalCode.Item(i));
        if (PostalCode != "")
        {
            while ((PostalCode.Length < 5))
                PostalCode = "0" + PostalCode;
        }

        Country = Trim(HR_Country.Item(i));
        needUpdateEmployeeID = false;
        if (HR_reqID.Item(i) == "need set EEID for reqID")
        {
            if (!needStop)
                needUpdateEmployeeID = true;
        }

        Description = "";
        otherTelephone = "";

        string needAccountStr = "";
        if (Strings.UCase(NeedAccount) == noNeedAccountValueHR)
            needAccountStr = "(No need account)";

        tmpStr = FirstName;
        if (NickName != "")
            tmpStr = NickName;
        // 300661 last name including ","
        strGeneral = "\"" + employeeID + "\"," + samAccountName + "," + tmpStr + ",\"" + LastName + "\"," + Status + ",\"" + HomeDepartment + "\",\"" + Jobtitle + "\"";
        if (tmpStr == "" && LastName == "")
        {
            string DNOU = Strings.UCase(DistinguishedName).Replace("," + Strings.UCase(strRoot), "");
            // lineStr = DNOU & " """ & employeeID & """," & samAccountName & " (###### no " & hrMIS & " record, if noshow please delete the waste account! ######)"
            string needDisabledText = Trim(HR_needDisabled(i));
            int ADindex = samIDlist.IndexOf(samAccountName);
            lineStr = DNOU + " \"" + employeeID + "\"," + samAccountName + " (" + firstNameList(ADindex) + "," + lastNameList(ADindex) + " ###### " + needDisabledText + " ######)";
        }
        else
            lineStr = "\"" + employeeID + "\"," + samAccountName + "(" + tmpStr + " " + LastName + ")," + Status + needAccountStr + ",\"" + HomeDepartment + "\",\"" + Jobtitle + "\",\"" + LocationAux1 + "\"";

        if (samAccountName == "")
        {
            swReport.WriteLine("row#" + intRowCount + ": " + lineStr);
            swReport.Close();
            swLog.WriteLine("The row " + intRowCount + "th user samAccountName is blank! Please check and correct first, then run the program again. ");
            return false;
        }
        if (employeeID == "")
        {
            swReport.WriteLine("row#" + intRowCount + ": " + lineStr);
            swReport.Close();
            swLog.WriteLine("The row " + intRowCount + "th person employeeID is blank! Please check and correct first, then run the program again. ");
            return false;
        }
        else if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
        {
            actualIgnoredEEIDs += employeeID + "; ";
            // swReport.WriteLine("row#" & intRowCount & ": " & lineStr)
            // swReport.WriteLine("    Ignore updating because the employeeID is in ignore update list")
            // swReport.WriteLine("")
            continue;
        }

        if (statusAD != "Active" && DistinguishedName.Contains("ou=alumni") && !needUpdateTerminated)
            continue;


        strReport = "";
        hasUpdated = false;
        tmpStr = HR_reqID.Item(i);
        if (tmpStr.Contains("Name mismatch:"))
        {
            if (tmpStr.Contains(major))
            {
                if (Status != "Terminated")
                {
                    swReport.WriteLine("====> " + lineStr);
                    swReport.WriteLine("    " + tmpStr.Replace(major, ""));
                    swReport.WriteLine("    No further updating. Please confirm the AD account found is correct first!");
                    swReport.WriteLine("");
                }
                continue;
            }
        }

        // the 2nd time should not do set disabled updating, otherwise the On Leave returnning will be set disabled again
        if (ignoreRehireSamIDs.Contains("." + samAccountName + "."))
            continue;

        if (!updateProperties(needUpdateTerminated))
            continue;

        if (hasUpdated)
        {
            intOKcount = intOKcount + 1;
            swReport.WriteLine("#" + intOKcount + ": " + lineStr);
            swReport.WriteLine(strReport);

            swriterGeneralToday.WriteLine(DateTime.Now + "," + appNameVersion + " " + runnerName + "," + strGeneral);
        }
    }
    tmpStr = intRowCount + " persons processed (" + intOKcount + " updated)";
    if (intErrorCount > 0)
        tmpStr = tmpStr + " but has " + intErrorCount + " update failed";
    Console.WriteLine(tmpStr);
    swLog.WriteLine(tmpStr);
    swLog.WriteLine("See " + strReportfileName);

    if (alreadyExpiredForLeave != "")
    {
        swReport.WriteLine("");
        swReport.WriteLine("AD account already set expired for Leave:");
        swReport.WriteLine(alreadyExpiredForLeave);
        swReport.WriteLine("");
    }

    if (alreadyExpiredForInactivate != "")
    {
        swReport.WriteLine("");
        swReport.WriteLine("AD account already set expired for inactive over " + daysInactiveLimit + " days:");
        if (forLeaveOnly)
            swReport.WriteLine("(just list for on leave return only)");
        swReport.WriteLine(alreadyExpiredForInactivate);
        swReport.WriteLine("");
    }

    if (actualIgnoredEEIDs != "")
    {
        swReport.WriteLine("");
        swReport.WriteLine("Ignore updating employeeID list:");
        swReport.WriteLine(actualIgnoredEEIDs);
    }

    return true;
}

private static bool updateProperties(bool needSetDisabled)
{
    // here alredy is minor name mismatch

    DistinguishedNameFound = "";
    employeeIDfound = "";
    samAccountNameFound = "";
    samAccountNameFound2 = "";
    samID2 = "";

    findBySamAccountName(samAccountName, "User");
    if (DistinguishedNameFound == "")
    {
        // only update found users after checking, so this error will never happen
        swLog.WriteLine("   =====>cannot find account samAccountName=" + samAccountName);
        return false;
    }
    objUser = Interaction.GetObject("LDAP://" + DistinguishedNameFound);
    bool isWDID = false;
    // 30155<===>030155 will consider is equivalent
    if (employeeIDfound != "" && employeeID != employeeIDfound && employeeID != "0" + employeeIDfound)
    {
        int preIndex = previousID.IndexOf(employeeID);
        if (preIndex >= 0 && WDID(preIndex) == employeeIDfound && objUser.employeeNumber == employeeID)
            // allow WDID set into employeeID and previousID set into employeeNumber
            isWDID = true;
        else if (Strings.Trim(employeeIDfound) == employeeID || Strings.LCase(employeeIDfound) == Strings.LCase(employeeID) || Conversion.Val(employeeIDfound) == Conversion.Val(employeeID))
        {
            if (!needStop)
                needUpdateEmployeeID = true;// force update employeeID without space and to upper case
        }
        else
        {
            strReport += "   samAccountName=" + samAccountName + ", employeeID mismatch: \"" + employeeIDfound + "\"(AD)<===>(" + hrMIS + ")\"" + employeeID + "\"" + privateants.vbCrLf;
            return false; // if employeeID mismatch, do not continue updating
        }
    }

    // If InStr(UCase(DistinguishedNameFound), "OU=EHI,") < 1 Then 'only update users in OU=EHI
    // Return
    // End If

    partialNeedUpdated = false;
    if (needUpdateEmployeeID && (!needStop) && !isWDID)
    {
        if (employeeID != "" && objUser.employeeID != employeeID)
        {
            arg1 = "    update employeeID " + "\"" + objUser.employeeID + "\"" + " to " + "\"" + employeeID + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.employeeID = employeeID;
            arg2 = "";
            if (needClearEmployeeNumber)
            {
                if (objUser.employeeNumber != "")
                {
                    arg2 = "    clear employeeNumber(reqID) " + "\"" + objUser.employeeNumber + "\"" + " to " + "\"\"";
                    strReport += arg2 + privateants.vbCrLf;
                    objUser.PutEx(ADS_PROPERTY_CLEAR, "employeeNumber", privateants.vbNull);
                }
            }

            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                employeeIDfound = employeeID;
                strGeneral += ",\"" + Strings.Trim(arg1.Replace("\"", "'")) + "\"";
                if (arg2 != "")
                    strGeneral += ",\"" + Strings.Trim(arg2.Replace("\"", "'")) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>employeeID update failed!" + privateants.vbCrLf + ex.ToString();
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    if (needUpdateAccountControl)
    {
        if (Strings.LCase(Status).Contains("leave"))
        {
            // if no "On Leave" record for the user, insert "L" record, regardless AD is already disabled or not, for reactivate use
            if (!hasOnLeaveRecord(objUser.samAccountName))
                // set “Password never expires” 'CTEWS-5225
                insertOnLeaveRecord(objUser.samAccountName);
        }

        // after first time updating strExportFileNeedDisabled, objUser.AccountDisabled will become true
        // the second time updating strExportFileFound will no longer go into this block

        if (Strings.LCase(Status) != "active" || Strings.UCase(NeedAccount) == noNeedAccountValueHR || needDisabledForInactive || needDisabledForNotLogon)
        {
            // Terminated or On Leave:
            // If objUser.AccountDisabled, do nothing here, but on sync action end, always deleteNonOnLeave()
            if (needSetDisabled && (needDisabledForInactive || needDisabledForNotLogon || !objUser.AccountDisabled))
            {
                // need disable mailbox before disable AD account
                // per Vishal, we should not disable mailbox for all location terminated employee
                if (needUpdateMailbox)
                {
                    if ((Strings.LCase(Status) == "terminated"))
                    {
                        if (Strings.LCase(runnerName) == serviceAccount)
                        {
                            // only do disable mailbox in server SJentMSutil01 running by service account
                            if (IsExistMailBox(samAccountName))
                            {
                                if (DisableMailBox(samAccountName))
                                {
                                    tmpStr = "    set mailbox from enabled to disabled.";
                                    strReport += tmpStr + privateants.vbCrLf;
                                    strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                                }
                                else
                                    strReport += "    =====>PowerShell command failed: Disable-Mailbox -identity " + samAccountName + privateants.vbCrLf;
                            }
                        }
                        else
                            strReport += "    =====>need run PowerShell command: Disable-Mailbox -identity " + samAccountName + privateants.vbCrLf;
                    }
                }

                // Vishal email Sent: Thursday, July 30, 2020 12:51 AM: Leave and inactive over 30 days just set expiration date only, don't disable
                bool hasSetExpired = false;
                if ((Strings.LCase(Status) == "terminated"))
                {
                    string hireDateStr = objUser.extensionAttribute9;
                    if (hireDateStr == "")
                        tmpStr = "    need disable AD account, please double check and disable manually!";
                    else
                    {
                        DateTime hireDate = "2020-08-01";
                        try
                        {
                            hireDate = Strings.Trim(hireDateStr);
                        }
                        catch (Exception ex)
                        {
                        }

                        if (hireDate < DateTime.Today)
                            tmpStr = "    need disable AD account, please double check and disable manually!";
                    }
                }
                else if (Strings.LCase(Status).Contains("leave") && AccountExpirationDate == "Not set")
                {
                    tmpStr = "    set AD account expired for Leave";
                    objUser.AccountExpirationDate = DateTime.Now.ToString("MM/dd/yyyy");  // "01/01/1970"
                    hasSetExpired = true;
                }
                else if (needDisabledForInactive && Strings.LCase(Status) != "terminated")
                {
                    tmpStr = "    set AD account expired for inactive over " + daysInactiveLimit + " days";
                    objUser.AccountExpirationDate = DateTime.Now.ToString("MM/dd/yyyy");  // "01/01/1970"
                    hasSetExpired = true;
                }
                else if (needDisabledForNotLogon)
                {
                    tmpStr = "    need disable AD account";
                    tmpStr += " for " + newHireNotLogonStr + DaysInactiveLimitNewHire + " days, please double check and disable manually!";
                }
                strReport += tmpStr + privateants.vbCrLf;
                hasUpdated = true;

                try
                {
                    if (hasSetExpired && !forTest)
                        objUser.SetInfo();
                    strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                    if (hasSetExpired)
                    {
                        DB_DisabledSamIDs.Add(Strings.LCase(Strings.Trim(objUser.samAccountName))); // getDisabledSamidFromDB() only can read history record, the new disabled may not in the list, so add it here
                        if (!InsertOrUpdateDisabledRecord(objUser.samAccountName, objUser.DistinguishedName))
                        {
                        }
                    }
                }
                catch (Exception ex)
                {
                    intErrorCount = intErrorCount + 1;
                    tmpStr = "    =====> set user expired failed!";
                    strReport += tmpStr + privateants.vbCrLf;
                }

                if (needUpdateMOCaccount && Strings.LCase(Status) == "terminated")
                {
                    try
                    {
                        tmpStr = objUser.Get("msRTCSIP-UserEnabled");
                    }
                    catch (Exception ex)
                    {
                        tmpStr = "";
                    }
                    if (tmpStr == "True")
                    {
                        tmpStr = "    set MOC account disabled";
                        strReport += tmpStr + privateants.vbCrLf;
                        objUser.put("msRTCSIP-UserEnabled", false);
                        hasUpdated = true;
                        try
                        {
                            if (!forTest)
                                objUser.SetInfo();
                            strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                        }
                        catch (Exception ex)
                        {
                            intErrorCount = intErrorCount + 1;
                            tmpStr = "    =====> set user MOC disabled failed!";
                            strReport += tmpStr + privateants.vbCrLf;
                        }
                    }
                }


                // Charles email 11/10/2014 When disabling the account of terminating employee, be sure to remove terminated user from all AD groups.
                // When disabling the account of someone who has left, be sure to move their account to the Alumni container
                // If LCase(Status) = "terminated" Then 'On Leave will not clear all groups

                // End If

                // for set disabled employee, don't update other properties further
                // since first time is updating strExportFileNeedDisabled
                // second time is updating strExportFileFound
                strReport += privateants.vbCrLf;
                return true;
            } // Not objUser.AccountDisabled
        }
        else
        {
            // If Not objUser.AccountDisabled, do nothing here, but on sync action end, always deleteNonOnLeave()

            if (AccountExpirationDate == "Need remove")
            {
                // Vishal email Sent: Thursday, July 30, 2020 12:51 AM: just remove expiration date for returning on Leave user
                if (hasOnLeaveRecord(objUser.samAccountName))
                {
                    tmpStr = "    set AD account Never expires for Leave return";
                    strReport += tmpStr + privateants.vbCrLf;
                    // objUser.AccountExpirationDate = "12/30/2099" 'Now.AddDays(365).ToString("MM/dd/yyyy") 'CTEWS-5250
                    objUser.AccountExpires = 0; // 0 or 0x7FFFFFFFFFFFFFFF (9223372036854775807) indicates that the account never expires.
                    hasUpdated = true;
                    try
                    {
                        if (!forTest)
                            objUser.SetInfo();
                        strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                        changeStatusLeaveReturning(objUser.samAccountName);
                    }
                    catch (Exception ex)
                    {
                        intErrorCount = intErrorCount + 1;
                        tmpStr = "    =====> set AD account Never expires for Leave returning user failed!";
                        strReport += tmpStr + privateants.vbCrLf;
                    }
                }
            }
            else if (AccountExpirationDate == "need set Never expires")
            {
                tmpStr = "    set AD account Never expires";
                strReport += tmpStr + privateants.vbCrLf;
                objUser.AccountExpires = 0;
                hasUpdated = true;
                try
                {
                    if (!forTest)
                        objUser.SetInfo();
                    strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                }
                catch (Exception ex)
                {
                    intErrorCount = intErrorCount + 1;
                    tmpStr = "    =====> set AD account Expires failed!";
                    strReport += tmpStr + privateants.vbCrLf;
                }
            }

            if (objUser.AccountDisabled)
            {
                // if has "On Leave" record for the user, reactivate AD account automatically
                if (hasOnLeaveRecord(objUser.samAccountName))
                {
                    tmpStr = "    reactivate AD account for Leave return";
                    strReport += tmpStr + privateants.vbCrLf;

                    objUser.AccountDisabled = false;
                    objUser.AccountExpires = 0;
                    hasUpdated = true;
                    try
                    {
                        if (!forTest)
                            objUser.SetInfo();
                        strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                        changeStatusLeaveReturning(objUser.samAccountName);
                    }
                    catch (Exception ex)
                    {
                        intErrorCount = intErrorCount + 1;
                        tmpStr = "    =====> reactivate AD account for Leave return user failed!";
                        strReport += tmpStr + privateants.vbCrLf;
                    }
                }
            }
            else if (needUpdatePhoneNotes && employeeID != "")
            {
                // copy EEID for showing on Outlook properties tab
                PhoneNotes = objUser.info;
                if (Strings.InStr(PhoneNotes, employeeID) < 1)
                {
                    arg1 = "";
                    arg2 = "";
                    partialNeedUpdated = false;
                    if (!needStop)
                    {
                        if (objUser.employeeID == "" && objUser.employeeNumber == "")
                        {
                            // also update employeeID
                            arg1 = "    update employeeID " + "\"" + objUser.employeeID + "\"" + " to " + "\"" + employeeID + "\"";
                            strReport += arg1 + privateants.vbCrLf;
                            objUser.employeeID = employeeID;
                        }
                    }

                    if (PhoneNotes == "EmployeeID:")
                        // avoiding updating phoneNotes from "EmployeeID:" to "EmployeeID: 301031  EmployeeID:"
                        PhoneNotes = "EmployeeID: " + employeeID;
                    else
                        PhoneNotes = "EmployeeID: " + employeeID + privateants.vbCrLf + PhoneNotes;
                    arg2 = "    update phoneNotes " + "\"" + objUser.info + "\"" + " to " + "\"" + PhoneNotes.Replace(privateants.vbCrLf, "  ") + "\"";
                    arg2 = arg2.Replace("  \"", "\"");
                    strReport += arg2 + privateants.vbCrLf;
                    objUser.info = PhoneNotes;

                    if (partialNeedUpdated)
                    {
                        hasUpdated = true;
                        try
                        {
                            if (!forTest)
                                objUser.SetInfo();
                            tmpStr = "";
                            if (Strings.Trim(arg1) != "")
                                tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                            if (Strings.Trim(arg2) != "")
                                tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                            strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                        }
                        catch (Exception ex)
                        {
                            intErrorCount = intErrorCount + 1;
                            tmpStr = "    =====>user employeeID and phoneNotes update failed!";
                            strReport += tmpStr + privateants.vbCrLf;
                        }
                    }
                }
            }
        }
    }

    if (needSetDisabled)
        // for those in need disabled list, just set disable and report separately, no further update this time
        return true;

    if (Status == "Terminated" && isNeedDisableTerminateRightNow())
    {
        if (DB_DisabledSamIDs.IndexOf(samAccountName) < 0)
        {
            string hireDateStr = objUser.extensionAttribute9;
            if (hireDateStr != "")
            {
                DateTime hireDate = "2020-08-01";
                try
                {
                    hireDate = Strings.Trim(hireDateStr);
                }
                catch (Exception ex)
                {
                }

                if (hireDate < DateTime.Today)
                {
                    if (needHideEmailForTerminated && employeeID != "" && Workemail != "")
                    {
                        if (objUser.msExchHideFromAddressLists == null || !objUser.msExchHideFromAddressLists)
                        {
                            // hide email address to show on Outlook properties tab
                            tmpStr = "    hide email address list on Outlook GAL";
                            strReport += tmpStr + privateants.vbCrLf;
                            objUser.msExchHideFromAddressLists = true;
                            objUser.PutEx(ADS_PROPERTY_CLEAR, "showInAddressBook", privateants.vbNull);
                            hasUpdated = true;
                            if (!forTest)
                                objUser.SetInfo();
                            strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                        }
                    }

                    if (needRemoveManagerForTerminated)
                    {
                        // CTEWS-2977 remove users from all Groups and clear organization reporting from AD upon termination
                        string currentDN = objUser.manager;
                        string shortCurrentDN = "";
                        string keyManagerDN = "";
                        string shortManagerDN = "";
                        if (currentDN != "")
                            shortCurrentDN = Microsoft.VisualBasic.Mid(currentDN, 4, Strings.InStr(currentDN, ",") - 4);
                        if (shortCurrentDN != "")
                        {
                            try
                            {
                                objUser.PutEx(ADS_PROPERTY_CLEAR, "manager", privateants.vbNull);
                                tmpStr = "    remove manager " + "\"" + shortCurrentDN + "\"";
                                strReport += tmpStr + privateants.vbCrLf;
                                hasUpdated = true;
                                if (!forTest)
                                    objUser.SetInfo();
                                strGeneral += ",\"" + Strings.Trim(tmpStr.Replace("\"", "'")) + "\"";
                            }
                            catch (Exception ex)
                            {
                                intErrorCount = intErrorCount + 1;
                                tmpStr = "    =====>remove manager failed!";
                                strReport += tmpStr + privateants.vbCrLf;
                            }
                        }
                    }
                }
            }
        }
        if (!needUpdaeTerminatedProperties)
            return true;
    }

    // here already is minor mismatch
    if (employeeID != "" && employeeID == employeeIDfound)
    {
        arg1 = "";
        arg2 = "";
        arg3 = "";
        arg4 = "";
        arg5 = "";
        partialNeedUpdated = false;
        if (needUpdateFirstName)
        {
            if (NickName != "")
            {
                // should use InStr , after removing " " and "-" ?
                if (objUser.givenName != NickName && !hasPreferredLastnameEEIDs.Contains("." + employeeID + "."))
                {
                    arg1 = "    update FirstName " + "\"" + objUser.givenName + "\"" + " to " + "\"" + NickName + "\"";
                    strReport += arg1 + privateants.vbCrLf;
                    objUser.givenName = NickName;
                    partialNeedUpdated = true;
                }
            }
            else if (FirstName != "" && objUser.givenName != FirstName)
            {
                arg1 = "    update FirstName " + "\"" + objUser.givenName + "\"" + " to " + "\"" + FirstName + "\"";
                strReport += arg1 + privateants.vbCrLf;
                objUser.givenName = FirstName;
                partialNeedUpdated = true;
            }
        }
        if (needUpdateMiddleName)
        {
            if (MiddleName == "")
            {
                if (objUser.initials != "")
                {
                    arg2 = "    clear initials " + "\"" + objUser.initials + "\"" + " to " + "\"" + MiddleName + "\"";
                    strReport += arg2 + privateants.vbCrLf;
                    objUser.PutEx(ADS_PROPERTY_CLEAR, "initials", privateants.vbNull);
                    partialNeedUpdated = true;
                }
                if (objUser.middleName != "")
                {
                    arg2 = "    clear MiddleName " + "\"" + objUser.middleName + "\"" + " to " + "\"" + MiddleName + "\"";
                    strReport += arg2 + privateants.vbCrLf;
                    objUser.PutEx(ADS_PROPERTY_CLEAR, "middleName", privateants.vbNull);
                    partialNeedUpdated = true;
                }
            }
            else if (objUser.initials != MiddleName)
            {
                // OrElse objUser.middleName <> MiddleName
                if (objUser.middleName != MiddleName)
                {
                    arg2 = "    update MiddleName " + "\"" + objUser.middleName + "\"" + " to " + "\"" + MiddleName + "\"";
                    strReport += arg2 + privateants.vbCrLf;
                    objUser.middleName = MiddleName;
                    partialNeedUpdated = true;
                }
                if (Strings.Len(MiddleName) > 6)
                {
                    if (objUser.initials != MiddleName.Substring(0, 1) + ".")
                    {
                        string newInitial = Strings.UCase(MiddleName.Substring(0, 1)) + ".";
                        arg2 = "    update initials " + "\"" + objUser.initials + "\"" + " to " + "\"" + newInitial + "\"";
                        strReport += arg2 + privateants.vbCrLf;
                        // MiddleName = newInitial
                        objUser.initials = newInitial;
                        partialNeedUpdated = true;
                    }
                }
                else
                {
                    arg2 = "    update initials " + "\"" + objUser.initials + "\"" + " to " + "\"" + MiddleName + "\"";
                    strReport += arg2 + privateants.vbCrLf;
                    objUser.initials = MiddleName;
                    partialNeedUpdated = true;
                }
            }
        }
        if (needUpdateLastName)
        {
            if (LastName != "" && objUser.sn != LastName && !hasPreferredLastnameEEIDs.Contains("." + employeeID + "."))
            {
                arg3 = "    update LastName " + "\"" + objUser.sn + "\"" + " to " + "\"" + LastName + "\"";
                strReport += arg3 + privateants.vbCrLf;
                objUser.sn = LastName;
                partialNeedUpdated = true;
            }
        }
        if (needUpdateLegalFirstName)
        {
            if (FirstName != "" && objUser.comment != FirstName)
            {
                arg4 = "    update comment(legal FirstName) " + "\"" + objUser.comment + "\"" + " to " + "\"" + FirstName + "\"";
                strReport += arg4 + privateants.vbCrLf;
                objUser.comment = FirstName;
                partialNeedUpdated = true;
            }
        }
        if (needUpdateDisplayName)
        {
            if (NickName != "")
                DisplayName = NickName + " " + LastName;
            else
                DisplayName = FirstName + " " + LastName;
            string dspName = objUser.DisplayName;
            if (dspName != DisplayName && (!dspName.Contains("(")) && !hasPreferredLastnameEEIDs.Contains("." + employeeID + "."))
            {
                arg5 = "    update DisplayName " + "\"" + objUser.DisplayName + "\"" + " to " + "\"" + DisplayName + "\"";
                strReport += arg5 + privateants.vbCrLf;
                objUser.DisplayName = DisplayName;
                partialNeedUpdated = true;
            }
        }

        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg3) != "")
                    tmpStr += Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg4) != "")
                    tmpStr += Strings.Trim(arg4.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg5) != "")
                    tmpStr += Strings.Trim(arg5.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>user name update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }


    arg1 = "";
    arg2 = "";
    arg3 = "";
    partialNeedUpdated = false;
    if (needUpdateJobTitle)
    {
        if (Jobtitle != "" && objUser.Title != Jobtitle)
        {
            // ENGR-288803 Delay Title Change for AD sync
            bool isMeetDelay = false;
            if (needUpdateJobTitleDelay)
            {
                if (jobTitleUpdateDelayLocationStr.Contains("." + LocationAux1 + "."))
                {
                    DateTime effDate = new DateTime(), jobTitleUpdateDelayEndDay = new DateTime();
                    try
                    {
                        effDate = JobEffDate;
                        jobTitleUpdateDelayEndDay = jobTitleUpdateDelayEndDayStr;
                    }
                    catch (Exception ex)
                    {
                    }
                    if (DateTime.Today >= jobTitleUpdateDelayEndDay || (DateTime.Today < jobTitleUpdateDelayEndDay && DateTime.Today >= effDate.AddDays(jobTitleUpdateDelayDays)))
                        isMeetDelay = true;
                }
                else
                    isMeetDelay = true;
            }
            if ((!needUpdateJobTitleDelay) || (needUpdateJobTitleDelay && isMeetDelay))
            {
                // if Today >= effDate then
                arg1 = "    update Title " + "\"" + objUser.Title + "\"" + " to " + "\"" + Jobtitle + "\"";
                strReport += arg1 + privateants.vbCrLf;
                objUser.Title = Jobtitle;
                partialNeedUpdated = true;
            }
        }
    }

    if (needUpdateHomeDeptment)
    {
        if (HomeDepartment != "" && objUser.Department != HomeDepartment)
        {
            arg2 = "    update Department " + "\"" + objUser.Department + "\"" + " to " + "\"" + HomeDepartment + "\"";
            strReport += arg2 + privateants.vbCrLf;
            objUser.Department = HomeDepartment;
            partialNeedUpdated = true;
        }
    }

    if (needUpdateUserSegment)
    {
        if (UserSegment != "" && objUser.employeeType != UserSegment)
        {
            arg3 = "    update employeeType " + "\"" + objUser.employeeType + "\"" + " to " + "\"" + UserSegment + "\"";
            strReport += arg3 + privateants.vbCrLf;
            objUser.employeeType = UserSegment;
            partialNeedUpdated = true;
        }
    }

    if (needUpdateJobTitle || needUpdateHomeDeptment || needUpdateUserSegment)
    {
        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg3) != "")
                    tmpStr += Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>JobTitle or HomeDepartment or UserSegment(employeeType) update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    // update ManagerDN should make sure the user is already existed in AD (if not, may need create first)
    if (needUpdateManager)
    {
        // there may be a lag, AD already set disabled with JIRA but HR system export file still show active
        // if HR system is On Leave or Active, do update
        if (Status != "Terminated")
        {
            // Manager record may be missing "today" due to WFN defect
            if (HR_eeID0.IndexOf(ManagerID) >= 0)
            {
                if (ManagerDN != "" && LCase(HR_FTE0(HR_eeID0.IndexOf(ManagerID))) != "non-employee" && !(objUser.AccountDisabled && Strings.UCase(DistinguishedNameFound).Contains("OU=ALUMNI")))
                {
                    string currentDN = objUser.manager;
                    string shortCurrentDN = "";
                    string keyManagerDN = "";
                    string shortManagerDN = "";
                    if (currentDN != "")
                        shortCurrentDN = Microsoft.VisualBasic.Mid(currentDN, 4, Strings.InStr(currentDN, ",") - 4);

                    keyManagerDN = Microsoft.VisualBasic.Left(ManagerDN, Strings.InStr(ManagerDN, "DC=") - 1);
                    shortManagerDN = Microsoft.VisualBasic.Mid(keyManagerDN, 4, Strings.InStr(keyManagerDN, ",") - 4);
                    if (currentDN != keyManagerDN + strDomain)
                    {
                        try
                        {
                            objUser.manager = keyManagerDN + strDomain;
                            tmpStr = "    update manager " + "\"" + shortCurrentDN + "\"" + " to " + "\"" + shortManagerDN + "\"";
                            strReport += tmpStr + privateants.vbCrLf;
                            hasUpdated = true;
                            if (!forTest)
                                objUser.SetInfo();
                            strGeneral += ",\"" + Strings.Trim(tmpStr.Replace("\"", "'")) + "\"";
                        }
                        catch (Exception ex)
                        {
                            intErrorCount = intErrorCount + 1;
                            tmpStr = "    =====>manager update failed!";
                            strReport += tmpStr + privateants.vbCrLf;
                        }
                    }
                }
            }
        }
    }

    arg1 = "";
    arg2 = "";
    arg3 = "";
    arg4 = "";
    arg5 = "";
    arg6 = "";
    partialNeedUpdated = false;
    if (needUpdateAddress)
    {
        tmpStr = objUser.streetAddress;
        if (Address != "" && tmpStr != Address)
        {
            arg1 = "    update street \"" + tmpStr + "\" to \"" + Address + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.streetAddress = Address;
            partialNeedUpdated = true;
        }

        if (PObox != "" && objUser.postOfficeBox != PObox)
        {
            arg2 = "    update P.O.Box " + "\"" + objUser.postOfficeBox + "\"" + " to " + "\"" + PObox + "\"";
            strReport += arg2 + privateants.vbCrLf;
            objUser.postOfficeBox = PObox;
            partialNeedUpdated = true;
        }
        if (City != "" && objUser.l != City)
        {
            arg3 = "    update City " + "\"" + objUser.l + "\"" + " to " + "\"" + City + "\"";
            strReport += arg3 + privateants.vbCrLf;
            objUser.l = City;
            partialNeedUpdated = true;
        }
        if (State != "" && objUser.st != State)
        {
            arg4 = "    update State " + "\"" + objUser.st + "\"" + " to " + "\"" + State + "\"";
            strReport += arg4 + privateants.vbCrLf;
            objUser.st = State;
            partialNeedUpdated = true;
        }
        if (PostalCode != "" && objUser.postalCode != PostalCode)
        {
            arg5 = "    update Zip Code " + "\"" + objUser.PostalCode + "\"" + " to " + "\"" + PostalCode + "\"";
            strReport += arg5 + privateants.vbCrLf;
            objUser.postalCode = PostalCode;
            partialNeedUpdated = true;
        }

        if (Country != "" && objUser.c != Country)
        {
            arg6 = "    update Country " + "\"" + objUser.c + "\"" + " to " + "\"" + Country + "\"";
            strReport += arg6 + privateants.vbCrLf;
            objUser.c = Country;
            partialNeedUpdated = true;
        }

        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg3) != "")
                    tmpStr += Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg4) != "")
                    tmpStr += Strings.Trim(arg4.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg5) != "")
                    tmpStr += Strings.Trim(arg5.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg6) != "" && tmpStr.Length < 200)
                    // DB field length is only 240
                    // if one transfer from CN to US, the address change string may be too long, so ignore recording the country change if it's too long
                    tmpStr += Strings.Trim(arg6.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>address, city, state, zip code or country update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    arg1 = "";
    arg2 = "";
    partialNeedUpdated = false;
    if (needUpdateDescription)
    {
        if (Description != "" && objUser.description != Description)
        {
            arg1 = "    update Description \"" + objUser.description + "\" to \"" + Description + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.description = Description;
            partialNeedUpdated = true;
        }
    }

    if (needUpdateOffice)
    {
        tmpStr = objUser.physicalDeliveryOfficeName;
        if (tmpStr == null)
            tmpStr = "";
        if (PhysicalDeliveryOfficeName != "" && tmpStr != PhysicalDeliveryOfficeName)
        {
            // do not update Xiamen Office A,B,C,D, "Xiamen 8B" "Xiamen 9B" "Xiamen 10B" "Xiamen 9A" 
            if (tmpStr == "" || PhysicalDeliveryOfficeName != "Xiamen" || (PhysicalDeliveryOfficeName == "Xiamen" && !(Strings.LCase(tmpStr).Contains("xiamen") || Strings.LCase(tmpStr).Contains("office "))))
            {
                arg2 = "    update Office " + "\"" + tmpStr + "\"" + " to " + "\"" + PhysicalDeliveryOfficeName + "\"";
                strReport += arg2 + privateants.vbCrLf;
                objUser.physicalDeliveryOfficeName = PhysicalDeliveryOfficeName;
                partialNeedUpdated = true;
            }
        }
    }

    if (needUpdateDescription || needUpdateOffice)
    {
        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>Description or Office update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    if (needUpdateMobile)
    {
        if (Mobile != "")
        {
            Mobile = Mobile.Replace(" ", "").Replace("(", "").Replace(")", "");
            if (objUser.Mobile != Mobile)
            {
                arg3 = "    update Mobile " + "\"" + objUser.Mobile + "\"" + " to " + "\"" + Mobile + "\"";
                strReport += arg3 + privateants.vbCrLf;
                objUser.Mobile = Mobile;

                hasUpdated = true;
                try
                {
                    if (!forTest)
                        objUser.SetInfo();
                    if (Strings.Trim(arg3) != "")
                        tmpStr = Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                    strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                }
                catch (Exception ex)
                {
                    intErrorCount = intErrorCount + 1;
                    tmpStr = "    =====>Mobile update failed!";
                    strReport += tmpStr + privateants.vbCrLf;
                }
            }
        }
    }

    arg1 = "";
    arg2 = "";
    arg3 = "";
    arg4 = "";
    arg5 = "";
    arg6 = "";
    arg7 = "";
    partialNeedUpdated = false;
    if (needUpdatePhone)
    {
        // only update if AD workphone is blank
        tmpStr = objUser.telephoneNumber;
        if (tmpStr == "" && Workphone != "" && tmpStr != Workphone)
        {
            arg1 = "    update Workphone " + "\"" + tmpStr + "\"" + " to " + "\"" + Workphone + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.telephoneNumber = Workphone;
            partialNeedUpdated = true;
        }
        tmpStr = objUser.homePhone;
        if (HomePhone != "" && tmpStr != HomePhone)
        {
            arg2 = "    update HomePhone " + "\"" + tmpStr + "\"" + " to " + "\"" + HomePhone + "\"";
            strReport += arg2 + privateants.vbCrLf;
            objUser.homePhone = HomePhone;
            partialNeedUpdated = true;
        }
        if (ipPhone != "" && objUser.ipPhone != ipPhone)
        {
            arg4 = "    update ipPhone " + "\"" + objUser.ipPhone + "\"" + " to " + "\"" + ipPhone + "\"";
            strReport += arg4 + privateants.vbCrLf;
            objUser.ipPhone = ipPhone;
            partialNeedUpdated = true;
        }

        if (otherTelephone != "")
        {
            object otherTel = objUser.otherTelephone;
            if (otherTel == null)
            {
                arg5 = "    update otherTelephone \"" + objUser.otherTelephone + "\" to \"" + otherTelephone + "\"";
                strReport += arg5 + privateants.vbCrLf;
                objUser.otherTelephone = otherTelephone;
                partialNeedUpdated = true;
            }
            else if (otherTel.GetType().ToString() == "System.String")
            {
                if (objUser.otherTelephone != otherTelephone)
                {
                    arg5 = "    update otherTelephone \"" + objUser.otherTelephone + "\" to \"" + otherTelephone + "\"";
                    strReport += arg5 + privateants.vbCrLf;
                    objUser.otherTelephone = otherTelephone;
                    partialNeedUpdated = true;
                }
            }
            else
            {
                tmpStr = "";
                for (int other = 0; other <= otherTel.length - 1; other++)
                    tmpStr += otherTel(other) + "; ";
                if (Strings.InStr(tmpStr, otherTelephone) < 1)
                {
                    arg5 = "    =====>has 2+ otherTelephone: (" + tmpStr + "). Do not update to \"" + otherTelephone + "\"";
                    strReport += arg5 + privateants.vbCrLf;
                    partialNeedUpdated = true;
                }
            }
        }
    }
    if (needUpdateFax)
    {
        // only update if AD workfax is blank
        tmpStr = objUser.facsimileTelephoneNumber;
        if (tmpStr == "" && Fax != "" && objUser.facsimileTelephoneNumber != Fax)
        {
            arg7 = "    update Fax " + "\"" + objUser.facsimileTelephoneNumber + "\"" + " to " + "\"" + Fax + "\"";
            strReport += arg7 + privateants.vbCrLf;
            objUser.facsimileTelephoneNumber = Fax;
            partialNeedUpdated = true;
        }
    }

    if (needUpdatePhone || needUpdateFax)
    {
        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg3) != "")
                    tmpStr += Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg4) != "")
                    tmpStr += Strings.Trim(arg4.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg5) != "")
                    tmpStr += Strings.Trim(arg5.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg6) != "")
                    tmpStr += Strings.Trim(arg6.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg7) != "")
                    tmpStr += Strings.Trim(arg7.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>Phone or Fax update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    arg1 = "";
    arg2 = "";
    arg3 = "";
    arg4 = "";
    partialNeedUpdated = false;
    if (needUpdateScriptPath)
    {
        if (ScriptPath != "" && LCase(objUser.scriptPath) != Strings.LCase(ScriptPath))
        {
            arg1 = "    NOT update Logon Script \"" + objUser.ScriptPath + "\" to \"" + ScriptPath + "\", please check and update manually!";
            strReport += arg1 + privateants.vbCrLf;
        }
        if (ProfilePath != "" && LCase(objUser.profilePath) != Strings.LCase(ProfilePath))
        {
            arg2 = "    NOT update Profile Path \"" + objUser.ProfilePath + "\" to \"" + ProfilePath + "\", please check and update manually!";
            strReport += arg2 + privateants.vbCrLf;
        }
    }

    if (needUpdateHomeDirectory)
    {
        if (HomeDirectory != "" && LCase(objUser.homeDirectory) != Strings.LCase(HomeDirectory))
        {
            arg3 = "    NOT update Home folder \"" + objUser.HomeDirectory + "\" to \"" + HomeDirectory + "\", please check and update manually!";
            strReport += arg3 + privateants.vbCrLf;
        }
        if (HomeDrive != "" && LCase(objUser.homeDrive) != Strings.LCase(HomeDrive))
        {
            arg4 = "    NOT update Home folder drive \"" + objUser.HomeDrive + "\" to \"" + HomeDrive + "\", please check and update manually!";
            strReport += arg4 + privateants.vbCrLf;
        }
    }

    if (needUpdateScriptPath || needUpdateHomeDirectory)
    {
        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg3) != "")
                    tmpStr += Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg4) != "")
                    tmpStr += Strings.Trim(arg4.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>ScriptPath, Profile Path, HomeDirectory or folder drive update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    if (CostCenter != "")
    {
        string shortCC = CostCenter;
        if (CostCenter.Contains(","))
        {
            shortCC = CostCenter.Substring(0, 1);
            shortCC += CostCenter.Substring(CostCenter.LastIndexOf(",") + 1, 3);
            shortCC = shortCC.Replace(",", "");
        }

        string DepartmentCode = shortCC;
        if (shortCC.Length > 3)
            DepartmentCode = CostCenter.Substring(shortCC.Length - 3, 3);

        arg1 = "";
        if (needUpdateDepartmentCode && objUser.DepartmentNumber != DepartmentCode)
        {
            arg1 = "    update DepartmentNumber " + "\"" + objUser.DepartmentNumber + "\"" + " to " + "\"" + DepartmentCode + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.DepartmentNumber = DepartmentCode;
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>DepartmentNumber(DepartmentCode) update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }

        arg1 = "";
        if (needUpdateCostcenterLocationStr.Contains("." + LocationAux1 + "."))
        {
            if (objUser.pager != shortCC)
            {
                arg1 = "    update Pager(CostCenter) " + "\"" + objUser.Pager + "\"" + " to " + "\"" + shortCC + "\"";
                strReport += arg1 + privateants.vbCrLf;
                objUser.pager = shortCC;
                hasUpdated = true;
                try
                {
                    if (!forTest)
                        objUser.SetInfo();
                    tmpStr = "";
                    if (Strings.Trim(arg1) != "")
                        tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                    strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
                }
                catch (Exception ex)
                {
                    intErrorCount = intErrorCount + 1;
                    tmpStr = "    =====>Pager(CostCenter) update failed!";
                    strReport += tmpStr + privateants.vbCrLf;
                }
            }
        }
    }

    arg1 = "";
    if (needUpdateDivisionCode && DivisionCode != "")
    {
        if (objUser.Division != DivisionCode)
        {
            arg1 = "    update Division " + "\"" + objUser.Division + "\"" + " to " + "\"" + DivisionCode + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.Division = DivisionCode;
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>Division(DivisionCode) update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }


    arg1 = "";
    if (needUpdateCompany && Status != "Terminated")
    {
        // 'CTPR-9: FTEs that come through WFN should be defaulted to “eHealth”
        // Non-FTEs such as temps and contractors that come through WFN should be defaulted to blank
        // All others that do not come through WFN such as vendors and the like should be blank. They will not come through WFN.
        // Also, in the event that we have part time employees and seasonal workers, they will come through WFN and also be defaulted to “eHealth.”
        // for NonFTE, should let company blank
        if (EmpClass.Contains("Regular"))
            Company = "eHealth";
        else
            Company = "";

        if (Company != "" && objUser.Company != Company)
        {
            arg1 = "    update Company " + "\"" + objUser.Company + "\"" + " to " + "\"" + Company + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.Company = Company;
            hasUpdated = true;
        }
        else if (Company == "" && objUser.Company == "eHealth")
        {
            arg1 = "    update Company " + "\"" + objUser.Company + "\"" + " to " + "\"" + Company + "\"";
            strReport += arg1 + privateants.vbCrLf;
            objUser.PutEx(ADS_PROPERTY_CLEAR, "Company", null);
            hasUpdated = true;
        }
        if (hasUpdated)
        {
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg1) != "")
                    tmpStr = Strings.Trim(arg1.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>company update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    arg2 = "";
    arg3 = "";
    arg4 = "";
    partialNeedUpdated = false;
    if (needUpdateOthers)
    {
        // may always not update AD email, since AD is correct, HR system may not be up to date
        // If Workemail <> "" Then
        // objUser.PutEx(ADS_PROPERTY_CLEAR, "mail", vbNull)
        // Else
        // objUser.put "mail", Workemail
        // End If
        if (wWWHomePage != "" && objUser.wWWHomePage != wWWHomePage)
        {
            arg2 = "    update Web page " + "\"" + objUser.wWWHomePage + "\"" + " to " + "\"" + wWWHomePage + "\"";
            strReport += arg2 + privateants.vbCrLf;
            objUser.wWWHomePage = wWWHomePage;
            partialNeedUpdated = true;
        }

        if (URL != "" && objUser.url != URL)
        {
            arg3 = "    update other Web page \"" + objUser.url + "\" to \"" + URL + "\"";
            strReport += arg3 + privateants.vbCrLf;
            objUser.url = URL; // should allow multi-lines
            partialNeedUpdated = true;
        }
    }

    if (needUpdateOthers)
    {
        if (partialNeedUpdated)
        {
            hasUpdated = true;
            try
            {
                if (!forTest)
                    objUser.SetInfo();
                tmpStr = "";
                if (Strings.Trim(arg2) != "")
                    tmpStr += Strings.Trim(arg2.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg3) != "")
                    tmpStr += Strings.Trim(arg3.Replace("\"", "'")) + "; ";
                if (Strings.Trim(arg4) != "")
                    tmpStr += Strings.Trim(arg4.Replace("\"", "'")) + "; ";
                strGeneral += ",\"" + Strings.Trim(tmpStr) + "\"";
            }
            catch (Exception ex)
            {
                intErrorCount = intErrorCount + 1;
                tmpStr = "    =====>other user properties update failed!";
                strReport += tmpStr + privateants.vbCrLf;
            }
        }
    }

    strReport += privateants.vbCrLf;
    return true;
}

private static void sendAlertEmail(string to_address, string CC_address, string subject, string body, string aFileName)
{
    if (!needSendAlertEmail)
        return;
    if (forTest)
    {
        body = "(To:" + to_address + "; CC:" + CC_address + ")</br></br>" + body;
        to_address = "Changtu.Wang@ehealth.com";
        CC_address = "";
    }

        private var from_address = "IT_automation@ehealth.com";
Console.WriteLine(DateTime.Now + " Sending email \"" + subject + "\"...");

MailMessage MailMsg = new MailMessage();
MailMsg.Subject = subject;  // strSubject
MailMsg.BodyEncoding = Encoding.Default;
MailMsg.IsBodyHtml = true;
MailMsg.Body = body; // strMessage
MailMsg.Priority = MailPriority.High; // Normal    Low

MailMsg.From = new MailAddress(from_address);

Strings.Replace(to_address, " ", "");
if (to_address != "")
{
    parseAdd(to_address, SendTo);
    foreach (string item in SendTo)
    {
        if (Strings.Trim(item) == null)
            break;
        MailMsg.To.Add(new MailAddress(item));
    }
}

Strings.Replace(CC_address, " ", "");
if (CC_address != "")
{
    parseAdd(CC_address, ccTo);
    foreach (string item in ccTo)
    {
        if (Strings.Trim(item) == null)
            break;
        MailMsg.CC.Add(new MailAddress(item));
    }
}

// MailMsg.Bcc.Add(New MailAddress("changtu.wang@ehealth.com"))
MailMsg.ReplyTo = new MailAddress(entOpsEmail);
// MailMsg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess 'OnFailure

Strings.Replace(aFileName, " ", "");
if (aFileName != "")
{
    string[] filenames = new string[11];
    parseFilenames(aFileName, filenames);
    foreach (string item in filenames)
    {
        if (Strings.Trim(item) == null)
            break;
        try
        {
            Attachment MsgAttach = new Attachment(item);
            MailMsg.Attachments.Add(MsgAttach);
        }
        catch (Exception ex)
        {
        }
    }
}

SmtpClient client = new SmtpClient();
client.Host = "awengsmtp01.ehealthinsurance.com"; // "sjengsmtp01.ehealthinsurance.com" '"outlook.ehealthinsurance.com"
                                                  // SmtpMail.Port = 25
                                                  // SmtpMail.Timeout = 
try
{
    client.Send(MailMsg);
    Console.WriteLine(DateTime.Now + " Sent email \"" + subject + "\"");
    if (!swLog == null)
        swLog.WriteLine(DateTime.Now + " Sent email \"" + subject + "\"");
}
catch (Exception ex)
{
    Console.WriteLine(DateTime.Now + " Sending email failed \"" + subject + "\"");
    if (!swLog == null)
        swLog.WriteLine(DateTime.Now + " Sending email failed \"" + subject + "\"");
}
return;
    }

    private static void GetUserMemberOf(string TargetUsername)
{
    memberOfList.Clear();
    memberOf = "";
    if (TargetUsername == "")
        return;
    DirectoryEntry entry = new DirectoryEntry("LDAP://" + strDomain);
    DirectorySearcher Searcher = new DirectorySearcher(entry);
    Searcher.Filter = "(&(samAccountName=" + TargetUsername + "))";
    // Searcher.PropertiesToLoad.Add("Description")
    // Searcher.PropertiesToLoad.Add("otherTelephone")
    Searcher.PropertiesToLoad.Add("memberOf");
    // Searcher.PropertiesToLoad.Add("accountExpires") 'don't use AccountExpirationDate
    // Searcher.PropertiesToLoad.Add("msRTCSIP-UserEnabled")
    SearchResult UserAccount = Searcher.FindOne;
    // tmpStr = UserAccount.Properties("adspath")(0)
    // Dim theUser As DictionaryEntry = New DirectoryEntry(tmpStr)
    // Dim RawExpires As Int64 = 0
    // Dim ExpirationDate As DateTime
    if (UserAccount == null)
        return;
    for (var j = 0; j <= UserAccount.Properties("memberOf").Count - 1; j++)
    {
        tmpStr = UserAccount.Properties("memberOf")(j);
        memberOfList.Add(tmpStr);
        tmpStr = Microsoft.VisualBasic.Left(tmpStr, Strings.InStr(tmpStr, ",") - 1);
        tmpStr = Strings.Replace(tmpStr, "CN=", "");
        memberOf += tmpStr + "; ";
    }
}

private static string GetUserProperty(string TargetUsername, string fieldname)
{
    DirectoryEntry entry = new DirectoryEntry("LDAP://" + strDomain);
    DirectorySearcher Searcher = new DirectorySearcher(entry);
    Searcher.Filter = "(&(samAccountName=" + TargetUsername + "))";
    Searcher.PropertiesToLoad.Add(fieldname);
    SearchResult UserAccount = Searcher.FindOne;
    if (UserAccount == null)
        return "";
    if (UserAccount.Properties(fieldname).Count > 0)
        return UserAccount.Properties(fieldname)(0);
    return "";
}

private static void parseAdd(string Recipients, string[] strTo)
{
    // clear the array first
    for (int s = 0; s <= strTo.Length - 1; s++)
        strTo[s] = "";

    Recipients.replace(",", ";");
    int sendToIndex = 0;

    while (Strings.InStr(Recipients, ";") > 0)
    {
        tmpStr = (Microsoft.VisualBasic.Left(Recipients, Strings.InStr(Recipients, ";") - 1));
        Recipients = (Microsoft.VisualBasic.Right(Recipients, Strings.Len(Recipients) - Strings.InStr(Recipients, ";")));
        if (Strings.InStr(tmpStr, "@") > 0)
        {
            strTo[sendToIndex] = Strings.Trim(tmpStr);
            sendToIndex += 1;
            if (sendToIndex >= maxAllowedRecipient)
                break;
        }
        else
            swLog.WriteLine("====>Invalid email address: " + tmpStr);
    }
    if (sendToIndex >= maxAllowedRecipient)
        return;
    if (Strings.Len(Recipients) > 0)
    {
        if (Strings.InStr(Recipients, "@") > 0)
        {
            strTo[sendToIndex] = Strings.Trim(Recipients);
            Recipients = Strings.Trim(Strings.Replace(Recipients, strTo[sendToIndex], ""));
            sendToIndex += 1;
        }
        else
            swLog.WriteLine("====>Invalid email address: " + Recipients);
    }
}

private static void parseFilenames(string fname, string[] fnames)
{
    // clear the array first
    for (int s = 0; s <= fnames.Length - 1; s++)
        fnames[s] = "";

    int index = 0;
    while (Strings.InStr(fname, ",") > 0)
    {
        tmpStr = (Microsoft.VisualBasic.Left(fname, Strings.InStr(fname, ",") - 1));
        fname = (Microsoft.VisualBasic.Right(fname, Strings.Len(fname) - Strings.InStr(fname, ",")));
        fname = tmpStr + ";" + fname;
    }

    while (Strings.InStr(fname, ";") > 0)
    {
        tmpStr = (Microsoft.VisualBasic.Left(fname, Strings.InStr(fname, ";") - 1));
        fname = (Microsoft.VisualBasic.Right(fname, Strings.Len(fname) - Strings.InStr(fname, ";")));
        fnames[index] = Strings.Trim(tmpStr);
        index += 1;
        if (index >= maxAllowedAttachment)
            break;
    }
    if (index >= maxAllowedAttachment)
        return;
    if (Strings.Len(fname) > 0)
    {
        fnames[index] = Strings.Trim(fname);
        fname = Strings.Trim(Strings.Replace(fname, fnames[index], ""));
        index += 1;
    }
}

private static void readCSV2table(string csvFileName, DataTable DataTbl)
{
    DataTbl.Rows.Clear();
    DataTbl.Columns.Clear();

    hasError = false;

    object input = null;
    try
    {
        input = My.Computer.FileSystem.OpenTextFieldParser(csvFileName);
    }
    catch (Exception ex)
    {
        swLog.WriteLine("readCSV2table failed! " + csvFileName + ex.ToString());
    }
    if (input == null)
        return;

    input.SetDelimiters(",");
    string[] fieldStrs = input.ReadFields();
    if (fieldStrs == null)
    {
        input.Close();
        return;
    }
    DataTbl.Columns.Add("row#");
    foreach (string aTitle in fieldStrs)
    {
        try
        {
            DataTbl.Columns.Add(aTitle);
        }
        catch (Exception ex)
        {
            hasError = true;
            swLog.WriteLine("Add table column \"" + aTitle + "\" failed! Please check column names is all valid in " + privateants.vbCrLf + csvFileName);
            break;
        }
    }
    if (hasError)
    {
        input.Close();
        return;
    }
    int rowCount = 0;
    while ((!input.EndOfData))
    {
        try
        {
            string[] row = new string[DataTbl.Columns.Count - 1 + 1];
            rowCount += 1;
            row[0] = rowCount;
            fieldStrs = input.ReadFields;
            for (int i = 0; i <= fieldStrs.Length - 1; i++)
                row[i + 1] = fieldStrs[i];
            DataTbl.Rows.Add(row);
        }
        catch (Exception ex)
        {
            swLog.WriteLine("line " + rowCount + " in " + csvFileName + " has problem.");
        }
    }
    input.Close();
}

private static bool checkSQLDB()
{
    if (forTest || !Strings.UCase(strDomain).Contains("DC=EHI,DC=EHEALTH,DC=COM"))
    {
        if (!reportTalbeName.Contains("Test"))
            reportTalbeName += "Test";
        if (!onLeaveTalbeName.Contains("Test"))
            onLeaveTalbeName += "Test";
        if (!disabledTalbeName.Contains("Test"))
            disabledTalbeName += "Test";
    }

    // don't check again and again on PROD DB
    return true;

    string sql;
    sql = "create table " + reportTalbeName + " (AutomateTime datetime not null DEFAULT (getdate()), Tool char(32), employeeID char(12) null, samAccountName char(22), firstName char(25) not null, lastName char(25) not null";
    sql += ", status char(12) not null, department char(60) null, jobTitle char(70) null, note1 char(240) null, note2 char(240) null, note3 char(240) null, note4 char(240) null, note5 char(240) null";
    sql += ", note6 char(240) null, note7 char(240) null, note8 char(240) null, note9 char(240) null, note10 char(240) null, note11 char(240) null, note12 char(240) null )";

    sql = "create table " + onLeaveTalbeName + " (AutomateTime datetime not null DEFAULT (getdate()), Tool char(32)";
    sql += ", eeID char(12) null, samID char(22), firstName char(25) not null, lastName char(25) not null";
    sql += ", status char(12) not null, statusEffDate datetime null";
    sql += ", ignoredCount numeric(8) not null DEFAULT (0), location char(30) null, Notes char(30) null )";

    sql = "create table " + contactITtalbeName + " (LocationAux1 char(6) not null, ITcontactEmail char(60) not null";
    sql += ", updateWhen datetime not null DEFAULT (getdate() ))";

    sql = "create table " + disabledTalbeName + " (AutomateTime datetime not null DEFAULT (getdate()), Tool char(32)";
    sql += ", eeID char(12) null, samID char(22), firstName char(25) not null, lastName char(25) not null";
    sql += ", disabledTime datetime not null DEFAULT (getdate())";
    sql += ", DN char(240) null, Notes char(30) null )";
}

private static bool copyCSVtoSQLDB(string CSVfileName)
{
    intOKcount = 0;
    try
    {
        readCSV2table(CSVfileName, tableTMP);
    }
    catch (Exception ex)
    {
        return false;
    }
    if (hasError)
        return false;

    hasError = false;
    ClassExecuteSQL executeSQL = new ClassExecuteSQL();

    Console.WriteLine("Recording report in SQL DB, please wait...");
    string strSql;
    for (int i = 0; i <= tableTMP.Rows.Count - 1; i++)
    {
        strSql = "insert into " + reportTalbeName + " (AutomateTime,Tool,employeeID,samAccountName,firstName,lastName,status, department,jobTitle";
        strSql += ",note1,note2,note3,note4,note5,note6,note7,note8,note9,note10,note11,note12) values (";
        DataRow drow;
        drow = tableTMP.Rows(i);
        for (int j = 1; j <= tableTMP.Columns.Count - 1; j++) // 1st is row#, ignore it
        {
            if (IsDBNull(drow(j)))
                strSql += "'',";
            else if (drow(j) == "")
                strSql += "'',";
            else
                strSql += "'" + drow(j).Replace("'", "''") + "',";
        }
        strSql = Microsoft.VisualBasic.Left(strSql, strSql.Length - 1) + ")";

        executeSQL.executeSQL(strSql, "insert");
        if (executeSQL.errStr != "")
        {
            hasError = true;
            errStr = executeSQL.errStr;
            Console.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        }
        intOKcount = executeSQL.intCount;
    }
    return !hasError;
}

private static void pickNewUser()
{
    return;
    // no longer use this method, it's out of date. If need use again, need update as new columns in current reportfileNotFound

    reportfileNotFound = new StreamWriter(strExportFileNewUser, false);
    reportfileNotFound.WriteLine("employeeID,samAccountName,FirstName,MiddleName,LastName,NickName,description,JobTitle,HomeDepartment,Status,statusEffDate,EmpClass,HireDate,managerID,ManagerDN,ManagerWorkEmail,LocationAux1,container,Address,City,State,PostalCode,Country,Office,CostCenter,NeedAccount,profilePath,scriptPath,homeDirectory,homeDrive,MOC_enabled,NeedMailbox,defaultGroup,workEmail,EmployeeID_mismatch,Notes");

    DataRow drow;
    string NeedMailbox;
    hasError = false;
    intNotFoundCount = 0;
    for (int i = 0; i <= tableTMP.Rows.Count - 1; i++)
    {
        drow = tableTMP.Rows(i);
        string eeIDmismatch = drow(52);
        // CTPR-65 Charles comment: EmpClass COBRA-Inactive do not have an AD account created
        // if eeID mismatch found between AD and HR system, should not report new hire
        if (drow(33) != "Active" || drow(46) == noNeedAccountValueHR || drow(35) == "COBRA-Inactive" || eeIDmismatch.Contains("(AD)<===>(" + hrMIS + ")"))
            continue;
        if (drow(43) == "RMT")
            // usually remote users do not require a mailbox
            NeedMailbox = "No";
        else
            NeedMailbox = "Yes";

        tmpStr = drow(56); // notes
        if (tmpStr.Contains("Invalid:"))
        {
            tmpStr = tmpStr.Replace("Invalid:", "");
            tmpStr = tmpStr.Replace("?", "");
        }
        else
        {
            // tmpStr = "already has existing account but different EED?"
            isXM = false;
            if (drow(43) == "XM")
                isXM = true;
            // This will always get new samAccountName
            tmpStr = newSAMID(drow(26), drow(28), isXM);
            if (tmpStr.Length > 20)
                tmpStr = newSAMID(drow(26), Microsoft.VisualBasic.Left(drow(28), drow(28).Length - 1), isXM);
            if (samAccountName.Length > 20)
                tmpStr = newSAMID(drow(26), Microsoft.VisualBasic.Left(drow(28), drow(28).Length - 2), isXM);
            // if still >20
            if (tmpStr.Length > 20)
                tmpStr = Microsoft.VisualBasic.Left(tmpStr, 18) + "zz";
        }
        samAccountName = Strings.Trim(tmpStr);
        lineStr = "\"" + drow(2) + "\","; // employeeID
        lineStr += "\"" + samAccountName + "\","; // samAccountName
        lineStr += "\"" + drow(26) + "\","; // firstName
        lineStr += "\"" + drow(27) + "\","; // MiddleName
        lineStr += "\"" + drow(28) + "\","; // lastName
        lineStr += "\"" + drow(29) + "\","; // nickName
        lineStr += "\"" + "\","; // description
        lineStr += "\"" + drow(30) + "\","; // jobTitle
        lineStr += "\"" + drow(32) + "\","; // Department
        lineStr += "\"" + drow(33) + "\","; // status
        lineStr += "\"" + drow(34) + "\","; // statusEffDate
        lineStr += "\"" + drow(35) + "\","; // EmpClass
        lineStr += "\"" + drow(37) + "\","; // HireDate
        lineStr += "\"" + drow(20) + "\","; // managerID
        lineStr += "\"" + drow(21) + "\","; // managerDN
        lineStr += "\"" + drow(25) + "\","; // managerWorkemail
        LocationAux1 = drow(43); // LocationAux1 is used in set_ScriptPath_HomeDirectory()
        lineStr += "\"" + LocationAux1 + "\","; // locationAUX1
        lineStr += "\"" + drow(57) + "\","; // container
        lineStr += "\"" + drow(38) + "\","; // address
        lineStr += "\"" + drow(39) + "\","; // city
        lineStr += "\"" + drow(40) + "\","; // state
        lineStr += "\"" + drow(41) + "\","; // postalCode
        lineStr += "\"" + drow(42) + "\","; // country
        lineStr += "\"" + drow(44) + "\","; // Office
        lineStr += "\"" + drow(45) + "\","; // CostCenter
        lineStr += "\"" + drow(46) + "\","; // needAccount
        lineStr += "\"" + drow(58) + "\","; // profilePath
        lineStr += "\"" + drow(59) + "\","; // scriptPath
        HomeDirectory = drow(60); // homeDirectory
        if (HomeDirectory == "")
            set_ScriptPath_HomeDirectory(samAccountName);
        lineStr += "\"" + HomeDirectory + "\",";

        lineStr += "\"" + drow(61) + "\","; // homeDrive
        lineStr += "\"" + drow(62) + "\","; // MOC_enabled
        lineStr += "\"" + NeedMailbox + "\","; // NeedMailbox
        lineStr += "\"" + drow(54) + "\","; // defaultGroup
        lineStr += "\"" + drow(63) + "\","; // workEmail
        lineStr += "\"" + drow(52) + "\","; // EmployeeID_mismatch
        tmpStr = drow(56); // notes
        tmpStr = tmpStr.Replace("Invalid: " + samAccountName + " ?", "");
        lineStr += "\"" + tmpStr + "\""; // remain portion of notes
        reportfileNotFound.WriteLine(lineStr);
        intNotFoundCount += 1;
    }
    reportfileNotFound.Close();
}

private static string newSAMID(string aFirstName, string aLastName, bool isXM)
{
    // First initial, last name (jgarcia)
    // First two initials, last name (jugarcia)
    // NOT First initial, middle initial, last name (jmgarcia)
    // NOT First initial, last name, number (jgarcia2)

    string tmpSAMID, DNfound;
    aFirstName = aFirstName.Replace(" ", "");
    aLastName = aLastName.Replace(" ", "");
    aFirstName = aFirstName.Replace("'", "");
    aLastName = aLastName.Replace("'", ""); // James O'Brien 101227
                                            // if contain "-" ...
                                            // From: Philip Zinn Sent: Tuesday, August 26, 2014 7:44 AM  RE: non-standard AD account and email address
                                            // “when an AD user account has a hyphen in it, Back Office accounts fail during creation”
                                            // so we should always remove "-"
                                            // Charles email 30 June, 2015 07:31 : Let’s go with no hyphen for the email.  If an employee asks for a hyphen, let’s add that to the alias email address.  
    aFirstName = aFirstName.Replace("-", "");
    aLastName = aLastName.Replace("-", "");

    if (isXM)
        tmpSAMID = Strings.LCase(aFirstName) + Strings.LCase(Microsoft.VisualBasic.Left(aLastName, 1));
    else
        tmpSAMID = Strings.LCase(Microsoft.VisualBasic.Left(aFirstName, 1)) + Strings.LCase(aLastName);

    if (tmpSAMID.Contains("-"))
    {
        if (tmpSAMID.Length > 20)
            tmpSAMID = tmpSAMID.Replace("-", "");
    }
    if (tmpSAMID.Length > 20)
        tmpSAMID = Microsoft.VisualBasic.Left(tmpSAMID, 20);

    int tryTime = 1;
    do
    {
        DNfound = GetUserProperty(tmpSAMID, "DistinguishedName");
        // here should not use findBySamAccountName(tmpSAMID, "User"), because public DistinguishedNameFound, employeeIDfound will affects other process
        if (DNfound == "")
            break;
        else
        {
            tryTime = tryTime + 1;
            objUser = Interaction.GetObject("LDAP://" + DNfound);
            if (Strings.LCase(aFirstName) == LCase(objUser.givenName) && Strings.LCase(aLastName) == LCase(objUser.sn))
                // tmpStr = "    " & tmpSAMID & " already used for " & objUser.employeeID & "(" & objUser.givenName & " " & objUser.sn & "), please double check!"
                return "";
            if (isXM)
            {
                if (tryTime > Strings.Len(aLastName))
                    tmpSAMID = Strings.LCase(aFirstName) + Strings.LCase(aLastName) + tryTime;
                else
                    tmpSAMID = Strings.LCase(aFirstName) + Strings.LCase(Microsoft.VisualBasic.Left(aLastName, tryTime));
            }
            else if (tryTime > Strings.Len(aFirstName))
                tmpSAMID = Strings.LCase(aFirstName) + Strings.LCase(aLastName) + tryTime;
            else
                tmpSAMID = Strings.LCase(Microsoft.VisualBasic.Left(aFirstName, tryTime)) + Strings.LCase(aLastName);
        }
    }
    while (DNfound != "");
    return tmpSAMID;
}

private static bool getOnLeaveEffDateFromDB()
{
    DB_OnLeaveEEID.Clear();
    DB_OnLeaveSamID.Clear();
    DB_OnLeaveEffDate.Clear();

    DB_LeaveReturningSamID.Clear();
    DB_LeaveReturningEffDate.Clear();

    string strSql;
    strSql = "select eeID,samID,statusEffDate,status from " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";
    strSql += "where eeID is not null and samID is not null and statusEffDate is not null";
    strSql += " and (status='On Leave' or status='returning')";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "select"))
    {
        tmpStr = "MS SQL database has problem, please check: " + executeSQL.errStr.Replace("<br/>", privateants.vbCrLf);
        swLog.WriteLine(tmpStr);
        return false;
    }

    if (executeSQL.tableResult.Rows.Count > 0)
    {
        for (int db = 0; db <= executeSQL.tableResult.Rows.Count - 1; db++)
        {
            DataRow drow;
            drow = executeSQL.tableResult.Rows(db);
            if (Trim(drow(3)) == "On Leave")
            {
                DB_OnLeaveEEID.Add(Trim(drow(0)));
                DB_OnLeaveSamID.Add(LCase(Trim(drow(1))));
                DB_OnLeaveEffDate.Add(Trim(drow(2)));
            }
            else
            {
                DB_LeaveReturningSamID.Add(LCase(Trim(drow(1))));
                DB_LeaveReturningEffDate.Add(Trim(drow(2)));
            }
        }
    }
    return true;
}

private static bool deleteLeaveReturning(string aSamID)
{
    string strSql;
    strSql = "delete from " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";
    strSql += "where samid='" + aSamID + "'";
    // strSql &= " and status='returning'"
    if (forXM_US != "")
    {
        if (forXM_US == "XM")
            strSql += " and (location='XM' or location='XIAMEN')";
        else
            strSql += " and location<>'XM' and location<>'XIAMEN'";
    }

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "delete"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return false;
    }
    return true;
}

private static void deleteNonOnLeave()
{
    if (onLeaveSamIDs != "")
        onLeaveSamIDs = onLeaveSamIDs.Substring(0, Strings.Len(onLeaveSamIDs) - 1);

    string strSql;
    strSql = "delete from " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";

    if (onLeaveSamIDs != "")
    {
        strSql += "where samid not in (" + onLeaveSamIDs + ")";
        strSql += " and status='On Leave'";
    }
    if (forXM_US != "")
    {
        if (forXM_US == "XM")
            strSql += " and (location='XM' or location='XIAMEN')";
        else
            strSql += " and location<>'XM' and location<>'XIAMEN'";
    }

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "delete"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return;
    }
    return;
}

private static bool hasOnLeaveRecord(string aSamID)
{
    // if already exists in DB, just return true, but new insert records will not be reflected in the array, so still select from DB
    int OnLeaveIndex = DB_OnLeaveSamID.IndexOf(Strings.LCase(aSamID));
    if (OnLeaveIndex >= 0)
    {
        statusEffDateStr = DB_OnLeaveEffDate(OnLeaveIndex);
        return true;
    }
    return false;

    bool hasOnLeave = false;
    string strSql;
    strSql = "select samID,statusEffDate from " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";
    strSql += "where samid = '" + aSamID + "'";
    strSql += " and statusEffDate is not null";
    strSql += " and status='On Leave'";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "select"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return hasOnLeave;
    }

    if (executeSQL.tableResult.Rows.Count > 0)
    {
        hasOnLeave = true;
        DataRow drow;
        drow = executeSQL.tableResult.Rows(0);
        statusEffDateStr = Trim(drow(1));
    }
    return hasOnLeave;
}

private static bool insertOnLeaveRecord(string aSamID)
{
    if (statusEffDateStr == "")
        statusEffDateStr = DateTime.Now.ToShortDateString();
    string strSql;
    strSql = "insert into " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";
    strSql += "(Tool,eeID,samID,firstName,lastName,status,statusEffDate,location,Notes)";
    strSql += " values (";
    strSql += "'" + appNameVersion + " " + runnerName + "','" + employeeID + "','" + aSamID + "','" + FirstName.Replace("'", "''") + "','" + LastName.Replace("'", "''") + "',";
    strSql += "'" + Status + "','" + statusEffDateStr + "','" + LocationAux1 + "','')";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "insert"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return false;
    }
    return true;
}

private static bool isNeedDisableOnLeave(string aSamID)
{
    int ignoredCount = 0;

    bool needDisable = false;
    string strSql;
    strSql = "select ignoredCount from " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";
    strSql += "where samid = '" + aSamID + "'";
    strSql += " and status='On Leave'";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "select"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return needDisable;
    }

    if (executeSQL.tableResult.Rows.Count > 0)
    {
        DataRow drow;
        drow = executeSQL.tableResult.Rows(0);
        if (!IsDBNull(drow(0)))
            ignoredCount = drow(0);

        if (ignoredCount + 1 > OnLeaveIgnoreTimesAllowed)
            needDisable = true;

        strSql = "update " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] set ";
        strSql += "ignoredCount = ignoredCount + 1";
        strSql += ",Notes='" + DateTime.Now + "'";
        strSql += "where samid = '" + aSamID + "'";
        strSql += " and status='On Leave'";

        if (!executeSQL.executeSQL(strSql, "update"))
        {
            hasError = true;
            swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        }
    }
    else if (Status.Contains("Leave"))
    {
        needDisable = true;
        strSql = "insert into " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] ";
        strSql += "(Tool,eeID,samID,firstName,lastName,status, Notes)";
        strSql += " values (";
        strSql += "'" + appNameVersion + " " + runnerName + "','" + employeeID + "','" + samAccountName + "','" + FirstName.Replace("'", "''") + "','" + LastName.Replace("'", "''") + "',";
        strSql += "'" + Status + "','')";

        if (!executeSQL.executeSQL(strSql, "insert"))
        {
            hasError = true;
            swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        }
    }

    return needDisable;
}

private static bool changeStatusLeaveReturning(string aSamID)
{
    string strSql = "update " + "[" + ITDBname + "].[dbo].[" + onLeaveTalbeName + "] set ";
    strSql += "status='returning',statusEffDate='" + DateTime.Now.ToShortDateString() + "'";
    strSql += "where samid = '" + aSamID + "'";
    strSql += " and status='On Leave'";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(strSql, "update"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return false;
    }
    return true;
}

private static bool hasLeaveReturningRecord(string aSamID)
{
    int LeaveReturningIndex = DB_LeaveReturningSamID.IndexOf(Strings.LCase(aSamID));
    if (LeaveReturningIndex >= 0)
    {
        leaveReturningEffDateStr = DB_LeaveReturningEffDate(LeaveReturningIndex);
        return true;
    }
    return false;
}

private static bool isNeedDisableOnLeaveByEffDate()
{
    bool needDisable = true;
    DateTime effDate = new DateTime();
    try
    {
        effDate = statusEffDateStr;
        if (DateTime.Today >= effDate.AddDays(OnLeaveCheckDays))
        {
            // If change the condition as effDate > Today ? , then if HR system not set status back to Active in time,
            // the AD account already enabled by location IT guys will be set disabled again and again
            DateTime theHireDate = HireDate;
            if (theHireDate < effDate)
                needDisable = false;
        }
    }
    catch (Exception ex)
    {
    }
    return needDisable;
}

private static bool isNeedDisableTerminateRightNow()
{
    // Return False 'stop disalbe after Workday + Okta integration

    if (!needDisableAtEndOfLastWorkingDay)
        return true;

    bool needDisable = true;
    DateTime effDate = new DateTime(), nowTime = new DateTime();
    try
    {
        effDate = statusEffDateStr;
        effDate = effDate.AddHours(17); // 17: 5:00 PM
        if (LocationAux1 == "XM")
            effDate = effDate.AddHours(-timeDifferenceXM);
        else if (LocationAux1 == "MA")
            effDate = effDate.AddHours(-timeDifferenceMA);
        else if (LocationAux1 == "UT")
            effDate = effDate.AddHours(-timeDifferenceUT);
        nowTime = DateTime.Now;
        if (forTest)
            nowTime = DateTime.Now.AddHours(-timeDifferenceXM);// always test in XM
        if (nowTime < effDate)
            needDisable = false;
    }
    catch (Exception ex)
    {
    }
    return needDisable;
}

private static int crossCheckHR()
{
    correctCount = 0;
    eeIDs.Clear();
    wkdID = "";
    for (int n = 0; n <= eeIDlist.Count - 1; n++)
    {
        employeeID = eeIDlist.Item(n);
        reqidAD = reqIDlist.Item(n);
        reqidAD = reqidAD.Replace(" ", ""); // reqidHR already Replace(" ", "")
        if (reqidAD != "")
        {
            while ((reqidAD.Length < 6))
                reqidAD = "0" + reqidAD;
        }
        if (employeeID == "")
            index = -1;
        else
            index = HR_eeID.IndexOf(employeeID);

        if (employeeID == "")
        {
            string reqidHR = "";
            int indexByEmail = -1;
            if (WorkemailListKey.Item(n) != "")
                indexByEmail = HR_workEmailKey.IndexOf(WorkemailListKey.Item(n));
            DateTime createDate = new DateTime();
            try
            {
                createDate = whenCreatedList(n);
            }
            catch (Exception ex)
            {
            }

            if (reqidAD != "")
            {
                // Charles Escoto email Sent: 29 April, 2016 15:19
                // If there is a Req ID but no EEID in AD, then trigger discrepancy “AD record needs to be updated with EEID”

                index = HR_reqID.IndexOf(reqidAD);
                if (index < 0)
                {
                    // HR system may has record for the newhire but missing reqID, try find reqID by workemail
                    if (indexByEmail < 0)
                    {
                        // HR system still has no record for the newhire, but AD already created account with reqID
                        if (createDate.AddDays(noShow_ThresholdDays) < DateTime.Today)
                        {
                            // may be "no show", HR record never created or already deleted
                            // readHireDate
                            objUser = GetObject("LDAP://" + userDNlist(n));
                            if (objUser.extensionAttribute9 == "")
                            {
                                bool needReport = false;
                                string Ulocation = UCase(cityList.Item(n));
                                if (forXM_US == "US")
                                {
                                    if (!Ulocation.Contains("XIAMEN"))
                                    {
                                        needReport = true;
                                        discrepancyCount += 1;
                                    }
                                }
                                else if (forXM_US == "XM")
                                {
                                    if (Ulocation.Contains("XIAMEN"))
                                    {
                                        needReport = true;
                                        discrepancyCount += 1;
                                        discrepancyCountXM += 1;
                                    }
                                }
                                else
                                {
                                    needReport = true;
                                    discrepancyCount += 1;
                                    if (Ulocation.Contains("XIAMEN"))
                                        discrepancyCountXM += 1;
                                }
                                if (needReport)
                                    reportfileDiscrepancy.WriteLine(",reqID:" + reqidAD + ",\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",No " + hrMIS + " record,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " is a no show?");
                            }
                            else
                            {
                            }
                        }
                    }
                    else
                    {
                        reqidHR = HR_reqID.Item(indexByEmail);
                        employeeID = HR_eeID.Item(indexByEmail);
                        if (reqidAD == employeeID)
                        {
                            // EEID set into reqID correcting
                            if (!needStop)
                            {
                                try
                                {
                                    objUser = GetObject("LDAP://" + userDNlist(n));
                                    string oldValue = objUser.employeeID;
                                    objUser.employeeID = employeeID;
                                    if (needClearEmployeeNumber)
                                        objUser.PutEx(ADS_PROPERTY_CLEAR, "employeeNumber", privateants.vbNull);
                                    if (!forTest)
                                        objUser.SetInfo();
                                    correctCount += 1;
                                    swReport.WriteLine("correcting #" + correctCount + ": " + samIDlist.Item(n) + "(" + HR_nickName.Item(indexByEmail) + " " + HR_firstName.Item(indexByEmail) + " " + HR_lastName.Item(indexByEmail) + "): " + "'EEID set into reqID' problem correcting: ");
                                    swReport.WriteLine("               update employeeID " + "\"" + oldValue + "\"" + " to " + "\"" + employeeID + "\"");
                                    if (needClearEmployeeNumber)
                                        swReport.WriteLine("               clear employeeNumber(reqID) " + "\"" + reqidAD + "\"" + " to " + "\"\"");
                                    swReport.WriteLine("");
                                }
                                catch (Exception ex)
                                {
                                    swLog.WriteLine("    =====>EEID set into reqID problem correcting failed: " + samIDlist.Item(n) + " " + userDNlist(n));
                                    // only report this discrepancy if failed to correct it automatically
                                    reportfileDiscrepancy.WriteLine(employeeID + ", (reqID AD:" + reqidAD + "<==>" + hrMIS + ":" + reqidHR + "),\"" + HR_firstName.Item(indexByEmail) + "\",\"" + HR_lastName.Item(indexByEmail) + "\",\"" + HR_nickName.Item(indexByEmail) + "\",\"" + HR_locationAux1.Item(indexByEmail) + "\"," + statusList.Item(n) + ",AD EEID set into reqID,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need update EEID");
                                    addDiscrepancyCount(HR_locationAux1.Item(indexByEmail));
                                }
                            }
                        }
                        else
                        {
                            // if HR system record found, should display with HR firstname and nickname
                            if (reqidHR == "")
                                // reportfileDiscrepancy.WriteLine(employeeID & ", (reqID AD:" & reqidAD & "<==>" & hrMIS & ":" & reqidHR & "),""" & HR_firstName.Item(indexByEmail) & """,""" & HR_lastName.Item(indexByEmail) & """,""" & HR_nickName.Item(indexByEmail) & """,""" & HR_locationAux1.Item(indexByEmail) & """," & statusList.Item(n) & ",AD EEID is blank and " & hrMIS & " has no reqID," & hrMIS & " need enter requisitionID")
                                reportfileDiscrepancy.WriteLine(employeeID + ", (reqID AD:" + reqidAD + "<==>" + hrMIS + ":" + reqidHR + "),\"" + HR_firstName.Item(indexByEmail) + "\",\"" + HR_lastName.Item(indexByEmail) + "\",\"" + HR_nickName.Item(indexByEmail) + "\",\"" + HR_locationAux1.Item(indexByEmail) + "\"," + statusList.Item(n) + ",AD EEID is blank but set reqID,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need update EEID");
                            else
                                reportfileDiscrepancy.WriteLine(employeeID + ", (reqID AD:" + reqidAD + "<==>" + hrMIS + ":" + reqidHR + "),\"" + HR_firstName.Item(indexByEmail) + "\",\"" + HR_lastName.Item(indexByEmail) + "\",\"" + HR_nickName.Item(indexByEmail) + "\",\"" + HR_locationAux1.Item(indexByEmail) + "\"," + statusList.Item(n) + ",AD EEID is blank and reqID mismatch,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need correct reqID");
                            addDiscrepancyCount(HR_locationAux1.Item(indexByEmail));
                        }
                    }
                }
                else
                    // reqidAD found in HR, but may be wrongly set as other employee's reqID
                    if (indexByEmail >= 0)
                {
                    reqidHR = HR_reqID.Item(indexByEmail);
                    employeeID = HR_eeID.Item(indexByEmail);
                    if (reqidAD != reqidHR)
                    {
                        reportfileDiscrepancy.WriteLine(employeeID + ", (reqID AD:" + reqidAD + "<==>" + hrMIS + ":" + reqidHR + "),\"" + HR_firstName.Item(indexByEmail) + "\",\"" + HR_lastName.Item(indexByEmail) + "\",\"" + HR_nickName.Item(indexByEmail) + "\",\"" + HR_locationAux1.Item(indexByEmail) + "\"," + statusList.Item(n) + ",AD EEID is blank and reqID mismatch,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need correct reqID");
                        addDiscrepancyCount(HR_locationAux1.Item(indexByEmail));
                    }
                }
            }
            else
                // do not report terminated account
                if (!needStop)
            {
                if (statusList.Item(n) != "Terminated")
                {
                    if (indexByEmail >= 0)
                    {
                        reqidHR = HR_reqID.Item(indexByEmail);
                        employeeID = HR_eeID.Item(indexByEmail);
                        // if employee length < 6 then loop add leading "0"
                        string firstNameHR = HR_firstName(indexByEmail);
                        string lastNameHR = HR_lastName(indexByEmail);
                        string nickNameHR = HR_nickName(indexByEmail);
                        // CN=8831 SIPEndpoint,OU=Users,OU=SL,OU=EHI,DC=ehi,DC=ehealth,DC=com samID: 2103591 firstNameAD: 8831 lastNameAD: SIPEndpoint
                        if ((employeeID != "" || reqidHR != "") && lastNameHR == lastNameList.Item(n) && (firstNameList.Item(n) == firstNameHR || firstNameList.Item(n) == nickNameHR))
                        {
                            // emailKey already match, here is safe to auto set EEID regardless reqidHR <> "", because some cases HR has no reqID
                            // If reqidHR <> "" Then
                            // CTPR-9: if emailKey match and AD SAMAccountName = WFN WindowsAccountName, auto set reqID
                            if (needAutoSetReqID && WorkemailListKey(n) == HR_workEmailKey(indexByEmail))
                            {
                                try
                                {
                                    objUser = GetObject("LDAP://" + userDNlist(n));
                                    if (employeeID != "")
                                        objUser.employeeID = employeeID;
                                    else if (reqidHR != "")
                                        objUser.employeeNumber = reqidHR;
                                    if (!forTest)
                                        objUser.SetInfo();
                                    correctCount += 1;
                                    swReport.WriteLine("correcting #" + correctCount + ": " + samIDlist.Item(n) + "(" + HR_nickName.Item(indexByEmail) + " " + HR_firstName.Item(indexByEmail) + " " + HR_lastName.Item(indexByEmail) + "): " + "'both AD reqID and EEID are blank' correcting: ");

                                    if (employeeID != "")
                                        swReport.WriteLine("               set employeeID " + "\"" + employeeID + "\"");
                                    else if (reqidHR != "")
                                        swReport.WriteLine("               set employeeNumber(reqID) " + "\"" + reqidHR + "\"");
                                    swReport.WriteLine("");
                                }
                                catch (Exception ex)
                                {
                                    swLog.WriteLine("    =====>'both AD reqID and EEID are blank' correcting failed: " + samIDlist.Item(n) + " " + userDNlist(n));
                                    // only report this discrepancy if failed to correct it automatically
                                    // if HR record found, should display with HR firstname and nickname
                                    reportfileDiscrepancy.WriteLine(employeeID + ", (reqID:" + reqidHR + "),\"" + firstNameHR + "\",\"" + lastNameHR + "\",\"" + nickNameHR + "\",\"" + HR_locationAux1.Item(indexByEmail) + "\"," + statusList.Item(n) + ",both AD reqID and EEID are blank,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need update EEID");
                                    addDiscrepancyCount(HR_locationAux1.Item(indexByEmail));
                                }
                            }
                            else
                            {
                                // if HR record found, should display with HR firstname and nickname
                                reportfileDiscrepancy.WriteLine(employeeID + ", (reqID:" + reqidHR + "),\"" + firstNameHR + "\",\"" + lastNameHR + "\",\"" + nickNameHR + "\",\"" + HR_locationAux1.Item(indexByEmail) + "\"," + statusList.Item(n) + ",both AD reqID and EEID are blank,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need update EEID");
                                addDiscrepancyCount(HR_locationAux1.Item(indexByEmail));
                            }
                        }
                    }
                }
            }
            continue;
        }

        // employeeID <> ""
        if (needClearEmployeeNumber)
        {
            if (reqidAD != "")
            {
                if (statusList.Item(n) != "Terminated")
                {
                    // active User has both ReqID and EEID
                    // remove reqID automatically
                    try
                    {
                        objUser = GetObject("LDAP://" + userDNlist(n));
                        objUser.PutEx(ADS_PROPERTY_CLEAR, "employeeNumber", privateants.vbNull);
                        if (!forTest)
                            objUser.SetInfo();
                        correctCount += 1;
                        swReport.WriteLine("correcting #" + correctCount + ": " + samIDlist.Item(n) + "(" + firstNameList.Item(n) + " " + lastNameList.Item(n) + "): " + "'both ReqID and EEID' problem correcting: ");
                        swReport.WriteLine("               clear employeeNumber(reqID) " + "\"" + reqidAD + "\"" + " to " + "\"\"");
                        swReport.WriteLine("");
                    }
                    catch (Exception ex)
                    {
                        // swLog.WriteLine("    =====>clear employeeNumber(reqID) failed: " & samIDlist.Item(n) & " " & userDNlist(n))
                        // only report this discrepancy if failed to clear it automatically
                        if (index < 0)
                        {
                            reportfileDiscrepancy.WriteLine(employeeID + ", (reqID:" + reqidAD + "),\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",has both reqID and EEID,need remove reqID from AD user \"" + samIDlist.Item(n) + "\"");
                            addDiscrepancyCount(cityList.Item(n));
                        }
                        else
                        {
                            // if HR record found, should display with HR firstname and nickname
                            reportfileDiscrepancy.WriteLine(employeeID + ", (reqID:" + reqidAD + "),\"" + HR_firstName.Item(index) + "\",\"" + HR_lastName.Item(index) + "\",\"" + HR_nickName.Item(index) + "\",\"" + HR_locationAux1.Item(index) + "\"," + statusList.Item(n) + ",has both reqID and EEID,need remove reqID from AD user \"" + samIDlist.Item(n) + "\"");
                            addDiscrepancyCount(HR_locationAux1.Item(index));
                        }
                    }
                }
            }
        }

        if (index < 0)
        {
            // Also need set AD disabled if HR is Active but No need account
            index = HR_reqID.IndexOf(employeeID);
            DateTime createDate = new DateTime();
            try
            {
                createDate = whenCreatedList(n);
            }
            catch (Exception ex)
            {
            }

            if (statusList.Item(n) != "Terminated")
            {
                intNeedDisabledCount = intNeedDisabledCount + 1;
                lineStr = "N/A(" + n + 1 + "),";
                lineStr += "\"" + employeeID + "\",";
                lineStr += "\"" + samIDlist.Item(n) + "\",";
                findBySamAccountName(samIDlist.Item(n), "User");
                lineStr += "\"" + DistinguishedNameFound + "\",";
                lineStr += "\"" + legalFirstNameList(n) + "\",";
                lineStr += "\"" + firstNameList.Item(n) + "\",";
                lineStr += "\"" + lastNameList.Item(n) + "\",";
                lineStr += "\"" + middleNameList.Item(n) + "\",";
                lineStr += statusList.Item(n) + ",";
                lineStr += "\"" + jobtitleList.Item(n) + "\",";
                lineStr += "\"" + HomeDepartmentList.Item(n) + "\",";
                lineStr += "\"" + WorkemailList.Item(n) + "\",";
                lineStr += "\"" + whenCreatedList.Item(n) + "\",";
                lineStr += "\"" + whenChangedList.Item(n) + "\",";
                lineStr += "\"" + LastLogonTimeStampList(n) + "\",";
                lineStr += "\"" + employeeTypeList.Item(n) + "\",";
                lineStr += "\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\",";
                lineStr += "\"" + managerList.Item(n) + "\",";
                lineStr += ",,,"; // managerID_AD,managerID,ManagerDN,
                                  // HR data will be listed as new data in report file:
                lineStr += ",,,"; // ManagerWorkEmail,UserSegment,PrimaryWorkMobile,
                lineStr += ",,,,,"; // FirstName,MiddleName,LastName,NickName,Jobtitle,
                lineStr += ",,,,,"; // JobEffDate,BusinessUnit,HomeDepartment,Status,statusEffDate,
                lineStr += ",,,,,,"; // EmpClass,OriginalHireDate,HireDate,Address,City,State,
                lineStr += ",,,,,,"; // PostalCode,Country,LocationAux1,Office,CostCenter,NeedAccount,
                lineStr += ",,,,"; // Name_mismatch,FTE_mismatch,Title_mismatch,Department_mismatch,
                lineStr += ",,,,"; // Manager_mismatch,eeID_mismatch,samID_mismatch,defaultGroup

                GetUserMemberOf(samIDlist.Item(n));

                lineStr += "\"" + memberOf + "\",,"; // Notes,needDisabled

                try
                {
                    if (index < 0)
                    {
                        if (isArchived(employeeID, firstNameList.Item(n), lastNameList.Item(n)))
                        {
                            noteStr = "\"" + hrMIS + " archived. Need set disabled!\"";
                            lineStr += noteStr;
                            // reportfileNeedDisabled.WriteLine(lineStr)
                            if (forTest)
                            {
                                bool needReport = false;
                                string Ulocation = UCase(cityList.Item(n));
                                if (forXM_US == "US")
                                {
                                    if (!Ulocation.Contains("XIAMEN"))
                                    {
                                        needReport = true;
                                        discrepancyCount += 1;
                                    }
                                }
                                else if (forXM_US == "XM")
                                {
                                    if (Ulocation.Contains("XIAMEN"))
                                    {
                                        needReport = true;
                                        discrepancyCount += 1;
                                        discrepancyCountXM += 1;
                                    }
                                }
                                else
                                {
                                    needReport = true;
                                    discrepancyCount += 1;
                                    if (Ulocation.Contains("XIAMEN"))
                                        discrepancyCountXM += 1;
                                }
                                if (needReport)
                                {
                                    objUser = GetObject("LDAP://" + userDNlist(n));
                                    string hireDateStr = objUser.extensionAttribute9;
                                    if (hireDateStr != "")
                                    {
                                        DateTime hireDate = "2020-08-01";
                                        try
                                        {
                                            hireDate = Strings.Trim(hireDateStr);
                                        }
                                        catch (Exception ex)
                                        {
                                        }

                                        if (hireDate > DateTime.Now)
                                        {
                                            needReport = false;
                                            discrepancyCount -= 1;
                                            if (Ulocation.Contains("XIAMEN"))
                                                discrepancyCountXM -= 1;
                                        }
                                    }
                                    if (needReport)
                                        reportfileDiscrepancy.WriteLine("," + employeeID + ",\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",HR archived (may be rehire?),AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need disabled?");
                                }
                            }
                        }
                        else if (WDID.IndexOf(employeeID) < 0)
                        {
                            if (createDate.AddDays(noShow_ThresholdDays) < DateTime.Today || (forTest && needShowNoHRrecord))
                            {
                                // a fail-safe here: just send discrepancy alert email, don't set disabled!!! There is an incident hapened on 3/2/2016
                                if (!EEIDincorrect.Contains("." + employeeID + "."))
                                {
                                    // readHireDate
                                    objUser = GetObject("LDAP://" + userDNlist(n));
                                    if (objUser.extensionAttribute9 == "")
                                    {
                                        bool needReport = false;
                                        string Ulocation = UCase(cityList.Item(n));
                                        if (forXM_US == "US")
                                        {
                                            if (!Ulocation.Contains("XIAMEN"))
                                            {
                                                needReport = true;
                                                discrepancyCount += 1;
                                            }
                                        }
                                        else if (forXM_US == "XM")
                                        {
                                            if (Ulocation.Contains("XIAMEN"))
                                            {
                                                needReport = true;
                                                discrepancyCount += 1;
                                                discrepancyCountXM += 1;
                                            }
                                        }
                                        else
                                        {
                                            needReport = true;
                                            discrepancyCount += 1;
                                            if (Ulocation.Contains("XIAMEN"))
                                                discrepancyCountXM += 1;
                                        }
                                        if (needReport)
                                        {
                                            if (HRactiveEEIDlist.IndexOf(employeeID) >= 0)
                                                reportfileDiscrepancy.WriteLine("," + employeeID + ",\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",No " + hrMIS + " data today (last time is Active),AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need disable?");
                                            else if (HRterminatedEEIDlist.IndexOf(employeeID) >= 0)
                                                reportfileDiscrepancy.WriteLine("," + employeeID + ",\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",No " + hrMIS + " data today (last time is Terminated),AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need disable?");
                                            else
                                                reportfileDiscrepancy.WriteLine("," + employeeID + ",\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",No " + hrMIS + " record,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need disable?");
                                        }
                                    }
                                    else
                                    {
                                    }
                                }
                            }
                        }
                    }
                    else
                        reqID2EEIDcorrecting(index, n, createDate);
                    for (int t = 0; t <= FTEnonFTEgroupNamesList.Count - 1; t++)
                    {
                        string FTEnonFTEname = FTEnonFTEgroupNamesList.Item(t);
                        if (Strings.UCase(memberOf).Contains(Strings.UCase(FTEnonFTEname)))
                            FTE_mismatch += "remove " + FTEnonFTEname + ". ";
                    }
                }
                catch (Exception ex)
                {
                    hasError = true;
                    hasWriteEntireError = true;
                    swLog.WriteLine("Write report file failed, please check if there is invalid text character: " + privateants.vbCrLf + lineStr);
                }
                eeIDs.Add(employeeID);
            }
            else
                // if AD disabled and reqID set into EEID, also need correcting
                if (index >= 0)
                reqID2EEIDcorrecting(index, n, createDate);
        }
    }
    return eeIDs.Count;
}

private static bool isArchived(string eeid, string fn, string ln)
{
    if (HRterminatedEEIDlist.IndexOf(eeid) >= 0)
        return true;
    return false;
}

private static void reqID2EEIDcorrecting(int index, int n, DateTime createDate)
{
    if (needStop)
        return;

    string firstNameHR = HR_firstName(index);
    string lastNameHR = HR_lastName(index);
    string nickNameHR = HR_nickName(index);
    if (lastNameHR == lastNameList.Item(n) && (firstNameList.Item(n) == firstNameHR || firstNameList.Item(n) == nickNameHR))
    {
        // reqID set into EEID correcting
        try
        {
            objUser = GetObject("LDAP://" + userDNlist(n));
            string oldReqID = objUser.employeeNumber;
            objUser.employeeNumber = employeeID;
            objUser.PutEx(ADS_PROPERTY_CLEAR, "employeeID", privateants.vbNull);
            if (!forTest)
                objUser.SetInfo();
            correctCount += 1;
            swReport.WriteLine("correcting #" + correctCount + ": " + samIDlist.Item(n) + "(" + firstNameList.Item(n) + " " + lastNameList.Item(n) + "): " + "'reqID set into EEID' problem correcting: ");
            swReport.WriteLine("               update employeeNumber(reqID) " + "\"" + oldReqID + "\"" + " to " + "\"" + employeeID + "\"");
            swReport.WriteLine("               clear employeeID " + "\"" + employeeID + "\"" + " to " + "\"\"");
            swReport.WriteLine("");
        }
        catch (Exception ex)
        {
            // swLog.WriteLine("    =====>reqID set into EEID problem correcting failed: " & samIDlist.Item(n) & " " & userDNlist(n))
            // only report this discrepancy if failed to correct it automatically
            reportfileDiscrepancy.WriteLine(wkdID + "," + preID + ",\"" + firstNameList.Item(n) + "\",\"" + lastNameList.Item(n) + "\",,\"" + mapLocationAux1ByCity(cityList.Item(n)) + "\"," + statusList.Item(n) + ",AD reqID set into EEID,AD user \"" + samIDlist.Item(n) + "\" created on " + createDate.ToString("MM/dd/yyyy") + " need correct EEID");
            addDiscrepancyCount(cityList.Item(n));
        }
    }
}

private static string getFileTimeCMD(string CMDStr)
{
    string output = "";
    string outputWithSize = "";
    // Dim monthStr, dayStr As String
    DateTime hrbTime = new DateTime(), scheduledTime = new DateTime();
    scheduledTime = DateTime.Now;

    if (isPDT(scheduledTime))
        timeDifferenceXM -= 1;

    if (forTest)
        // Xiamen is 15 hours ahead of US
        // manual upload at 8am to 19pm, sync running at next day 7:15am, minus 15 hours is also reasonable
        scheduledTime = scheduledTime.AddHours(-timeDifferenceXM);
    // monthStr = Month(scheduledTime)
    // dayStr = Day(scheduledTime)
    // scheduledTime = monthStr & "/" & dayStr & "/" & Year(scheduledTime) & " 08:15"

    try
    {
        Process p = new Process();
        p.StartInfo.FileName = @"C:\Windows\system32\cmd.exe";
        p.StartInfo.Arguments = "/c " + CMDStr;
        p.StartInfo.UseShellExecute = false;
        p.StartInfo.RedirectStandardInput = true;
        p.StartInfo.RedirectStandardOutput = true;
        p.StartInfo.RedirectStandardError = true;
        p.StartInfo.CreateNoWindow = false;
        p.StartInfo.WorkingDirectory = strWorkFolder;
        p.Start();
        p.StandardInput.WriteLine("Exit");
        output = p.StandardOutput.ReadToEnd();
        p.Close();
        if (output == null)
            output = "";
        if (output.Length > 27)
        {
            outputWithSize = Strings.Trim(output.Substring(output.Length - 36));
            output = Strings.Trim(output.Substring(output.Length - 27));
            if (output.Contains(HRsftpFileName))
            {
                output = Strings.Trim(output.Substring(0, Strings.InStr(output, HRsftpFileName) - 1));
                hrbTime = DateTime.Year(scheduledTime) + " " + output;
            }
        }

        if (hrbTime.AddHours(outOfDate_ThresholdHours) > scheduledTime)
        {
            if (hrbTime.AddHours(outOfDate_ThresholdHours * (-2)) < scheduledTime)
                hrFileReady = true;
            else if (scheduledTime.ToString("MM/dd/yyyy").StartsWith("01/01") && hrbTime.ToString("MM/dd/yyyy").StartsWith("12/31"))
                hrFileReady = true;
        }
    }
    catch (Exception ex)
    {
        string ccAddress = "Steve.Arnoldus@ehealth.com";
        tmpStr = "Reading " + hrMIS + " file received time from SFTP failed!" + privateants.vbCrLf + ex.ToString();
        swLog.WriteLine(tmpStr);
        swLog.WriteLine("Sending alert email for reading " + hrMIS + " file received time failed...");
        sendAlertEmail("Changtu.Wang@ehealth.com", ccAddress, "AD sync: Download " + hrMIS + " file from SFTP failed!", tmpStr, "");
    }

    return outputWithSize;
}

private static void DoGroupMemberAction()
{
    tmpStr = getYYYYMMDD_hhtt();
    strReportfileName = strWorkFolder + @"\" + strReportSubFolder + @"\" + "groupReport" + tmpStr + ".txt";
    try
    {
        swReport = new System.IO.StreamWriter(strReportfileName, true); // append for 2nd time
    }
    catch (Exception ex)
    {
        swLog.WriteLine("Cannot open the file for writing, please close it first: " + strReportfileName);
        return;
    }

    object input = null;
    int tryTimes = 0;
    while (input == null)
    {
        try
        {
            tryTimes += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(strExportFileGroupAction);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTimes > 900)
        {
            tmpStr = "Open file time out!" + strExportFileGroupAction;
            swLog.WriteLine(tmpStr);
            sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: Read CSV file time out", tmpStr, "");
            break;
        }
    }
    if (input == null)
        return;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] CSVtitle;
    CSVtitle = input.readfields();

    intRowCount = 0;
    intOKcount = 0;
    intErrorCount = 0;

    string groupDN = "";
    string action = "";
    string memberDN = "";
    while ((!input.endofdata))
    {
        rows.Add(input.readfields);

        groupDN = "";
        try
        {
            groupDN = rows.Item(intRowCount)(0);
        }
        catch (Exception ex)
        {
        }
        action = "";
        try
        {
            action = LCase(rows.Item(intRowCount)(1));
        }
        catch (Exception ex)
        {
        }
        memberDN = "";
        try
        {
            memberDN = rows.Item(intRowCount)(2);
        }
        catch (Exception ex)
        {
        }

        intRowCount = intRowCount + 1;

        if (!memberDN.Contains(strDomain))
        {
            if (!memberDN.EndsWith(","))
                memberDN += ",";
            memberDN += strDomain;
        }

        try
        {
            objUser = Interaction.GetObject("LDAP://" + memberDN);
        }
        catch (Exception ex)
        {
            tmpStr = "#" + intRowCount + ": ";
            swReport.WriteLine(tmpStr);
            swReport.WriteLine("    ====>Cannot find " + memberDN);
            continue;
        }

        employeeID = objUser.employeeID;
        samAccountName = objUser.samAccountName;
        FirstName = objUser.givenName;
        LastName = objUser.sn;
        HomeDepartment = objUser.Department;
        Jobtitle = objUser.Title;

        if (objUser.AccountDisabled)
        {
            Status = "Terminated";
            if (memberDN.Contains("OU=Alumni") && action == "add")
                continue;
        }
        else
            Status = "Active";

        lineStr = "\"" + employeeID + "\"," + samAccountName + "(" + FirstName + " " + LastName + ")," + Status;
        if (employeeID != "")
        {
            if (HR_eeID0.IndexOf(employeeID) >= 0)
            {
                if (LCase(HR_FTE0(HR_eeID0.IndexOf(employeeID))) == "non-employee")
                    lineStr += "(Non-employee)";
            }
        }
        lineStr += ",\"" + HomeDepartment + "\",\"" + Jobtitle + "\"," + memberDN.Replace("," + strDomain, "");

        lineStr = "#" + intRowCount + ": " + lineStr;

        swReport.WriteLine(lineStr);

        if (groupDN == "" || action == "")
        {
            swReport.WriteLine("    ====>group DN or action is blank.");
            continue;
        }
        if (!groupDN.Contains(strDomain))
        {
            if (!groupDN.EndsWith(","))
                groupDN += ",";
            groupDN += strDomain;
        }

        hasUpdated = false;
        if (action == "add")
        {
            try
            {
                DirectoryEntry objgroup = new DirectoryEntry("LDAP://" + groupDN);
                if (!forTest)
                {
                    objgroup.Properties("member").Add(memberDN);
                    objgroup.CommitChanges();
                }
                tmpStr = Microsoft.VisualBasic.Mid(groupDN, 4, Strings.InStr(groupDN, ",") - 4);
                lineStr = "    add to DL group: " + tmpStr;
                hasUpdated = true;
            }
            catch (Exception ex)
            {
                intErrorCount += 1;
                lineStr = "    ===>add to DL group failed: " + groupDN + " : " + ex.ToString();
            }
        }
        else if (action == "remove")
        {
            try
            {
                DirectoryEntry objgroup = new DirectoryEntry("LDAP://" + groupDN);
                if (!forTest)
                {
                    objgroup.Properties("member").Remove(memberDN);
                    objgroup.CommitChanges();
                }
                tmpStr = Microsoft.VisualBasic.Mid(groupDN, 4, Strings.InStr(groupDN, ",") - 4);
                lineStr = "    remove from DL group: " + tmpStr;
                hasUpdated = true;
            }
            catch (Exception ex)
            {
                intErrorCount += 1;
                lineStr = "    ===>remove from DL group failed: " + groupDN + " : " + ex.ToString();
            }
        }
        if (hasUpdated)
            intOKcount = intOKcount + 1;
        swReport.WriteLine(lineStr);
    }
    input.Close();

    swReport.Close();

    if (intRowCount < 1)
    {
        try
        {
            System.IO.File.Delete(strReportfileName);
            System.IO.File.Delete(strExportFileGroupAction);
        }
        catch (Exception ex)
        {
        }
        return;
    }

    tmpStr = intRowCount + " records processed ( " + intOKcount + " executed)";
    if (intErrorCount > 0)
        tmpStr = tmpStr + " but has " + intErrorCount + " update failed";
    swLog.WriteLine("group actiion: " + tmpStr + ". See " + strReportfileName);

    if (intOKcount > 0)
    {
        swLog.WriteLine("");
        swLog.WriteLine("Sending alert email for AD user email distribution list updated...");

        string subjectStr = intOKcount + " email DL or security group updated by the " + hrMIS + " into AD sync";

        string toAddress = entOpsEmail + ";";
        string ccAddress = "";
        ccAddress += checkEntOps("Nirmal.Mehta@ehealth.com");
        ccAddress += checkEntOps("Changtu.Wang@ehealth.com");

        string strMessage = "Hello IT Team Members,<br/><br/>";
        strMessage += subjectStr + ".<br/>";
        strMessage += "Please see the attachment for details.<br/>";
        strMessage += "Please double check.";

        sendAlertEmail(toAddress, ccAddress, subjectStr, strMessage, strReportfileName);
    }
}

private static ArrayList getAssociatedITemailFromDB(string ITlocation)
{
    ArrayList emailList = new ArrayList();
    emailList.Clear();
    if (ITlocation == "")
        return emailList;
    // if first time, insert records from CSV

    string sql;
    sql = "select ITcontactEmail ";
    sql += "from [ITauto].[dbo].[" + contactITtalbeName + "] "; // always read from ITauto
    sql += "where LocationAux1 = '" + ITlocation + "'";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    executeSQL.executeSQL(sql, "select");
    if (executeSQL.errStr != "")
    {
        hasError = true;
        return emailList;
    }
    if (executeSQL.intCount < 1)
        return emailList;

    DataTable tableTMP = new DataTable();
    tableTMP = executeSQL.tableResult;
    for (int i = 0; i <= tableTMP.Rows.Count - 1; i++)
    {
        DataRow drow;
        drow = tableTMP.Rows(i);
        if (IsDBNull(drow(0)))
            continue;
        emailList.Add(Trim(drow(0)));
    }
    return emailList;
}

private static void listFTEnonFTEgroupMemberShip()
{
    if (FTEnonFTEgroupNamesList.Count < 1)
        return;

    var filename = strWorkFolder + @"\" + strReportSubFolder + @"\FTE_nonFTE_" + getYYYYMMDD_hhtt() + ".csv";
    System.IO.StreamWriter swGroupMembership;

    try
    {
        swGroupMembership = new System.IO.StreamWriter(filename, false); // always create new file
    }
    catch (Exception ex)
    {
        Interaction.MsgBox("Cannot open the file for writing, please close it first: " + filename);
        return;
    }
    tmpStr = "groupDN,memberDN";
    swGroupMembership.WriteLine(tmpStr);

    for (int i = 0; i <= FTEnonFTEgroupNamesList.Count - 1; i++)
    {
        DirectoryEntry entry = new DirectoryEntry("LDAP://" + strDomain);
        DirectorySearcher Searcher = new DirectorySearcher(entry);
        // All distribution groups
        Searcher.Filter = "(&(objectCategory=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648)))";
        Searcher.Filter = "(&(" + Searcher.Filter + ")(sAMAccountName=" + FTEnonFTEgroupNamesList.Item(i) + "))";
        Searcher.PropertiesToLoad.Clear();
        Searcher.PropertiesToLoad.Add("distinguishedName");
        Searcher.PropertiesToLoad.Add("member");

        SearchResultCollection queryResults = Searcher.FindAll();
        SearchResult result;

        int index = 0;
        foreach (var result in queryResults)
        {
            index += 1;
            lineStr = "\"" + result.Properties("distinguishedName")(0).ToString + "\",";
            if (result.Properties("member").Count > 0)
            {
                // loop members
                for (int j = 0; j <= result.Properties("member").Count - 1; j++)
                {
                    DistinguishedNameFound = result.Properties("member")(j).ToString;
                    swGroupMembership.WriteLine(lineStr + "\"" + DistinguishedNameFound + "\"");
                }
            }
            else
                swGroupMembership.WriteLine(lineStr + ",");
        }
    }
    swGroupMembership.Close();
    swLog.WriteLine("List group membership in " + filename);
}

private static string readSettings(string theKeyName)
{
    string theKeyValue = "";
    for (int i = 0; i <= tableTMP.Rows.Count - 1; i++)
    {
        DataRow drow;
        drow = tableTMP.Rows(i);
        // drow(0) is row#
        if (IsDBNull(drow(1)) || IsDBNull(drow(2)))
            continue;
        string keyName = Trim(drow(1));
        if (Strings.LCase(keyName) == Strings.LCase(theKeyName))
        {
            theKeyValue = Trim(drow(2));
            break;
        }
    }
    return theKeyValue;
}

public static void setSyncSettings() // for ehiAD call
{
    string csvFileName = strWorkFolder + @"\" + strDataSubfolder + @"\SyncSettings.csv";
    if (System.IO.File.Exists(csvFileName))
    {
        readCSV2table(csvFileName, tableTMP);
        if (!hasError)
        {
            try
            {
                needUpdateJobTitle = readSettings("needUpdateJobTitle");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateJobTitle has problem");
            }
            try
            {
                needUpdateJobTitleDelay = readSettings("needUpdateJobTitleDelay");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateJobTitleDelay has problem");
            }
            try
            {
                jobTitleUpdateDelayLocationStr = readSettings("jobTitleUpdateDelayLocationStr");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for jobTitleUpdateDelayLocationStr has problem");
            }
            try
            {
                daysInactiveLimit = readSettings("daysInactiveLimit");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for daysInactiveLimit has problem");
            }
            try
            {
                jobTitleUpdateDelayDays = readSettings("jobTitleUpdateDelayDays");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for jobTitleUpdateDelayDays has problem");
            }
            try
            {
                jobTitleUpdateDelayEndDayStr = readSettings("jobTitleUpdateDelayEndDayStr");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for jobTitleUpdateDelayEndDayStr has problem");
            }
            try
            {
                needUpdateHomeDeptment = readSettings("needUpdateHomeDeptment");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateHomeDeptment has problem");
            }
            try
            {
                needUpdateManager = readSettings("needUpdateManager");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateManager has problem");
            }
            try
            {
                needUpdateCompany = readSettings("needUpdateCompany");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateCompany has problem");
            }
            try
            {
                needUpdateAddress = readSettings("needUpdateAddress");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateAddress has problem");
            }
            try
            {
                needUpdateOffice = readSettings("needUpdateOffice");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateOffice has problem");
            }
            // use Skype now, no longer need update MOC
            // Try
            // needUpdateMOCaccount = readSettings("needUpdateMOCaccount")
            // Catch ex As Exception
            // swLog.WriteLine("read settings for needUpdateMOCaccount has problem")
            // End Try
            try
            {
                needUpdateMiddleName = readSettings("needUpdateMiddleName");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateMiddleName has problem");
            }
            try
            {
                needEmailMismatchReport = readSettings("needEmailMismatchReport");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needEmailMismatchReport has problem");
            }
            try
            {
                needDiscrepancyEmailDomainLocationStr = readSettings("needDiscrepancyEmailDomainLocationStr");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needDiscrepancyEmailDomainLocationStr has problem");
            }
            try
            {
                NoNeedDiscrepancyNewHireEmailLocationStr = readSettings("NoNeedDiscrepancyNewHireEmailLocationStr");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for NoNeedDiscrepancyNewHireEmailLocationStr has problem");
            }
            try
            {
                needUpdateCostcenterLocationStr = readSettings("needUpdateCostcenterLocationStr");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needUpdateCostcenterLocationStr has problem");
            }
            try
            {
                needListFTEnonFTEgroupMemberShip = readSettings("needListFTEnonFTEgroupMemberShip");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for needListFTEnonFTEgroupMemberShip has problem");
            }
            try
            {
                ignoreOnLeaveEEIDs = readSettings("ignoreOnLeaveEEIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for ignoreOnLeaveEEIDs has problem");
            }
            try
            {
                ignoreUpdateEEIDs = readSettings("ignoreUpdateEEIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for ignoreUpdateEEIDs has problem");
            }
            try
            {
                hasPreferredLastnameEEIDs = readSettings("hasPreferredLastnameEEIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings for hasPreferredLastnameEEIDs has problem");
            }
            try
            {
                svctestSamIDs = readSettings("svctestSamIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings svctestSamIDs has problem");
            }
            try
            {
                NONsvctestSamIDs = readSettings("NONsvctestSamIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings NONsvctestSamIDs has problem");
            }
            try
            {
                noShowSamIDs = readSettings("noShowSamIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings noShowSamIDs has problem");
            }
            try
            {
                SOXsamIDs = readSettings("SOXsamIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings SOXsamIDs has problem");
            }
            try
            {
                ignoreRehireSamIDs = readSettings("ignoreRehireSamIDs");
            }
            catch (Exception ex)
            {
                swLog.WriteLine("read settings ignoreRehireSamIDs has problem");
            }
        }
    }
}

private static void addDiscrepancyCount(string location)
{
    if (forXM_US == "US")
    {
        if (Strings.UCase(location) != "XM" && !Strings.LCase(location).Contains("xiamen"))
            discrepancyCount += 1;
    }
    else if (forXM_US == "XM")
    {
        if (Strings.UCase(location) == "XM" || Strings.LCase(location).Contains("xiamen"))
        {
            discrepancyCount += 1;
            discrepancyCountXM += 1;
        }
    }
    else
    {
        discrepancyCount += 1;
        if (Strings.UCase(location) == "XM" || Strings.LCase(location).Contains("xiamen"))
            discrepancyCountXM += 1;
    }
}
private static void addEmailMismatchCount(string location)
{
    if (forXM_US == "US")
    {
        if (Strings.UCase(location) != "XM" && !Strings.LCase(location).Contains("xiamen"))
            emailMismatchCount += 1;
    }
    else if (forXM_US == "XM")
    {
        if (Strings.UCase(location) == "XM" || Strings.LCase(location).Contains("xiamen"))
        {
            emailMismatchCount += 1;
            emailMismatchCountXM += 1;
        }
    }
    else
    {
        emailMismatchCount += 1;
        if (Strings.UCase(location) == "XM" || Strings.LCase(location).Contains("xiamen"))
            emailMismatchCountXM += 1;
    }
}

private static string mapLocationAux1ByCity(string city)
{
    string aux1 = city;
    switch (Strings.UCase(city))
    {
        case "REMOTE" // USA Remote   TPA: Remote    TPA: Hobart, IN
       :
            {
                aux1 = "RMT";
                break;
            }

        case "SANTA CLARA":
            {
                aux1 = "SCS";
                break;
            }

        case "SAN FRANCISCO":
            {
                aux1 = "SF";
                break;
            }

        case "GOLD RIVER":
            {
                aux1 = "GDRV";
                break;
            }

        case "SALT LAKE CITY":
            {
                aux1 = "UT";
                break;
            }

        case "XIAMEN":
            {
                aux1 = "XM";
                break;
            }

        case "WASHINGTON":
            {
                aux1 = "DC";
                break;
            }

        case "AUSTIN":
            {
                aux1 = "AUSTIN";
                break;
            }

        case "INDIANAPOLIS":
            {
                aux1 = "IN";
                break;
            }

        case "MOUNTAIN VIEW":
            {
                aux1 = "SCS";
                break;
            }

        case "WESTFORD":
            {
                aux1 = "MA";
                break;
            }
    }
    return aux1;
}

public static bool listAllUsersCSV() // for ehiAD call
{
    hasError = false;
    Console.WriteLine(DateTime.Now + " query all AD users, please wait...");
    if (swLog != null)
        swLog.WriteLine(DateTime.Now + " query all AD users, please wait...");
    if (strRoot == "")
    {
        objRootDSE = Interaction.GetObject("LDAP://RootDSE");
        strRoot = objRootDSE.GET("DefaultNamingContext");
        strDomain = strRoot;
        strLDAP = "LDAP://" + strDomain;
    }

    System.IO.StreamWriter swr, swrHR;
    AllUsersCSVFileName = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADallusers" + getYYYYMMDD_hhtt() + ".csv"; // ".xls"
    ADusersHRfilenameCSV = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADusersCommon_" + getYYYYMMDD_hhtt() + ".csv";

    // writeUTF8Head(AllUsersCSVFileName)
    try
    {
        swr = new System.IO.StreamWriter(AllUsersCSVFileName, false); // for overwrite
    }
    catch (Exception ex)
    {
        swLog.WriteLine("Open file for writing error. Please close this file first: " + privateants.vbCrLf + AllUsersCSVFileName);
        return false;
    }

    // writeUTF8Head(ADusersHRfilenameCSV)
    try
    {
    }
    // swrHR = New System.IO.StreamWriter(ADusersHRfilenameCSV, False) 'for overwrite
    catch (Exception ex)
    {
        ADusersHRfilenameCSV = "";
        swLog.WriteLine("Open file for writing error. Please close this file first: " + privateants.vbCrLf + ADusersHRfilenameCSV);
    }
    connQueryAllUsers = Interaction.CreateObject("ADODB.Connection");
    cmdQueryAllUsers = Interaction.CreateObject("ADODB.Command");
    connQueryAllUsers.Provider = "ADsDSOObject";
    connQueryAllUsers.Open("Active Directory Provider");
    // objConnection.Open "Provider=ADsDSOObject;"
    cmdQueryAllUsers.ActiveConnection = connQueryAllUsers;

    cmdQueryAllUsers.Properties("Page Size") = 1000;
    strfilter = "(&(objectCategory=Person)(objectClass=User))";
        // strfilter = "(&(&(objectCategory=Person)(objectClass=User))(sAMAccountName=bchissie))"
        private string strAttributesCSV = "employeeID,employeeNumber,sAMAccountName,DistinguishedName,comment,givenName,sn,initials,title,department,mail,proxyAddresses,whenCreated,whenChanged,employeeType,l,manager,userAccountControl";
private string strTitleCSVall = "employeeID,reqID,sAMAccountName,DistinguishedName,LegalFirstName,FirstName,LastName,MiddleName,JobTitle,HomeDepartment,WorkEmail,proxyAddresses,WhenCreated,WhenChanged,employeeType,City,Manager,Status,LastLogonTimeStamp,inactiveDays,needDisableInactive,daysCreated,Test_Service_Account";
private string strTitleHR = "employeeID,reqID,WindowsAccountName,LegalFirstName,FirstName,LastName,MiddleName,JobTitle,HomeDepartment,WorkEmail,WhenCreated,WhenChanged,employeeType,City,Status,daysCreated,Test_Service_Account";
private int proxyAddressesCol = 11;
cmdQueryAllUsers.commandtext = "<LDAP://OU=EHI," + strRoot + ">;" + strfilter + ";" + strAttributesCSV + ";" + strScope;
try
{
    rsQueryAllUsers = cmdQueryAllUsers.EXECUTE;
}
catch (Exception ex)
{
    swLog.WriteLine("Error: " + ex.ToString());
    return false;
}
if (rsQueryAllUsers.RecordCount < 1)
{
    swLog.WriteLine("No AD users found on " + strRoot + " for " + strfilter);
    return false;
}
swr.WriteLine(strTitleCSVall);
if (ADusersHRfilenameCSV != "")
{
}

// rsall2CSVstatusConverted
// Dim lineHR As String = ""
intRow = 0;
rsQueryAllUsers.MoveFirst();
try
{
    while (!(rsQueryAllUsers.EOF))
    {
        intRow += 1;
        employeeidAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields(0).value))
            employeeidAD = Strings.Trim(rsQueryAllUsers.Fields(0).value);
        eeIDlist.Add(employeeidAD);
        reqidAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("employeeNumber").value))
            reqidAD = Strings.Trim(rsQueryAllUsers.Fields("employeeNumber").value);
        reqIDlist.Add(reqidAD);

        samAccountName = rsQueryAllUsers.Fields("samAccountName").value;
        samIDlist.Add(Strings.LCase(samAccountName));

        DistinguishedName = rsQueryAllUsers.Fields("DistinguishedName").value;
        userDNlist.Add(DistinguishedName);
        LegalFirstNameAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("comment").value))
            LegalFirstNameAD = rsQueryAllUsers.Fields("comment").value;
        legalFirstNameList.Add(LegalFirstNameAD);
        firstNameAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("givenName").value))
            firstNameAD = rsQueryAllUsers.Fields("givenName").value;
        firstNameList.Add(firstNameAD);
        lastNameAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("sn").value))
            lastNameAD = rsQueryAllUsers.Fields("sn").value;
        lastNameList.Add(lastNameAD);
        middleNameAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("initials").value))
            middleNameAD = rsQueryAllUsers.Fields("initials").value;
        middleNameList.Add(middleNameAD);
        jobtitleAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("title").value))
            jobtitleAD = rsQueryAllUsers.Fields("title").value;
        jobtitleList.Add(jobtitleAD);
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("department").value))
            HomeDepartmentList.Add(rsQueryAllUsers.Fields("department").value);
        else
            HomeDepartmentList.Add("");
        tmpStr = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("mail").value))
            tmpStr = rsQueryAllUsers.Fields("mail").value;

        if (Strings.InStr(tmpStr, "@") > 0)
        {
            WorkemailList.Add(tmpStr);
            WorkemailListKey.Add(Microsoft.VisualBasic.Left(Strings.LCase(tmpStr), Strings.InStr(tmpStr, "@"))); // Only WorkemailListKey need LCase
        }
        else
        {
            WorkemailList.Add("");
            WorkemailListKey.Add("");
        }

        if (!Information.IsDBNull(rsQueryAllUsers.Fields("proxyAddresses").value))
        {
            // proxyAddresses is an object
            object aliasObj = rsQueryAllUsers.Fields("proxyAddresses").value;
            if (!Information.IsDBNull(aliasObj))
            {
                tmpStr = "\"";
                foreach (var emailAlias in aliasObj)
                    tmpStr += "[" + emailAlias + "] ";
                tmpStr += "\"";
            }
            proxyAddressesList.Add(tmpStr);
        }
        else
            proxyAddressesList.Add("");

        DateTime createTime, LastLogonTime, runTime;
        runTime = DateTime.Now;
        int inactiveDays;
        string inactiveValue = "";

        // always has whenCreated value in AD
        createTime = rsQueryAllUsers.Fields("whenCreated").value;
        whenCreatedList.Add(createTime);
        LastLogonTime = GetUserLastLogonTime(samAccountName); // also read AccountExpirationDate
        if (LastLogonTime < createTime)
            inactiveDays = DateTime.DateDiff("d", createTime, runTime);
        else
            inactiveDays = DateTime.DateDiff("d", LastLogonTime, runTime);

        if (!Information.IsDBNull(rsQueryAllUsers.Fields("whenChanged").value))
            whenChangedList.Add(rsQueryAllUsers.Fields("whenChanged").value);
        else
            whenChangedList.Add("");
        PhysicalDeliveryOfficeNameList.Add("");

        if (!Information.IsDBNull(rsQueryAllUsers.Fields("employeetype").value))
            employeeTypeList.Add(rsQueryAllUsers.Fields("employeetype").value);
        else
            employeeTypeList.Add("");

        if (!Information.IsDBNull(rsQueryAllUsers.Fields("l").value))
            cityList.Add(rsQueryAllUsers.Fields("l").value);
        else
            cityList.Add("");
        ManagerDN = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("manager").value))
            ManagerDN = rsQueryAllUsers.Fields("manager").value;
        managerList.Add(ManagerDN);

        lineStr = "\"" + employeeidAD + "\","; // if begins "0", will recorded in .CSV
        lineStr += "\"" + reqidAD + "\",";
        lineStr += "\"" + samAccountName + "\",";
        lineStr += "\"" + DistinguishedName + "\",";
        // lineHR = """" & employeeidAD & """,""" & reqidAD & """,""" & samAccountName & ""","

        for (var j = 4; j <= rsQueryAllUsers.Fields.Count - 3; j++)
        {
            if (!Information.IsDBNull(rsQueryAllUsers.Fields(j).value))
            {
                if (j == proxyAddressesCol)
                {
                    object aliasObj = rsQueryAllUsers.Fields(j).value;
                    if (!Information.IsDBNull(aliasObj))
                    {
                        lineStr += "\"";
                        foreach (var emailAlias in aliasObj)
                            lineStr += "[" + emailAlias + "] ";
                        lineStr += "\",";
                    }
                }
                else
                    lineStr += "\"" + rsQueryAllUsers.Fields(j).value + "\",";
            }
            else
                lineStr += "\"" + "\",";
        }
        lineStr += "\"" + ManagerDN + "\",";

        statusAD = "";
        if (!Information.IsDBNull(rsQueryAllUsers.Fields("userAccountControl").value))
        {
            intUAC = rsQueryAllUsers.Fields("userAccountControl").value;
            if ((intUAC & ADS_UF_ACCOUNTDISABLE) == 0)
            {
                statusAD = "Active";
                if ((employeeidAD != "" || reqidAD != ""))
                {
                    if (AccountExpirationDate == null)
                    {
                        if (inactiveDays >= daysInactiveLimit || (needCheckNewUsersNotLogon && inactiveDays >= DaysInactiveLimitNewHire && LastLogonTime == "01/01/1900"))
                            // will loop DC in checkNewUsersNotLogon(), just mark Yes here
                            // If LastLogonTime < createTime Then
                            // Dim LastLogon As DateTime
                            // LastLogon = loopDCfindLastLogon(samAccountName)
                            // If LastLogonTime < LastLogon Then
                            // LastLogonTime = LastLogon
                            // inactiveDays = DateDiff("d", LastLogonTime, runTime)
                            // End If
                            // If LastLogonTime < createTime Then
                            // inactiveDays = DateDiff("d", createTime, runTime)
                            // End If
                            // End If
                            // If (inactiveDays >= daysInactiveLimit OrElse LastLogonTime < createTime) Then
                            inactiveValue = inactiveYesStr;
                    }
                    else
                        inactiveValue = "expired:" + AccountExpirationDate;
                }
            }
            else
                statusAD = "Terminated";
        }
        statusList.Add(statusAD);
        LastLogonTimeStampList.Add(LastLogonTime);
        inactiveDaysList.Add(inactiveDays);
        inactiveList.Add(inactiveValue);

        lineStr += "\"" + statusAD + "\"," + "\"" + LastLogonTime + "\"";
        lineStr += "," + inactiveDays + "," + inactiveValue;
        // lineHR &= """" & statusAD & """"

        string isServiceTestAccount = "";
        if (DistinguishedName.Contains("OU=ServiceAccounts"))
            isServiceTestAccount = "Yes";
        else if (employeeidAD == "" && reqID == "")
        {
            string LsamID = Strings.LCase(samAccountName);
            if (firstNameAD == "" || lastNameAD == "")
            {
                if ((NONsvctestSamIDs == ""))
                    isServiceTestAccount = "Yes";
                else if (!NONsvctestSamIDs.Contains("." + LsamID + "."))
                    isServiceTestAccount = "Yes";
            }
            else if (LsamID.Contains("svc") || LsamID.Contains("admin") || LsamID.StartsWith("conf") || LsamID.Contains("test") || LsamID.Contains("train") || LsamID.StartsWith("rrs") || LsamID.Contains("cisco") || LsamID.Contains("scan") || LsamID.Contains("service") || LsamID.Contains("citrix") || (svctestSamIDs != "" && svctestSamIDs.Contains("." + LsamID + ".")))
                isServiceTestAccount = "Yes";
            else if (jobtitleAD.StartsWith("Shared"))
                isServiceTestAccount = "Yes";
            else
            {
                string Lln = Strings.LCase(lastNameAD);
                if (Lln == "a" || Lln.Contains("endpoint") || Lln.Contains("test") || Lln.Contains("orion") || Lln.Contains("svc") || Lln.Contains("svn") || Lln.Contains("auto") || Lln.Contains("label") || Lln.Contains("transfer"))
                    isServiceTestAccount = "Yes";
                else
                {
                    string Lfn = Strings.LCase(firstNameAD);
                    if (Lfn.StartsWith("unity") || Lfn.StartsWith("protiv") || Lfn == "ppi" || Lfn == "psi" || Lfn == "rate" || Lfn == "wbe" || Lfn == "fp&a" || Lfn == "remote" || Lfn == "backup" || Lfn == "compliance" || Lfn == "bulk" || Lfn == "websense" || Lfn.StartsWith("tableau") || Lfn.Contains("test") || Lfn.Contains("events") || Lfn.Contains("eyitteam"))
                        // OrElse Lfn.Contains("websense") 'Lastname already blank
                        isServiceTestAccount = "Yes";
                }
            }
        }

        string in90days = "";
        if (isServiceTestAccount == "")
        {
            float hours;
            hours = DateTime.DateDiff("h", createTime, runTime);
            if (hours < 0)
                hours = 0;
            if (Math.Round(hours / (double)24, 1) < 91)
                in90days = Math.Round(hours / (double)24, 1);
        }

        lineStr += "," + in90days + ",\"" + isServiceTestAccount + "\"";
        swr.WriteLine(lineStr);
        // If isServiceTestAccount = "" Then
        // lineHR &= "," & in90days & ",""" & isServiceTestAccount & """"
        // If ADusersHRfilenameCSV <> "" Then
        // swrHR.WriteLine(lineHR)
        // End If
        // End If

        rsQueryAllUsers.MoveNext();
    }
}
catch (Exception ex)
{
    hasError = true;
    swLog.WriteLine("Error read row #" + intRow + ": " + samAccountName + ", " + ex.ToString());
    swr.Close();
    rsQueryAllUsers = null;
    connQueryAllUsers.close();
    try
    {
        File.Delete(AllUsersCSVFileName);
    }
    catch (Exception ex2)
    {
    }
    return false;
}

swr.Close();
if (ADusersHRfilenameCSV != "")
{
}
rsQueryAllUsers = null;
connQueryAllUsers.close();
Console.WriteLine(DateTime.Now + " list " + intRow + " users in " + AllUsersCSVFileName);
if (swLog != null)
    swLog.WriteLine(DateTime.Now + " list " + intRow + " users in " + AllUsersCSVFileName);
return true;
    }

    private static void readAllCSV2ArrayList()
{
    string aFileName = AllUsersCSVFileName;
    object inputAll = null;
    inputAll = My.Computer.FileSystem.OpenTextFieldParser(aFileName);

    inputAll.setDelimiters(",");
    ArrayList rows = new ArrayList();
    string[] title;
    title = inputAll.readfields();

    int col_employeeid = -1;
    int col_reqid = -1;
    int col_samaccountname = -1;
    int col_LegalFirstName = -1;
    int col_firstname = -1;
    int col_lastname = -1;
    int col_middlename = -1;
    int col_distinguishedname = -1;
    int col_jobtitle = -1;
    int col_homedepartment = -1;
    int col_workemail = -1;
    int col_proxyAddresses = -1;
    int col_whencreated = -1;
    int col_whenchanged = -1;
    int col_office = -1;
    int col_employeeType = -1;
    int col_city = -1;
    int col_manager = -1;
    int col_status = -1;
    int col_LastLogonTimeStamp = -1;
    int col_inactiveDays = -1;
    int col_disableInactive = -1;

    int intColCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        if (aTitle == "employeeid")
            col_employeeid = intColCount;
        else if (aTitle == "reqid" || aTitle == "requisitionid" || aTitle == "employeenumber")
            col_reqid = intColCount;
        else if (aTitle == "samaccountname" || aTitle == "windowsaccountname")
            col_samaccountname = intColCount;
        else if (aTitle == "distinguishedname")
            col_distinguishedname = intColCount;
        else if (aTitle == "legalfirstname")
            col_LegalFirstName = intColCount;
        else if (aTitle == "firstname")
            col_firstname = intColCount;
        else if (aTitle == "lastname")
            col_lastname = intColCount;
        else if (aTitle == "middlename")
            col_middlename = intColCount;
        else if (aTitle == "jobtitle")
            col_jobtitle = intColCount;
        else if (aTitle == "homedepartment")
            col_homedepartment = intColCount;
        else if (aTitle == "workemail")
            col_workemail = intColCount;
        else if (aTitle == "proxyaddresses")
            col_proxyAddresses = intColCount;
        else if (aTitle == "whencreated")
            col_whencreated = intColCount;
        else if (aTitle == "whenchanged")
            col_whenchanged = intColCount;
        else if (aTitle == "office")
            col_office = intColCount;
        else if (aTitle == "employeetype")
            col_employeeType = intColCount;
        else if (aTitle == "city")
            col_city = intColCount;
        else if (aTitle == "manager")
            col_manager = intColCount;
        else if (aTitle == "status")
            col_status = intColCount;
        else if (aTitle == "lastlogontimestamp")
            col_LastLogonTimeStamp = intColCount;
        else if (aTitle == "inactivedays")
            col_inactiveDays = intColCount;
        else if (aTitle == "needdisableinactive")
            col_disableInactive = intColCount;
        intColCount = intColCount + 1;
    }
    rows.Clear();
    intRowCount = 0;
    while ((!inputAll.endofdata))
    {
        rows.Add(inputAll.readfields);
        if (col_employeeid < 0)
            eeIDlist.Add("");
        else
            eeIDlist.Add(rows.Item(intRowCount)(col_employeeid));
        if (col_reqid < 0)
            reqIDlist.Add("");
        else
            reqIDlist.Add(rows.Item(intRowCount)(col_reqid));
        if (col_samaccountname < 0)
            samIDlist.Add("");
        else
            samIDlist.Add(LCase(rows.Item(intRowCount)(col_samaccountname)));
        if (col_LegalFirstName < 0)
            legalFirstNameList.Add("");
        else
            legalFirstNameList.Add(rows.Item(intRowCount)(col_LegalFirstName));
        if (col_firstname < 0)
            firstNameList.Add("");
        else
            firstNameList.Add(rows.Item(intRowCount)(col_firstname));
        if (col_lastname < 0)
            lastNameList.Add("");
        else
            lastNameList.Add(rows.Item(intRowCount)(col_lastname));
        if (col_middlename < 0)
            middleNameList.Add("");
        else
            middleNameList.Add(rows.Item(intRowCount)(col_middlename));
        if (col_distinguishedname < 0)
            userDNlist.Add("");
        else
            userDNlist.Add(rows.Item(intRowCount)(col_distinguishedname));
        if (col_jobtitle < 0)
            jobtitleList.Add("");
        else
        {
            // RFS-95926 has invalid character "–":  Position: Intern – Medicare
            // IT may copy this invalid character into AD
            tmpStr = rows.Item(intRowCount)(col_jobtitle);
            jobtitleList.Add(tmpStr.Replace("?C", "--"));
        }
        if (col_homedepartment < 0)
            HomeDepartmentList.Add("");
        else
            HomeDepartmentList.Add(rows.Item(intRowCount)(col_homedepartment));
        if (col_workemail < 0)
        {
            WorkemailList.Add("");
            WorkemailListKey.Add("");
        }
        else
        {
            tmpStr = rows.Item(intRowCount)(col_workemail);
            if (Strings.InStr(tmpStr, "@") > 0)
            {
                WorkemailList.Add(tmpStr);
                WorkemailListKey.Add(Microsoft.VisualBasic.Left(Strings.LCase(tmpStr), Strings.InStr(tmpStr, "@"))); // only WorkemailListKey need LCase
            }
            else
            {
                WorkemailList.Add("");
                WorkemailListKey.Add("");
            }
        }
        if (col_proxyAddresses < 0)
            proxyAddressesList.Add("");
        else
            proxyAddressesList.Add(rows.Item(intRowCount)(col_proxyAddresses));
        if (col_whencreated < 0)
            whenCreatedList.Add("");
        else
            whenCreatedList.Add(rows.Item(intRowCount)(col_whencreated));
        if (col_whenchanged < 0)
            whenChangedList.Add("");
        else
            whenChangedList.Add(rows.Item(intRowCount)(col_whenchanged));
        if (col_office < 0)
            PhysicalDeliveryOfficeNameList.Add("");
        else
            PhysicalDeliveryOfficeNameList.Add(rows.Item(intRowCount)(col_office));
        if (col_employeeType < 0)
            employeeTypeList.Add("");
        else
            employeeTypeList.Add(rows.Item(intRowCount)(col_employeeType));
        if (col_city < 0)
            cityList.Add("");
        else
            cityList.Add(rows.Item(intRowCount)(col_city));
        if (col_manager < 0)
            managerList.Add("");
        else
            managerList.Add(rows.Item(intRowCount)(col_manager));
        if (col_status < 0)
            statusList.Add("");
        else
            statusList.Add(rows.Item(intRowCount)(col_status));
        if (col_LastLogonTimeStamp < 0)
            LastLogonTimeStampList.Add("");
        else
            LastLogonTimeStampList.Add(rows.Item(intRowCount)(col_LastLogonTimeStamp));
        if (col_inactiveDays < 0)
            inactiveDaysList.Add("");
        else
            inactiveDaysList.Add(rows.Item(intRowCount)(col_inactiveDays));
        if (col_disableInactive < 0)
            inactiveList.Add("");
        else
            inactiveList.Add(rows.Item(intRowCount)(col_disableInactive));
        intRowCount += 1;
    }
    inputAll.Close();
    Console.WriteLine(aFileName + ": " + intRowCount + " records");
    if (swLog != null)
        swLog.WriteLine(aFileName + ": " + intRowCount + " records");
}

private static int checkNewUsersNotLogon()
{
    if (!needCheckNewUsersNotLogon)
        return 0;
    // may cause ContextSwitchDeadlock exception
    Console.WriteLine(DateTime.Now + " checke new users not logon, please wait...");
    swLog.WriteLine(DateTime.Now + " checke new users not logon, please wait...");
    int NewHireInactiveCount = 0;
    for (int n = 0; n <= eeIDlist.Count - 1; n++)
    {
        if ((n % 100) == 0)
            Console.Write(".");

        if (statusList.Item(n) != "Active")
            continue;
        employeeID = eeIDlist.Item(n);
        reqidAD = reqIDlist.Item(n);
        if (employeeID == "" && reqidAD == "")
            continue;

        if (ignoreRehireSamIDs.Contains("." + samIDlist.Item(n) + "."))
            continue;

        int HRindex = HR_eeID.IndexOf(employeeID);

        bool isRehire = false;
        if (HRindex >= 0)
        {
            OriginalHireDate = HR_OriginalHireDate(HRindex);
            HireDate = HR_HireDate(HRindex);
            if (OriginalHireDate != HireDate)
                isRehire = true;
        }

        DateTime createTime = new DateTime(), LastLogonTime = new DateTime(), LastLogon = new DateTime(), Runtime = new DateTime();
        try
        {
            createTime = whenCreatedList(n);
            LastLogonTime = LastLogonTimeStampList(n);
        }
        catch (Exception ex)
        {
        }
        Runtime = DateTime.Now;

        // rehire has old logon time
        if (LastLogonTime != DateTime.DateSerial(1900, 1, 1) && !isRehire)
            continue;

        DateTime theHireDate = new DateTime();
        if (HRindex < 0)
        {
            theHireDate = createTime;
            // d, h, n, s  7.98 days will show as 7 days!
            int inactiveDays = DateTime.DateDiff("d", theHireDate, Runtime);
            if (inactiveDays < DaysInactiveLimitNewHire)
                continue;
        }
        else
        {
            theHireDate = HireDate;
            if (theHireDate.AddDays(DaysInactiveLimitNewHire - 1) != DateTime.Today)
                continue;
        }

        if (LastLogonTime >= theHireDate)
            continue;
        LastLogon = loopDCfindLastLogon(samIDlist.Item(n));
        // Application.DoEvents()
        if (LastLogonTime < LastLogon)
            LastLogonTime = LastLogon;
        // rehire has old logon time
        if (LastLogonTime >= theHireDate)
            continue;
        if (HRindex < 0)
        {
        }
        else
        {
            // has HR record
            string newHireOrRehireText = "new hire";
            if (isRehire)
                newHireOrRehireText = "rehire(" + HireDate + ")";

            // don't need report this in discrepancy, but in set disabled report
            if (employeeID == "")
                lineStr = "NA,\"reqID:" + reqidAD + "\",";
            else
                lineStr = "NA,\"" + employeeID + "\",";
            lineStr += "\"" + samIDlist.Item(n) + "\",";
            lineStr += "\"" + userDNlist.Item(n) + "\",";
            lineStr += "\"" + legalFirstNameList.Item(n) + "\",";
            lineStr += "\"" + firstNameList.Item(n) + "\",";
            lineStr += "\"" + lastNameList.Item(n) + "\",";
            lineStr += "\"" + middleNameList.Item(n) + "\",";
            lineStr += statusList.Item(n) + ",";
            lineStr += "\"" + jobtitleList(n) + "\",";
            lineStr += "\"" + HomeDepartmentList(n) + "\",";
            lineStr += "\"" + WorkemailList(n) + "\",";
            lineStr += "\"" + whenCreatedList(n) + "\",";
            lineStr += "\"" + whenChangedList(n) + "\",";
            lineStr += "\"" + LastLogonTimeStampList(n) + "\",";

            lineStr += "\"" + employeeTypeList(n) + "\",";
            lineStr += "\"" + cityList(n) + "\",";
            lineStr += "\"" + managerList(n) + "\",";
            if (Trim(managerList(n)) != "")
            {
                int indexM = userDNlist.IndexOf(managerList(n));
                if (indexM < 0)
                    lineStr += "\"Not_Found\",";
                else
                    lineStr += "\"" + eeIDlist.Item(indexM) + "\",";
            }
            else
                lineStr += ",";
            lineStr += "\"" + HR_managerID(HRindex) + "\",";
            lineStr += ","; // """" & HR_managerDN(HRindex) & ""","

            lineStr += "\"" + HR_managerWorkEmail(HRindex) + "\",";
            lineStr += "\"" + HR_UserSegment(HRindex) + "\",";
            lineStr += "\"" + HR_mobile(HRindex) + "\",";

            lineStr += "\"" + HR_firstName(HRindex) + "\",";
            lineStr += "\"" + HR_middleName(HRindex) + "\",";
            lineStr += "\"" + HR_lastName(HRindex) + "\",";
            lineStr += "\"" + HR_nickName(HRindex) + "\",";
            lineStr += "\"" + HR_Jobtitle(HRindex) + "\",";
            lineStr += "\"" + HR_JobEffDate(HRindex) + "\",";
            lineStr += "\"" + HR_BusinessUnit(HRindex) + "\",";
            lineStr += "\"" + HR_HomeDepartment(HRindex) + "\",";
            lineStr += "\"" + HR_status(HRindex) + "\",";
            lineStr += "\"" + HR_statusEffDate(HRindex) + "\",";
            lineStr += "\"" + HR_FTE(HRindex) + "\",";
            lineStr += "\"" + HR_OriginalHireDate(HRindex) + "\",";
            lineStr += "\"" + HR_HireDate(HRindex) + "\",";

            // LocationAux1 = HR_locationAux1(HRindex)
            // Address = ""
            // City = ""
            // State = ""
            // PostalCode = ""
            // Country = ""
            // Fax = ""
            // mapOU_Address()

            lineStr += ",,,,,";
            // lineStr &= """" & Address & ""","
            // lineStr &= """" & City & ""","
            // lineStr &= """" & State & ""","
            // lineStr &= """" & PostalCode & ""","
            // lineStr &= """" & Country & ""","
            lineStr += "\"" + HR_locationAux1(HRindex) + "\",";
            lineStr += ","; // """" & PhysicalDeliveryOfficeName & ""","

            lineStr += "\"" + HR_CostCenter(HRindex) + "\",";
            lineStr += "\"" + HR_needAccount(HRindex) + "\",";
            lineStr += ",,,,,,,,,,";
            lineStr += "\"" + newHireNotLogonStr + DaysInactiveLimitNewHire + " days, need set disabled!\"";
            reportfileNeedDisabled.WriteLine(lineStr);
            NewHireInactiveCount += 1;
        }
    }
    return NewHireInactiveCount;
}

private static SearchResultCollection getDCs()
{
    var objRootDSE = Interaction.GetObject("LDAP://RootDSE");
    string configurationNamingContext = objRootDSE.Get("configurationNamingContext");
    string defaultNamingContext = objRootDSE.GET("DefaultNamingContext");
    strRoot = defaultNamingContext;
    // strDomain = strRoot

    DirectoryEntry dEntry = new DirectoryEntry();
    // dEntry.Path = "LDAP://" & strRoot
    dEntry.Path = "LDAP://" + configurationNamingContext;

    // dEntry.AuthenticationType = AuthenticationTypes.Secure
    DirectorySearcher deSearch = new DirectorySearcher();
    deSearch.SearchRoot = dEntry;
    deSearch.Filter = "(objectClass=nTDSDSA)";
    SearchResultCollection results = deSearch.FindAll;
    return results;
}

private static DateTime loopDCfindLastLogon(string samID)
{
    // D:\techPage\AD\Find LastLogon Across All Windows Domain Controllers - CodeProject.htm
    DateTime LastLogon = "01/01/1900";
    if ((ehiDCs.Count < 1))
        return LastLogon;

    Int64 RawLastLogonMAX = 0;
    SearchResult OneSearchResult;
    foreach (var OneSearchResult in ehiDCs)
    {
        DirectoryEntry deDomain = OneSearchResult.GetDirectoryEntry();
        if (!deDomain == null)
        {
            string dnsHostName;
            dnsHostName = deDomain.Parent.Properties("DNSHostName").Value.ToString;

            // GetUserLastLogon(samID)
            DirectoryEntry deUsers = new DirectoryEntry("LDAP://" + dnsHostName + "/" + strRoot);
            DirectorySearcher dsUsers = new DirectorySearcher(deUsers);
            dsUsers.Filter = "(&(objectClass=user)(samAccountName=" + samID + "))";
            // dsUsers.PropertiesToLoad.Add(" msDS-LastSuccessfulInteractiveLogonTime")
            dsUsers.PropertiesToLoad.Add("lastLogon");

            Int64 RawLastLogon = 0;
            try
            {
                SearchResult UserAccount = dsUsers.FindOne;
                RawLastLogon = System.Convert.ToInt64(UserAccount.Properties("lastLogon")(0));
                if (RawLastLogon > RawLastLogonMAX)
                    RawLastLogonMAX = RawLastLogon;
            }
            catch (Exception ex)
            {
            }
        }
    }
    LastLogon = DateTime.FromFileTime(RawLastLogonMAX);
    return LastLogon;
}

private static bool setDisable(string samID, string userDN, string reason)
{
    objUser = Interaction.GetObject("LDAP://" + userDN);
    if (objUser.AccountDisabled)
        return true;
    objUser.AccountDisabled = true;
    try
    {
        if (!forTest)
            objUser.SetInfo();
    }
    catch (Exception ex)
    {
        swLog.WriteLine("    =====> set account disabled failed: " + samID);
        return false;
    }

    return true;
}

public static bool createLogFile(string logFilename, bool isAppend)
{
    try
    {
        swLog = new System.IO.StreamWriter(logFilename, isAppend); // create new file for 1st time
    }
    catch (Exception ex)
    {
        return false;
    }
    return true;
}

private static string checkEntOps(string anEmail)
{
    // this function is for avoiding duplicated alert email sent since 09/27/2017 18:30(US time)

    // From:   Philip Gao
    // Sent:   Saturday, September 30, 2017 10: 08 AM
    // To: Sam Zhao < Sam.Zhao@ehealth.com>; Changtu Wang <Changtu.Wang@ehealth.com>; Michael Lin <Michael.Lin@ehealth.com>
    // about duplicated email sent, has configuration changes
    // Before: sjengsmtp01 -> sjexsmtp04 -> outlook.ehealthinsurance.com
    // Now: removed outlook.ehealthinsurance.com，all route to Office365
    // now do not specify delivery host，all route via DNS, query MX record
    // dig ehealthinsurance.com MX +short
    // 0 ehealthinsurance-com.mail.protection.outlook.com

    if (!EntOps.Contains(Strings.LCase(anEmail)))
        return anEmail + ";";
    return "";
}

private static bool CSV2XLS(string afileName)
{
    if (Strings.LCase(afileName).EndsWith(".xls"))
        return true;

    object input = null;
    while (input == null)
    {
        try
        {
            input = My.Computer.FileSystem.OpenTextFieldParser(afileName);
        }
        catch (Exception ex)
        {
            swLog.WriteLine(DateTime.Now + " read file error:  " + afileName);
            return false;
        }
    }
    if (input == null)
        return false;

    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    colHR_eeID = 0;
    int colHR_reqID = -1;
    colHR_reqID = 1;
    colHR_samAccountName = 2;
    int colHR_legalFirstName = 3;
    colHR_firstName = 4;
    colHR_lastName = 5;
    colHR_middleName = 6;
    colHR_Jobtitle = 7;
    colHR_HomeDepartment = 8;
    colHR_workEmail = 9;
    int col_WhenCreated = 10;
    int col_WhenChanged = 11;

    colHR_office = 12;
    colHR_City = 13;
    colHR_status = 14;
    int col_days = 15;
    int col_TestService = 16;

    var afileNameXLS = afileName.Replace(".csv", ".xls");
    if (File.Exists(afileNameXLS))
    {
        try
        {
            File.Delete(afileNameXLS);
            System.Threading.Thread.Sleep(1000);
        }
        catch (Exception ex)
        {
        }
    }

    string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + afileNameXLS + ";Extended Properties='Excel 8.0;HDR=No;IMEX=0;'"; // for write
    OleDbConnection conn = new OleDbConnection(strConn);
    try
    {
        conn.Open();
    }
    catch (Exception ex)
    {
        swLog.WriteLine(ex.ToString());
        return false;
    }
    string sheetName = "ADusers";
    string sql;
    sql = "CREATE TABLE " + sheetName + " ([employeeID] VarChar,[reqID] VarChar,[WindowsAccountName] VarChar,[LegalFirstName] VarChar,[FirstName] VarChar,[LastName] VarChar,[MiddleName] VarChar,[JobTitle] VarChar,[HomeDepartment] VarChar,[WorkEmail] VarChar,[WhenCreated] DateTime,[WhenChanged] DateTime,[Office] VarChar,[City] VarChar,[Status] VarChar,[daysCreated] VarChar,[Test_Service_Account] VarChar)";
    OleDbCommand olecommand = new OleDbCommand(sql, conn);
    try
    {
        olecommand.ExecuteNonQuery();
    }
    catch (Exception ex)
    {
        swLog.WriteLine(ex.ToString());
        return false;
    }

    DateTime runTime = DateTime.Now;
    HRrowCount = -1;
    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);
            HRrowCount = HRrowCount + 1;

            string legalFirstname = rows.Item(HRrowCount)(colHR_legalFirstName);
            legalFirstname = legalFirstname.Replace("'", "''");
            string lastname = rows.Item(HRrowCount)(colHR_lastName);
            lastname = lastname.Replace("'", "''");
            string firstname = rows.Item(HRrowCount)(colHR_firstName);
            firstname = firstname.Replace("'", "''");
            string middlename = rows.Item(HRrowCount)(colHR_middleName);
            middlename = middlename.Replace("'", "''");
            string eeID = rows.Item(HRrowCount)(colHR_eeID);
            string reqID = rows.Item(HRrowCount)(colHR_reqID);
            string jobtitle = rows.Item(HRrowCount)(colHR_Jobtitle);
            jobtitle = jobtitle.Replace("'", "''");
            string samAccountName = rows.Item(HRrowCount)(colHR_samAccountName);

            string in90days = rows.Item(HRrowCount)(col_days);
            string isServiceTestAccount = rows.Item(HRrowCount)(col_TestService);

            string workemail = rows.Item(HRrowCount)(colHR_workEmail);
            workemail = workemail.Replace("'", "''");
            string status = rows.Item(HRrowCount)(colHR_status);
            string dept = rows.Item(HRrowCount)(colHR_HomeDepartment);
            dept = dept.Replace("'", "''");
            string city = rows.Item(HRrowCount)(colHR_City);
            string office = rows.Item(HRrowCount)(colHR_office);
            string whenCreated = rows.Item(HRrowCount)(col_WhenCreated);
            string whenChanged = rows.Item(HRrowCount)(col_WhenChanged);

            sql = "insert into [" + sheetName + "$](employeeID,reqID,WindowsAccountName,LegalFirstName,FirstName,LastName,MiddleName,JobTitle,HomeDepartment,WorkEmail,WhenCreated,WhenChanged,Office,City,Status,daysCreated,Test_Service_Account)";
            sql += " values (";
            sql += "'" + eeID + "','" + reqID + "','" + samAccountName + "','" + legalFirstname + "','" + firstname + "','" + lastname + "','" + middlename + "','" + jobtitle + "','" + dept + "','" + workemail + "','" + whenCreated + "','" + whenChanged + "','" + office + "','" + city + "','" + status + "','" + in90days + "','" + isServiceTestAccount + "'";
            sql += ")";

            olecommand = new OleDbCommand(sql, conn);
            try
            {
                olecommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                conn.Close();
                swLog.WriteLine(ex.ToString());
                return false;
            }
        }
    }
    catch (Exception ex)
    {
        input.Close();
        swLog.WriteLine(ex.ToString());
        return false;
    }
    input.Close();
    conn.Close();
    // System.Threading.Thread.Sleep(5000)

    ADusersHRfilenameXLS = afileNameXLS;
    swLog.WriteLine(DateTime.Now + " generate Excel file OK: " + ADusersHRfilenameXLS);

    return true;
}

private static void sendADusersHRfile(string aFileName)
{
    if (!needSendADfileToHR)
        return;
    if (aFileName == "")
        return;
    if (!File.Exists(aFileName))
        return;
    string toAddress = "sameer.desai@ehealth.com";
    string ccAddress = "";
    try
    {
        sendAlertEmail(toAddress, ccAddress, "AD common users report", "Attached is all common AD users under OU=EHI （Excluding service/test accounts）.", aFileName);
    }
    catch (Exception ex)
    {
        Console.WriteLine(DateTime.Now + " Sending AD user report email failed!");
        sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: Sending AD user report email failed!", "Sending " + aFileName + " failed!", "");

        swLog.Close();
        if (!createLogFile(logFilename, true))
        {
            Console.WriteLine("Cannot reopen the file for writing, please close it first: " + logFilename);
            return;
        }
        if (Strings.LCase(aFileName).EndsWith(".xls"))
        {
            // if sending XLS file failed, just sending CSV file
            hasError = true;
            sendADusersHRfile(ADusersHRfilenameCSV);
        }
    }
}

private static bool preparePrime()
{
    if (hasWriteEntireError)
    {
    }
    if (!File.Exists(strExportFileEntire))
        return false;
    object input = null;
    int tryTime = 0;
    while (input == null)
    {
        try
        {
            tryTime += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(strExportFileEntire);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTime > 10)
        {
            tmpStr = "AD sync: Open file time out!" + strExportFileEntire;
            // swLog.WriteLine(tmpStr)
            // sendAlertEmail("Changtu.Wang@ehealth.com", "", tmpStr, tmpStr, "")
            break;
        }
    }
    if (input == null)
        return false;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    colHR_firstName = -1;
    colHR_lastName = -1;
    colHR_workEmail = -1;
    // Dim colHR_WDID As Integer = -1
    colHR_eeID = -1;
    colHR_samAccountName = -1;
    colHR_OriginalHireDate = -1;
    colHR_HireDate = -1;
    colHR_statusEffDate = -1;
    colHR_status = -1;
    colHR_managerID = -1;
    int colAD_managerID = -1;
    colHR_Jobtitle = -1;
    int colHR_BusinessUnit = -1;
    int colName_mismatch = -1;
    int colNoNeedAccount = -1;

    colHR_HomeDepartment = -1;
    colHR_City = -1;
    int colAD_City = -1;
    colHR_State = -1;
    colHR_locationAux1 = -1;

    colHR_FTE = -1;
    colHR_nickName = -1;

    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        // aTitle = aTitle.Replace("#", "")
        if (aTitle == "lastname")
            colHR_lastName = HRcolCount;
        else if (aTitle == "firstname")
            colHR_firstName = HRcolCount;
        else if (aTitle == "nickname")
            colHR_nickName = HRcolCount;
        else if (aTitle == "workemail_ad")
            colHR_workEmail = HRcolCount;
        else if (aTitle == "employeeid")
            // always use WDID now
            colHR_eeID = HRcolCount;
        else if (aTitle == "samaccountname")
            colHR_samAccountName = HRcolCount;
        else if (aTitle == "originalhiredate")
            colHR_OriginalHireDate = HRcolCount;
        else if (aTitle == "hiredate")
            colHR_HireDate = HRcolCount;
        else if (aTitle == "statuseffdate")
            colHR_statusEffDate = HRcolCount;
        else if (aTitle == "status")
            colHR_status = HRcolCount;
        else if (aTitle == "managerid")
            colHR_managerID = HRcolCount;
        else if (aTitle == "managerid_ad")
            colAD_managerID = HRcolCount;
        else if (aTitle == "jobtitle")
            colHR_Jobtitle = HRcolCount;
        else if (aTitle == "businessunit")
            colHR_BusinessUnit = HRcolCount;
        else if (aTitle == "homedepartment")
            colHR_HomeDepartment = HRcolCount;
        else if (aTitle == "city")
            colHR_City = HRcolCount;
        else if (aTitle == "city_ad")
            colAD_City = HRcolCount;
        else if (aTitle == "state")
            colHR_State = HRcolCount;
        else if (aTitle == "locationaux1")
            colHR_locationAux1 = HRcolCount;
        else if (aTitle == "empclass")
            colHR_FTE = HRcolCount;
        else if (aTitle == "name_mismatch")
            colName_mismatch = HRcolCount;
        else if (aTitle == "needaccount")
            colNoNeedAccount = HRcolCount;
        HRcolCount += 1;
    }
    if (HRcolCount < 1)
        // swLog.WriteLine(aFileName & ": " & " has no column.")
        return false;

    strReportPrime = strWorkFolder + @"\" + strPrimeSubfolder + @"\CaptivatePrimeupload.csv";
    strReportPrimeFailed = strWorkFolder + @"\" + strPrimeSubfolder + @"\CaptivatePrime_Failed" + DateTime.Now.ToString("yyyyMMddhhmm") + ".csv";
    if (File.Exists(strReportPrime))
    {
        DateTime fileTime = FileSystem.FileDateTime(strReportPrime);
        string backupFileName = strReportPrime.Replace(".csv", fileTime.ToString("yyyyMMddHHmm") + ".csv");
        if (!File.Exists(backupFileName))
        {
            try
            {
                File.Copy(strReportPrime, backupFileName);
            }
            catch (Exception ex)
            {
            }
        }
    }

    StreamWriter swPrime, swFailed;
    try
    {
        swPrime = new StreamWriter(strReportPrime, false); // always create new file
        swFailed = new StreamWriter(strReportPrimeFailed, false); // always create new file
    }
    catch (Exception ex)
    {
        return false;
    }

    ArrayList uniqueIDs = new ArrayList(), uniqueEmails = new ArrayList();
    string columnNamePrime = "Student.Full Name,Student.Email,Student.EmpID,Student.LoginCode,Student.StartDate,Date of Hire/Rehire,Student.Status,Supervisor.Email,Student.Position.Title,OrgLevel_1.Title,OrgLevel_2.Title,Home Department,Location";
    swPrime.WriteLine(columnNamePrime);
    swFailed.WriteLine(columnNamePrime + ",notes");

    HRrowCount = 0;
    int rootCount = 0;
    int failedCount = 0;
    string CircularReportingEmails = "";
    int col_managerID;
    int col_city;
    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);
            string jobTitle = LCase(rows.Item(HRrowCount)(colHR_Jobtitle));
            if (jobTitle == "independent contractor" || jobTitle == "board of directors" || jobTitle == "cobra special dependents")
            {
                HRrowCount += 1;
                continue;
            }
            if (colHR_FTE > 0)
            {
                if (LCase(rows.Item(HRrowCount)(colHR_FTE)) == "non-employee" || LCase(rows.Item(HRrowCount)(colHR_FTE)) == "independent contractor")
                {
                    HRrowCount += 1;
                    continue;
                }
            }
            if (colNoNeedAccount > 0)
            {
                if (rows(HRrowCount)(colNoNeedAccount) == noNeedAccountValueHR)
                {
                    HRrowCount += 1;
                    continue;
                }
            }
            statusAD = "";
            string linePrime = "";
            if (colHR_nickName < 0)
                NickName = "";
            else
                NickName = Trim(rows.Item(HRrowCount)(colHR_nickName));
            string fullName = "";
            if (NickName == "")
            {
                // no nickname, so use legal first
                if (colHR_firstName >= 0)
                    fullName = rows.Item(HRrowCount)(colHR_firstName);
            }
            else
                fullName = NickName;
            if (colHR_lastName >= 0)
                fullName += " " + rows.Item(HRrowCount)(colHR_lastName);
            linePrime += "\"" + fullName + "\"";
            if (colHR_workEmail < 0)
            {
                Workemail = "";
                linePrime += ",";
            }
            else
            {
                Workemail = rows.Item(HRrowCount)(colHR_workEmail);
                // If Workemail = "" Then
                // has no AD email address, copy from WFN email address
                // If colHR_eeID >= 0 Then
                // Dim HRindex As Integer = HR_eeID.IndexOf(rows.Item(HRrowCount)(colHR_eeID))
                // If HRindex >= 0 Then
                // Dim HRemail As String = HR_workEmail.Item(HRindex)
                // If Not HRemail.Contains("@ehealth") Then
                // just copy TPA email address
                // Workemail = Trim(HRemail)
                // End If
                // End If
                // End If
                // End If
                linePrime += ",\"" + Workemail + "\"";
            }

            col_managerID = colHR_managerID;
            col_city = colHR_City;

            if (colHR_eeID < 0)
            {
                employeeID = "";
                linePrime += ",";
            }
            else
            {
                // always use WDID now
                employeeID = rows(HRrowCount)(colHR_eeID);
                linePrime += ",\"" + employeeID + "\"";
                if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                {
                    // users in ignore update list will use current AD property for manager and city
                    if (colAD_managerID >= 0)
                        col_managerID = colAD_managerID;
                    if (colAD_City >= 0)
                        col_city = colAD_City;
                }
            }

            if (colHR_samAccountName < 0)
            {
                samAccountName = "";
                linePrime += ",";
            }
            else
            {
                samAccountName = rows.Item(HRrowCount)(colHR_samAccountName);
                linePrime += ",\"" + samAccountName + "\"";
            }

            if (colHR_OriginalHireDate < 0)
                linePrime += ",";
            else
                linePrime += ",\"" + rows.Item(HRrowCount)(colHR_OriginalHireDate) + "\"";
            if (colHR_HireDate < 0)
                linePrime += ",";
            else
                linePrime += ",\"" + rows.Item(HRrowCount)(colHR_HireDate) + "\"";
            // If colHR_statusEffDate < 0 Then
            // linePrime &= ","
            // Else
            // linePrime &= ",""" & rows.Item(HRrowCount)(colHR_statusEffDate) & """"
            // End If
            if (colHR_status < 0)
                linePrime += ",";
            else
            {
                statusAD = rows.Item(HRrowCount)(colHR_status);
                linePrime += ",\"" + statusAD + "\"";
            }
            if (col_managerID < 0)
                linePrime += ",";
            else
            {
                if (fullName == rootFullName || fullName == CEO)
                {
                    ManagerWorkEmail = "root";
                    rootCount += 1;
                }
                else
                {
                    // should consider manager account has email, but AD may not set manager?
                    ManagerID = rows.Item(HRrowCount)(col_managerID);
                    while ((ManagerID.Length < 6))
                        ManagerID = "0" + ManagerID;
                    int ADindex = eeIDlist.IndexOf(ManagerID);
                    if (ADindex < 0)
                        ManagerWorkEmail = "";
                    else
                        ManagerWorkEmail = WorkemailList(ADindex);
                }
                linePrime += ",\"" + ManagerWorkEmail + "\"";
            }
            if (colHR_Jobtitle < 0)
            {
                linePrime += ",";
                jobTitle = "";
            }
            else
            {
                jobTitle = rows.Item(HRrowCount)(colHR_Jobtitle);
                linePrime += ",\"" + jobTitle + "\"";
            }
            if (colHR_locationAux1 < 0)
                LocationAux1 = "";
            else
                LocationAux1 = Trim(rows.Item(HRrowCount)(colHR_locationAux1));
            if (LocationAux1 == "XM")
                linePrime += ",\"eHealth China\"";
            else
                linePrime += ",\"eHealth US\"";
            if (colHR_BusinessUnit < 0)
                linePrime += ",";
            else
                linePrime += ",\"" + rows.Item(HRrowCount)(colHR_BusinessUnit) + "\"";

            if (colHR_HomeDepartment < 0)
                linePrime += ",";
            else
                linePrime += ",\"" + rows.Item(HRrowCount)(colHR_HomeDepartment) + "\"";
            if (col_city < 0)
                linePrime += ",";
            else
                linePrime += ",\"" + rows.Item(HRrowCount)(col_city) + "\"";

            if (statusAD == "Active" || statusAD == "On Leave")
            {
                // Anthony email Sent: Friday, September 20, 2019 3:32 AM: Please start to include users whose status is “On Leave” on this report.
                string notes = "";
                bool passed = true;

                if (samAccountName == "")
                {
                    passed = false;
                    notes += "AD account not found; ";
                }

                if (employeeID == "")
                {
                    passed = false;
                    notes += "missing employee ID; ";
                }
                else
                {
                    if (uniqueIDs.IndexOf(employeeID) < 0)
                        uniqueIDs.Add(employeeID);
                    else
                    {
                        passed = false;
                        notes += "duplicate employee ID; ";
                    }
                    if (employeeID.Length > 32)
                    {
                        passed = false;
                        notes += "employee ID length is greater than 32; ";
                    }
                }
                if (fullName == "")
                {
                    passed = false;
                    notes += "missing employee name; ";
                }
                if (jobTitle == "")
                {
                    passed = false;
                    notes += "missing job title; ";
                }
                if (Workemail == "")
                {
                    passed = false;
                    notes += "missing employee email; ";
                }
                else
                {
                    if (uniqueEmails.IndexOf(Workemail) < 0)
                        uniqueEmails.Add(Workemail);
                    else
                    {
                        passed = false;
                        notes += "duplicate employee email; ";
                    }
                    if (!isValidEmail(Workemail))
                    {
                        passed = false;
                        notes += "invalid employee email; ";
                    }
                }
                if (ManagerWorkEmail == "")
                {
                    passed = false;
                    notes += "missing manager email; ";
                }
                else if (ManagerWorkEmail != "root")
                {
                    if (!isValidEmail(ManagerWorkEmail))
                    {
                        passed = false;
                        notes += "invalid manager email; ";
                    }
                    if (Workemail == ManagerWorkEmail)
                    {
                        passed = false;
                        notes += "Employee and Manager email is the same; ";
                    }

                    // ManagerWorkEmail is lookup from AD WorkemailList, so if ManagerWorkEmail <> "", Manager always is employee, except root

                    // if Circular reporting issue with employee aa to xxx
                    while ((ManagerWorkEmail != ""))
                    {
                        if (ManagerWorkEmail == Workemail || CircularReportingEmails.Contains(ManagerWorkEmail))
                        {
                            passed = false;
                            notes += "Circular reporting issue; ";
                            if (!CircularReportingEmails.Contains(ManagerWorkEmail))
                                CircularReportingEmails += ManagerWorkEmail + "; ";
                            break;
                        }
                        string preMngEmail = ManagerWorkEmail;
                        ManagerWorkEmail = ADfindManagerEmail(ManagerWorkEmail);
                        if (ManagerWorkEmail == preMngEmail)
                        {
                            // some one report to him/her self
                            if (!CircularReportingEmails.Contains(ManagerWorkEmail))
                                CircularReportingEmails += ManagerWorkEmail + "; ";
                        }
                    }
                }

                if (passed)
                    swPrime.WriteLine(linePrime);
                else
                {
                    // validate Missing employee email or Missing employee manager
                    failedCount += 1;
                    // report problem record in separated file
                    swFailed.WriteLine(linePrime + ",\"" + notes + "\"");
                }
            }
            // swLog.WriteLine(HR_eeID.Item(HRrowCount) & " " & HR_firstName.Item(HRrowCount) & " " & HR_lastName.Item(HRrowCount))
            HRrowCount += 1;
        }
    }

    catch (Exception ex)
    {
        input.Close();
        swLog.WriteLine(ex.ToString());
        return false;
    }

    if (rootCount < 1)
        swFailed.WriteLine(",,,\"" + "Root not found; " + "\"");
    else if (rootCount > 1)
        swFailed.WriteLine(",,,\"" + "Multiple " + rootCount + " roots (" + rootFullName + "); " + "\"");


    input.Close();

    swPrime.Close();
    swFailed.Close();

    Console.WriteLine(strExportFileEntire + ": " + HRrowCount + " records convert to " + strReportPrime + " successfully (validation failed count: " + failedCount + " ).");
    // Console.ReadKey()
    return true;
}

private static bool prepareOomnitza()
{
    if (!File.Exists(strExportFileFound))
        return false;
    object input = null;
    int tryTime = 0;
    while (input == null)
    {
        try
        {
            tryTime += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(strExportFileFound);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTime > 10)
        {
            tmpStr = "AD sync: Open file time out!" + strExportFileFound;
            // swLog.WriteLine(tmpStr)
            // sendAlertEmail("Changtu.Wang@ehealth.com", "", tmpStr, tmpStr, "")
            break;
        }
    }
    if (input == null)
        return false;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    colHR_firstName = -1;
    colHR_lastName = -1;
    colHR_workEmail = -1;
    // Dim colHR_WDID As Integer = -1
    colHR_eeID = -1;
    colHR_samAccountName = -1;
    colHR_OriginalHireDate = -1;
    colHR_HireDate = -1;
    colHR_statusEffDate = -1;
    colHR_status = -1;
    colHR_managerID = -1;
    int colAD_managerID = -1;
    colHR_Jobtitle = -1;
    int colHR_BusinessUnit = -1;
    int colName_mismatch = -1;
    int colNoNeedAccount = -1;

    colHR_HomeDepartment = -1;
    colHR_City = -1;
    int colAD_City = -1;
    colHR_State = -1;
    colHR_locationAux1 = -1;

    colHR_FTE = -1;
    colHR_nickName = -1;
    int colHR_employeeType = -1;

    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        // aTitle = aTitle.Replace("#", "")
        if (aTitle == "lastname")
            colHR_lastName = HRcolCount;
        else if (aTitle == "firstname")
            colHR_firstName = HRcolCount;
        else if (aTitle == "nickname")
            colHR_nickName = HRcolCount;
        else if (aTitle == "workemail_ad")
            colHR_workEmail = HRcolCount;
        else if (aTitle == "employeeid")
            // always use WDID now
            colHR_eeID = HRcolCount;
        else if (aTitle == "samaccountname")
            colHR_samAccountName = HRcolCount;
        else if (aTitle == "hiredate")
            colHR_HireDate = HRcolCount;
        else if (aTitle == "statuseffdate")
            colHR_statusEffDate = HRcolCount;
        else if (aTitle == "status")
            colHR_status = HRcolCount;
        else if (aTitle == "managerid")
            colHR_managerID = HRcolCount;
        else if (aTitle == "managerid_ad")
            colAD_managerID = HRcolCount;
        else if (aTitle == "jobtitle")
            colHR_Jobtitle = HRcolCount;
        else if (aTitle == "businessunit")
            colHR_BusinessUnit = HRcolCount;
        else if (aTitle == "homedepartment")
            colHR_HomeDepartment = HRcolCount;
        else if (aTitle == "city")
            colHR_City = HRcolCount;
        else if (aTitle == "city_ad")
            colAD_City = HRcolCount;
        else if (aTitle == "state")
            colHR_State = HRcolCount;
        else if (aTitle == "locationaux1")
            colHR_locationAux1 = HRcolCount;
        else if (aTitle == "empclass")
            colHR_FTE = HRcolCount;
        else if (aTitle == "name_mismatch")
            colName_mismatch = HRcolCount;
        else if (aTitle == "needaccount")
            colNoNeedAccount = HRcolCount;
        else if (aTitle == "usersegment")
            colHR_employeeType = HRcolCount;
        HRcolCount += 1;
    }
    if (HRcolCount < 1)
        // swLog.WriteLine(aFileName & ": " & " has no column.")
        return false;

    strReportOomnitza = strWorkFolder + @"\" + strPrimeSubfolder + @"\" + CSVfilenameOomnitza;
    if (File.Exists(strReportOomnitza))
    {
        DateTime fileTime = FileSystem.FileDateTime(strReportOomnitza);
        string backupFileName = strReportOomnitza.Replace(".csv", fileTime.ToString("yyyyMMddHHmm") + ".csv");
        if (!File.Exists(backupFileName))
        {
            try
            {
                File.Copy(strReportOomnitza, backupFileName);
            }
            catch (Exception ex)
            {
            }
        }
    }

    StreamWriter swOomnitza;
    try
    {
        swOomnitza = new StreamWriter(strReportOomnitza, false); // always create new file
    }
    catch (Exception ex)
    {
        return false;
    }

    swOomnitza.WriteLine("userAccount,firstName,lastName,FullName,HRstatus,HireDate,terminatedDate,workemail,location,managerName,employeeType,workerCategory,ehiComputer");

    HRrowCount = 0;
    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);

            if (colHR_eeID >= 0)
            {
                employeeID = rows(HRrowCount)(colHR_eeID);
                if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                {
                    HRrowCount += 1;
                    continue;
                }
            }

            if (colHR_samAccountName < 0)
                samAccountName = "";
            else
                samAccountName = rows.Item(HRrowCount)(colHR_samAccountName);
            if (samAccountName == "")
            {
                HRrowCount += 1;
                continue;
            }

            string lineOomnitza = "\"" + samAccountName + "\"";

            FirstName = "";
            if (colHR_firstName >= 0)
                FirstName = rows.Item(HRrowCount)(colHR_firstName);
            NickName = "";
            if (colHR_nickName >= 0)
                NickName = rows.Item(HRrowCount)(colHR_nickName);
            LastName = "";
            if (colHR_lastName >= 0)
                LastName = rows.Item(HRrowCount)(colHR_lastName);

            if (NickName == "")
                // lineOomnitza &= ",""" & FirstName & " " & LastName & """"
                lineOomnitza += ",\"" + FirstName + "\",\"" + LastName + "\",\"" + FirstName + " " + LastName + "\"";
            else
                // lineOomnitza &= ",""" & NickName & " " & LastName & """"
                lineOomnitza += ",\"" + NickName + "\",\"" + LastName + "\",\"" + NickName + " " + LastName + "\"";

            // Dim jobTitle As String = LCase(rows.Item(HRrowCount)(colHR_Jobtitle))
            if (colHR_status < 0)
                Status = "";
            else
                Status = rows.Item(HRrowCount)(colHR_status);
            lineOomnitza += ",\"" + Status.Replace("On Leave", "Active") + "\"";

            if (colHR_HireDate < 0)
                lineOomnitza += ",";
            else
                lineOomnitza += ",\"" + rows.Item(HRrowCount)(colHR_HireDate) + "\"";

            if (Status == "Terminated")
            {
                if (colHR_statusEffDate < 0)
                    lineOomnitza += ",";
                else
                    lineOomnitza += ",\"" + rows.Item(HRrowCount)(colHR_statusEffDate) + "\"";
            }
            else
                lineOomnitza += ",";

            if (colHR_workEmail < 0)
                Workemail = "";
            else
                Workemail = rows.Item(HRrowCount)(colHR_workEmail);
            lineOomnitza += ",\"" + Workemail + "\"";

            if (colHR_City < 0)
                lineOomnitza += ",";
            else
                lineOomnitza += ",\"" + rows.Item(HRrowCount)(colHR_City) + "\"";

            if (colHR_eeID < 0)
                employeeID = "";
            else
            {
                // always use WDID now
                employeeID = rows(HRrowCount)(colHR_eeID);
                if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                {
                    // users in ignore update list will use current AD property for manager and city
                    if (colAD_managerID >= 0)
                        colHR_managerID = colAD_managerID;
                    if (colAD_City >= 0)
                        colHR_City = colAD_City;
                }
            }

            string managerName = "";
            ManagerWorkEmail = "";
            if (colHR_managerID >= 0)
            {
                // should consider manager account has email, but AD may not set manager?
                ManagerID = rows.Item(HRrowCount)(colHR_managerID);
                while ((ManagerID.Length < 6))
                    ManagerID = "0" + ManagerID;
                int ADindex = eeIDlist.IndexOf(ManagerID);
                if (ADindex >= 0)
                {
                    // ManagerWorkEmail = WorkemailList(ADindex)
                    string mFirstName = firstNameList(ADindex);
                    string mLastName = lastNameList(ADindex);
                    managerName = mFirstName + " " + mLastName;
                }
            }

            // If InStr(ManagerWorkEmail, "@") > 0 Then
            // managerName = Microsoft.VisualBasic.Left(LCase(ManagerWorkEmail), InStr(ManagerWorkEmail, "@") - 1)
            // Else
            // managerName = ManagerWorkEmail
            // End If
            // managerName = managerName.Replace(".", " ")

            lineOomnitza += ",\"" + managerName + "\"";
            if (colHR_employeeType < 0)
                lineOomnitza += ",";
            else
                lineOomnitza += ",\"" + rows.Item(HRrowCount)(colHR_employeeType) + "\"";
            if (colHR_FTE < 0)
                lineOomnitza += ",,";
            else
            {
                string workerType = rows(HRrowCount)(colHR_FTE);
                string ehiComputer = "";
                if (workerType == "Consultant_TPA" || workerType == "Board_of_Directors")
                    ehiComputer = "No";
                else
                    // Regular Full_time, Regular Part_time, Seasonal Full_time, Intern Full_time
                    // If workerType = "Independent_Contractor" Then ehiComputer = ""
                    ehiComputer = "Yes";
                lineOomnitza += ",\"" + workerType + "\",\"" + ehiComputer + "\"";
            }

            swOomnitza.WriteLine(lineOomnitza);

            HRrowCount += 1;
            continue;



            if (colHR_locationAux1 < 0)
                LocationAux1 = "";
            else
                LocationAux1 = Trim(rows.Item(HRrowCount)(colHR_locationAux1));


            if (colHR_nickName < 0)
                NickName = "";
            else
                NickName = Trim(rows.Item(HRrowCount)(colHR_nickName));
            string fullName = "";
            if (NickName == "")
            {
                // no nickname, so use legal first
                if (colHR_firstName >= 0)
                    fullName = rows.Item(HRrowCount)(colHR_firstName);
            }
            else
                fullName = NickName;
            if (colHR_lastName >= 0)
                fullName += " " + rows.Item(HRrowCount)(colHR_lastName);


            if (colHR_OriginalHireDate < 0)
            {
            }
            else
            {
            }
            if (colHR_HireDate < 0)
            {
            }
            else
            {
            }


            if (colHR_Jobtitle < 0)
                // lineOomnitza &= ","
                Jobtitle = "";
            else
                Jobtitle = rows.Item(HRrowCount)(colHR_Jobtitle);
            if (colHR_BusinessUnit < 0)
            {
            }
            else
            {
            }

            if (colHR_HomeDepartment < 0)
            {
            }
            else
            {
            }
            HRrowCount += 1;
        }
    }

    catch (Exception ex)
    {
        input.Close();
        if (swLog != null)
            swLog.WriteLine(ex.ToString());
        return false;
    }


    input.Close();

    swOomnitza.Close();

    Console.WriteLine(strExportFileFound + ": " + HRrowCount + " records convert to " + strReportOomnitza + " successfully.");
    // Console.ReadKey()
    return true;
}

private static string ADfindManagerEmail(string email)
{
    int ADindex = WorkemailList.IndexOf(email);
    if (ADindex < 0)
        return "";
    ManagerDN = managerList(ADindex);
    if (ManagerDN == "" || ManagerDN.Contains(CEO))
        return "";
    int Mindex = userDNlist.IndexOf(ManagerDN);
    if (Mindex < 0)
        return "";
    return (WorkemailList(Mindex));
}

private static bool isValidEmail(string pAddress)
{
    if (pAddress == "root")
        return true;
    if (!pAddress.Contains("@"))
        return false;
    if (pAddress.Contains(" "))
        return false;
    pAddress = Strings.LCase(pAddress);
        private string strRFC2822 = @"[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!" + "#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:" + @"[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:" + "[a-z0-9-]*[a-z0-9])?";
Regex reg = new Regex(strRFC2822);

bool vOK = false;
vOK = reg.Match(pAddress).Success;

return vOK;
    }

    private static string checkMinorMismatch(string fnAD, string lnAD, string fnHR, string lnHR, string nickHR)
{
    fnAD = Strings.Trim(fnAD).Replace("  ", " ").Replace("'", "");
    fnAD = Strings.UCase(Strings.Trim(fnAD).Replace("  ", " ").Replace(" ", "-"));
    lnAD = Strings.Trim(lnAD).Replace("  ", " ").Replace("'", "");
    lnAD = Strings.UCase(Strings.Trim(lnAD).Replace("  ", " ").Replace(" ", "-"));
    fnHR = Strings.Trim(fnHR).Replace("  ", " ").Replace("'", "");
    fnHR = Strings.UCase(Strings.Trim(fnHR).Replace("  ", " ").Replace(" ", "-"));
    lnHR = Strings.Trim(lnHR).Replace("  ", " ").Replace("'", "");
    lnHR = Strings.UCase(Strings.Trim(lnHR).Replace("  ", " ").Replace(" ", "-"));
    nickHR = Strings.Trim(nickHR).Replace("  ", " ").Replace("'", "");
    nickHR = Strings.UCase(Strings.Trim(nickHR).Replace("  ", " ").Replace(" ", "-"));

    if ((fnAD == fnHR || fnAD == nickHR))
    {
        if (lnAD == lnHR)
            return minor;
    }
    else
        return major;

    return "";
}

private static void getHRactiveEEIDlist()
{
    string Sql;
    ClassExecuteSQL exeSQL = new ClassExecuteSQL();
    exeSQL.ITDBname = ITDBname;

    HRactiveEEIDlist.Clear();
    HRactiveREQIDlist.Clear();
    HRterminatedEEIDlist.Clear();
    HRterminatedREQIDlist.Clear();
    // always use PROD DB for OnLeave
    Sql = "select employeeID,requisitionID,status from [ITauto].[dbo].[ADemployee]";
    // Sql &= "where status='Active' or status='Leave'"

    exeSQL.executeSQL(Sql, "select");
    if (exeSQL.errStr != "")
        return;
    if (exeSQL.intCount < 1)
        return;

    for (int i = 0; i <= exeSQL.tableResult.Rows.Count - 1; i++)
    {
        DataRow drow;
        drow = exeSQL.tableResult.Rows(i);

        string eeID = "";
        string reqID = "";
        string status = "";
        if (!IsDBNull(drow(0)))
            eeID = UCase(Trim(drow(0)));
        if (!IsDBNull(drow(1)))
            reqID = UCase(Trim(drow(1)));
        if (!IsDBNull(drow(2)))
            status = LCase(Trim(drow(2)));
        if (status == "active" || status == "leave")
        {
            HRactiveEEIDlist.Add(eeID);
            HRactiveREQIDlist.Add(reqID);
        }
        else if (status != "")
        {
            HRterminatedEEIDlist.Add(eeID);
            HRterminatedREQIDlist.Add(reqID);
        }
    }
    return;
}

private static bool InsertOrUpdateDisabledRecord(string aSamID, string aDN)
{
    string sql;
    // if exists, just update
    sql = "select samID,disabledTime from " + "[" + ITDBname + "].[dbo].[" + disabledTalbeName + "] ";
    sql += "where samID='" + aSamID + "'";
    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(sql, "select"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return false;
    }
    if (executeSQL.tableResult.Rows.Count > 0)
    {
        // just update
        sql = "update " + "[" + ITDBname + "].[dbo].[" + disabledTalbeName + "] set ";
        sql += "disabledTime='" + DateTime.Now + "'";
        sql += " where samID='" + aSamID + "'";

        if (!executeSQL.executeSQL(sql, "update"))
        {
            hasError = true;
            swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
            return false;
        }
        return true;
    }

    sql = "insert into " + "[" + ITDBname + "].[dbo].[" + disabledTalbeName + "] ";
    sql += "(Tool,eeID,samID,firstName,lastName,disabledTime,DN)";
    sql += " values (";
    sql += "'" + appNameVersion + " " + runnerName + "','" + employeeID + "','" + aSamID + "','" + FirstName.Replace("'", "''") + "','" + LastName.Replace("'", "''") + "',";
    sql += "'" + DateTime.Now + "','" + aDN.Replace("'", "''") + "')";

    if (!executeSQL.executeSQL(sql, "insert"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return false;
    }
    return true;
}

private static void deleteDisabledDaysAgo()
{
    string sql;
    sql = "delete from " + "[" + ITDBname + "].[dbo].[" + disabledTalbeName + "] ";
    sql += "where disabledTime < (GETDATE()-" + (DaysCanRemoveManager + 1) + ")";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(sql, "delete"))
    {
        hasError = true;
        swLog.WriteLine(executeSQL.errStr.Replace("<br/>", privateants.vbCrLf));
        return;
    }
    return;
}

private static bool getDisabledSamidFromDB()
{
    DB_DisabledSamIDs.Clear();

    string sql;
    sql = "select eeID,samID,disabledTime from " + "[" + ITDBname + "].[dbo].[" + DisabledTalbeName + "] ";
    sql += "where samID is not null and disabledTime >= (GETDATE()- " + DaysCanRemoveManager + ")";

    ClassExecuteSQL executeSQL = new ClassExecuteSQL();
    if (!executeSQL.executeSQL(sql, "select"))
    {
        tmpStr = "MS SQL database has problem, please check: " + executeSQL.errStr.Replace("<br/>", privateants.vbCrLf);
        swLog.WriteLine(tmpStr);
        return false;
    }

    if (executeSQL.tableResult.Rows.Count > 0)
    {
        for (int db = 0; db <= executeSQL.tableResult.Rows.Count - 1; db++)
        {
            DataRow drow;
            drow = executeSQL.tableResult.Rows(db);
            DB_DisabledSamIDs.Add(LCase(Trim(drow(1))));
        }
    }
    return true;
}

private static bool saveDB3(bool XMonly)
{
    errStr = "";
    if (!needSaveEEinfo2DB)
        return true;
    if (WDID.Count < 1)
    {
        errStr = "No HR record that already read for saving to DB.";
        return false;
    }
    // create table [ITtest].[dbo].[ADemployee] (employeeID char(15) null, requisitionID char(18) null,
    // firstName char(30) not null, lastName char(30) not null, preferredName char(25) not null,
    // status char(12) not null, Location char(6) null, notes char(50) null, updatedTime datetime not null DEFAULT (getdate()))

    swLog.WriteLine(DateTime.Now + " save to DB for " + HRrowCount + " records...");
    bool hasError = false;
    string sql = "";
    ClassExecuteSQL exeSQL = new ClassExecuteSQL();
    exeSQL.ITDBname = ITDBname;
    sql = "select [employeeID],[requisitionID],[firstName],[lastName],[preferredName],[status],[Location]";
    sql += " from [" + ITDBname + "].[dbo].[ADemployee]";
    if (!exeSQL.executeSQL(sql, "select"))
    {
        errStr = exeSQL.errStr;
        hasError = true;
        return false;
    }
    int DBcount = exeSQL.tableResult.Rows.Count;
    ArrayList eeIDlist = new ArrayList(), preIDlist = new ArrayList(), fnameList = new ArrayList(), lNameList = new ArrayList(), pNameList = new ArrayList(), statusList = new ArrayList(), locationList = new ArrayList();
    for (int i = 0; i <= DBcount - 1; i++)
    {
        DataRow drow;
        drow = exeSQL.tableResult.Rows(i);
        if (!IsDBNull(drow(0)))
        {
            eeIDlist.Add(UCase(Trim(drow(0))));
            if (IsDBNull(drow(1)))
                preIDlist.Add("");
            else
                preIDlist.Add(UCase(Trim(drow(1))));
            if (IsDBNull(drow(2)))
                fnameList.Add("");
            else
                fnameList.Add(Trim(drow(2)));
            if (IsDBNull(drow(3)))
                lNameList.Add("");
            else
                lNameList.Add(Trim(drow(3)));
            if (IsDBNull(drow(4)))
                pNameList.Add("");
            else
                pNameList.Add(Trim(drow(4)));
            if (IsDBNull(drow(5)))
                statusList.Add("");
            else
                statusList.Add(Trim(drow(5)));
            if (IsDBNull(drow(6)))
                locationList.Add("");
            else
                locationList.Add(Trim(drow(6)));
        }
    }

    hasError = false;

    int totalCount = 0;
    int intOKcount = 0;
    int insertCount = 0;
    string insertLog = "";
    string updateLog = "";

    for (int i = 0; i <= WDID.Count - 1; i++)
    {
        string locationAux1 = "";
        locationAux1 = UCase(HR_locationAux1(i));
        if (locationAux1 != "")
            locationAux1 = mapLocation_Aux1(locationAux1);
        if (locationAux1 == "")
            continue;
        else if (forXM_US != "")
        {
            if (forXM_US == "US")
            {
                if (locationAux1 == "XM" || locationAux1.Contains("XIAMEN"))
                    continue;
            }
            else if (forXM_US == "XM")
            {
                if (locationAux1 != "XM" && !locationAux1.Contains("XIAMEN"))
                    continue;
            }
        }

        string status = "";
        status = Trim(HR_status(i));

        string pID = "";
        pID = Trim(UCase(previousID(i)));

        string aWDID = "";
        aWDID = Trim(UCase(WDID(i)));
        if (aWDID != "")
        {
            while ((aWDID.Length < 6))
                aWDID = "0" + aWDID;
        }

        string reqID = "";
        reqID = Trim(UCase(previousID(i)));

        string firstname = "";
        firstname = Trim(HR_firstName(i));

        string lastname = "";
        lastname = Trim(HR_lastName(i));
        string preferredname = "";
        preferredname = Trim(HR_nickName(i));
        int index = eeIDlist.IndexOf(aWDID);
        // If index < 0 Then
        // index = eeIDlist.IndexOf(reqID)
        // End If
        if (index < 0)
        {
            sql = "insert into [" + ITDBname + "].[dbo].[ADemployee] ([employeeID],[requisitionID],[firstName]";
            sql += ",[lastName],[preferredName],[status],[Location]";
            sql += ") values (";
            sql += "'" + aWDID + "'";
            sql += ",'" + reqID + "'";
            sql += ",'" + firstname.Replace("'", "''") + "'"; // firstName
            sql += ",'" + lastname.Replace("'", "''") + "'"; // lastName
            sql += ",'" + preferredname.Replace("'", "''") + "'"; // preferredName
            sql += ",'" + status + "'";
            sql += ",'" + locationAux1 + "'";
            sql += ")" + privateants.vbCrLf;

            insertCount += 1;
            totalCount += 1;
            insertLog += aWDID + "(" + firstname;
            if (preferredname != "")
                insertLog += "[" + preferredname + "]";
            insertLog += " " + lastname + "," + locationAux1 + ")" + ". ";

            intOKcount = exeSQL.executeSQL(sql, "insert");
            if (exeSQL.errStr != "")
            {
                errStr = exeSQL.errStr;
                hasError = true;
                swLog.WriteLine(exeSQL.errStr);
            }
        }
        else
        {
            if ((reqID == "" || reqID == preIDlist(index)) && firstname == fnameList(index) && lastname == lNameList(index) && preferredname == pNameList(index) && status == statusList(index) && locationAux1.Contains(locationList(index)))
                continue;
            sql = "update [" + ITDBname + "].[dbo].[ADemployee] set ";
            // sql &= "[employeeID]='" & aWDID & "'," 'replace previousID with aWDID
            sql += "[requisitionID]='" + reqID + "'";
            sql += ",[firstName]='" + firstname.Replace("'", "''") + "'";
            sql += ",[lastName]='" + lastname.Replace("'", "''") + "'";
            sql += ",[preferredName]='" + preferredname.Replace("'", "''") + "'";
            sql += ",[status]='" + status + "'";
            sql += ",[Location]='" + locationAux1 + "'";
            sql += ",[updatedTime]='" + DateTime.Now + "'";
            sql += " where employeeID='" + aWDID + "'";
            if (reqID != "")
                sql += " or employeeID='" + reqID + "'";

            totalCount += 1;
            updateLog += aWDID + "(" + firstname;
            if (preferredname != "")
                updateLog += "[" + preferredname + "]";
            updateLog += " " + lastname + "," + status + "," + locationAux1 + ")" + ". ";

            intOKcount = exeSQL.executeSQL(sql, "update");
            if (exeSQL.errStr != "")
            {
                hasError = true;
                swLog.WriteLine("execute SQL error");
                errStr = exeSQL.errStr;

                errStr += "<br/><br/>DBcount= " + DBcount + "<br/>";
                errStr += "HRcount= " + HRrowCount + "<br/>";
                errStr += "DB - HR= " + (DBcount - HRrowCount) + "<br/><br/>";
                errStr += "insertCount= " + insertCount + ":<br/>";
                errStr += insertLog + "<br/><br/>";
                errStr += "updateCount= " + (totalCount - insertCount) + ":<br/>";
                errStr += updateLog + "<br/><br/>";
                errStr += "NoChangeCount= " + (DBcount - totalCount) + "<br/>";
            }
        }
    }

    if (hasError)
        return false;

    swLog.WriteLine(DateTime.Now + " save EE info to " + ITDBname + " OK");
    string bodyText = "";
    if (XMonly)
        bodyText += "for XM only " + "<br/>";
    else
        bodyText += "for US only " + "<br/>";

    bodyText += "DBcount= " + DBcount + "<br/>";
    bodyText += "HRcount= " + HRrowCount + "<br/>";
    bodyText += "DB - HR= " + (DBcount - HRrowCount) + "<br/><br/>";
    bodyText += "insertCount= " + insertCount + ":<br/>";
    bodyText += insertLog + "<br/><br/>";
    bodyText += "updateCount= " + (totalCount - insertCount) + ":<br/>";
    bodyText += updateLog + "<br/><br/>";
    bodyText += "NoChangeCount= " + (DBcount - totalCount) + "<br/>";

    sendAlertEmail("Changtu.Wang@ehealth.com", "", hrMIS + " employee info save to DB " + ITDBname + " successfully", bodyText, "");

    return true;
}

private static string mapAux1_location(object aux1)
{
    if (aux1 != "")
    {
        switch (aux1)
        {
            case "RMT":
                {
                    aux1 = "REMOTE"; // USA Remote; TPA: Remote
                    break;
                }

            case "SCS":
            case "MTNV":
                {
                    aux1 = "SANTA CLARA";
                    break;
                }

            case "SF":
                {
                    aux1 = "SAN FRANCISCO";
                    break;
                }

            case "GDRV":
                {
                    aux1 = "GOLD RIVER";
                    break;
                }

            case "UT":
                {
                    aux1 = "SALT LAKE CITY";
                    break;
                }

            case "XM":
                {
                    aux1 = "XIAMEN";
                    break;
                }

            case "DC":
                {
                    aux1 = "WASHINGTON";
                    break;
                }

            case "IN":
                {
                    aux1 = "INDIANAPOLIS";
                    break;
                }

            case "MA":
                {
                    aux1 = "WESTFORD";
                    break;
                }

            default:
                {
                    if (aux1.startswith("TPA:"))
                        aux1 = "REMOTE";
                    break;
                }
        }
    }
    return aux1;
}

private static string mapLocation_Aux1(object location)
{
    if (location != "")
    {
        switch (location)
        {
            case "USA REMOTE":
            case "TPA: REMOTE":
                {
                    location = "RMT";
                    break;
                }

            case "SANTA CLARA, CA" // ,"MOUNTAIN VIEW, CA"
     :
                {
                    location = "SCS";
                    break;
                }

            case "SAN FRANCISCO, CA":
                {
                    location = "SF";
                    break;
                }

            case "GOLD RIVER, CA":
                {
                    location = "GDRV";
                    break;
                }

            case "SALT LAKE CITY, UT":
                {
                    location = "UT";
                    break;
                }

            case "XIAMEN, CHINA":
                {
                    location = "XM";
                    break;
                }

            case "WASHINGTON D.C.":
                {
                    location = "DC";
                    break;
                }

            case "AUSTIN, TX":
                {
                    location = "AUSTIN";
                    break;
                }

            case "INDIANAPOLIS, IN":
                {
                    location = "IN";
                    break;
                }

            default:
                {
                    break;
                }
        }
    }
    return location;
}

public static bool openCSVFileForNoNeedAD(string aFileName)
{
    NoNeedAD_EEIDs.Clear();

    if (aFileName == null)
        return false;
    object input = null;
    int tryTimes = 0;
    while (input == null)
    {
        try
        {
            tryTimes += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(aFileName);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTimes > 9)
        {
            tmpStr = "AD sync: Open file time out!" + aFileName;
            swLog.WriteLine(tmpStr);
            sendAlertEmail("Changtu.Wang@ehealth.com", "", tmpStr, tmpStr, "");
            break;
        }
    }
    if (input == null)
        return false;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    int col_EEID = -1;

    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        // aTitle = aTitle.Replace("#", "")
        if (aTitle == "employeeid")
            col_EEID = HRcolCount;
        HRcolCount = HRcolCount + 1;
    }

    if (HRcolCount < 1)
    {
        swLog.WriteLine(aFileName + ": " + " has no column.");
        return false;
    }
    int rowCount = 0;
    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);
            if (col_EEID < 0)
                NoNeedAD_EEIDs.Add("");
            else
            {
                // deal with strange EEID that may be contains ","
                tmpStr = rows.Item(rowCount)(col_EEID).replace(",", "_");
                while ((tmpStr.Length < 6))
                    tmpStr = "0" + tmpStr;
                NoNeedAD_EEIDs.Add(tmpStr);
            }
            rowCount += 1;
        }
    }
    catch (Exception ex)
    {
        input.Close();
        sendAlertEmail("Changtu.Wang@ehealth.com", "", "AD sync: reading file NoNeedAD.csv failed!", "please check NoNeedAD.csv", "");
        swLog.WriteLine(tmpStr);
        return false;
    }
    input.Close();
    return true;
}

private static bool isPDT(DateTime theDate)
{
    // 美国太平洋时区的夏令时间始于每年三月的第二个星期六深夜二时正，并终于每年十一月的第一个星期日深夜二时正
    DateTime beginDay = new DateTime(), endDay = new DateTime();
    // theDate = "2020/3/10 1:59"
    string yyyy = DateTime.Year(theDate);

    // determine the 2nd Saturday of March
    int wantDayCount = 0;
    for (int i = 1; i <= 14; i++)
    {
        beginDay = yyyy + "/3/" + i + " 2:00";
        if (DateTime.Weekday(beginDay) == 3)
        {
            wantDayCount += 1;
            if (wantDayCount == 2)
                break;
        }
    }

    // determine the 1st Sunday of November
    for (int i = 1; i <= 7; i++)
    {
        endDay = yyyy + "/11/" + i + " 2:00";
        if (DateTime.Weekday(endDay) == 1)
            break;
    }

    if (theDate >= beginDay && theDate < endDay)
        // MsgBox(beginDay & " - " & endDay & ": " & theDate & " is " & timeDifferenceXM)
        return true;

    return false;
}

private static bool saveDB4()
{
    // if no-show, Workday no longer export record, then DB should delete the no-show record

    errStr = "";
    if (!needSaveEEinfo2DB)
        return true;
    if (WDID.Count < 1)
    {
        errStr = "No HR record that already read for saving to DB.";
        return false;
    }
    // create table [ITtest].[dbo].[ADemployee] (employeeID char(15) null, requisitionID char(18) null,
    // firstName char(30) not null, lastName char(30) not null, preferredName char(25) not null,
    // status char(12) not null, Location char(6) null, notes char(50) null, updatedTime datetime not null DEFAULT (getdate()))

    swLog.WriteLine(DateTime.Now + " save to DB for " + HRrowCount + " records...");
    bool hasError = false;
    string sql = "";
    ClassExecuteSQL exeSQL = new ClassExecuteSQL();
    exeSQL.ITDBname = ITDBname;
    sql = "select [employeeID],[requisitionID],[firstName],[lastName],[preferredName],[status],[Location]";
    sql += " from [" + ITDBname + "].[dbo].[ADemployee]";
    if (!exeSQL.executeSQL(sql, "select"))
    {
        errStr = exeSQL.errStr;
        hasError = true;
        return false;
    }
    int DBcount = exeSQL.tableResult.Rows.Count;
    ArrayList eeIDlist = new ArrayList(), preIDlist = new ArrayList(), fnameList = new ArrayList(), lNameList = new ArrayList(), pNameList = new ArrayList(), statusList = new ArrayList(), locationList = new ArrayList();
    for (int i = 0; i <= DBcount - 1; i++)
    {
        DataRow drow;
        drow = exeSQL.tableResult.Rows(i);
        if (!IsDBNull(drow(0)))
        {
            eeIDlist.Add(UCase(Trim(drow(0))));
            if (IsDBNull(drow(1)))
                preIDlist.Add("");
            else
                preIDlist.Add(UCase(Trim(drow(1))));
            if (IsDBNull(drow(2)))
                fnameList.Add("");
            else
                fnameList.Add(Trim(drow(2)));
            if (IsDBNull(drow(3)))
                lNameList.Add("");
            else
                lNameList.Add(Trim(drow(3)));
            if (IsDBNull(drow(4)))
                pNameList.Add("");
            else
                pNameList.Add(Trim(drow(4)));
            if (IsDBNull(drow(5)))
                statusList.Add("");
            else
                statusList.Add(Trim(drow(5)));
            if (IsDBNull(drow(6)))
                locationList.Add("");
            else
                locationList.Add(Trim(drow(6)));
        }
    }

    hasError = false;

    int totalCount = 0;
    int intOKcount = 0;
    int insertCount = 0;
    int deleteCount = 0;
    string insertLog = "";
    string updateLog = "";
    string deleteLog = "";

    for (int i = 0; i <= WDID.Count - 1; i++)
    {
        string locationAux1 = "";
        locationAux1 = UCase(HR_locationAux1(i));
        if (locationAux1 != "")
            locationAux1 = mapLocation_Aux1(locationAux1);
        if (locationAux1 == "")
            continue;
        else if (forXM_US != "")
        {
            if (forXM_US == "US")
            {
                if (locationAux1 == "XM" || locationAux1.Contains("XIAMEN"))
                    continue;
            }
            else if (forXM_US == "XM")
            {
                if (locationAux1 != "XM" && !locationAux1.Contains("XIAMEN"))
                    continue;
            }
        }

        string status = "";
        status = Trim(HR_status(i));

        string pID = "";
        pID = Trim(UCase(previousID(i)));

        string aWDID = "";
        aWDID = Trim(UCase(WDID(i)));
        if (aWDID != "")
        {
            while ((aWDID.Length < 6))
                aWDID = "0" + aWDID;
        }

        string reqID = "";
        reqID = Trim(UCase(previousID(i)));

        string firstname = "";
        firstname = Trim(HR_firstName(i));

        string lastname = "";
        lastname = Trim(HR_lastName(i));
        string preferredname = "";
        preferredname = Trim(HR_nickName(i));
        int index = eeIDlist.IndexOf(aWDID);
        // If index < 0 Then
        // index = eeIDlist.IndexOf(reqID)
        // End If
        if (index < 0)
        {
            sql = "insert into [" + ITDBname + "].[dbo].[ADemployee] ([employeeID],[requisitionID],[firstName]";
            sql += ",[lastName],[preferredName],[status],[Location]";
            sql += ") values (";
            sql += "'" + aWDID + "'";
            sql += ",'" + reqID + "'";
            sql += ",'" + firstname.Replace("'", "''") + "'"; // firstName
            sql += ",'" + lastname.Replace("'", "''") + "'"; // lastName
            sql += ",'" + preferredname.Replace("'", "''") + "'"; // preferredName
            sql += ",'" + status + "'";
            sql += ",'" + locationAux1 + "'";
            sql += ")" + privateants.vbCrLf;

            insertCount += 1;
            totalCount += 1;
            insertLog += aWDID + "(" + firstname;
            if (preferredname != "")
                insertLog += "[" + preferredname + "]";
            insertLog += " " + lastname + "," + locationAux1 + ")" + ". ";

            intOKcount = exeSQL.executeSQL(sql, "insert");
            if (exeSQL.errStr != "")
            {
                errStr = exeSQL.errStr;
                hasError = true;
                swLog.WriteLine(exeSQL.errStr);
            }
        }
        else
        {
            if ((reqID == "" || reqID == preIDlist(index)) && firstname == fnameList(index) && lastname == lNameList(index) && preferredname == pNameList(index) && status == statusList(index) && locationAux1.Contains(locationList(index)))
                continue;
            sql = "update [" + ITDBname + "].[dbo].[ADemployee] set ";
            // sql &= "[employeeID]='" & aWDID & "'," 'replace previousID with aWDID
            sql += "[requisitionID]='" + reqID + "'";
            sql += ",[firstName]='" + firstname.Replace("'", "''") + "'";
            sql += ",[lastName]='" + lastname.Replace("'", "''") + "'";
            sql += ",[preferredName]='" + preferredname.Replace("'", "''") + "'";
            sql += ",[status]='" + status + "'";
            sql += ",[Location]='" + locationAux1 + "'";
            sql += ",[updatedTime]='" + DateTime.Now + "'";
            sql += " where employeeID='" + aWDID + "'";
            if (reqID != "")
                sql += " or employeeID='" + reqID + "'";

            totalCount += 1;
            updateLog += aWDID + "(" + firstname;
            if (preferredname != "")
                updateLog += "[" + preferredname + "]";
            updateLog += " " + lastname + "," + status + "," + locationAux1 + ")" + ". ";

            intOKcount = exeSQL.executeSQL(sql, "update");
            if (exeSQL.errStr != "")
            {
                hasError = true;
                swLog.WriteLine("execute SQL error");
                errStr = exeSQL.errStr;

                errStr += "<br/><br/>DBcount= " + DBcount + "<br/>";
                errStr += "HRcount= " + HRrowCount + "<br/>";
                errStr += "DB - HR= " + (DBcount - HRrowCount) + "<br/><br/>";
                errStr += "insertCount= " + insertCount + ":<br/>";
                errStr += insertLog + "<br/><br/>";
                errStr += "updateCount= " + (totalCount - insertCount) + ":<br/>";
                errStr += updateLog + "<br/><br/>";
                errStr += "NoChangeCount= " + (DBcount - totalCount) + "<br/>";
            }
        }
    }

    if (hasError)
        return false;

    for (int i = 0; i <= eeIDlist.Count - 1; i++)
    {
        string theEEID = eeIDlist(i);
        int index = WDID.IndexOf(theEEID);
        if (index >= 0)
            continue;
        string theStatus = statusList(i);
        if (theStatus != "Active")
            continue;
        if (index < 0)
        {
            if (preIDlist(i) != "")
            {
                index = previousID.IndexOf(preIDlist(i));
                if (index >= 0)
                    continue;
            }

            sql = "delete from [" + ITDBname + "].[dbo].[ADemployee]";
            sql += " where employeeID='" + theEEID + "'";

            deleteCount += 1;
            deleteLog += theEEID + "(" + fnameList(i);
            if (pNameList(i) != "")
                deleteLog += "[" + pNameList(i) + "]";
            deleteLog += " " + lNameList(i) + "," + statusList(i) + "," + locationList(i) + ")" + ". ";
            intOKcount = exeSQL.executeSQL(sql, "delete");
            if (exeSQL.errStr != "")
            {
                hasError = true;
                swLog.WriteLine("execute SQL error");
                errStr = exeSQL.errStr;
            }
        }
    }

    swLog.WriteLine(DateTime.Now + " save EE info to " + ITDBname + " OK");
    string bodyText = "";

    bodyText += "DBcount= " + DBcount + "<br/>";
    bodyText += "HRcount= " + HRrowCount + "<br/>";
    bodyText += "DB - HR= " + (DBcount - HRrowCount) + "<br/><br/>";
    bodyText += "insertCount= " + insertCount + ":<br/>";
    bodyText += insertLog + "<br/><br/>";
    bodyText += "updateCount= " + (totalCount - insertCount) + ":<br/>";
    bodyText += updateLog + "<br/><br/>";

    bodyText += "deleteCount= " + deleteCount + ":<br/>";
    bodyText += deleteLog + "<br/><br/>";

    bodyText += "NoChangeCount= " + (DBcount - totalCount - deleteCount) + "<br/>";
    bodyText += "<br/>" + errStr + "<br/>";

    sendAlertEmail("Changtu.Wang@ehealth.com", "", hrMIS + " employee info save to DB " + ITDBname + " successfully", bodyText, "");

    if (insertCount > 0 || deleteCount > 0 || ((totalCount - insertCount) > 0))
    {
        StreamWriter DBlog;
        try
        {
            DBlog = new StreamWriter(strWorkFolder + @"\" + strLogSubFolder + @"\" + "DBlog" + DateTime.Now.ToString("yyyyMMddHHmm") + ".txt", false);
            DBlog.WriteLine(bodyText.Replace("<br/>", privateants.vbCrLf));
            DBlog.Close();
        }
        catch (Exception ex)
        {
        }
    }

    return true;
}

private static bool putFileToSharedOomnitza()
{
    if (forTest)
        return true;
    if (!File.Exists(strReportOomnitza))
    {
        Console.WriteLine(DateTime.Now + " file does not exist: " + strReportOomnitza);
        return false;
    }
    string remoteFilePath = @"\\xmprspapp01/CSV_Files/" + CSVfilenameOomnitza;
    try
    {
        if (File.Exists(remoteFilePath))
        {
            DateTime fileTime = FileSystem.FileDateTime(remoteFilePath);
            string backupFileName = remoteFilePath.Replace(".csv", fileTime.ToString("yyyyMMddHHmm") + ".csv");
            if (!File.Exists(backupFileName))
            {
                try
                {
                    File.Copy(remoteFilePath, backupFileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(DateTime.Now + " backup Oomnitza AD file error: " + ex.ToString());
                }
            }

            try
            {
                File.Delete(remoteFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now + " delete existing file failed! " + ex.ToString());
                return false;
            }
        }

        // File.Copy(strReportOomnitza, "H:/" & CSVfilenameOomnitza)
        // File.Copy(strReportOomnitza, "\\xmprspapp01/CSV_Files/" & CSVfilenameOomnitza)
        File.Copy(strReportOomnitza, remoteFilePath);
        // File.Copy(strReportOomnitza, "T:\\IT\\" & CSVfilenameOomnitza) 'OK
        Console.WriteLine(DateTime.Now + " put file to Oomnitza OK: " + CSVfilenameOomnitza);
    }
    catch (Exception ex)
    {
        Console.WriteLine(DateTime.Now + " put file to Oomnitza error: " + ex.ToString());
        try
        {
            File.Copy(strReportOomnitza, @"\\xmprspapp01\\CSV_Files\\OomnitzaAD" + DateTime.Now.ToString("yyyyMMddHHmm") + ".csv");
        }
        catch (Exception ex2)
        {
        }
        return false;
    }
    return true;
}

private static bool putFileToSharedCoupa()
{
    if (forTest)
        return true;
    if (!File.Exists(strReportCoupa))
    {
        Console.WriteLine(DateTime.Now + " file does not exist: " + strReportCoupa);
        return false;
    }
    string remoteFilePath = @"\\awsentgpdb01\Dynshare\Import\Scribe\Coupa\Production\" + CSVfilenameCoupa; // "\\sjentgpdb01\Dynshare\Import\Scribe\Coupa\Production\" & CSVfilenameCoupa
                                                                                                           // Dim remoteFilePath As String = "\\awprnas01\EHI_Shares\EHI\Private\Dept\IT\Coupa\Production\" & CSVfilenameCoupa
    try
    {
        if (File.Exists(remoteFilePath))
        {
            DateTime fileTime = FileSystem.FileDateTime(remoteFilePath);
            string backupFileName = remoteFilePath.Replace(".csv", fileTime.ToString("yyyyMMddHHmm") + ".csv");
            if (!File.Exists(backupFileName))
            {
                try
                {
                    File.Copy(remoteFilePath, backupFileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(DateTime.Now + " backup Coupa AD file error: " + ex.ToString());
                }
            }

            try
            {
                File.Delete(remoteFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now + " delete existing file failed! " + ex.ToString());
                return false;
            }
        }

        // File.Copy(strReportCoupa, "H:/" & CSVfilenameCoupa)
        // File.Copy(strReportCoupa, "\\xmprspapp01/CSV_Files/" & CSVfilenameCoupa)
        File.Copy(strReportCoupa, remoteFilePath);
        // File.Copy(strReportCoupa, "T:\\IT\\" & CSVfilenameCoupa) 'OK
        Console.WriteLine(DateTime.Now + " put file to Coupa OK: " + CSVfilenameCoupa);
    }
    catch (Exception ex)
    {
        Console.WriteLine(DateTime.Now + " put file to Coupa error: " + ex.ToString());
        return false;
    }
    return true;
}

private static bool prepareCoupa()
{
    if (hasWriteEntireError)
    {
    }

    if (!File.Exists(strExportFileEntire))
        return false;
    object input = null;
    int tryTime = 0;
    while (input == null)
    {
        try
        {
            tryTime += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(strExportFileEntire);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTime > 10)
        {
            tmpStr = "AD sync: Open file time out!" + strExportFileEntire;
            // swLog.WriteLine(tmpStr)
            // sendAlertEmail("Changtu.Wang@ehealth.com", "", tmpStr, tmpStr, "")
            break;
        }
    }
    if (input == null)
        return false;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    colHR_firstName = -1;
    colHR_lastName = -1;
    colHR_workEmail = -1;
    // Dim colHR_WDID As Integer = -1
    colHR_eeID = -1;
    colHR_samAccountName = -1;
    colHR_OriginalHireDate = -1;
    colHR_HireDate = -1;
    colHR_statusEffDate = -1;
    colHR_status = -1;
    colHR_managerID = -1;
    int colAD_managerID = -1;
    colHR_Jobtitle = -1;
    int colCostCenter = -1;
    int colNoNeedAccount = -1;

    colHR_HomeDepartment = -1;
    int colAD_City = -1;
    colHR_State = -1;
    colHR_locationAux1 = -1;

    colHR_FTE = -1;
    colHR_nickName = -1;

    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        // aTitle = aTitle.Replace("#", "")
        if (aTitle == "lastname")
            colHR_lastName = HRcolCount;
        else if (aTitle == "firstname")
            colHR_firstName = HRcolCount;
        else if (aTitle == "nickname")
            colHR_nickName = HRcolCount;
        else if (aTitle == "middlename")
            colHR_middleName = HRcolCount;
        else if (aTitle == "workemail_ad")
            colHR_workEmail = HRcolCount;
        else if (aTitle == "employeeid")
            // always use WDID now
            colHR_eeID = HRcolCount;
        else if (aTitle == "samaccountname")
            colHR_samAccountName = HRcolCount;
        else if (aTitle == "status")
            colHR_status = HRcolCount;
        else if (aTitle == "managerid")
            colHR_managerID = HRcolCount;
        else if (aTitle == "managerid_ad")
            colAD_managerID = HRcolCount;
        else if (aTitle == "jobtitle")
            colHR_Jobtitle = HRcolCount;
        else if (aTitle == "homedepartment")
            colHR_HomeDepartment = HRcolCount;
        else if (aTitle == "city_ad")
            colAD_City = HRcolCount;
        else if (aTitle == "locationaux1")
            colHR_locationAux1 = HRcolCount;
        else if (aTitle == "empclass")
            colHR_FTE = HRcolCount;
        else if (aTitle == "costcenter")
            colCostCenter = HRcolCount;
        else if (aTitle == "needaccount")
            colNoNeedAccount = HRcolCount;
        HRcolCount += 1;
    }
    if (HRcolCount < 1)
        // swLog.WriteLine(aFileName & ": " & " has no column.")
        return false;

    strReportCoupa = strWorkFolder + @"\" + strPrimeSubfolder + @"\" + CSVfilenameCoupa;
    if (File.Exists(strReportCoupa))
    {
        DateTime fileTime = FileSystem.FileDateTime(strReportCoupa);
        string backupFileName = strReportCoupa.Replace(".csv", fileTime.ToString("yyyyMMddHHmm") + ".csv");
        if (!File.Exists(backupFileName))
        {
            try
            {
                File.Copy(strReportCoupa, backupFileName);
            }
            catch (Exception ex)
            {
            }
        }
    }

    StreamWriter swCoupa;
    try
    {
        swCoupa = new StreamWriter(strReportCoupa, false); // always create new file
    }
    catch (Exception ex)
    {
        return false;
    }

    ArrayList uniqueIDs = new ArrayList(), uniqueEmails = new ArrayList();
    string columnNameCoupa = "employeeID,WindowsAccountName,LegalFirstName,PreferredName,LastName,MiddleName,JobTitle,HomeDepartment,DepartmentCode,WorkEmail,HR_Status,ManagerEmail,City,LocationAuxiliary1,LocationCode,DivisionCode,CompanyCode,ManagementLevel";
    swCoupa.WriteLine(columnNameCoupa);

    HRrowCount = 0;
    int failedCount = 0;
    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);

            if (colHR_FTE > 0)
            {
                string FTE = LCase(rows.Item(HRrowCount)(colHR_FTE));
                if (FTE == "independent contractor" || FTE == "regular part_time" || FTE == "temporary" || FTE == "consultant_tpa" || FTE == "non-employee" || FTE == "cobra special dependents" || FTE == "consultant" || FTE == "board_of_directors" || FTE == "intern full_time")
                {
                    HRrowCount += 1;
                    continue;
                }
            }
            if (colNoNeedAccount >= 0)
            {
                if (rows(HRrowCount)(colNoNeedAccount) == noNeedAccountValueHR)
                {
                    HRrowCount += 1;
                    continue;
                }
            }

            if (colHR_eeID >= 0)
            {
                employeeID = rows(HRrowCount)(colHR_eeID);
                if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                {
                    HRrowCount += 1;
                    continue;
                }
            }

            string lineCoupa = "";

            samAccountName = "";
            if (colHR_samAccountName >= 0)
                samAccountName = rows.Item(HRrowCount)(colHR_samAccountName);
            if (samAccountName == "")
            {
                HRrowCount += 1;
                continue;
            }

            if (colHR_eeID < 0)
                employeeID = "";
            else
            {
                // always use WDID now
                employeeID = rows(HRrowCount)(colHR_eeID);
                if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                {
                    // users in ignore update list will use current AD property for manager and city
                    if (colAD_managerID >= 0)
                        colHR_managerID = colAD_managerID;
                }
            }
            lineCoupa += "\"" + employeeID + "\"";
            lineCoupa += ",\"" + samAccountName + "\"";
            if (colHR_firstName < 0)
                FirstName = "";
            else
                FirstName = Trim(rows.Item(HRrowCount)(colHR_firstName));
            if (colHR_nickName < 0)
                NickName = "";
            else
                NickName = Trim(rows.Item(HRrowCount)(colHR_nickName));
            if (colHR_lastName >= 0)
                LastName = rows.Item(HRrowCount)(colHR_lastName);
            if (colHR_middleName >= 0)
                MiddleName = rows.Item(HRrowCount)(colHR_middleName);
            lineCoupa += ",\"" + FirstName + "\"";
            lineCoupa += ",\"" + NickName + "\"";
            lineCoupa += ",\"" + LastName + "\"";
            lineCoupa += ",\"" + MiddleName + "\"";

            string jobTitle = "";
            if (colHR_Jobtitle >= 0)
                jobTitle = LCase(rows.Item(HRrowCount)(colHR_Jobtitle));
            lineCoupa += ",\"" + jobTitle + "\"";

            HomeDepartment = "";
            if (colHR_HomeDepartment >= 0)
                HomeDepartment = LCase(rows.Item(HRrowCount)(colHR_HomeDepartment));
            lineCoupa += ",\"" + HomeDepartment + "\"";

            // DepartmentCode
            CostCenter = "";
            if (colCostCenter >= 0)
            {
                CostCenter = rows.Item(HRrowCount)(colCostCenter);
                int len = CostCenter.Length;
                if (len >= 4)
                    CostCenter = CostCenter.Substring(len - 3, 3);
            }
            lineCoupa += ",\"" + CostCenter + "\"";

            Workemail = "";
            if (colHR_workEmail >= 0)
                Workemail = rows.Item(HRrowCount)(colHR_workEmail);
            lineCoupa += ",\"" + Workemail + "\"";

            Status = "";
            if (colHR_status >= 0)
                Status = rows.Item(HRrowCount)(colHR_status);
            lineCoupa += ",\"" + Status + "\"";

            ManagerWorkEmail = "";
            if (colHR_managerID >= 0)
            {
                // should consider manager account has email, but AD may not set manager?
                ManagerID = rows.Item(HRrowCount)(colHR_managerID);
                while ((ManagerID.Length < 6))
                    ManagerID = "0" + ManagerID;
                int ADindex = eeIDlist.IndexOf(ManagerID);
                if (ADindex >= 0)
                    ManagerWorkEmail = WorkemailList(ADindex);
            }
            lineCoupa += ",\"" + ManagerWorkEmail + "\"";

            cityAD = "";
            if (colAD_City >= 0)
                cityAD = rows.Item(HRrowCount)(colAD_City);
            lineCoupa += ",\"" + cityAD + "\"";

            LocationAux1 = "";
            if (colHR_locationAux1 >= 0)
                LocationAux1 = Trim(rows.Item(HRrowCount)(colHR_locationAux1));
            lineCoupa += ",\"" + LocationAux1 + "\"";

            string LocationCode = "";
            string DivisionCode = "";
            string CompanyCode = "";
            string ManagementLevel = "";
            // locate HR LocationCode, DivisionCode, HR CompanyCode
            if (employeeID != "")
            {
                int HRindex = HR_eeID.IndexOf(employeeID);
                if (HRindex >= 0)
                {
                    LocationCode = HR_LocationCode(HRindex);
                    DivisionCode = HR_DivisionCode(HRindex);
                    CompanyCode = HR_CompanyCode(HRindex);
                    ManagementLevel = HR_ManagementLevel(HRindex);
                }
            }

            lineCoupa += ",\"" + LocationCode + "\"";
            lineCoupa += ",\"" + DivisionCode + "\"";
            lineCoupa += ",\"" + CompanyCode + "\"";
            lineCoupa += ",\"" + ManagementLevel + "\"";

            // swLog.WriteLine(HR_eeID.Item(HRrowCount) & " " & HR_firstName.Item(HRrowCount) & " " & HR_lastName.Item(HRrowCount))
            swCoupa.WriteLine(lineCoupa);
            HRrowCount += 1;
        }
    }
    catch (Exception ex)
    {
        input.Close();
        swLog.WriteLine(ex.ToString());
        return false;
    }

    input.Close();

    swCoupa.Close();

    Console.WriteLine(strExportFileEntire + ": " + HRrowCount + " records convert to " + strReportCoupa + " successfully (validation failed count: " + failedCount + " ).");
    // Console.ReadKey()
    return true;
}

private static bool prepareADHRactive()
{
    if (hasWriteEntireError)
    {
    }

    if (!File.Exists(strExportFileFound))
        return false;
    object input = null;
    int tryTime = 0;
    while (input == null)
    {
        try
        {
            tryTime += 1;
            input = My.Computer.FileSystem.OpenTextFieldParser(strExportFileFound);
        }
        catch (Exception ex)
        {
            System.Threading.Thread.Sleep(3000);
        }
        if (tryTime > 10)
        {
            tmpStr = "AD sync: Open file time out!" + strExportFileFound;
            // swLog.WriteLine(tmpStr)
            // sendAlertEmail("Changtu.Wang@ehealth.com", "", tmpStr, tmpStr, "")
            break;
        }
    }
    if (input == null)
        return false;
    input.setDelimiters(",");
    ArrayList rows = new ArrayList();
    rows.Clear();
    string[] title;
    title = input.readfields();

    colHR_firstName = -1;
    colHR_lastName = -1;
    colHR_workEmail = -1;
    // Dim colHR_WDID As Integer = -1
    colHR_eeID = -1;
    int col_preID = -1;
    colHR_samAccountName = -1;
    colHR_OriginalHireDate = -1;
    colHR_HireDate = -1;
    colHR_statusEffDate = -1;
    colHR_status = -1;
    colHR_managerID = -1;
    int colAD_managerID = -1;
    colHR_Jobtitle = -1;
    int colHR_BusinessUnit = -1;
    int colName_mismatch = -1;
    int colNoNeedAccount = -1;

    colHR_HomeDepartment = -1;

    int colWhenCreated = -1;
    int colWhenChanged = -1;

    colHR_City = -1;
    int colAD_City = -1;
    int colEmployeeType_AD = -1;
    colHR_locationAux1 = -1;

    colHR_FTE = -1;
    colHR_nickName = -1;

    int HRcolCount = 0;
    foreach (string aTitle in title)
    {
        aTitle = Strings.LCase(aTitle.Replace(" ", ""));
        aTitle = aTitle.Replace(".", "");
        // aTitle = aTitle.Replace("#", "")
        if (aTitle == "lastname")
            colHR_lastName = HRcolCount;
        else if (aTitle == "firstname")
            colHR_firstName = HRcolCount;
        else if (aTitle == "nickname")
            colHR_nickName = HRcolCount;
        else if (aTitle == "middlename")
            colHR_middleName = HRcolCount;
        else if (aTitle == "workemail_ad")
            colHR_workEmail = HRcolCount;
        else if (aTitle == "employeeid")
            // always use WDID now
            colHR_eeID = HRcolCount;
        else if (aTitle == "previousid")
            col_preID = HRcolCount;
        else if (aTitle == "samaccountname")
            colHR_samAccountName = HRcolCount;
        else if (aTitle == "status")
            colHR_status = HRcolCount;
        else if (aTitle == "managerid")
            colHR_managerID = HRcolCount;
        else if (aTitle == "managerid_ad")
            colAD_managerID = HRcolCount;
        else if (aTitle == "jobtitle")
            colHR_Jobtitle = HRcolCount;
        else if (aTitle == "businessunit")
            colHR_BusinessUnit = HRcolCount;
        else if (aTitle == "homedepartment")
            colHR_HomeDepartment = HRcolCount;
        else if (aTitle == "whencreated_ad")
            colWhenCreated = HRcolCount;
        else if (aTitle == "whenchanged_ad")
            colWhenChanged = HRcolCount;
        else if (aTitle == "city")
            colHR_City = HRcolCount;
        else if (aTitle == "city_ad")
            colAD_City = HRcolCount;
        else if (aTitle == "employeetype_ad")
            colEmployeeType_AD = HRcolCount;
        else if (aTitle == "locationaux1")
            colHR_locationAux1 = HRcolCount;
        else if (aTitle == "empclass")
            colHR_FTE = HRcolCount;
        else if (aTitle == "name_mismatch")
            colName_mismatch = HRcolCount;
        else if (aTitle == "needaccount")
            colNoNeedAccount = HRcolCount;
        HRcolCount += 1;
    }
    if (HRcolCount < 1)
        // swLog.WriteLine(aFileName & ": " & " has no column.")
        return false;

    ADusersHRactiveCSV = strWorkFolder + @"\" + strReportSubFolder + @"\" + "ADusers_" + getYYYYMMDD_hhtt() + ".csv";
    StreamWriter swADHR;
    try
    {
        swADHR = new StreamWriter(ADusersHRactiveCSV, false); // always create new file
    }
    catch (Exception ex)
    {
        return false;
    }

    swADHR.WriteLine("employeeID,reqID,WindowsAccountName,LegalFirstName,FirstName,LastName,MiddleName,JobTitle,HomeDepartment,WorkEmail,WhenCreated,WhenChanged,employeeType,City,Status,daysCreated");

    DateTime createTime, runTime;
    runTime = DateTime.Now;

    HRrowCount = 0;
    try
    {
        while ((!input.endofdata))
        {
            rows.Add(input.readfields);
            if (colHR_status < 0)
                Status = "";
            else
                Status = rows.Item(HRrowCount)(colHR_status);

            if (Status != "Active" && Status != "On Leave")
            {
                HRrowCount += 1;
                continue;
            }

            if (colHR_eeID < 0)
                employeeID = "";
            else
                employeeID = rows.Item(HRrowCount)(colHR_eeID);
            if (col_preID < 0)
                preID = "";
            else
                preID = rows.Item(HRrowCount)(col_preID);

            if (colHR_samAccountName < 0)
                samAccountName = "";
            else
                samAccountName = rows.Item(HRrowCount)(colHR_samAccountName);
            if (samAccountName == "")
            {
                HRrowCount += 1;
                continue;
            }

            string lineHRactive = "\"" + employeeID + "\",\"" + preID + "\",\"" + samAccountName + "\"";

            FirstName = "";
            if (colHR_firstName >= 0)
                FirstName = rows.Item(HRrowCount)(colHR_firstName);
            NickName = "";
            if (colHR_nickName >= 0)
                NickName = rows.Item(HRrowCount)(colHR_nickName);
            LastName = "";
            if (colHR_lastName >= 0)
                LastName = rows.Item(HRrowCount)(colHR_lastName);
            MiddleName = "";
            if (colHR_middleName >= 0)
                MiddleName = rows.Item(HRrowCount)(colHR_middleName);

            lineHRactive += ",\"" + FirstName + "\",\"" + NickName + "\",\"" + LastName + "\",\"" + MiddleName + "\"";

            Jobtitle = "";
            if (colHR_Jobtitle >= 0)
                Jobtitle = rows.Item(HRrowCount)(colHR_Jobtitle);
            lineHRactive += ",\"" + Jobtitle + "\"";
            HomeDepartment = "";
            if (colHR_HomeDepartment >= 0)
                HomeDepartment = rows.Item(HRrowCount)(colHR_HomeDepartment);
            lineHRactive += ",\"" + HomeDepartment + "\"";

            if (colHR_workEmail < 0)
                Workemail = "";
            else
                Workemail = rows.Item(HRrowCount)(colHR_workEmail);
            lineHRactive += ",\"" + Workemail + "\"";

            if (colWhenCreated < 0)
                whenCreated = "";
            else
            {
                whenCreated = rows.Item(HRrowCount)(colWhenCreated);
                createTime = whenCreated;
            }
            if (colWhenChanged < 0)
                whenChanged = "";
            else
                whenChanged = rows.Item(HRrowCount)(colWhenChanged);
            lineHRactive += ",\"" + whenCreated + "\"";
            lineHRactive += ",\"" + whenChanged + "\"";

            string employeeType = "";
            if (colEmployeeType_AD >= 0)
                employeeType = rows.Item(HRrowCount)(colEmployeeType_AD);
            lineHRactive += ",\"" + employeeType + "\"";

            if (colHR_City < 0)
                lineHRactive += ",";
            else
                lineHRactive += ",\"" + rows.Item(HRrowCount)(colHR_City) + "\"";

            lineHRactive += ",\"" + Status + "\"";

            string in90days = "";
            float hours;
            hours = DateTime.DateDiff("h", createTime, runTime);
            if (hours < 0)
                hours = 0;
            if (Math.Round(hours / (double)24, 1) < 91)
                in90days = Math.Round(hours / (double)24, 1);


            if (colHR_eeID < 0)
                employeeID = "";
            else
            {
                // always use WDID now
                employeeID = rows(HRrowCount)(colHR_eeID);
                if (ignoreUpdateEEIDs.Contains("." + employeeID + "."))
                {
                    // users in ignore update list will use current AD property for manager and city
                    if (colAD_managerID >= 0)
                        colHR_managerID = colAD_managerID;
                    if (colAD_City >= 0)
                        colHR_City = colAD_City;
                }
            }

            lineHRactive += "," + in90days + ",";

            swADHR.WriteLine(lineHRactive);

            HRrowCount += 1;
            continue;
        }
    }
    catch (Exception ex)
    {
        input.Close();
        if (swLog != null)
            swLog.WriteLine(ex.ToString());
        return false;
    }


    input.Close();

    swADHR.Close();

    Console.WriteLine(strExportFileFound + ": " + HRrowCount + " records convert to " + ADusersHRactiveCSV + " successfully.");
    // Console.ReadKey()
    return true;
}

private static void testS3()
{
    if (!needUploadS3)
        return;
    S3base s3 = new S3base();
    // s3.S3key = "ADusers_20200701_1515.csv"

    ADusersHRactiveCSV = @"C:\ehiAD\report\ADusers_20200701_0951.csv";
    tmpStr = strWorkFolder + @"\" + strReportSubFolder + @"\";
    s3.S3key = ADusersHRactiveCSV.Replace(tmpStr, ""); // "ADusers_20200701_1515.csv"

    if (s3.ListObject)
    {
        string resultStr = s3.resultStr;
        Console.WriteLine(resultStr);
    }
    else
        Console.WriteLine(s3.errStr);
    Console.ReadKey();

    if (s3.upload)
    {
        string resultStr = s3.resultStr;
        Console.WriteLine(resultStr);
    }
    else
        Console.WriteLine(s3.errStr);
    Console.ReadKey();

    if (s3.ListObject)
    {
        string resultStr = s3.resultStr;
        Console.WriteLine(resultStr);
    }
    else
        Console.WriteLine(s3.errStr);
    Console.ReadKey();

    if (s3.GetObject)
    {
        string resultStr = s3.resultStr;
        Console.WriteLine(resultStr);
    }
    else
        Console.WriteLine(s3.errStr);
    Console.ReadKey();

    if (s3.DeleteObject)
    {
        string resultStr = s3.resultStr;
        Console.WriteLine(resultStr);
    }
    else
        Console.WriteLine(s3.errStr);
    Console.ReadKey();

    if (s3.ListObject)
    {
        string resultStr = s3.resultStr;
        Console.WriteLine(resultStr);
    }
    else
        Console.WriteLine(s3.errStr);
    Console.ReadKey();
    return;
}
}
