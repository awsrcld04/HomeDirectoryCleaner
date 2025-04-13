﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using Microsoft.Win32;
using System.Management;
using System.Reflection;
using System.Diagnostics;

namespace HomeDirectoryCleaner
{
    class HDCMain
    {
        struct HomeDirectoryParams
        {
            public string strHomeDirectoryLocation;
            public string strSendActionList;
            public string strActionListSender;
            public string strActionListRecipient;
            public List<string> lstExcludePrefix;
            public List<string> lstExclude;
        }

        struct CMDArguments
        {
            public bool bParseCmdArguments;
        }

        static bool funcLicenseCheck()
        {
            try
            {
                string strLicenseString = "";
                bool bValidLicense = false;

                TextReader tr = new StreamReader("sotfwlic.dat");

                try
                {
                    strLicenseString = tr.ReadLine();

                    if (strLicenseString.Length > 0 & strLicenseString.Length < 29)
                    {
                        // [DebugLine] Console.WriteLine("if: " + strLicenseString);
                        Console.WriteLine("Invalid license");

                        tr.Close(); // close license file

                        return bValidLicense;
                    }
                    else
                    {
                        tr.Close(); // close license file
                        // [DebugLine] Console.WriteLine("else: " + strLicenseString);

                        string strMonthTemp = ""; // to convert the month into the proper number
                        string strDate;

                        //Month
                        strMonthTemp = strLicenseString.Substring(7, 1);
                        if (strMonthTemp == "A")
                        {
                            strMonthTemp = "10";
                        }
                        if (strMonthTemp == "B")
                        {
                            strMonthTemp = "11";
                        }
                        if (strMonthTemp == "C")
                        {
                            strMonthTemp = "12";
                        }
                        strDate = strMonthTemp;

                        //Day
                        strDate = strDate + "/" + strLicenseString.Substring(16, 1);
                        strDate = strDate + strLicenseString.Substring(6, 1);

                        // Year
                        strDate = strDate + "/" + strLicenseString.Substring(24, 1);
                        strDate = strDate + strLicenseString.Substring(4, 1);
                        strDate = strDate + strLicenseString.Substring(1, 2);

                        // [DebugLine] Console.WriteLine(strDate);
                        // [DebugLine] Console.WriteLine(DateTime.Today.ToString());
                        DateTime dtLicenseDate = DateTime.Parse(strDate);
                        // [DebugLine]Console.WriteLine(dtLicenseDate.ToString());

                        if (dtLicenseDate >= DateTime.Today)
                        {
                            bValidLicense = true;
                        }
                        else
                        {
                            Console.WriteLine("License expired.");
                        }

                        return bValidLicense;
                    }

                } //end of try block on tr.ReadLine

                catch
                {
                    // [DebugLine] Console.WriteLine("catch on tr.Readline");
                    Console.WriteLine("Invalid license");
                    tr.Close();
                    return bValidLicense;

                } //end of catch block on tr.ReadLine

            } // end of try block on new StreamReader("sotfwlic.dat")

            catch (System.Exception ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());

                // [DebugLine] System.Console.WriteLine(ex.Message);

                if (ex.Message.StartsWith("Could not find file"))
                {
                    Console.WriteLine("License file not found.");
                }
                else
                {
                    MethodBase mb1 = MethodBase.GetCurrentMethod();
                    funcGetFuncCatchCode(mb1.Name, ex);
                }

                return false;

            } // end of catch block on new StreamReader("sotfwlic.dat")
        }

        static bool funcLicenseActivation()
        {
            try
            {
                if (funcCheckForFile("TurboActivate.dll"))
                {
                    if (funcCheckForFile("TurboActivate.dat"))
                    {
                        TurboActivate.VersionGUID = "4935355894e0da3d4465e86.37472852";

                        if (TurboActivate.IsActivated())
                        {
                            return true;
                        }
                        else
                        {
                            Console.WriteLine("A license for this product has not been activated.");
                            return false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("TurboActivate.dat is required and could not be found.");
                        return false;
                    }
                }
                else
                {
                    Console.WriteLine("TurboActivate.dll is required and could not be found.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("Parameters must be specified properly to run HomeDirectoryCleaner.");
            Console.WriteLine("Run HomeDirectoryCleaner -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("HomeDirectoryCleaner v1.0 (c) 2011 SystemsAdminPro.com");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following required parameters in the following order:");
            Console.WriteLine("-run                     required parameter");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("HomeDirectoryCleaner -run");
        }

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                if (cmdargs[0] == "-run" & cmdargs.Length == 1)
                {
                    objCMDArguments.bParseCmdArguments = true;
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static void funcProgramRegistryTag(string strProgramName)
        {
            try
            {
                string strRegistryProfilesPath = "SOFTWARE";
                RegistryKey objRootKey = Microsoft.Win32.Registry.LocalMachine;
                RegistryKey objSoftwareKey = objRootKey.OpenSubKey(strRegistryProfilesPath, true);
                RegistryKey objSystemsAdminProKey = objSoftwareKey.OpenSubKey("SystemsAdminPro", true);
                if (objSystemsAdminProKey == null)
                {
                    objSystemsAdminProKey = objSoftwareKey.CreateSubKey("SystemsAdminPro");
                }
                if (objSystemsAdminProKey != null)
                {
                    if (objSystemsAdminProKey.GetValue(strProgramName) == null)
                        objSystemsAdminProKey.SetValue(strProgramName, "1", RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static DirectorySearcher funcCreateDSSearcher()
        {
            try
            {
                System.DirectoryServices.DirectorySearcher objDSSearcher = new DirectorySearcher();
                // [Comment] Get local domain context

                string rootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                rootDSE = objrootDSESearcher.SearchRoot.Path;
                //Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
                objDSSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                //Console.WriteLine(objDSSearcher.SearchRoot.Path);

                return objDSSearcher;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }
        }

        static PrincipalContext funcCreatePrincipalContext()
        {
            PrincipalContext newctx = new PrincipalContext(ContextType.Machine);

            try
            {
                //Console.WriteLine("Entering funcCreatePrincipalContext");
                Domain objDomain = Domain.GetComputerDomain();
                string strDomain = objDomain.Name;
                DirectorySearcher tempDS = funcCreateDSSearcher();
                string strDomainRoot = tempDS.SearchRoot.Path.Substring(7);
                // [DebugLine] Console.WriteLine(strDomainRoot);
                // [DebugLine] Console.WriteLine(strDomainRoot);

                newctx = new PrincipalContext(ContextType.Domain,
                                    strDomain,
                                    strDomainRoot);

                // [DebugLine] Console.WriteLine(newctx.ConnectedServer);
                // [DebugLine] Console.WriteLine(newctx.Container);



                //if (strContextType == "Domain")
                //{

                //    PrincipalContext newctx = new PrincipalContext(ContextType.Domain,
                //                                    strDomain,
                //                                    strDomainRoot);
                //    return newctx;
                //}
                //else
                //{
                //    PrincipalContext newctx = new PrincipalContext(ContextType.Machine);
                //    return newctx;
                //}
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            if (newctx.ContextType == ContextType.Machine)
            {
                Exception newex = new Exception("The Active Directory context did not initialize properly.");
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, newex);
            }

            return newctx;
        }

        static void funcToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            try
            {
                string strLogName;

                strLogName = "Application";

                if (!EventLog.SourceExists(strAppName))
                    EventLog.CreateEventSource(strAppName, strLogName);

                //EventLog.WriteEntry(strAppName, strEventMsg);
                EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcGetFuncCatchCode", "f0");
            dCatchTable.Add("funcLicenseCheck", "f1");
            dCatchTable.Add("funcPrintParameterWarning", "f2");
            dCatchTable.Add("funcPrintParameterSyntax", "f3");
            dCatchTable.Add("funcParseCmdArguments", "f4");
            dCatchTable.Add("funcProgramExecution", "f5");
            dCatchTable.Add("funcProgramRegistryTag", "f6");
            dCatchTable.Add("funcCreateDSSearcher", "f7");
            dCatchTable.Add("funcCreatePrincipalContext", "f8");
            dCatchTable.Add("funcCheckNameExclusion", "f9");
            dCatchTable.Add("funcMoveDisabledAccounts", "f10");
            dCatchTable.Add("funcFindAccountsToDisable", "f11");
            dCatchTable.Add("funcCheckLastLogin", "f12");
            dCatchTable.Add("funcRemoveUserFromGroup", "f13");
            dCatchTable.Add("funcToEventLog", "f14");
            dCatchTable.Add("funcCheckForFile", "f15");
            dCatchTable.Add("funcCheckForOU", "f16");
            dCatchTable.Add("funcWriteToErrorLog", "f17");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);

        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                FileStream newFileStream = new FileStream("Err-HomeDirectoryCleaner.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static TextWriter funcOpenOutputLog()
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat2 = "MMddyyyy"; // for log file directory creation

                string strPath = Directory.GetCurrentDirectory();

                string strLogFileName = strPath + "\\HomeDirectoryCleaner" + dtNow.ToString(dtFormat2) + ".log";

                FileStream newFileStream = new FileStream(strLogFileName, FileMode.Append, FileAccess.Write);
                TextWriter twOuputLog = new StreamWriter(newFileStream);

                return twOuputLog;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }

        }

        static void funcWriteToOutputLog(TextWriter twCurrent, string strOutputMessage)
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twCurrent.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strOutputMessage);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCloseOutputLog(TextWriter twCurrent)
        {
            try
            {
                twCurrent.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcRecurse(DirectoryInfo directory)
        {
            foreach (FileInfo fi in directory.GetFiles())
            {
                fi.Attributes = FileAttributes.Normal;
            }

            foreach (DirectoryInfo di in directory.GetDirectories())
            {
                di.Attributes = FileAttributes.Normal;
            }

            foreach (DirectoryInfo subdir2 in directory.GetDirectories())
            {
                funcRecurse(subdir2);
            }

        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                if (funcCheckForFile("configHomeDirectoryCleaner.txt"))
                {
                    funcToEventLog("HomeDirectoryCleaner", "HomeDirectoryCleaner started.", 1001);

                    funcProgramRegistryTag("HomeDirectoryCleaner");

                    HomeDirectoryParams newParams = new HomeDirectoryParams();
                    newParams.lstExclude = new List<string>();
                    newParams.lstExcludePrefix = new List<string>();

                    TextReader trConfigFile = new StreamReader("configHomeDirectoryCleaner.txt");

                    using (trConfigFile)
                    {
                        string strNewLine = "";

                        while ((strNewLine = trConfigFile.ReadLine()) != null)
                        {

                            if (strNewLine.StartsWith("HomeDirectoryLocation="))
                            {
                                newParams.strHomeDirectoryLocation = strNewLine.Substring(22);
                                //[DebugLine] Console.WriteLine(newParams.strHomeDirectoryLocation);
                            }
                            if (strNewLine.StartsWith("SendActionList="))
                            {
                                newParams.strSendActionList = strNewLine.Substring(15);
                                //[DebugLine] Console.WriteLine(newParams.strSendActionList);
                            }
                            if (strNewLine.StartsWith("ActionListSender="))
                            {
                                newParams.strActionListSender = strNewLine.Substring(17);
                                //[DebugLine] Console.WriteLine(newParams.strActionListSender);
                            }
                            if (strNewLine.StartsWith("ActionListRecipient="))
                            {
                                newParams.strActionListRecipient = strNewLine.Substring(20);
                                //[DebugLine] Console.WriteLine(newParams.strActionListRecipient);
                            }
                            if (strNewLine.StartsWith("ExcludePrefix="))
                            {
                                newParams.lstExcludePrefix.Add(strNewLine.Substring(14));
                                //[DebugLine] Console.WriteLine(strNewLine.Substring(14));
                            }
                            if (strNewLine.StartsWith("Exclude="))
                            {
                                newParams.lstExclude.Add(strNewLine.Substring(8));
                                //[DebugLine] Console.WriteLine(strNewLine.Substring(8));
                            }
                        }
                    }

                    //[DebugLine] Console.WriteLine("# of Exclude= : {0}", newParams.lstExclude.Count.ToString());
                    //[DebugLine] Console.WriteLine("# of ExcludePrefix= : {0}", newParams.lstExcludePrefix.Count.ToString());

                    trConfigFile.Close();

                    TextWriter twCurrent = funcOpenOutputLog();
                    string strOutputMsg = "";

                    string[] strDirectories = Directory.GetDirectories(@newParams.strHomeDirectoryLocation);

                    //[DebugLine] Console.WriteLine(strDirectories.Count<string>().ToString());

                    foreach (string strDirectoryName in strDirectories)
                    {
                        //[DebugLine] Console.WriteLine(strDirectoryName);
                        //[DebugLine] Console.WriteLine(strDirectoryName.Substring(9));

                        // Create the principal context for the usr object.
                        PrincipalContext ctx = funcCreatePrincipalContext();

                        // Create the principal user object from the context
                        UserPrincipal usr = new UserPrincipal(ctx);

                        string strUserName = strDirectoryName.Substring(9);

                        usr.SamAccountName = strUserName;

                        // Create a PrincipalSearcher object.
                        PrincipalSearcher ps = new PrincipalSearcher(usr);
                        Principal pr = ps.FindOne();

                        if (pr != null)
                        {
                            UserPrincipal u = (UserPrincipal)pr;
                            //[DebugLine] Console.WriteLine(u.Name);
                            //[DebugLine] if (u.SamAccountName != u.Name)
                            //[DebugLine]     Console.WriteLine(u.SamAccountName);
                            //[DebugLine] Console.WriteLine(u.DistinguishedName);

                            strOutputMsg = "Matching account found for directory " + strUserName +
                                           " (" + u.DistinguishedName + ")";                           
                            funcWriteToOutputLog(twCurrent, strOutputMsg);
                        }
                        else
                        {
                            //[DebugLine] Console.WriteLine("User " + strUserName + " not found");
                            strOutputMsg = "User " + strUserName + " not found";
                            funcWriteToOutputLog(twCurrent, strOutputMsg);
                            strOutputMsg = "Deleting home directory: " + strUserName;
                            funcWriteToOutputLog(twCurrent, strOutputMsg);
                            DirectoryInfo tmpDirectoryInfo = new DirectoryInfo(strDirectoryName);
                            funcRecurse(tmpDirectoryInfo);
                            System.IO.Directory.Delete(strDirectoryName, true);
                            if (!Directory.Exists(strDirectoryName))
                            {
                                strOutputMsg = "Successfully deleted home directory " + strUserName;
                                funcWriteToOutputLog(twCurrent, strOutputMsg);
                            }
                            else
                            {
                                strOutputMsg = "Home directory for " + strUserName + " was not successfully deleted";
                                funcWriteToOutputLog(twCurrent, strOutputMsg);
                            }
                        }

                    }

                    funcCloseOutputLog(twCurrent);

                    funcToEventLog("HomeDirectoryCleaner", "HomeDirectoryCleaner stopped.", 1002);
                }
                else
                {
                    Console.WriteLine("configHomeDirectoryCleaner.txt is required and could not be found.");
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    funcPrintParameterWarning();
                }
                else
                {
                    if (args[0] == "-?")
                    {
                        funcPrintParameterSyntax();
                    }
                    else
                    {
                        string[] arrArgs = args;
                        CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                        if (objArgumentsProcessed.bParseCmdArguments)
                        {
                            funcProgramExecution(objArgumentsProcessed);
                        }
                        else
                        {
                            funcPrintParameterWarning();
                        } // check objArgumentsProcessed.bParseCmdArguments
                    } // check args[0] = "-?"
                } // check args.Length == 0
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }
    }
}
