//
// Licensed Materials - Property of IBM
//
// 5725-C15
// © Copyright IBM Corp. 1994, 2019 All Rights Reserved
// US Government Users Restricted Rights - Use, duplication or
// disclosure restricted by GSA ADP Schedule Contract with IBM Corp.
//
// This is an example of a .NET action for IBM Datacap using .NET 4.0.
// The compliled DLL needs to be placed into the RRS directory.
// The DLL does not need to be registered.  
// Datacap studio will find the RRX file that is embedded in the DLL, you do not need to place the RRX in the RRS directory.
// If you add references to other DLLs, such as 3rd party, you may need to place those DLLs in C:\RRS so they are found at runtime.
// If Datacap references are not found at compile time, add a reference path of C:\Datacap\DCShared\NET to the project to locate the DLLs while building.
// This template has been tested with IBM Datacap 9.0.  
//Initial version is created by Muhammet Ali Hancer.
//Modified date: 12/03/2023

using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;

namespace MRZLibrary
{
    public class MRZParserActions // This class must be a base class for .NET 4.0 Actions.
    {
        #region ExpectedByRRS
        /// <summary/>
        ~MRZParserActions()
        {
            DatacapRRCleanupTime = true;
        }

        /// <summary>
        /// Cleanup: This property is set right before the object is released by RRS
        /// The Dispose method is not called by RRS.
        /// </summary>
        public bool DatacapRRCleanupTime
        {
            set
            {
                if (value)
                {
                    CleanUp();
                    CurrentDCO = null;
                    DCO = null;
                    RRLog = null;
                    RRState = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        protected PILOTCTRLLib.IBPilotCtrl BatchPilot = null;
        public PILOTCTRLLib.IBPilotCtrl DatacapRRBatchPilot { set { this.BatchPilot = value; GC.Collect(); GC.WaitForPendingFinalizers(); } get { return this.BatchPilot; } }

        protected TDCOLib.IDCO DCO = null;
        /// <summary/>
        public TDCOLib.IDCO DatacapRRDCO
        {
            get { return this.DCO; }
            set
            {
                DCO = value;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected dcrroLib.IRRState RRState = null;
        /// <summary/>
        public dcrroLib.IRRState DatacapRRState
        {
            get { return this.RRState; }
            set
            {
                RRState = value;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public TDCOLib.IDCO CurrentDCO = null;
        /// <summary/>
        public TDCOLib.IDCO DatacapRRCurrentDCO
        {
            get { return this.CurrentDCO; }
            set
            {
                CurrentDCO = value;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public dclogXLib.IDCLog RRLog = null;
        /// <summary/>
        public dclogXLib.IDCLog DatacapRRLog
        {
            get { return this.RRLog; }
            set
            {
                RRLog = value;
                LogAssemblyVersion();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region CommonActions

        void OutputToLog(int nLevel, string strMessage)
        {
            if (null == RRLog)
                return;
            RRLog.WriteEx(nLevel, strMessage);
        }

        public void WriteLog(string sMessage)
        {
            OutputToLog(5, sMessage);
        }

        private bool versionWasLogged = false;

        // Log the version of the library that was running to help with diagnosis.
        // Hooked this method to be called after the log object is assigned.  Also put in
        // a check that this action runs only once, just in case it gets called multiple times.
        protected bool LogAssemblyVersion()
        {
            try
            {
                if (versionWasLogged == false)
                {
                    FileVersionInfo fv = System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                    WriteLog(Assembly.GetExecutingAssembly().Location +
                             ". AssemblyVersion: " + Assembly.GetExecutingAssembly().GetName().Version.ToString() +
                             ". AssemblyFileVersion: " + fv.FileVersion.ToString() + ".");
                    versionWasLogged = true;
                }
            }
            catch (Exception ex)
            {
                WriteLog("Version logging exception: " + ex.Message);
            }

            // We can always return true.  If getting the version fails, we can try to continue anyway.
            return true;
        }

        #endregion


        // implementation of the Dispose method to release managed resources
        // There is no guarentee that dispose will be called.  Also note, class distructors are also not called.  CleanupTime is called by RRS.        
        public void Dispose()
        {
            CleanUp();
        }

        /// <summary>
        /// Everthing that should be cleaned up on exit
        /// It is recommended to avoid logging during cleanup.
        /// </summary>
        protected void CleanUp()
        {
            try
            {
                // Cleanup and release any allocated objects here. This will be called before the DLL is released.
            }
            catch { } // Ignore any errors.
        }

        private struct Level
        {
            internal const int Batch = 0;
            internal const int Document = 1;
            internal const int Page = 2;
            internal const int Field = 3;
        }

        private struct Status
        {
            internal const int Hidden = -1;
            internal const int OK = 0;
            internal const int Fail = 1;
            internal const int Over = 3;
            internal const int RescanPage = 70;
            internal const int VerificationFailed = 71;
            internal const int PageOnHold = 72;
            internal const int PageOverridden = 73;
            internal const int NoData = 74;
            internal const int DeletedPage = 75;
            internal const int ExportComplete = 76;
            internal const int DeleteApproved = 77;
            internal const int ReviewPage = 79;
            internal const int DeletedDoc = 128;
        }

        /// <summary/>
        /// This is an example custom .NET action that takes multiple parameters with multiple types.
        /// The parameter order and types must match the definition in the RRX file.

        public bool CreatePassportVariables(string MRZFirstLine, string MRZSecondLine)
        {
            bool passportVariables = true;
            dcSmart.SmartNav localSmartObj = null;
            try
            {
                localSmartObj = new dcSmart.SmartNav(this);
                string str1 = localSmartObj.MetaWord(MRZFirstLine).Replace(" ", String.Empty);
                string str2 = localSmartObj.MetaWord(MRZSecondLine).Replace(" ", String.Empty);

                if (this.CurrentDCO.ObjectType() == 0)
                {
                    this.WriteLog("Batch Level");
                }
                else
                {

                    if (this.CurrentDCO.ObjectType() == 1)
                    {
                        this.WriteLog("Document Level");
                    }
                    else
                    {

                        if (this.CurrentDCO.ObjectType() == 2)
                        {
                            this.WriteLog("Page Level");
                            this.WriteLog("Image name: " + this.CurrentDCO.ImageName);
                            char[] separator = new char[1] { '<' };
                            string pVal1 = str1.Substring(0, 1);
                            string pVal2 = str1.Substring(2, 3);
                            string[] strArray = str1.Substring(5, str1.Length - 5).Split(separator, 2, StringSplitOptions.RemoveEmptyEntries);
                            string pVal3 = strArray[0];
                            string pVal4 = strArray[1].Replace("<", " ").Trim();
                            string pVal5 = str2.Substring(0, 9);
                            string pVal6 = str2.Substring(10, 3);
                            string str3 = str2.Substring(13, 6);
                            string pVal7 = str3.Substring(4, 2) + "/" + str3.Substring(2, 2) + "/" + str3.Substring(0, 2);
                            string pVal8 = str2.Substring(20, 1);
                            string str4 = str2.Substring(21, 6);
                            string pVal9 = str4.Substring(4, 2) + "/" + str4.Substring(2, 2) + "/" + str4.Substring(0, 2);
                            string pVal10 = str2.Substring(28, 11);

                            this.CurrentDCO.set_Variable("docType", pVal1);

                            this.CurrentDCO.set_Variable("countryCode", pVal2);

                            this.CurrentDCO.set_Variable("personSurname", pVal3);

                            this.CurrentDCO.set_Variable("personName", pVal4);

                            this.CurrentDCO.set_Variable("passportNo", pVal5);

                            this.CurrentDCO.set_Variable("nationality", pVal6);

                            this.CurrentDCO.set_Variable("birthDate", pVal7);

                            this.CurrentDCO.set_Variable("sex", pVal8);

                            this.CurrentDCO.set_Variable("expiryDate", pVal9);

                            this.CurrentDCO.set_Variable("personalNumber", pVal10);
                        }
                        else
                        {

                            if (this.CurrentDCO.ObjectType() == 3)
                                this.WriteLog("Field Level");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.WriteLog("There was an exception: " + ex.Message);
            }
            localSmartObj = null;
            return passportVariables;
        }

        public bool CreateIDDocVariables(string MRZFirstLine, string MRZSecondLine, string MRZThirdLine)
        {
            bool idDocVariables = true;
            dcSmart.SmartNav localSmartObj = null;
            try
            {
                localSmartObj = new dcSmart.SmartNav(this);
                string str1 = localSmartObj.MetaWord(MRZFirstLine).Replace(" ", String.Empty);
                string str2 = localSmartObj.MetaWord(MRZSecondLine).Replace(" ", String.Empty);
                string str3 = localSmartObj.MetaWord(MRZThirdLine).Replace(" ", String.Empty);

                if (this.CurrentDCO.ObjectType() == 0)
                {
                    this.WriteLog("Batch Level");
                }
                else
                {

                    if (this.CurrentDCO.ObjectType() == 1)
                    {
                        this.WriteLog("Document Level");
                    }
                    else
                    {

                        if (this.CurrentDCO.ObjectType() == 2)
                        {
                            this.WriteLog("Page Level");
                            string pVal1 = str1.Substring(0, 2);
                            string pVal2 = str1.Substring(2, 3);
                            string pVal3 = str1.Substring(5, 9);
                            str1.Substring(14, 1);
                            string pVal4 = str1.Substring(15, 15);
                            string pVal5 = DateTime.ParseExact(str2.Substring(0, 6), "yyMMdd", (IFormatProvider)CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                            str2.Substring(6, 1);
                            string pVal6 = str2.Substring(7, 1);
                            string pVal7 = DateTime.ParseExact(str2.Substring(8, 6), "yyMMdd", (IFormatProvider)CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                            str2.Substring(14, 1);
                            string pVal8 = str2.Substring(15, 3);
                            str2.Substring(18, 11);
                            str2.Substring(29, 1);
                            string[] strArray = str3.Split(new string[1] { "<<" }, 2, StringSplitOptions.None);
                            string pVal9 = strArray[0];
                            string pVal10 = strArray[1].Replace('<', ' ').Trim();

                            this.CurrentDCO.set_Variable("docType", pVal1);
                            this.CurrentDCO.set_Variable("countryCode", pVal2);
                            this.CurrentDCO.set_Variable("docNumber", pVal3);
                            this.CurrentDCO.set_Variable("optionalData1", pVal4);
                            this.CurrentDCO.set_Variable("dateOfBirth", pVal5);
                            this.CurrentDCO.set_Variable("gender", pVal6);
                            this.CurrentDCO.set_Variable("expiryDate", pVal7);
                            this.CurrentDCO.set_Variable("nationality", pVal8);
                            this.CurrentDCO.set_Variable("optionalData2", pVal7);
                            this.CurrentDCO.set_Variable("lastName", pVal9);
                            this.CurrentDCO.set_Variable("firstName", pVal10);
                        }
                        else
                        {

                            if (this.CurrentDCO.ObjectType() == 3)
                                this.WriteLog("Field Level");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.WriteLog("There was an exception: " + ex.Message);
            }
            localSmartObj = null;
            return idDocVariables;
        }



    }
}
