using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Connectivity.WebServices;
using Autodesk.Connectivity.WebServicesTools;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties;
using Inventor;
using iLogicAutomation = Autodesk.iLogic.Automation;
using iLogicInterface = Autodesk.iLogic.Interfaces;
using AcInterop = Autodesk.AutoCAD.Interop;
using AcInteropCom = Autodesk.AutoCAD.Interop.Common;

namespace VDSUtils
{
    public class VltHelpers

    {
        //UserCredentials1 and UserCredentials2 differentiate overloads as powershell can't handle 
        //UserCredentials1 returns read-write loginuser object
        public Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials UserCredentials1(string server, string vault, string user, string pw)
        {
            ServerIdentities mServer = new ServerIdentities();
            mServer.DataServer = server;
            mServer.FileServer = server;
             Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials mCred = new Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials(mServer, vault, user, pw);
            return mCred;
        }

        //UserCredentials2 returns readonly loginuser object
        public Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials UserCredentials2(string server, string vault, string user, string pw, bool rw = true)
        {
            ServerIdentities mServer = new ServerIdentities();
            mServer.DataServer = server;
            mServer.FileServer = server;
            Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials mCred = new Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials(mServer, vault, user, pw, rw);
            return mCred;
        }

        // - change in 2016 - the overload is removed (only the update routine introduced in 2015 R2 remains.
        //public Boolean UpdateFolderProp1(WebServiceManager svc, long[] FldIds, long[] PropDefIds, System.Object[] Values)
        //{
        //    try
        //    {
        //        svc.DocumentServiceExtensions.UpdateFolderProperties(FldIds, PropDefIds, Values);
        //        return true;
        //    }
        //    catch 
        //    {
        //        return false;
        //    }

        //}

        public Boolean UpdateFolderProp2(WebServiceManager svc, long[] FldIds, PropInstParamArray[] m_PropArray)
        {
            try
            {
                svc.DocumentServiceExtensions.UpdateFolderProperties(FldIds, m_PropArray);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public class InvHelpers
    {
        Inventor.Application m_Inv = null;
        Inventor.Document m_Doc = null;
        Inventor.DrawingDocument m_DrawDoc = null;
        Inventor.PresentationDocument m_IpnDoc = null;
        String m_ModelPath = null;
        Inventor.CommandManager m_InvCmdMgr = null;

        [System.Runtime.InteropServices.DllImport("User32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
                    
        public object m_GetMainViewModelPropValue(String m_ViewModelFullName, String m_PropName)
        {
            try
            {
                m_Doc = m_Inv.Documents.Open(m_ViewModelFullName,false);
                        foreach (PropertySet m_PropSet in m_Doc.PropertySets)
                        {
                            foreach (Property m_Prop in m_PropSet)
                            {
                                if (m_Prop.Name == m_PropName)
                                {
                                    return m_Prop.Value;
                                }
                            }
                        }
            }
            catch
            {
                return "Error retrieving iProperty";
            }
            return "iProperty not found";
        }

        public String m_GetMainViewModelPath()
        {
            try
            {
                if (m_ConnectInv()==true)
                {
                    if (m_Inv.ActiveDocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    {
                        m_DrawDoc = (DrawingDocument)m_Inv.ActiveDocument;
                        Sheet m_Sheet = m_DrawDoc.Sheets[1];
                        DrawingView m_DrwView = m_Sheet.DrawingViews[1];
                        m_ModelPath = m_DrwView.ReferencedDocumentDescriptor.FullDocumentName;
                        return m_ModelPath;
                    }
                    if (m_Inv.ActiveDocumentType == DocumentTypeEnum.kPresentationDocumentObject)
                    {
                        m_IpnDoc = (PresentationDocument)m_Inv.ActiveDocument;
                        m_ModelPath = m_IpnDoc.ReferencedDocuments[1].FullDocumentName;
                        return m_ModelPath;
                    }
                }
            }
            catch
                {
                    return m_ModelPath="";
                }
            return m_ModelPath = "";
        }

        // insert provided component (ipt/iam) into current assembly - waiting for user interaction to finalize
        public void m_PlaceComponent(String m_CompFullFileName)
        {
            if (m_ConnectInv() == true)
            {
                if (m_Inv.ActiveDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    try
                    {
                        m_InvCmdMgr = m_Inv.CommandManager;
                        m_InvCmdMgr.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, m_CompFullFileName);
                        Inventor.ControlDefinition m_InvCtrlDef = (ControlDefinition)m_InvCmdMgr.ControlDefinitions["AssemblyPlaceComponentCmd"];
                        //bring Inventor to front
                        IntPtr mWinPt = (IntPtr)m_Inv.MainFrameHWND;
                        SwitchToThisWindow(mWinPt, true);
                        m_InvCtrlDef.Execute();
                    }
                    catch
                    {

                    }

                }
            }
        }

        //create new file and insert selected component as derived - waiting for user interaction to finalize
        public void m_DeriveComponent(String m_CompFullFileName)
        {
            //create new file and set file name & path accordingly
            if (m_ConnectInv() == true)
            {
                Inventor.PartDocument m_NewPart = (PartDocument)m_Inv.Documents.Add(DocumentTypeEnum.kPartDocumentObject, "", true);

                //insert selected component as derived interactively (show dialog of derive options instead of direct placement)
                if (m_Inv.ActiveDocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    try
                    {
                        m_InvCmdMgr = m_Inv.CommandManager;
                        m_InvCmdMgr.PostPrivateEvent(PrivateEventTypeEnum.kFileNameEvent, m_CompFullFileName);
                        Inventor.ControlDefinition m_InvCtrlDef = (ControlDefinition)m_InvCmdMgr.ControlDefinitions["PartDerivedComponentCmd"];
                        //bring Inventor to front
                        IntPtr mWinPt = (IntPtr)m_Inv.MainFrameHWND;
                        SwitchToThisWindow(mWinPt, true);
                        m_InvCtrlDef.Execute();
                    }
                    catch
                    {
                    }
                }
            }
        }

        //run an iLogic rule providing the rule's name and internal/external option
        public String m_RunRule(String m_RuleName) //, Boolean m_External=false)
        {
            string m_RunRuleResult = null;

            if (m_ConnectInv() == true)
            {            
                //iLogic is also an addin which has its guid
                string iLogicAddinGuid = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";

                Inventor.ApplicationAddIn addin = null;
                try
                {
                    // try to get iLogic addin
                    addin = m_Inv.ApplicationAddIns.get_ItemById(iLogicAddinGuid);
                }
                catch
                {
                    // any error...
                }

                if (addin != null)
                {
                    // activate the addin
                    if (!addin.Activated)
                        addin.Activate();

                    // entrance of iLogic
                    iLogicAutomation.iLogicAutomation _iLogicAutomation = addin.Automation;
                    
                    Document oCurrentDoc = m_Inv.ActiveDocument;

                    dynamic myRule = null;
                    //dump all rules
                    //foreach (dynamic eachRule in _iLogicAutomation.Rules(oCurrentDoc))
                        foreach (Autodesk.iLogic.Interfaces.iLogicRule eachRule in _iLogicAutomation.get_Rules(oCurrentDoc))
                    {
                        if (eachRule.Name == m_RuleName)
                        {
                            myRule = eachRule;
                            //list the code of rule to the list box
                            //MessageBox.Show(myRule.Text);
                            break;
                        }
                    }
                    if (myRule != null)
                    {
                        try
                        {
                            //if (m_External == false) 
                            _iLogicAutomation.RunRule(oCurrentDoc, m_RuleName);
                            //if (m_External == true) _iLogicAutomation.RunExternalRule(oCurrentDoc, m_RuleName);
                            return m_RunRuleResult= "Rule successfully executed";
                        }
                        catch
                        {
                            return m_RunRuleResult = "Rule execution failed";
                        }
                    }
                    else
                    {
                        return m_RunRuleResult = "Indicated Rule not found!";
                    }   
                }
                else
                {
                    return m_RunRuleResult = "iLogic Module not loaded; rule could not be called";
                }
            }
            else
            {
                return m_RunRuleResult = "Inventor Application not found; rule could not be called";
            }
   
        }

        public Boolean m_ConnectInv ()
        {
            // Try to get an active instance of Inventor
            try
                {
                    m_Inv = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
                    return true;
                }
            catch
                {
                    return false;
                }
        }
       
    }

    public class AcadHelpers
    {
        AcInterop.AcadApplication m_Acad = null;
        private const string progID = "AutoCAD.Application";
       
        [System.Runtime.InteropServices.DllImport("User32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
       

        public Boolean m_ConnectAcad ()
        {
            try
            {
                m_Acad = (AcInterop.AcadApplication)System.Runtime.InteropServices.Marshal.GetActiveObject(progID);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void m_GoToAcad ()
        {
            if (m_ConnectAcad() == true)
            {
                try
                {
                    AcInterop.AcadDocument m_AcDoc = m_Acad.ActiveDocument;
                    IntPtr mWinPt = (IntPtr)m_Acad.HWND;
                    SwitchToThisWindow(mWinPt, true);
                    //String m_Command = @"(Command ""_Insert"" ""C:/AB_Vault/Konstruktion/01-0080.dwg"") "";"" ";
                    //m_AcDoc.SendCommand(m_Command);
                 }
                catch
                {
                }
            }            
        }
    }
}

