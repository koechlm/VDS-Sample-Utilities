﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Connectivity.WebServices;
using Autodesk.Connectivity.WebServicesTools;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.PersistentId;
using Inventor;
using AcInterop = Autodesk.AutoCAD.Interop;
using AcInteropCom = Autodesk.AutoCAD.Interop.Common;

/// <summary>
/// Quickstart Library to extend VDS scripting capabilities.
/// </summary>
namespace VDSUtils
{
    /// <summary>
    /// Quickstart Library to extend VDS scripting capabilities.
    /// </summary>
    public class VltHelpers

    {
        /// <summary>
        /// UserCredentials1 and UserCredentials2 differentiate overloads as powershell can't handle
        /// UserCredentials1 returns read-write loginuser object
        /// </summary>
        /// <param name="server">IP Address or DNS Name of ADMS Server</param>
        /// <param name="vault">Name of vault to connect to</param>
        /// <param name="user">User name</param>
        /// <param name="pw">Password</param>
        /// <returns>User Credentials</returns>
        public Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials UserCredentials1(string server, string vault, string user, string pw)
        {
            ServerIdentities mServer = new ServerIdentities();
            mServer.DataServer = server;
            mServer.FileServer = server;
             Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials mCred = new Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials(mServer, vault, user, pw);
            return mCred;
        }

        /// <summary>
        /// UserCredentials1 and UserCredentials2 differentiate overloads as powershell can't handle
        /// UserCredentials2 returns readonly loginuser object
        /// </summary>
        /// <param name="server">IP Address or DNS Name of ADMS Server</param>
        /// <param name="vault">Name of vault to connect to</param>
        /// <param name="user">User name</param>
        /// <param name="pw">Password</param>
        /// <param name="rw">Set to "True" to allow Read/Write access</param>
        /// <returns></returns>
        public Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials UserCredentials2(string server, string vault, string user, string pw, bool rw = true)
        {
            ServerIdentities mServer = new ServerIdentities();
            mServer.DataServer = server;
            mServer.FileServer = server;
            Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials mCred = new Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials(mServer, vault, user, pw, rw);
            return mCred;
        }

        /// <summary>
        /// Deprecated - no longer required, as the overload is removed in 2017 API
        /// </summary>
        /// <param name="svc"></param>
        /// <param name="FldIds"></param>
        /// <param name="m_PropArray"></param>
        /// <returns></returns>
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

        /// <summary>
        /// LinkManager.GetLinkedChildren has an override list; the input is of type IEntity. 
        /// This wrapper allows to input commonly known object types, like Ids and entity names instead.
        /// </summary>
        /// <param name="con">The utility dll is not connected to Vault; 
        /// we need to leverage the established connection to call LinkManager methods</param>
        /// <param name="mId">The parent entity's id to get linked children of</param>
        /// <param name="mClsId">The parent entity's class name; allowed values are FILE FLDR and CUSTENT. 
        /// CO and ITEM cannot have linked children, as they use specific links to related child objects.</param>
        /// <param name="mFilter">Limit the search on links to a particular class; providing an empty value "" will result in a search on all types</param>
        /// <returns>List of entity Ids</returns>
        public List<long> mGetLinkedChildren1(Connection con, long mId, string mClsId, string mFilter)
        {
            IEnumerable<PersistableIdEntInfo> mEntInfo = new PersistableIdEntInfo[] { new PersistableIdEntInfo(mClsId, mId, true, false) };
            IDictionary<PersistableIdEntInfo,IEntity> mIEnts = con.EntityOperations.ConvertEntInfosToIEntities(mEntInfo) ;
            IEntity mIEnt = null;
            try
            {
                foreach (var item in mIEnts)
                {
                    mIEnt = item.Value;
                }
                IEnumerable<IEntity> mLinkedChldrn = con.LinkManager.GetLinkedChildren(mIEnt, mFilter);
                //return mLinkedChldrn;
                List<long> mLinkedIds = new List<long>();
                foreach (var item in mLinkedChldrn)
                {
                    mLinkedIds.Add(item.EntityIterationId);
                }
                return mLinkedIds;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Evaluation of overload 2; see mGetLinkedchildren1 for detailed description
        /// </summary>
        /// <param name="con"></param>
        /// <param name="mParEntIds"></param>
        /// <param name="mClsIds"></param>
        /// <returns></returns>
        private IEnumerable<IEntity> GetLinkedChildren2(Connection con, long[] mParEntIds,  string[] mClsIds)
        {
            List<PersistableIdEntInfo> mEntInfo = new List<PersistableIdEntInfo>();
            for (int i = 0; i < mParEntIds.Length; i++)
            {
                mEntInfo.Add(new PersistableIdEntInfo("CUSTENT", mParEntIds[i], true, false));
            }
            //List<CustEnt> mEnts = new List<CustEnt>();
            //CustEnt mEnt = new CustEnt();
            //foreach (var item in mParentEnts)
            //{
            //    mEnt = (CustEnt)item;
            //    mEnts.Add(mEnt);
            //}
            //List<PersistableIdEntInfo> mEntInfo = new List<PersistableIdEntInfo>();
            //foreach (var item in mEnts)

            //{
            //    mEntInfo.Add( new PersistableIdEntInfo(mClsIds[0], item.Id, true, false));
            //}

            IDictionary<PersistableIdEntInfo, IEntity> mIEnts = con.EntityOperations.ConvertEntInfosToIEntities(mEntInfo.AsEnumerable());
            List<IEntity> mIEnt = new List<IEntity>();
            try
            {
                foreach (var item in mIEnts)
                {
                    mIEnt.Add(item.Value);
                }
                IEnumerable<IEntity> mLinkedChldrn = con.LinkManager.GetLinkedChildren(mIEnt.AsEnumerable(), mClsIds.AsEnumerable());
                return mLinkedChldrn;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Library of VDS Quickstart calls to hosting Inventor session
    /// </summary>
    public class InvHelpers
    {
        Inventor.Application m_Inv = null;
        Inventor.Document m_Doc = null;
        Inventor.DrawingDocument m_DrawDoc = null;
        Inventor.PresentationDocument m_IpnDoc = null;
        String m_ModelPath = null;

        [System.Runtime.InteropServices.DllImport("User32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        /// <summary>
        /// Retrieve property value of main view referenced model
        /// </summary>
        /// <param name="m_InvApp">Connect to the hosting instance of the VDS dialog</param>
        /// <param name="m_ViewModelFullName"></param>
        /// <param name="m_PropName">Display Name</param>
        /// <returns></returns>
        public object m_GetMainViewModelPropValue(object m_InvApp, String m_ViewModelFullName, String m_PropName)
        {
            try
            {
                m_Inv = (Inventor.Application)m_InvApp;
                m_Doc = m_Inv.Documents.Open(m_ViewModelFullName, false);
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

        /// <summary>
        /// Gets the 3D model (ipt/iam/ipn) linked to the main view of the current (active) drawing.
        /// Gets the 3D model (iam) linked to the main view of the current (active) presentation.
        /// </summary>
        /// <param name="m_InvApp">Running host (instance of Inventor) of calling VDS Dialog.</param>
        /// <returns>Returns the fullfilename (path and filename incl. extension) of the referenced model as string.</returns>
        public String m_GetMainViewModelPath(object m_InvApp)
        {
            try
            {
                m_Inv = (Inventor.Application)m_InvApp;
                //if (m_ConnectInv()==true)
                //{
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
                //}
            }
            catch
            {
                return m_ModelPath = "";
            }
            return m_ModelPath = "";
        }

    }
    /// <summary>
    /// Library of VDS Quickstart calls to hosting AutoCAD session
    /// </summary>
    public class AcadHelpers
    {
        AcInterop.AcadApplication m_Acad = null;
        private const string progID = "AutoCAD.Application";
       
        [System.Runtime.InteropServices.DllImport("User32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
       
        /// <summary>
        /// Get AutoCAD session hosting; deprecated as VDS 2017 dialogs share the hosting application object
        /// </summary>
        /// <returns></returns>
        private Boolean m_ConnectAcad ()
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

        /// <summary>
        /// Switch running AutoCAD application; requires updated - VDS 2017 shares application object in VDS Dialog
        /// </summary>
        private void m_GoToAcad ()
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

