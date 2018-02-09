using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using VCF = Autodesk.DataManagement.Client.Framework;

using AWS = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Currency;
using VB = Connectivity.Application.VaultBase;

namespace QuickstartiLogicLibrary
{
    public class QuickstartiLogicLib :IDisposable
    {
        public void Dispose()
        {
            //do clean up here if required
        }

        public VDF.Vault.Currency.Connections.Connection mGetVaultConn()
        {
            VDF.Vault.Currency.Connections.Connection mConn = VB.ConnectionManager.Instance.Connection;
            if (mConn != null)
            {
                return mConn;
            }
            return null;
        }

        public string mGetFileByFullFileName(VDF.Vault.Currency.Connections.Connection conn, string mVaultFile)
        {
            List<string> mFiles = new List<string>();
            mFiles.Add(mVaultFile);
            AWS.File[] wsFiles = conn.WebServiceManager.DocumentService.FindLatestFilesByPaths(mFiles.ToArray());
            VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn,(wsFiles[0]));

            VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
            settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
            settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
            settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
            VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
            mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
            mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
            settings.AddFileToAcquire(mFileIt, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
            VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
            if (results != null)
            {
                VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                return mFilesDownloaded.LocalPath.FullPath.ToString();
            }
            return "FileNotFound";
        }

        public string mGetFilebySearchCriteria(VDF.Vault.Currency.Connections.Connection conn, Dictionary<string,string> mSearchCriteria)
        {
            AWS.PropDef[] mFilePropDefs = conn.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
            //iterate mSearchcriteria to get property definitions and build AWS search criteria
            List<AWS.SrchCond> mSrchConds = new List<AWS.SrchCond>();
            int i = 0;
            List<AWS.File> totalResults = new List<AWS.File>();
            foreach (var item in mSearchCriteria)
            {
                AWS.PropDef mFilePropDef = mFilePropDefs.Single(n => n.DispName == item.Key);
                AWS.SrchCond mSearchCond = new AWS.SrchCond();
                {
                    mSearchCond.PropDefId = mFilePropDef.Id;
                    mSearchCond.PropTyp = AWS.PropertySearchType.SingleProperty;
                    mSearchCond.SrchOper = 1; //equals
                    if(i == 0) mSearchCond.SrchRule = AWS.SearchRuleType.May;
                    else mSearchCond.SrchRule = AWS.SearchRuleType.Must;
                    mSearchCond.SrchTxt = item.Value;
                }
                mSrchConds.Add(mSearchCond);
                i++;
            }
            string bookmark = string.Empty;
            AWS.SrchStatus status = null;
                
            while (status == null || totalResults.Count < status.TotalHits)
            {
                AWS.File[] mSrchResults = conn.WebServiceManager.DocumentService.FindFilesBySearchConditions(
                    mSrchConds.ToArray(), null, null, false, true, ref bookmark, out status);
                if (mSrchResults != null) totalResults.AddRange(mSrchResults);
                else break;
            }
            //if results not empty
            if (totalResults.Count == 1)
            {
                AWS.File wsFile = totalResults.First<AWS.File>();
                VDF.Vault.Currency.Entities.FileIteration mFileIt = new VDF.Vault.Currency.Entities.FileIteration(conn, (wsFile));

                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(conn);
                settings.DefaultAcquisitionOption = VCF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                settings.OptionsRelationshipGathering.IncludeLinksSettings.IncludeLinks = false;
                VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
                mResOpt.OverwriteOption = VCF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
                mResOpt.SyncWithRemoteSiteSetting = VCF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;
                settings.AddFileToAcquire(mFileIt, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
                VDF.Vault.Results.AcquireFilesResults results = conn.FileManager.AcquireFiles(settings);
                if (results != null)
                {
                    VDF.Vault.Results.FileAcquisitionResult mFilesDownloaded = results.FileResults.Last();
                    return mFilesDownloaded.LocalPath.FullPath.ToString();
                }
                else
                {
                    return "FileNotFound";
                }
            }
            else
            {
                return "FileNotFound";
            }
            
        }
    }
}
