package source;

import java.io.IOException;
//import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

import wsc.Config;
import wsc.MetadataLoginUtil;
import wsc.WSC;

import com.sforce.soap.partner.GetUserInfoResult;
import com.sforce.soap.partner.fault.ExceptionCode;
import com.sforce.soap.partner.fault.UnexpectedErrorFault;
//import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class UtilExportInfc {
	/**
	 * 
	 * Read MetaData
	 * @throws Exception 
	 */
	//選択したメタデータより、各メタデータの作成クラスを呼び出す
	public void ExportInfc () throws Exception{
		Util.logger.info("ExportInfc Start.");
		//Util ut = new Util();
		//Loop Types list
		Set<Map.Entry<String, List<String>>> entryseSet = UtilConnectionInfc.getExportMap().entrySet();
		List<String> objectsList1 = null;
		for (Map.Entry<String, List<String>> entry : entryseSet) {
			String type = entry.getKey();
			List<String> objectsList = entry.getValue();
			
			if(type.equals("NetworkBranding")){
				objectsList1 = objectsList;
				break;
			}
		}
		for (Map.Entry<String, List<String>> entry : entryseSet) {
			if(entry.getValue().size()>0){
				//if session time out, the re login
				//try {
					GetUserInfoResult result = MetadataLoginUtil.Connection.getUserInfo();
					//System.out.println(result.getUserFullName());
				/*} catch (UnexpectedErrorFault uef) {
				    if (uef.getExceptionCode() == ExceptionCode.INVALID_SESSION_ID) {
				        // Re-authenticate the user
				    	Properties propertiesout = new Properties();
				    	if(WSC.isBatch){
				    		propertiesout=Config.properties;
				    	}else{
				    		propertiesout=WSC.propertiesout;
				    	}
				    	String username = propertiesout.getProperty("sfdc.username");
				    	String password =Util.decrypt(propertiesout.getProperty("sfdc.password"));
				    	String url = propertiesout.getProperty("sfdc.endpoint");
						String proxyHost=propertiesout.getProperty("sfdc.proxyHost");
						String proxyPort=propertiesout.getProperty("sfdc.proxyPort");
						String proxyUsername=propertiesout.getProperty("sfdc.proxyUsername");
						String proxyPassword=Util.decrypt(propertiesout.getProperty("sfdc.proxyPassword"));
				    	MetadataLoginUtil reLogin = new MetadataLoginUtil(username,password.toCharArray(),url.contains("login")?true:false,proxyHost,proxyPort,proxyUsername,proxyPassword);
				    	MetadataLoginUtil.metadataConnection = reLogin.userLogin();
				    }
				}*/
				String type = entry.getKey();
				List<String> objectsList = entry.getValue();
				Util.logger.info("Starting type="+type);
				Util.logger.info("Starting objectsList="+objectsList);
				
				if(type.equals("AnalyticSnapshot")){
					ReadAnalyticSnapshotSync readAnalyticSnapshot = new ReadAnalyticSnapshotSync();
					readAnalyticSnapshot.readAnalyticSnapshot(type, objectsList);
				}
				if( type.equals("ApexClass")){
					ReadApexClassSync readApexClass = new ReadApexClassSync();
					readApexClass.readApexClass(type, objectsList);
				}
				if( type.equals("ApexComponent")){
					ReadApexComponentSync readApexComponent = new ReadApexComponentSync();
					readApexComponent.readApexComponent(type, objectsList);
				}
				if(type.equals("ApexPage")){
					ReadApexPageSync readApexPageSync = new ReadApexPageSync();
					readApexPageSync.readApexPage(type, objectsList);
				}
				if( type.equals("ApexTrigger")){
					ReadApexTriggerSync readApexTrigger = new ReadApexTriggerSync();
					readApexTrigger.readApexTrigger(type, objectsList);
				}
				if( type.equals("ApprovalProcess")){
					ReadApprovalProcessSync readApprovalProcess = new ReadApprovalProcessSync();
					readApprovalProcess.readApprovalProcess(type,objectsList);
				}
				if(type.equals("AssignmentRules")){
					ReadAssignmentRuleSync readAssignmentRule = new ReadAssignmentRuleSync();
					readAssignmentRule.readAssignmentRule(type,objectsList);
				}
				if(type.equals("AutoResponseRules")){
					ReadAutoResponseRuleSync readAutoResponseRule = new ReadAutoResponseRuleSync();
					readAutoResponseRule.readAutoResponseRule(type,objectsList);
				}
				if(type.equals("AuraDefinitionBundle")){
					ReadAuraDefinitionBundleSync readAuraDefinitionBundle = new ReadAuraDefinitionBundleSync();
					readAuraDefinitionBundle.readAuraDefinitionBundle(type,objectsList);
				}
				if( type.equals("CustomLabels")){
					ReadCustomLabelSync readCLObject = new ReadCustomLabelSync();
					readCLObject.readCustomLabel(type, objectsList);
				}
				if( type.equals("StandardObject")){
					ReadCustomObjectSync readCustomObject = new ReadCustomObjectSync();
					readCustomObject.readCustomObject(type, objectsList);
				}
				if( type.equals("CustomObject")){
					ReadCustomObjectSync readCustomObject = new ReadCustomObjectSync();
					readCustomObject.readCustomObject(type, objectsList);
				}
				if( type.equals("EXTERNALOBJECT")){
					ReadCustomObjectSync readCustomObject = new ReadCustomObjectSync();
					readCustomObject.readCustomObject(type, objectsList);
				}
				if(type.equals("CustomSetting")){
					ReadCustomSettingSync readCustomSettingSync = new ReadCustomSettingSync();
					readCustomSettingSync.readCustomSetting(type, objectsList);
				}
				if(type.equals("CustomTab")){
					ReadCustomTabSync readCustomTab = new ReadCustomTabSync();
					readCustomTab.readCustomTab(type, objectsList);
				}
				if(type.equals("ExternalDataSource")){
					ReadExternalDataSourceSync readCustomTab = new ReadExternalDataSourceSync();
					readCustomTab.readExternalDataSource(type, objectsList);
				}				
				if(type.equals("UserGroup")){
					ReadGroupRoleQueueSync readGroupRoleQueue = new ReadGroupRoleQueueSync();
					readGroupRoleQueue.readGroupRoleQueue(type, objectsList);
				}
				if( type.equals("HomePage")){
					ReadHomePageSync readHome = new ReadHomePageSync();
					readHome.readHomePage(type, objectsList);
				}
				if( type.equals("Layout")){
					ReadLayoutSync readLayout = new ReadLayoutSync();
					readLayout.readLayout(type,objectsList);					
				}
				if(type.equals("Profile")){
					ReadProfileSync readProfile = new ReadProfileSync();
					readProfile.readProfile(type, objectsList);
				}
				if( type.equals("Report")){
					ReadReportSync readReport = new ReadReportSync();
					readReport.readReportFloder(type, objectsList);
				}
				if( type.equals("Dashboard")){
					ReadDashboardSync readDashboard = new ReadDashboardSync();
					readDashboard.readDashboardFloder(type, objectsList);
				}				
				if(type.equals("ReportType")){
					ReadReportTypeSync reporttype=new ReadReportTypeSync();
					reporttype.readReportType(type, objectsList);
				}
				if( type.equals("Settings")){
					ReadSettingsSync readSettings = new ReadSettingsSync();
					readSettings.readSettings(type, objectsList);
				}
				if( type.equals("StaticResource")){
					ReadStaticResourceSync readStaticResource = new ReadStaticResourceSync();
					readStaticResource.ReadStaticResource(type,objectsList);
				}
				if( type.equals("Workflow")){
					ReadWorkFlowSync readWorkFlow = new ReadWorkFlowSync();
					readWorkFlow.readWorkFlow(type,objectsList);
				}
				if( type.equals("EntitlementProcess")){
					ReadEntitlementProcessesSync readEntitlementProcesses= new ReadEntitlementProcessesSync();
					readEntitlementProcesses.readEntitlementProcesses(type,objectsList);
				}
				if(type.equals("CustomObjectTranslation")){
					ReadCustomObjectTranslationSync customTranslation = new ReadCustomObjectTranslationSync();
					customTranslation.readCustomObjectTranslation(type, objectsList);
				}

				if( type.equals("EscalationRules")){
					ReadEscalationRuleSync readEscalationRule = new ReadEscalationRuleSync();
					readEscalationRule.readEscalationRule(type,objectsList);
				}		
				if(type.equals("Document")){
					ReadDocumentSync documentSync = new ReadDocumentSync();
					documentSync.readDocumentFolder(type, objectsList);
				}	
				
				if(type.equals("SharingSetting")){
					ReadSharingSettingsSync sharingSync = new ReadSharingSettingsSync();
					sharingSync.readSharingSettings(type, objectsList);
				}
				if(type.equals("UserGroup")){
					ReadGroupRoleQueueSync groupSync = new ReadGroupRoleQueueSync();
					groupSync.readGroupRoleQueue(type, objectsList);
				}
				if(type.equals("PermissionSet")){
					ReadPermissionSetSync PermissionSetSync = new ReadPermissionSetSync();
					PermissionSetSync.readpermissionSet(type, objectsList);
					
				}
				
				if(type.equals("Network")){
					ReadNetworkSync readNetwork = new ReadNetworkSync();
					readNetwork.readNetwork(type, objectsList, objectsList1);
				}
				if(type.equals("CustomMetadata")){
					ReadCustomMetadataSync customMetadataSync = new ReadCustomMetadataSync();
					customMetadataSync.ReadCustomMetadata(type, objectsList);
				}				
			}else{
				Util.logger.warn("***********No result to export!************");
			}
		}
		Util.logger.info("ExportInfc Complete.");	
	}

}
