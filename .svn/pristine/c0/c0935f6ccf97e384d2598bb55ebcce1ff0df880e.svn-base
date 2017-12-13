package source;

import java.util.List;
import java.util.Map;

import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.partner.PartnerConnection;
import com.sforce.soap.partner.sobject.SObject;

public class UtilConnectionInfc {
	
	public static MetadataConnection metadataConnection;
	
	public static SObject[] userInfo;
	
	public static PartnerConnection Connection; 
	
	public static String downloadPath;
	
	public static Long lastUpdateTime;
	
	public static String language;
	
	public static Map<String,List<String>> exportMap;
	
	public static Boolean settingSubfolder;    //by  wangyuxin 
	
	public static Boolean modifiedFlag;

	public static Map<String,String> apiToLabel;
	
	public static Map<String,String> eMailTemplateLabel;
	//データの受け渡し
	public static SObject[] getUserInfo() {
		return userInfo;
	}
	public static void setUserInfo(SObject[] userInfo) {
		UtilConnectionInfc.userInfo = userInfo;
	}
	public static Map<String,String> geteMailTemplateLabel() {
		return eMailTemplateLabel;
	}

	public static void seteMailTemplateLabel(Map<String,String> eMailTemplateLabel) {
		UtilConnectionInfc.eMailTemplateLabel = eMailTemplateLabel;
	}
	public static Map<String,String> getApiToLabel() {
		return apiToLabel;
	}

	public static void setApiToLabel(Map<String,String> apiToLabel) {
		UtilConnectionInfc.apiToLabel = apiToLabel;
	}	
	
	public static Boolean getModifiedFlag() {
		return modifiedFlag;
	}

	public static void setModifiedFlag(Boolean modifiedFlag) {
		UtilConnectionInfc.modifiedFlag = modifiedFlag;
	}
	
	public static Boolean getSettingSubfolder() {
		return settingSubfolder;
	}

	public static void setSettingSubfolder(Boolean settingSubfolder) {
		UtilConnectionInfc.settingSubfolder = settingSubfolder;
	}

	public static MetadataConnection getMetadataConnection() {
		return metadataConnection;
	}

	public static void setMetadataConnection(MetadataConnection metadataConnection) {
		UtilConnectionInfc.metadataConnection = metadataConnection;
	}

	public static PartnerConnection getConnection() {
		return Connection;
	}

	public static void setConnection(PartnerConnection connection) {
		Connection = connection;
	}

	public static String getDownloadPath() {
		return downloadPath;
	}

	public static void setDownloadPath(String downloadPath) {
		UtilConnectionInfc.downloadPath = downloadPath;
	}

	public static Long getLastUpdateTime() {
		return lastUpdateTime;
	}

	public static void setLastUpdateTime(Long lastUpdateTime) {
		UtilConnectionInfc.lastUpdateTime = lastUpdateTime;
	}

	public static String getLanguage() {
		return language;
	}

	public static void setLanguage(String language) {
		UtilConnectionInfc.language = language;
	}

	public static Map<String,List<String>> getExportMap() {
		return exportMap;
	}

	public static void setExportMap(Map<String,List<String>> exportMap) {
		UtilConnectionInfc.exportMap = exportMap;
	}
	
	
	
}
