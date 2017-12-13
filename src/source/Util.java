package source;

import java.io.BufferedInputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.security.InvalidKeyException;
import java.security.KeyFactory;
import java.security.NoSuchAlgorithmException;
import java.security.PublicKey;
import java.security.spec.X509EncodedKeySpec;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.TimeZone;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.crypto.BadPaddingException;
import javax.crypto.Cipher;
import javax.crypto.IllegalBlockSizeException;
import javax.crypto.NoSuchPaddingException;

import wsc.MetadataLoginUtil;
import wsc.WSC;

import com.sforce.soap.metadata.AuraBundleType;
import com.sforce.soap.metadata.AuraDefinitionBundle;
import com.sforce.soap.metadata.CustomField;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.DescribeMetadataObject;
import com.sforce.soap.metadata.DescribeMetadataResult;
import com.sforce.soap.metadata.EmailFolder;
import com.sforce.soap.metadata.EmailTemplate;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.FilterItem;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.ReadResult;
import com.sforce.soap.metadata.RecordType;
import com.sforce.soap.metadata.SaveResult;
import com.sforce.soap.metadata.SharedTo;
import com.sforce.soap.partner.PartnerConnection;
import com.sforce.soap.partner.QueryResult;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.soap.tooling.ToolingConnection;
import com.sforce.ws.ConnectionException;

import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.security.InvalidAlgorithmParameterException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.output.ByteArrayOutputStream;
//log4j
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Util {
	
	
	public static Logger logger = LogManager.getLogger(Util.class.getName());
	/**
    * static initializer による初期化
	*/
    static {
    	
    }	
		
 
    //MetadataConnection metadataConnection = MetadataLoginUtil.metadataConnection;

	PartnerConnection Connection = MetadataLoginUtil.Connection;
	public ToolingConnection Connection2 = MetadataLoginUtil.toolingConnection;
	public static final double API_VERSION = 41.0;
	public static final int RowIntervalNum = 2;
	public static java.util.Properties prop;

	/**
	 * 
	 * Get MetaData from SalesForce
	 * @param
	 * 		type			
	 * 			MetaData Type
	 * 		objectsList
	 * 			objects name's list
	 * @return
	 * 		List<Metadata>
	 */
	//SalesForceからメタデータを取得する
	public List<Metadata> readMateData(String type,List<String> objectsList){
		//Change type back to metadata defined
		List<String> derivedMetaDataType = Arrays.asList("CustomObject","StandardObject","SharingSetting","CustomSetting","ArticleType","EXTERNALOBJECT");
		if(derivedMetaDataType.contains(type)){
			type="CustomObject";
		}		
		List<Metadata> mdInfoList = new ArrayList<Metadata>();
		int objListSize = objectsList.size();
		int countTemp = 0;
		while(objListSize>countTemp){
			String[] objects;
			List<String> tempList;
			Integer MAXPERTIME=10;
			if(type=="CustomObject"){
				MAXPERTIME=1;
			}
			if( objListSize-countTemp>MAXPERTIME ){
				tempList = objectsList.subList(countTemp, countTemp+MAXPERTIME);
				objects = new String[MAXPERTIME];
				tempList.toArray(objects);
				countTemp += MAXPERTIME;
			}else{
				tempList = objectsList.subList(countTemp, objListSize);
				objects = new String[objListSize-countTemp];
				tempList.toArray(objects);
				countTemp = objListSize;
			}
			try {
				List<Metadata> mdInfoTempList = new ArrayList<Metadata>();
				ReadResult readResult = MetadataLoginUtil.metadataConnection.readMetadata(type, objects);
				Metadata[] mdInfos = readResult.getRecords();
				mdInfoTempList = Arrays.asList(mdInfos);
				mdInfoList.addAll(mdInfoTempList);
			} catch (ConnectionException e) {
				Util.logger.error("readMateData error"+e.getStackTrace());
				StringWriter sw = new StringWriter();
				e.printStackTrace(new PrintWriter(sw));
				String stackTrace = sw.toString();
				Util.logger.debug("Get the exception String"+stackTrace);
			}
		}
		return mdInfoList;
	}
	
	/**
	 * API Query method(toolingConnection)
	 * 
	 * @param
	 * 		String
	 * 			SQL
	 * return 
	 * 		SObject []
	 */
	//オブジェクトApexClass、ApexTriggerからデータを取得する
	public com.sforce.soap.tooling.sobject.SObject[] apiQuery2(String sql) throws ConnectionException{
		com.sforce.soap.tooling.QueryResult qr = Connection2.query(sql);
		return qr.getRecords();
	}
    
	
	//APIよりemailTemplateLabel値を戻る
    public String getEmailTemplateLabel(String api){
    	String label="";
    	if(UtilConnectionInfc.geteMailTemplateLabel().get(api)!=null){
    		label=UtilConnectionInfc.geteMailTemplateLabel().get(api)+"("+api+")";
    	}else if(api!=null&&!api.equals("")){
    		label=api+"("+api+")";
    	}
    	return label;
    }
	/**
	 * API Query method(partnerConnection)
	 * 
	 * @param
	 * 		String
	 * 			SQL
	 * return 
	 * 		SObject []
	 */
	//UserLicense、User、Group、CustomSettingの情報を取得する
	public SObject[] apiQuery(String sql) throws ConnectionException{

		Connection.setQueryOptions(500);
		QueryResult qr = Connection.query(sql);
		boolean done = false;
		SObject[] moreRecords = new SObject[qr.getSize()];
		int x=0;
		if(qr.getSize() > 0) {	
			while(!done) {
				SObject[] records = qr.getRecords();
				for(int i=0; i<records.length; i++){
					moreRecords[x] = records[i];
					x++;
				}
				if(qr.isDone()){
					done = true;
				}else{
					qr = Connection.queryMore(qr.getQueryLocator());
				}
			}
		}
		return moreRecords;
	}
	
	//EmailTemplateの情報を取得し、結果をUtilConnectionInfc.eMailTemplateLabelに設定する
	public static void getEMailTemplateFromServer() throws ConnectionException{
		List<String> allFolders  = new ArrayList<String>();
		Map<String,String> resultMap = new LinkedHashMap<String,String>();
		try{			
			ListMetadataQuery queryReportF = new ListMetadataQuery();
			queryReportF.setType("Email"+"Folder");
			FileProperties[] rfp = MetadataLoginUtil.metadataConnection.listMetadata(
					new ListMetadataQuery[] { queryReportF }, Util.API_VERSION);
			if (rfp != null) {	
				for(FileProperties f : rfp){
					//if(f.getNamespacePrefix()!=null&&f.getManageableState()!=null&&(f.getManageableState().toString().equals("unmanaged")||f.getManageableState().toString().equals("released"))){
					//	allFolders.add(f.getNamespacePrefix()+"__"+URLDecoder.decode(f.getFullName(),"utf-8"));
					//}else{
						allFolders.add(URLDecoder.decode(f.getFullName(),"utf-8"));
					//}					
				}
				Collections.sort(allFolders);
			}
//			ConnectionException|UnsupportedEncodingException ce
		} catch (Exception ce) {
			ce.getStackTrace();
		}
		Util ut = new Util();
		List<String> list = new ArrayList<String>();
		for (String str : allFolders) {
			ListMetadataQuery queries = new ListMetadataQuery();
			queries.setType("EmailTemplate");
			queries.setFolder(str);
			FileProperties[] fileProperties = MetadataLoginUtil.metadataConnection.listMetadata(
					new ListMetadataQuery[] { queries }, Util.API_VERSION);
			for (FileProperties f : fileProperties) {
				list.add(f.getFullName());
			}
			//resultMap = ut.getComparedResult(type,
			//		str, UtilConnectionInfc.getLastUpdateTime());
			//this.readReport(type, list);
		}
		List<Metadata> mdInfos = ut.readMateData("EmailTemplate", list);
		for (Metadata md : mdInfos) {
			if (md != null) {
				EmailTemplate et=(EmailTemplate)md;
				if(et.getName()!=null){
					resultMap.put(et.getFullName(),et.getName());
				}else{
					resultMap.put(et.getFullName(),et.getFullName());
				}
			}
		}
		UtilConnectionInfc.seteMailTemplateLabel(resultMap);
	}
	
	/*public static void  getTest(){
		Util ut = new Util();
		try {
		Map<String,String> resultMap = new LinkedHashMap<String,String>();
		List<String> k = new ArrayList<String>();
		ListMetadataQuery query = new ListMetadataQuery();
		//query.setFolder("AuraFolder");
		query.setType("AuraDefinitionBundle");
		//query.setType("CustomObject");
		FileProperties[] filePro = MetadataLoginUtil.metadataConnection.listMetadata(new ListMetadataQuery[] {query}, API_VERSION);			
		if (filePro != null) {
			List<String> allFile = new ArrayList<String>();
			for (FileProperties fPro : filePro) {
				allFile.add(URLDecoder.decode(fPro.getFullName(),"utf-8"));

			}
			Collections.sort(allFile);
			//List<String> tt=new ArrayList<String>();
			//tt.add("HelloWorld");
			List<Metadata> mdInfos= ut.readMateData("AuraDefinitionBundle",allFile);
			//ReadResult mdInfos =  MetadataLoginUtil.metadataConnection.readMetadata("AuraDefinitionBundle", new String[]{"mvc__CalendarComponent", "mvc__SelectEventMonthView", "mvc__SelectUserComponent", "mvc__SelectUserEvent"});
		//List<Metadata> mdInfos = ut.readMateData("CustomObject",k);
		for (Metadata md : mdInfos) {
			if (md != null) {
				// Create CustomObject object
				AuraDefinitionBundle obj = (AuraDefinitionBundle) md;
				resultMap.put(obj.getFullName(), obj.getFullName());
			}
		}}}catch (ConnectionException|UnsupportedEncodingException e) {
			Util.logger.error("readMateData error"+e.getStackTrace());
		}
	}*/
	public String getUserLabel(String fieldName,String searchInfo){
		String result=searchInfo;
		for(SObject obj : UtilConnectionInfc.getUserInfo()){
			if(Util.nullFilter(obj.getField(fieldName)).equals(searchInfo)){
				result=Util.nullFilter(obj.getField("Name"));
			}
		}
		return result;
	}
	public static void getUserInfoFromServer() throws ConnectionException{
		Util ut = new Util();
		String sql = "Select Id,Name,Username from User";
	    SObject[] objs = ut.apiQuery(sql);
	    UtilConnectionInfc.setUserInfo(objs);
	}
	/**
	 * get object label and api name
	 * result in UtilConnectionInfc.ApiToLabel
	 */
	//APIとLabelの情報を取得し、結果をUtilConnectionInfc.ApiToLabelに設定する
	public static void getApiToLabelFromServer(){
		Map<String,String> resultMap = new LinkedHashMap<String,String>();
		Util ut = new Util();
		try {
			ListMetadataQuery query = new ListMetadataQuery();
			query.setType("CustomObject");
			FileProperties[] filePro = MetadataLoginUtil.metadataConnection.listMetadata(new ListMetadataQuery[] {query}, API_VERSION);			
			if (filePro != null) {
				List<String> allFile = new ArrayList<String>();
				for (FileProperties fPro : filePro) {
					allFile.add(URLDecoder.decode(fPro.getFullName(),"utf-8"));

				}
				Collections.sort(allFile);
				List<Metadata> mdInfos = ut.readMateData("CustomObject", allFile);
				for (Metadata md : mdInfos) {
					if (md != null) {
						// Create CustomObject object
						CustomObject obj = (CustomObject) md;
						resultMap.put(obj.getFullName().toUpperCase(), obj.getLabel());
						if(obj.getFields().length>0){		
							for( Integer i=0; i<obj.getFields().length; i++ ){
								CustomField cf = (CustomField)obj.getFields()[i];
								if(obj.getFullName().equals("Activity")){
									resultMap.put("TASK"+"."+cf.getFullName().toUpperCase(), cf.getLabel());
									resultMap.put("EVENT"+"."+cf.getFullName().toUpperCase(), cf.getLabel());
								}
								resultMap.put(obj.getFullName().toUpperCase()+"."+cf.getFullName().toUpperCase(), cf.getLabel());
							}
						}
						if(obj.getRecordTypes().length>0){
							for(RecordType rt: obj.getRecordTypes()){
								resultMap.put(obj.getFullName().toUpperCase()+"."+rt.getFullName().toUpperCase(), rt.getLabel());
							}
						}
					}
				}
			}
			resultMap.put("Organization-Wide Defaults".toUpperCase(),String.valueOf(Util.getTranslate("COMMON","DEFAULTSHARING")));
			UtilConnectionInfc.setApiToLabel(resultMap);
		} catch (ConnectionException|UnsupportedEncodingException e) {
			Util.logger.error("readMateData error"+e.getStackTrace());
		}
	}
	/**
	 * @param
	 * 		objectsList
	 * 			objects name's list
	 * @return
	 * 		string of the objects' names
	 */
	//ApexClass、ApexTrigger、Groupのnamelistを文字列の形で戻る
	public String getObjectNames(List<String> objectsList) {
		String names = "";
		for(String name : objectsList ){
			if(name.contains("__")){
				name=name.split("__")[1];
			}
			names += "'"+name+"'"+",";
		}
		names = names.substring(0, names.length()-1);
		return names;
	}

	/**
	 * Method to export source file
	 * 
	 * @param
	 * 		String 
	 * 			FolderName
	 * 		exportList 
	 * 			String[]{FileName,FileBody}
	 * @return
	 * 		void
	 */
	//ApexClass、ApexTrigger、ApexPage、ApexComponentのソースファイル出力
	public void exportSourceFile(String FolderName,List<String []> exportList) throws IOException{
		String folderPath = UtilConnectionInfc.getDownloadPath()+"\\"+FolderName;
		File file = new File(folderPath);
		//check whether the folder exists,if not, create a new one.
		if (!file.exists()) {
			file.mkdir();
		}
		for(String [] source:exportList){
			//Create files downLoad path
			String path = folderPath+"\\"+ source[0];
			BufferedWriter  fos = new BufferedWriter (new FileWriter(path));
			fos.write(source[1]);
			fos.flush();
			fos.close();
		}
	}

	/**
	 * Method to load multi-language file 
	 * 
	 * @param language
	 * result in this.prop
	 * @throws IOException
	 */
	//multi-languageの配置ファイルをロードする
	public static void LoadProperties(String language) throws IOException{
		String path;
		if(WSC.isBatch){
			path = ".././common/properties/Application_"+language+".properties";
		}else{
			path = "./common/properties/Application_"+language+".properties";
		}
		InputStream input = new BufferedInputStream(new FileInputStream(path));
		prop = new Properties();
		prop.load(input); 
	}

	/**
	 * Method to get translations 
	 * 
	 * @param
	 * 		type
	 * 			metaDataType 
	 * 		name
	 * 			object's name
	 * return
	 * 		filed translation
	 */
	//LoadPropertiesでロードした多言語より訳す内容を戻る
	public static String getTranslate(String type,String name){
		String keyStr=type+"."+name.replace(" ", "");
		if(prop.get(keyStr.toUpperCase()) == null){
			return name;
		}
		return String.valueOf(prop.get(keyStr.toUpperCase()));
	}

	/**
	 * Method to check object whether be changed between two export
	 * 
	 * @param type
	 * @param folder
	 * @param lastUpdateTime
	 * @return checkResultMap<checkFieldName,checkResult>
	 * @throws ConnectionException
	 */
	//前回の実行情報を比較し、変更されるかどうかの情報を戻る
	public Map<String,String> getComparedResult(String type,String folder,Long lastUpdateTime) throws ConnectionException{
		DescribeMetadataResult d = MetadataLoginUtil.metadataConnection.describeMetadata(API_VERSION);

		Map<String,String> resultMap = new LinkedHashMap<String,String>();

		for (DescribeMetadataObject mObj : d.getMetadataObjects()){
			if(mObj.getXmlName().equals(type)){
				ListMetadataQuery query = new ListMetadataQuery();
				query.setType(type);
				query.setFolder(folder);
				FileProperties[] filePro = MetadataLoginUtil.metadataConnection.listMetadata(new ListMetadataQuery[] {query}, API_VERSION);
				for (FileProperties fPro : filePro) {
					resultMap.put(type+"."+fPro.getFullName(), fPro.getLastModifiedDate().getTimeInMillis()>lastUpdateTime?"TRUE":"FALSE");
				}
				if(mObj.getChildXmlNames().length>0){
					for(String child : mObj.getChildXmlNames()){			
						ListMetadataQuery queryChild = new ListMetadataQuery();
						if(child.equalsIgnoreCase("WorkflowFlowAction")){
							continue;
						}
						queryChild.setType(child);
						FileProperties[] fileProChild = MetadataLoginUtil.metadataConnection.listMetadata(new ListMetadataQuery[] {queryChild}, API_VERSION);
						for (FileProperties fPro : fileProChild) {
							resultMap.put(child+"."+fPro.getFullName(), fPro.getLastModifiedDate().getTimeInMillis()>lastUpdateTime?"TRUE":"FALSE");
						}
					}
				}
			}
		}
		return resultMap;
	}
	
	/**
	 * Method to check object whether be changed between two export
	 * 
	 * @param type
	 * @param default null folder
	 * @param lastUpdateTime
	 * @return checkResultMap<checkFieldName,checkResult>
	 * @throws ConnectionException
	 */
	public Map<String,String> getComparedResult(String type,Long lastUpdateTime) throws ConnectionException{
		return this.getComparedResult(type, "", lastUpdateTime);
	}
	
	/**
	 * Method to check object whether be changed between two export
	 * 
	 * @param type
	 * @param folder(multiple data)
	 * @param lastUpdateTime
	 * @return checkResultMap<checkFieldName,checkResult>
	 * @throws ConnectionException
	 */
	//LYU ADDED
	//for Document,email template, report, Dashboard 
	//Use example ====== Map yourMap= getFolderChilldComparedResult("Report",folderList)
	//無効なメソッド
	public Map<String,String> getFolderChilldComparedResult(String type,List<String> folderList,Long lastUpdateTime) throws ConnectionException{		
		DescribeMetadataResult d = MetadataLoginUtil.metadataConnection.describeMetadata(API_VERSION);		
		Map<String,String> resultMap = new LinkedHashMap<String,String>();
		for(String s:folderList){
			for (DescribeMetadataObject mObj : d.getMetadataObjects()){
				if(mObj.getXmlName().equals(type)){
					ListMetadataQuery query = new ListMetadataQuery();
					query.setType(s);
					query.setType(type);
					FileProperties[] filePro = MetadataLoginUtil.metadataConnection.listMetadata(new ListMetadataQuery[] {query}, API_VERSION);
					for (FileProperties fPro : filePro) {
						resultMap.put(type+"."+fPro.getFullName(), fPro.getLastModifiedDate().getTimeInMillis()>lastUpdateTime?"TRUE":"FALSE");
					}
				}
			}			
		}		
		return resultMap;
	}//LYU END
	
	/**
	 * get gmt time(calendar) to local time
	 * 
	 * @param Calendar cal
	 * @return localtime
	 */
	//日時データのフォーマット
	public String getLocalTime(Calendar cal){
		DateFormat format=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String time = format.format(cal.getTime());
		return time;
	}
	
	
	/**
	 * get utc time(string) to local time
	 * 
	 * @param Calendar cal
	 * @return localtime
	 */
	 public String getLocalTime(String utcTime) {		 
		 SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd\'T\'HH:mm:ss");
		 TimeZone tz = TimeZone.getTimeZone("GMT");
		 dateFormat.setTimeZone(tz);
		 Date utcDate =new Date();
		 try{
			 utcDate = dateFormat.parse(utcTime);
		 }catch(ParseException e){
			 e.getStackTrace();
		 }		 		  
		 GregorianCalendar cal = new GregorianCalendar();
		 cal.setTime(utcDate);
		 DateFormat format=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		 String time = format.format(cal.getTime());
		 return time;
	 }	
	 
	 /**
	 * get filterItem and make it to string
	 * 
	 * @param String objectName 
	 * @param FilterItem[] fis 
	 * 				fieldName
	 * @return filterItemString
	 */
	 //filterItemの内容をやすして戻る。
    public String getFilterItem(String objectName,FilterItem[] fis){
    	String string="";
    	if(fis!=null){
    		for(int i=0;i<fis.length;i++){
    			FilterItem fi= fis[i];
    			string+=(i+1)+". ";
    			String temName = "";
    			if(fi.getField().contains(".")){
    				String fiStr=fi.getField();
					if(fiStr.contains("$")){
				        if(fiStr.contains("$Source")){
				        	fiStr=fiStr.replace("Source", objectName);
				        }
						fiStr=fiStr.replace("$", "");
					}
					temName=getLabelforAll(fiStr);
    				//temName=getLabelforAll(fi.getField());
    			}else{
    				temName=getLabelforAll(objectName+"."+fi.getField());
    			}
    			string+=temName+" "+Util.getTranslate("FilterOperation",fi.getOperation().toString())+" ";
    			if(fi.getValueField()!=null){
    				if(fi.getValueField().contains(".")){
    					String fiStr=fi.getValueField();
    					if(fiStr.contains("$")){
    						if(fiStr.contains("$Source")){
    							fiStr=fiStr.replace("Source", objectName);
    				        }
    						fiStr=fiStr.replace("$", "");
    					}
    					string+=getLabelforAll(fiStr);
    					//string+=getLabelforAll(fi.getValueField());
    				}else{
    					string+=getLabelforAll(objectName+"."+fi.getValueField());
    				}
    			}else{
    				string+=fi.getValue();
    			}	
    			string+="\n";
    		}
    	}
    	return string;
    }
    
    public String getLabelforAll(String api){
    	String label="";
    	api=api.replace("$", ".");
    	if(Util.getTranslate("STANDARDOBJECTFIELD", api).equals(api)){
	    	if(api.contains(".")){
	    		String kk=api.replace(".", ",");
	    		String[] apiSpl= kk.split(",",0);
	    		if(UtilConnectionInfc.getApiToLabel().get(apiSpl[0].toUpperCase())!=null){
	    			label=UtilConnectionInfc.getApiToLabel().get(apiSpl[0].toUpperCase())+".";
	    		}else{
	    			label=apiSpl[0]+".";
	    		}
	    		if(UtilConnectionInfc.getApiToLabel().get(api.toUpperCase())!=null){
	        		label+=UtilConnectionInfc.getApiToLabel().get(api.toUpperCase());
	        	}else{
	    			label+=apiSpl[1];
	        	}
	    	}else{
	    		if(UtilConnectionInfc.getApiToLabel().get(api.toUpperCase())!=null){
	        		label=UtilConnectionInfc.getApiToLabel().get(api.toUpperCase());
	        	}else{
	    			label=api;
	        	}
	    	}
    	}else{
    		if(api.contains(".")){
	    		String kk=api.replace(".", ",");
	    		String[] apiSpl= kk.split(",",0);
	    		if(UtilConnectionInfc.getApiToLabel().get(apiSpl[0].toUpperCase())!=null){
	    			label=UtilConnectionInfc.getApiToLabel().get(apiSpl[0].toUpperCase())+".";
	    		}else{
	    			label=apiSpl[0]+".";
	    		}
    		}
    		label+=Util.getTranslate("STANDARDOBJECTFIELD", api);
    	}
    	return label;
    }
    
    /**
	 * get api label which saved in UtilConnectionInfc.ApiToLabel
	 * 
	 * @param String api
	 * @return labelString
	 */
    //APIよりLabel値を戻る
    public String getLabelApi(String api){
    	String label="";
    	if(UtilConnectionInfc.getApiToLabel().get(api.toUpperCase())!=null){
    		label=UtilConnectionInfc.getApiToLabel().get(api.toUpperCase())+"("+api+")";
    	}else{
			label=api+"("+api+")";
    	}
    	return label;
    }
    
    /**
	 * get sharedTo default format data translation
	 * 
	 * @param SharedTo st
	 * @return translated string
	 */
    //共有情報の訳す
	public String getSharedTo(SharedTo st){
		String string="";
		if(st!=null){
			if(st.getAllCustomerPortalUsers()!=null){
				string+=Util.getTranslate("SHAREDTO","ALLCUSTOMERPORTALUSERS")+"\n";			
			}
			if(st.getAllInternalUsers()!=null){
				string+=Util.getTranslate("SHAREDTO","ALLINTERNALUSERS")+"\n";			
			}
			if(st.getAllPartnerUsers()!=null){
				string+=Util.getTranslate("SHAREDTO","ALLPARTNERUSERS")+"\n";				
			}
			if(st.getGroup()!=null){
				for(String s:st.getGroup()){
					string+=Util.getTranslate("SHAREDTO","Group")+":"+s+"\n";
				}
			}
			if(st.getGroups()!=null){
				for(String s:st.getGroups()){
					string+=Util.getTranslate("SHAREDTO","Groups")+":"+s+"\n";
				}
			}
			if(st.getManagerSubordinates()!=null){
				for(String s:st.getManagerSubordinates()){
					string+=Util.getTranslate("SHAREDTO","ManagerSubordinates")+":"+s+"\n";
				}
			}
			if(st.getManagers()!=null){
				for(String s:st.getManagers()){
					string+=Util.getTranslate("SHAREDTO","Manager")+":"+s+"\n";
				}
			}
			if(st.getPortalRole()!=null){
				for(String s:st.getPortalRole()){
					string+=Util.getTranslate("SHAREDTO","PortalRole")+":"+s+"\n";
				}
			}
			if(st.getPortalRoleAndSubordinates()!=null){
				for(String s:st.getPortalRoleAndSubordinates()){
					string+=Util.getTranslate("SHAREDTO","PortalRoleAndSubordinates")+":"+s+"\n";
				}
			}	
			if(st.getRole()!=null){
				for(String s:st.getRole()){
					string+=Util.getTranslate("SHAREDTO","Role")+":"+s+"\n";
				}
			}
			if(st.getRoles()!=null){
				for(String s:st.getRoles()){
					string+=Util.getTranslate("SHAREDTO","Roles")+":"+s+"\n";
				}
			}
			if(st.getRoleAndSubordinates()!=null){
				for(String s:st.getRoleAndSubordinates()){
					string+=Util.getTranslate("SHAREDTO","RoleAndSubordinates")+":"+s+"\n";
				}
			}
			if(st.getRolesAndSubordinates()!=null){
				for(String s:st.getRolesAndSubordinates()){
					string+=Util.getTranslate("SHAREDTO","RolesAndSubordinates")+":"+s+"\n";
				}
			}
			if(st.getRoleAndSubordinatesInternal()!=null){
				for(String s:st.getRoleAndSubordinatesInternal()){
					string+=Util.getTranslate("SHAREDTO","RoleAndSubordinatesInternal")+":"+s+"\n";
				}
			}
			if(st.getTerritory()!=null){
				for(String s:st.getTerritory()){
					string+=Util.getTranslate("SHAREDTO","Territory")+":"+s+"\n";
				}
			}
			if(st.getTerritories()!=null){
				for(String s:st.getTerritories()){
					string+=Util.getTranslate("SHAREDTO","Territories")+":"+s+"\n";
				}
			}
			if(st.getTerritoryAndSubordinates()!=null){
				for(String s:st.getTerritoryAndSubordinates()){
					string+=Util.getTranslate("SHAREDTO","TERRITORYANDSUBORDINATES")+":"+s+"\n";
				}
			}
			if(st.getTerritoriesAndSubordinates()!=null){
				for(String s:st.getTerritoriesAndSubordinates()){
					string+=Util.getTranslate("SHAREDTO","TerritoriesAndSubordinates")+":"+s+"\n";
				}
			}
			if(st.getQueue()!=null){
				for(String s:st.getQueue()){
					string+=Util.getTranslate("SHAREDTO","Queue")+":"+s+"\n";
				}
			}
		}	
		return string;
	}

	/**
	 * Check User whether has the License to use this app
	 * 
	 * @param String path
	 *            path for the license file
	 * @return license info.
	 */
	//ライセンス情報のチェック
	public static String checkLicense(String path){
		Util.logger.info("Begin licence file check.");
		String license = null;
		try {
			X509EncodedKeySpec keySpec;
			if(WSC.isBatch){
				keySpec = new X509EncodedKeySpec(read(".././conf/publickey.txt"));
			}
			else{
				keySpec = new X509EncodedKeySpec(read("conf/publickey.txt"));
			}

			KeyFactory factory = KeyFactory.getInstance("RSA");
			PublicKey publicKey = factory.generatePublic(keySpec);
			Cipher cipher = Cipher.getInstance("RSA");
			cipher.init(Cipher.DECRYPT_MODE,publicKey);
			byte[] b = read(path);
			license = new String(cipher.doFinal(b));
		} catch (Exception e) {
			Util.logger.fatal("Error of licence file check.");
			Util.logger.fatal(e.getStackTrace());
		}
		Util.logger.info("End of licence file check.");
		return license;
	}
	
	/**
	 * get license file info from file
	 * 
	 * @param String path
	 *            path for the license file
	 * @return license info.(byte[])
	 */
	//Read License file "conf/publickey.txt"
	//publickey.txtの読み込み
	public static byte[] read(String path){
		File file = new File(path);
		if(file.exists()){
			InputStream in = null;
			try {
				byte[] bytes = new byte[(int)file.length()];
				in = new FileInputStream(file);
				in.read(bytes);				
				return bytes;
				
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}finally{
				//add by vp cheng 15-9-8 start
				try{
					in.close();
				}catch(IOException e){
					e.printStackTrace();
				}
				//add by vp cheng 15-9-8 end
			}
		}
		
		return null;
	
	}

	/**
	 * write license file info into file
	 * 
	 * @param String path
	 *            path for the license file
	 * @param byte[] bytes
	 *            license info.
	 * @return 
	 */
	//	"conf/publickey.txt"
	//無効なメソッド
	public static void write(byte[] bytes,String path){
		File file = new File(path);
		try {
			OutputStream out = new FileOutputStream(file);
			out.write(bytes, 0, bytes.length);
			out.flush();
			out.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		//checkLicense("conf/pass.txt");
	}

	//http://itmemo.net-luck.com/java-aes/
    public static final String ENCRYPT_KEY = "valueplusabcdefg";
	public static final String ENCRYPT_IV = "speclogger123456";


	/**
	 * 暗号化メソッド
	 *
	 * @param text 暗号化する文字列
	 * @return 暗号化文字列
	 */
	//パスワードを暗号化する
	public static  String encrypt(String text) {
		// 変数初期化
		String strResult = null;
 
		try {
			// 文字列をバイト配列へ変換
			byte[] byteText = text.getBytes("UTF-8");
 
			// 暗号化キーと初期化ベクトルをバイト配列へ変換
			byte[] byteKey = ENCRYPT_KEY.getBytes("UTF-8");
			byte[] byteIv = ENCRYPT_IV.getBytes("UTF-8");
 
			// 暗号化キーと初期化ベクトルのオブジェクト生成
			SecretKeySpec key = new SecretKeySpec(byteKey, "AES");
			IvParameterSpec iv = new IvParameterSpec(byteIv);
 
			// Cipherオブジェクト生成
			Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
 
			// Cipherオブジェクトの初期化
			cipher.init(Cipher.ENCRYPT_MODE, key, iv);
 
			// 暗号化の結果格納
			byte[] byteResult = cipher.doFinal(byteText);
 
			// Base64へエンコード
			strResult = Base64.encodeBase64String(byteResult);
 
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (NoSuchAlgorithmException e) {
			e.printStackTrace();
		} catch (NoSuchPaddingException e) {
			e.printStackTrace();
		} catch (InvalidKeyException e) {
			e.printStackTrace();
		} catch (IllegalBlockSizeException e) {
			e.printStackTrace();
		} catch (BadPaddingException e) {
			e.printStackTrace();
		} catch (InvalidAlgorithmParameterException e) {
			e.printStackTrace();
		}
 
		// 暗号化文字列を返却
		return strResult;
	}
 
	/**
	 * 復号化メソッド
	 *
	 * @param text 復号化する文字列
	 * @return 復号化文字列
	 */
	//パスワードを復号化する
	public static  String decrypt(String text) {
		// 変数初期化
		String strResult = null;
 
		try {
			// Base64をデコード
			byte[] byteText = Base64.decodeBase64(text);
 
			// 暗号化キーと初期化ベクトルをバイト配列へ変換
			byte[] byteKey = ENCRYPT_KEY.getBytes("UTF-8");
			byte[] byteIv = ENCRYPT_IV.getBytes("UTF-8");
 
			// 復号化キーと初期化ベクトルのオブジェクト生成
			SecretKeySpec key = new SecretKeySpec(byteKey, "AES");
			IvParameterSpec iv = new IvParameterSpec(byteIv);
 
			// Cipherオブジェクト生成
			Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
 
			// Cipherオブジェクトの初期化
			cipher.init(Cipher.DECRYPT_MODE, key, iv);
 
			// 復号化の結果格納
			byte[] byteResult = cipher.doFinal(byteText);
 
			// バイト配列を文字列へ変換
			strResult = new String(byteResult, "UTF-8");
 
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (NoSuchAlgorithmException e) {
			e.printStackTrace();
		} catch (NoSuchPaddingException e) {
			e.printStackTrace();
		} catch (InvalidKeyException e) {
			e.printStackTrace();
		} catch (IllegalBlockSizeException e) {
			e.printStackTrace();
		} catch (BadPaddingException e) {
			e.printStackTrace();
		} catch (InvalidAlgorithmParameterException e) {
			e.printStackTrace();
		}
 
		// 復号化文字列を返却
		return strResult;
	}
	
	/**
	 * Map for objects match templates
	 *
	 */
	public static Map<String, String> MetaDataTypeMap=new HashMap<String, String>() {
	    {
	        put("StandardObject","CustomObject");
	        put("CustomSetting", "CustomSetting");
	        put("ArticleType", "CustomObject");	    
	        put("EXTERNALOBJECT", "CustomObject");	
	    }
	};	
	
	/**
	 * get the name of which template to use
	 *
	 */
	//テンプレートの名前を戻る
	public static String getTemplateName(String originType){
		String retType=MetaDataTypeMap.get(originType);
		if(retType==null){
			retType=originType;
		}
		return retType;		
	}
	
	/**
	 * To deal with special Japanese char 
	 * http://www.utf8-chartable.de/unicode-utf8-table.pl?start=65280&number=128
	 *
	 */
	static String[][] UTF8MAP = { 
		{"%EF%BC%88", "（"},       
        {"%EF%BC%89", "）"}};
	//UTF-8の文字を戻る
	public static String translateSpecialChar(String originStr){
		String retStr=originStr.replaceAll(" ","_");
		try{
			retStr=URLDecoder.decode(originStr, "UTF-8");
		}catch(UnsupportedEncodingException e){
			
		}
		//for(int i=0;i<UTF8MAP.length;i++ ){
//
	//		retStr=retStr.replaceAll(UTF8MAP[i][0], UTF8MAP[i][1]);
		//}
		return retStr;		
	}	
	
	/**
	 * Make sure hyperlink name start with a letter or underscore and not contain spaces
	 * And to be unique
	 * for nameValue
	 */
	//hyperlinkの特殊文字処理
	public static int nameSequence=0;
	public static String makeNameValue(String originStr){
		String retStr=originStr;
		//retStr=retStr.replaceAll( "[¥$&-+=ー*\'\"(){}\\[\\]|%~@:;<>?^ 　（）/【】·・：]", "_");
		retStr=retStr.replaceAll("[`~!@#$%^&*()+=|{}':;',\\-\\[\\]<>/?~！@#￥%……&*（）——+|{}【】‘；：”“’。，、？]","_"); 
		retStr=retStr.replaceAll("\\s", "_");
		retStr=retStr.replaceAll( "\\Q"+"\\"+"\\E", "_");	
		retStr=retStr.replace("·", "_");
		retStr=retStr.replace("・", "_");
		retStr=retStr.replace(" ", "_");
		retStr=retStr.replace("　", "_");
		nameSequence++;
		retStr="L"+String.format("%04d",nameSequence)+retStr;
		if(retStr.length()>255){
			retStr.substring(0, 255);
		}	
		return retStr;		
	}
	
	/**
	 * Make sure hyperlink name start with a letter or underscore and not contain spaces
	 * And to be unique
	 * for SheetName
	 */
	public static int sheetSequence=0;	
	public static String makeSheetName(String originStr){
		String retStr=originStr;
		//retStr=retStr.replaceAll( "[¥$&-+=ー*\'\"{}\\[\\]|%~@:;<>?^ 　/()（）【】·・：]", "_");
		retStr=retStr.replace("＄", "_");
		retStr=retStr.replaceAll("[`~!@#$%^&*()+=|{}':;',\\-\\[\\]<>/?~！@#￥%……&*（）——+|{}【】‘；：”“’。，、？]","_");  
		retStr=retStr.replaceAll("\\s", "_");		
		retStr=retStr.replaceAll( "\\Q"+"\\"+"\\E", "_");
		retStr=retStr.replace("·", "_");
		retStr=retStr.replace("・", "_");
		retStr=retStr.replace(" ", "_");
		retStr=retStr.replace("　", "_");
		sheetSequence++;
		retStr="S"+String.format("%04d",sheetSequence)+"_"+retStr;
		
		return retStr;		
	}
	public static String cutSheetName(String retStr){
		
		if(retStr.length()>30){
			retStr=retStr.substring(0,30);
		}
		return retStr;	
	}
	
	/**
	 * 
	 * String Filter
	 * for special byte
	 */
	//特殊文字処理
	public static String UrlFilter(String str){
		if(str!=null){
			str=str.replaceAll("%2E", ".");   
			str=str.replaceAll("%2F", ".");  
		}
		return str;
	}	
	
	/**
	 * 
	 * String Filter
	 * for special byte
	 */
	//特殊文字処理
	public static String stringFilter(String str){
		String regEx="[`~!@#$%^&*()+=|{}':;',\\-\\[\\]<>/?~！@#￥%……&*（）——+|{}【】‘；：”“’。，、？]";  
		Pattern pat = Pattern.compile(regEx);     
		Matcher mat = pat.matcher(str.replaceAll(" ", "_"));     
		String ret = mat.replaceAll("").trim();  
		return ret;
	}	
	
	/**
	 * 
	 * NULL Filter
	 * for String
	 */
	//nullをスペースに変換する
	public static String nullFilter(String obj){
		String ret = "";  
		if(obj != null){
			ret=String.valueOf(obj);
		}
		return ret;
	}
	
	/**
	 * 
	 * NULL Filter
	 * for object 
	 */
	//nullをスペースに変換する
	public static String nullFilter(Object obj){
		String ret = "";  
		if(obj != null){
			ret=String.valueOf(obj);
		}
		return ret;
	}
	//add by vp cheng 15-9-8 start
	//変更のマークの共通化処理してください
	public String getUpdateFlag(Map<String, String> resultMap,String getname){
		String objUpdateFlag=resultMap.get(getname);
		String result="";
		if(objUpdateFlag!=null){
			result=Util.getTranslate("IsChanged",Util.nullFilter(objUpdateFlag));
		}else{
			result=Util.getTranslate("IsChanged",Util.nullFilter("NONE"));
		}
		return result;		
	}
	//add by vp cheng 15-9-8 end
	
	
	//create the excel file,if the file's size is lager than the maxSize,then create a new one to store.
	public Boolean createExcel (XSSFWorkbook workBook,CreateExcelTemplate excelTemplate,String type,Integer size,Integer lastIndex){
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		try {
			workBook.write(byteArrayOutputStream);
			if (byteArrayOutputStream.size() > excelTemplate.maxFileSize) {//判断是否溢出
				if(size == lastIndex && excelTemplate.wbNumber == 1){
					excelTemplate.exportExcel(type, "" + excelTemplate.wbNumber++);
				}else{
					excelTemplate.exportExcel(type, ".Part" + excelTemplate.wbNumber++);
					return true;
				}
			} else if (size == lastIndex) {//如果没有溢出，但是所有表格都已经读取完了
				if (excelTemplate.wbNumber == 1) {//判断是否之前拆分过表格，如果没有，则不生成part
					excelTemplate.exportExcel(type, "");
				} else {
					excelTemplate.exportExcel(type, ".Part" + excelTemplate.wbNumber);//如果拆分过，则按wbNumber生成part
				}
				return false;
			}
			//如果没有溢出而且CustObject还未读取完，则继续读取下一个CustObject
		} catch (Exception e) {
			logger.info("Export ErrorL:"+e.getMessage());
			throw new RuntimeException(e.getMessage());
		}
		return false;
	}
}
