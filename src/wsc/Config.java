package wsc;

import java.text.SimpleDateFormat;
import java.util.List;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.Map;
import java.util.Properties;


import source.Util;
import source.UtilConnectionInfc;



public class Config {
	private String username;
	private String password;
	private String loginURL;
    private String proxyHost;
    private String proxyPort;
    private String proxyUsername;
    private String proxyPassword;	
	private MetadataLoginUtil login;
	private UtilConnectionInfc uti;
	private static Util ut;
    private static WSC window;
    public static Properties properties;
	public static void main(String[] args){
		window.isBatch=Boolean.valueOf(args[0]);
		ut = new Util();
		window = new WSC();
		Config c = new Config();
		c.getConfig();
	}	
	//config.propertiesの情報を読み込む
	public void getConfig(){
		try {
			File file = new File(".././conf/config.properties");			
			InputStream inputStream =  new FileInputStream(file);
		    properties = new Properties();
			properties.load(inputStream);
			username=properties.getProperty("sfdc.username");
			password=Util.decrypt(properties.getProperty("sfdc.password"));
			loginURL=properties.getProperty("sfdc.endpoint");
			proxyHost=properties.getProperty("sfdc.proxyHost");
			proxyPort=properties.getProperty("sfdc.proxyPort");
			proxyUsername=properties.getProperty("sfdc.proxyUsername");
			proxyPassword=Util.decrypt(properties.getProperty("sfdc.proxyPassword"));
			//Application Properties
			ut = new Util();
			ut.LoadProperties(properties.getProperty("Language"));
			
			if(loginURL.contains("login")){
				login = new MetadataLoginUtil(username,password.toCharArray(),true,proxyHost,proxyPort,proxyUsername,proxyPassword);
			}else{
				login = new MetadataLoginUtil(username,password.toCharArray(),false,proxyHost,proxyPort,proxyUsername,proxyPassword);
			}
			try{
				
			    uti = new UtilConnectionInfc();
				uti.setDownloadPath(properties.getProperty("process.outputSuccess"));
				uti.setLanguage(properties.getProperty("Language"));
				uti.setMetadataConnection(login.userLogin());
				uti.setConnection(login.Connection);
				uti.setModifiedFlag(Boolean.valueOf(properties.getProperty("setting.modifiedFlag")));
				SelectMetadata sm=new SelectMetadata(login.userLogin(),uti,window);
				//uti.setLastUpdateTime(lastUpdateTime);
				Map<String,List<String>> selectedMap= sm.getCustomSetting();
				
				//to do output
				//sm.getMetadataList(selectedMap);
			    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				sm.lastRunTime = df.format(new Date());				
				WriteXML wx = new WriteXML(selectedMap,Util.API_VERSION,sm.lastRunTime);
				sm.exportToExcel(selectedMap);
			}catch(Exception ce){								
				ce.printStackTrace();
			}	
		} catch (IOException e) {
			e.printStackTrace();
		}		
	}
	

}
