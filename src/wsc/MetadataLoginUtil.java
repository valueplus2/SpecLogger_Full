package wsc;
import java.net.Authenticator;

import java.net.PasswordAuthentication;


import java.util.Properties;
import javax.xml.namespace.QName;

import javax.swing.JOptionPane;

import source.Util;

import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.partner.PartnerConnection;
import com.sforce.soap.tooling.ToolingConnection;
import com.sforce.ws.ConnectionException;
import com.sforce.ws.ConnectorConfig;

/**
 * Login utility.
 */
public class MetadataLoginUtil {
    private String USERNAME;
    private String PASSWORD;
    public String loginUrl;
    private String proxyHost;
    private String proxyPort;
    private String proxyUsername;
    private String proxyPassword;

    public static MetadataConnection metadataConnection;
    public static PartnerConnection Connection;
    public static ToolingConnection toolingConnection;
    //Salesforce情報伝達
    public MetadataLoginUtil(String username,char[] password,boolean isPro,String proxyHost, String proxyPort, String proxyUsername, String proxyPassword){
    	this.USERNAME=username;
    	this.PASSWORD=String.valueOf(password);
    	loginUrl="https://login.salesforce.com/services/Soap/u/"+Util.API_VERSION;
    	//is production or sandbox
    	if(!isPro){
    		loginUrl="https://test.salesforce.com/services/Soap/u/"+Util.API_VERSION;
    	}
    	if(proxyHost==null||proxyHost.isEmpty()){
    		this.proxyHost=null;
    	}else{
    		this.proxyHost = proxyHost;
    	}
    	if(proxyPort==null||proxyPort.isEmpty()){
    		this.proxyPort=null;
    	}else{
    		this.proxyPort = proxyPort;
    	}
    	if(proxyUsername==null||proxyUsername.isEmpty()){
    		this.proxyUsername=null;
    	}else{
    		this.proxyUsername = proxyUsername;
    	}
    	if(proxyPassword==null||proxyPassword.isEmpty()){
    		this.proxyPassword=null;
    	}else{
    		this.proxyPassword = proxyPassword;
    	}    
    }
    //User Login and Salesforceへログイン情報の認証
    public MetadataConnection userLogin() throws ConnectionException {
	    ConnectorConfig partnerConfig = new ConnectorConfig();
	    ConnectorConfig metadataConfig = new ConnectorConfig();
	    partnerConfig.setUsername(USERNAME);
	    partnerConfig.setPassword(PASSWORD);
	    partnerConfig.setAuthEndpoint(loginUrl);
	    partnerConfig.setServiceEndpoint(loginUrl);
	    
	    partnerConfig.setManualLogin(true);
	    Properties systemSettings = System.getProperties();
	    if(proxyHost!=null&&proxyHost.length() > 0&&proxyPort!=null){
        	 partnerConfig.setProxy(proxyHost, Integer.valueOf(proxyPort)); 
	         systemSettings.put("proxySet", "true");
	         systemSettings.put("http.proxyHost", proxyHost);
	         systemSettings.put("http.proxyPort", proxyPort);
        }
        // Set the username and password if your proxy must be authenticated
        if (proxyUsername != null && proxyUsername.length() > 0) {
        	partnerConfig.setProxyUsername(proxyUsername);
        	final String authUser = proxyUsername;
            String tempauthPassword = "";
            if (proxyPassword != null && proxyPassword.length() > 0) {
            	partnerConfig.setProxyPassword(proxyPassword);
            	tempauthPassword=proxyPassword;
            } else {
            	partnerConfig.setProxyPassword("");
            }
            final String authPassword = tempauthPassword;
            
            Authenticator.setDefault(
               new Authenticator() {
                  public PasswordAuthentication getPasswordAuthentication() {
            	   return new PasswordAuthentication(
                           authUser, authPassword.toCharArray());
                  }
               }
            ); 
            System.setProperty("http.proxyUser", authUser);
            System.setProperty("http.proxyPassword", authPassword);        
        }     
        partnerConfig.setConnectionTimeout(60 * 1000);
        partnerConfig.setReadTimeout((540 * 1000));
      //  try{
			PartnerConnection pc = com.sforce.soap.partner.Connector.newConnection(partnerConfig);
			com.sforce.soap.partner.LoginResult lr = pc.login(USERNAME, PASSWORD);
			// if password has expired, throw an exception
	        if (lr.getPasswordExpired()) { 
				JOptionPane.showMessageDialog(null, Util.getTranslate("Message","Password Expired"),"Password Expired", JOptionPane.ERROR_MESSAGE);										
	        }
			//YU start
			ConnectorConfig config = new ConnectorConfig();
	        config.setUsername(USERNAME);
	        config.setPassword(PASSWORD);
	        config.setAuthEndpoint(loginUrl);
	        config.setTraceMessage(true);
	        //config.setPrettyPrintXml(true);
	        config.setSessionRenewer(new SessionRenewer());
	        Connection = new PartnerConnection(config);
			//YU end
			metadataConfig.setSessionId(lr.getSessionId());
			metadataConfig.setServiceEndpoint(lr.getMetadataServerUrl());
		    metadataConnection = new MetadataConnection(metadataConfig);
		    
		    //toolings
		    ConnectorConfig toolingConfig = new ConnectorConfig();
		    toolingConfig.setSessionId(lr.getSessionId());
		    toolingConfig.setServiceEndpoint(lr.getServerUrl().replace('u', 'T'));
		    toolingConnection = com.sforce.soap.tooling.Connector.newConnection(toolingConfig);
		    
       // }catch(Exception e){
        //	JOptionPane.showMessageDialog(null, Util.getTranslate("Message","Timeout"),"Timeout", JOptionPane.ERROR_MESSAGE);										
        
       // }
	    //Util.getApiToLabel();
		Util.getUserInfoFromServer();
	    Util.getApiToLabelFromServer();
	    Util.getEMailTemplateFromServer();
	    //Util.getTest();
	    return metadataConnection;
    }
    
    public class SessionRenewer implements com.sforce.ws.SessionRenewer {
        @Override
        public SessionRenewalHeader renewSession(ConnectorConfig config) throws ConnectionException {
        	metadataConnection=userLogin();
            SessionRenewalHeader header = new SessionRenewalHeader();
            header.name = new QName("urn:enterprise.soap.sforce.com", "SessionHeader");
            header.headerElement = metadataConnection.getSessionHeader();
            return header;
        }
    }
    //Logged out
    public static void userLogout() {
    	   try {
    		  Connection.logout();
    	      System.out.println("Logged out.");
    	   } catch (ConnectionException ce) {
    	      ce.printStackTrace();
    	   }
    	}
}