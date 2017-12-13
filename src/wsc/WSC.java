package wsc;

import java.awt.*;

import javax.swing.*;


import com.sforce.soap.metadata.DescribeMetadataResult;

import com.sforce.ws.ConnectionException;
//import com.sforce.ws.ConnectionException;



import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;


import java.io.File;
import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.BufferedInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import java.util.Date;
import java.util.HashMap;

import java.util.Map;
import java.util.Properties;
import java.util.Vector;
import java.util.concurrent.ExecutionException;


import source.Util;
import source.UtilConnectionInfc;

public class WSC implements ActionListener {

	public JFrame frame;

	JTextField usernameField;
	JPasswordField passwordField;	
	JLabel errorLabel;
    JRadioButton rdbtnNewRadioButton;
    JRadioButton rdbtnNewRadioButton_1;
	public JButton btnLogIn;
	static WSC window;
	static MyOwnFocusTraversalPolicy newPolicy;
	public JFrame settingFrame;
	private JTextField hostField;
	private JTextField portField;
	private JTextField proxyUsername;
	private JPasswordField proxyPassword;	
	private JTextField localField;
	private JComboBox cb;
    public static Properties propertiesout;
	private Map<String, String> languageMap;
	public JDialog versionFrame;
	public static final int FRAME_HEIGHT = 600;
	public static final int FRAME_WIDTH = 400;
	
	private JLabel lblUsername;
	private JLabel lblPassword;
	private JLabel licenseDate;
	private JMenu menuSetting;
	private JMenuItem mItemSet;
	private JMenu menuHelp;
	private JMenuItem mAbout;
	//private JMenu menuExit;
	private JMenuItem mExit;
	private MetadataLoginUtil login;
	private String language;
	private JLabel lblLicense;
	private JTextField licenseField;
	private JButton btnLicense;
	private String errInfo;
	private JCheckBox settingSubfolder;
	private JCheckBox cbModified;
	private String licenseExpireDate;
	//private JButton btnSelectMetadata;
	private UtilConnectionInfc uti;
	JLabel lblProcess;
	String key;
	boolean isValid = true;
	boolean isLicensePath = true;
	boolean isErrorLicense = true;
	boolean isSuccess = true;
	boolean isFail = true;
	public static boolean isBatch=false;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		if(args.length!=0){
			isBatch=Boolean.valueOf(args[0]);
		}
		Util.logger.debug("isBatch: "+isBatch);
		Date now = new Date(); 
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
		String systemTime = dateFormat.format(now)+".log";
		File file;
		String c;
		
		//=new File("../ExportMetaData_NoPath/logs/normal_app.log"); 
		if(isBatch){
			file=new File(".././logs/normal_app.log");
			String d=file.getParent();
			c=d+File.separator+"logs";
		}else{
			file=new File("logs/normal_app.log");
			c=file.getParent();
		}	
		Util.logger.debug("exists: "+file.exists());
		Util.logger.debug("getPath: "+file.getPath());
		File mm=new File(c+File.separator+systemTime);
        Util.logger.debug("mm: "+mm);
	    if(file.renameTo(mm))  
	    {   
	    	Util.logger.debug("rename log file success!");   
	    }else   
	    {   
	    	Util.logger.debug("rename log file failure!");  
	    }  
		EventQueue.invokeLater(new Runnable() {
			
			public void run() {
				try {
					Util.logger.info("SpecLogger initialization Start.");
					window = new WSC();
					window.frame.setVisible(true);

				} catch (Exception e) {
					Util.logger.error("SpecLogger initialization failure!");
					Util.logger.fatal(e.getStackTrace());
				}
			}
		});
		
	}
	 
	
	/**
	 * Create the application.
	 */
	public WSC() {
		Util.logger.info("WSC initialization Start.");
		loginbefore();
		initLanguage();
		SettingWin(null);
		initialize();
		Util.logger.info("WSC initialization complete");
	}
	public void loginbefore(){
		File f;
		if(isBatch){
			f=new File(".././conf/config.properties");
		}else{
			f=new File("conf/config.properties");
		}		
	
		Properties props = new Properties();
        try {
    		if (!f.exists() || !f.isFile()) {
    			f.createNewFile();
    		}	    		
	         InputStream in = new BufferedInputStream(new FileInputStream(f));
	         props.load(in);
	         if(props.getProperty("sfdc.username")==null){
	        	 props.put("sfdc.username", "");
	         }
	         if(props.getProperty("Language")==null){
	        	 props.put("Language", "JP");
	         }
	         if(props.getProperty("sfdc.proxyHost")==null){
	        	 props.put("sfdc.proxyHost", "");
	         }
	         if(props.getProperty("LicensePath")==null){
	        	 props.put("LicensePath", "");
	         }
	         if(props.getProperty("sfdc.password")==null){
	        	 props.put("sfdc.password", "");
	         }
	         if(props.getProperty("sfdc.endpoint")==null){
	        	 props.put("sfdc.endpoint", "");
	         }
	         if(props.getProperty("setting.modifiedFlag")==null){
	        	 props.put("setting.modifiedFlag", "false");
	         }
	         if(props.getProperty("sfdc.proxyPort")==null){
	        	 props.put("sfdc.proxyPort", "");
	         }
	         if(props.getProperty("sfdc.proxyUsername")==null){
	        	 props.put("sfdc.proxyUsername", "");
	         }
	         if(props.getProperty("sfdc.proxyPassword")==null){
	        	 props.put("sfdc.proxyPassword", "");
	         }	         
	         if(props.getProperty("process.outputSuccess")==null){
	        	 props.put("process.outputSuccess", "");
	         }
	         if(props.getProperty("setting.Subfolder")==null){
	        	 props.put("setting.Subfolder", "true");
	         }
	         if(!props.getProperty("sfdc.password").isEmpty()&&!props.getProperty("sfdc.username").isEmpty()){
		    	 String username = props.getProperty("sfdc.username");
		    	 String password =Util.decrypt(props.getProperty("sfdc.password"));
		    	 key=username+password;
	         }
	         if(!props.getProperty("LicensePath").isEmpty()){
					String licensePath = props.getProperty("LicensePath");
					String licenseStr=Util.checkLicense(licensePath);
					String[] licenseInfo;
					if(licenseStr!=null){
						 licenseInfo = licenseStr.split(",",0);
				
					    try {
					    	
					    	Date lDate = new SimpleDateFormat("yyyy/MM/dd").parse(licenseInfo[1]);
					    	licenseExpireDate=licenseInfo[1];
					    }catch(Exception ec){
					    	
					    	licenseExpireDate =Util.getTranslate("Setting","ErrorLicense");
					    }
					}
				
	         }
			if(isBatch){
				props.store(new	FileOutputStream(".././conf/config.properties"), "Config File");
			}else{
				props.store(new FileOutputStream("conf/config.properties"),"Config File");
			}	         
        } catch (Exception e) {
         e.printStackTrace();
        }
	}
	public void loginafter(Properties Prop){
		String username = Prop.getProperty("sfdc.username");
		String password =Util.decrypt(Prop.getProperty("sfdc.password"));
		String loginkey=username+password;
		if(loginkey.equals(key)){	
			Util.logger.info("The user does not have the change!");
		}
		else{
			 try {    
			   File f = new File("../ExportMetaData_NoPath/conf/package.xml");
			   f.delete();
			   }catch (Exception e){
				   Util.logger.error("Empty the failure");
			   };
		}
		
	}
	public void initLanguage() {
		Util.logger.info("InitLanguage Start.");
		getConfig();
		try {
			Util.LoadProperties(propertiesout.get("Language").toString());
		} catch (IOException e) {
			Util.logger.error("InitLanguage failure.");
			Util.logger.fatal(e.getStackTrace());
		}

		languageMap = new HashMap<String, String>();
		languageMap.put(Util.getTranslate("language", "EN"), "EN");
		languageMap.put(Util.getTranslate("language", "JP"), "JP");
		languageMap.put(Util.getTranslate("language", "CN"), "CN");
		Util.logger.info("InitLanguage complete.");
	}

	public void savePro(String loginUrl) {
		try {
			Util.logger.info("SavePro start.");
			Properties properties = new Properties();
			properties.put("Language", language);
			
			properties.put("sfdc.password",Util.encrypt(String.valueOf(passwordField.getPassword())));
			properties.put("sfdc.username", usernameField.getText());
			properties.put("sfdc.endpoint", loginUrl);					
			properties.put("process.outputSuccess", localField.getText());
			properties.put("sfdc.proxyHost", hostField.getText());
			properties.put("sfdc.proxyPort", portField.getText());
			properties.put("sfdc.proxyUsername", proxyUsername.getText());
			properties.put("sfdc.proxyPassword", Util.encrypt(String.valueOf(proxyPassword.getPassword())));
			
			properties.put("setting.Subfolder", String.valueOf(settingSubfolder.isSelected()));     //by  wangyuxin
			properties.put("setting.modifiedFlag", String.valueOf(cbModified.isSelected()));  
			properties.put("LicensePath", licenseField.getText());
			loginafter(properties);
			if(isBatch){
				properties.store(new FileOutputStream(".././conf/config.properties"), "Config File");
			}else{
				properties.store(new FileOutputStream("conf/config.properties"),"Config File");
			}
			Util.logger.info("SavePro complete.");
			propertiesout=properties;//save to app
			uti.setLanguage(language);
			uti.setSettingSubfolder(settingSubfolder.isSelected());
			uti.setModifiedFlag(cbModified.isSelected());	
			uti.setDownloadPath(localField.getText());
			//btnLogIn.setEnabled(true);
		} catch (Exception e) {
			Util.logger.error("SavePro failure.");
			Util.logger.fatal(e.getStackTrace());
		}
	}

	public void getConfig() {
		try {
			Util.logger.info("getConfig Start.");
			File file;
			if(isBatch){
			   file = new File(".././conf/config.properties");
			}else{
			   file = new File("conf/config.properties");	
			}
			InputStream inputStream = new FileInputStream(file);
			propertiesout = new Properties();
			propertiesout.load(inputStream);
			Util.logger.info("getConfig complete.");
		} catch (IOException e) {
			Util.logger.error("getConfig failure.");
			Util.logger.fatal(e.getStackTrace());
		}
	}

	/**
	 * use thread update UI
	 */
	private void changeLanguage() {
		Util.logger.info("ChangeLanguage start.");
		try{
			new Thread(new Runnable() {
				@Override
				public void run() {
					frame.setTitle(Util.getTranslate("Setting", "Title"));
					/*
					if(usernameField.isEnabled()){
						btnLogIn.setText(Util.getTranslate("Setting", "Login"));
					}else{
						btnLogIn.setText(Util.getTranslate("Setting", "Logout"));
					}*/
					rdbtnNewRadioButton.setText(Util.getTranslate("Setting","Production"));
					rdbtnNewRadioButton_1.setText(Util.getTranslate("Setting","Sandbox"));
					lblUsername.setText(Util.getTranslate("Setting", "Username"));
					lblPassword.setText(Util.getTranslate("Setting", "Password"));
					menuSetting.setText(Util.getTranslate("Setting", "Setting"));
					mItemSet.setText(Util.getTranslate("Setting", "Setting"));
					menuHelp.setText(Util.getTranslate("Setting", "Help"));
					mAbout.setText(Util.getTranslate("Setting", "About"));
					if(!isValid)
						errorLabel.setText(Util.getTranslate("Setting","OutputPathValid"));
					if(!isLicensePath)
						errorLabel.setText(Util.getTranslate("Setting", "NeedLicense"));
					if(!isErrorLicense)
						errorLabel.setText(Util.getTranslate("Setting", "ErrorLicense"));
					/*if(!isSuccess)
						errorLabel.setText(Util.getTranslate("Setting","Success"));
					if(!isFail)
						errorLabel.setText(Util.getTranslate("Setting","Fail"));*/
					//btnSelectMetadata.setText(Util.getTranslate("Setting", "SelectsMetadata"));
				}
			}).start();
			Util.logger.info("ChangeLanguage start.");
		}catch (Exception e) {
			Util.logger.error("ChangeLanguage failure.");
			Util.logger.error(e.getStackTrace());
			
		}

	}
	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		Util.logger.info("Frame initialization Start.");
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());// change the look
		
			languageMap = new HashMap<String, String>();
			//languageMap.put(Util.getTranslate("language", "EN"), "EN");
			languageMap.put(Util.getTranslate("language", "JP"), "JP");
			//languageMap.put(Util.getTranslate("language", "CN"), "CN");
			frame = new JFrame(Util.getTranslate("Setting", "Title"));
			//left top LOGO image
			frame.setIconImage(new ImageIcon("./common/icons/logo.png").getImage());
			frame.setBounds(60, 100, FRAME_HEIGHT, FRAME_WIDTH);
			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
			frame.setResizable(false);
			frame.getContentPane().setLayout(null);
			// Menubar
			JMenuBar mb = new JMenuBar();
			menuSetting = new JMenu(Util.getTranslate("Setting", "Setting"));
			mItemSet = new JMenuItem(Util.getTranslate("Setting", "Setting"));
			mExit = new JMenuItem("Exit");
			mExit.addActionListener(this);
			// mExit = new JMenuItem(Util.getTranslate("Setting", "About"));
			mItemSet.addActionListener(this);
			menuSetting.add(mItemSet);
			menuSetting.add(mExit);
			mb.add(menuSetting);
			menuHelp = new JMenu(Util.getTranslate("Setting", "Help"));
			mAbout = new JMenuItem(Util.getTranslate("Setting", "About"));
			mAbout.addActionListener(this);
			menuHelp.add(mAbout);
			mb.add(menuHelp);
			frame.setJMenuBar(mb);
	
			lblUsername = new JLabel(Util.getTranslate("Setting", "Username"));
			lblUsername.setBounds(53, 65, 65, 25);
			frame.getContentPane().add(lblUsername);
	
			lblPassword = new JLabel(Util.getTranslate("Setting", "Password"));
			lblPassword.setBounds(53, 95, 65, 25);
			frame.getContentPane().add(lblPassword);
	
			usernameField = new JTextField();
			usernameField.setText(propertiesout.getProperty("sfdc.username"));		
			usernameField.setBounds(130, 68, 407, 20);
			frame.getContentPane().add(usernameField);
			usernameField.setColumns(10);
	
			passwordField = new JPasswordField();
			if(propertiesout.getProperty("sfdc.password")!=""){
				passwordField.setText(Util.decrypt(propertiesout.getProperty("sfdc.password")));	
			}
			passwordField.setBounds(130, 98, 407, 20);
			frame.getContentPane().add(passwordField);
			
			JList list = new JList();
			list.setBounds(223, 64, 1, 1);
			frame.getContentPane().add(list);
	
			btnLogIn = new JButton(Util.getTranslate("Setting", "Login"));
			btnLogIn.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					if(Util.getTranslate("Setting","Login").equals(e.getActionCommand())){
						
						SwingWorker<Boolean, Void> sw = new SwingWorker<Boolean, Void>() {
							@Override
							protected Boolean doInBackground()// This is called when you
																// .execute() the
																// SwingWorker instance
							{// Runs on its own thread, thus not "freezing" the
								// interface
								setInputEnable(false);						
								lblProcess.setVisible(true);								
								String outputPath = propertiesout.getProperty("process.outputSuccess");
								isValid = checkOutputPathValid(outputPath);
								if(!isValid){
									errInfo = Util.getTranslate("Setting","OutputPathValid");
									setInputEnable(true);
									return false;
								}
								
								String licensePath = propertiesout.getProperty("LicensePath");
								Util.logger.info(licensePath);
								if(licensePath == null || licensePath.equals("")){
									isLicensePath = false;
									errInfo = Util.getTranslate("Setting", "ErrorLicense");
									setInputEnable(true);
									return false;
								}
								uti = new UtilConnectionInfc();
								String licenseStr=Util.checkLicense(licensePath);
								Date curDate = new Date();
							    Date licensedDate=curDate;
							    String[] licenseInfo;
							    
								if(licenseStr==null){
									
									JOptionPane.showMessageDialog(null, Util.getTranslate("Setting","ErrorLicense"),"ErrorLicense",JOptionPane.ERROR_MESSAGE);
									
									return false;	
							    }
							    	 licenseInfo = licenseStr.split(",",0);
							    	
								 try {
									   
									licensedDate = new SimpleDateFormat("yyyy/MM/dd").parse(licenseInfo[1]);
									 
								 }catch( Exception e){
									 
									errInfo = Util.getTranslate("Setting", "ErrorLicense");
										
								  }
									
							    Util.logger.info("-------------licensedDate="+licensedDate);
							   	if(curDate.after(licensedDate)){
							   		
							   		JOptionPane.showMessageDialog(null, Util.getTranslate("Setting","ErrorLicense"),"ErrorLicense",JOptionPane.ERROR_MESSAGE);
							    	
									errInfo = Util.getTranslate("Setting", "ExpiredLicense");
									//isErrorLicense = false;
									Util.logger.fatal(errInfo);
									//setInputEnable(true);
									return false;								
								}
							    
							   	boolean dateOnlyLicense=false;
							   	if(licenseInfo[0].equals("ANYID")){
							   		dateOnlyLicense=true;
							   	}
							    Util.logger.info("-------------licensedDate="+licensedDate);			
							    
								if (rdbtnNewRadioButton.isSelected()) {
									// Disable license feature at first
									if(dateOnlyLicense ||(usernameField.getText().equals(licenseInfo[0]))){
										login = new MetadataLoginUtil(usernameField.getText(), passwordField.getPassword(), true,hostField.getText(),portField.getText(),proxyUsername.getText(),String.valueOf(proxyPassword.getPassword()));
										
									}else{
										errInfo = Util.getTranslate("Setting", "ErrorLicense");
										isErrorLicense = false;
										Util.logger.fatal(errInfo);
										setInputEnable(true);
										return false;
									}
								} else {
									if(dateOnlyLicense ||(usernameField.getText().startsWith(licenseInfo[0]))){
										login = new MetadataLoginUtil(usernameField.getText(), passwordField.getPassword(), false,hostField.getText(),portField.getText(),proxyUsername.getText(),String.valueOf(proxyPassword.getPassword()));
									}else{
										errInfo = Util.getTranslate("Setting", "ErrorLicense");
										isErrorLicense = false;
										Util.logger.fatal(errInfo);
										setInputEnable(true);
										return false;
									}
								}
							
								try {
									login.userLogin();
									errInfo = Util.getTranslate("Setting","Success");								
									//btnLogIn.setText(Util.getTranslate("Setting", "Logout"));
									//btnSelectMetadata.setEnabled(true);
									//show select metadata windows
									DescribeMetadataResult d = login.metadataConnection.describeMetadata(Util.API_VERSION);
									
									if(!d.getOrganizationNamespace().equals("")){
										JOptionPane.showMessageDialog(null, Util.getTranslate("Message","Not Support NameSpace"),"Not Support NameSpace", JOptionPane.ERROR_MESSAGE);										
										setInputEnable(true);								
										isFail = false;
										errInfo = Util.getTranslate("Setting","Fail");
										Util.logger.fatal(errInfo);										
										return false;
									}
									
									lblProcess.setVisible(true);
									savePro(login.loginUrl);
									uti.setLanguage(language);
									SelectMetadata sm;
									errorLabel.setText(errInfo);
									sm = new SelectMetadata(login.metadataConnection, uti,window);
									uti.setMetadataConnection(login.metadataConnection);
									uti.setConnection(login.Connection);
									uti.setSettingSubfolder(settingSubfolder.isSelected());                       //by wangyuxin
									uti.setModifiedFlag(cbModified.isSelected());
									sm.frame.setVisible(true);
									lblProcess.setVisible(false);								
									window.frame.setVisible(false);
									isSuccess = false;
									return true;
								} catch (ConnectionException ce) {					
									setInputEnable(true);								
									isFail = false;
									errInfo = Util.getTranslate("Setting","Fail");
									Util.logger.fatal(errInfo);
									Util.logger.fatal(ce.getStackTrace());
									StringWriter sw = new StringWriter();
									ce.printStackTrace(new PrintWriter(sw));
									String stackTrace = sw.toString();
									Util.logger.debug("Get the exception String"+stackTrace);
									return false;
								}
							
							}
							
							@Override
							protected void done()// this is called after
													// doInBackground() has finished
							{
								try {
									if (this.get()) {
										errorLabel.setText(errInfo);
										frame.setCursor(Cursor.DEFAULT_CURSOR);
										setInputEnable(false);
										lblProcess.setVisible(false);
										window.frame.setVisible(false);
									} else {
										errorLabel.setText(errInfo);
										frame.setCursor(Cursor.DEFAULT_CURSOR);
										usernameField.setCursor(new Cursor(Cursor.TEXT_CURSOR));
										usernameField.setFocusable(true);
										passwordField.setCursor(new Cursor(Cursor.TEXT_CURSOR));
										passwordField.setFocusable(true);
										lblProcess.setVisible(false);
									}
								} catch (InterruptedException|ExecutionException ce) {
									ce.printStackTrace();									
								}
							}
						};
						frame.setCursor(Cursor.WAIT_CURSOR);
						usernameField.setCursor(new Cursor(Cursor.WAIT_CURSOR));
						usernameField.setFocusable(false);
						passwordField.setCursor(new Cursor(Cursor.WAIT_CURSOR));
						passwordField.setFocusable(false);
						errorLabel.setText(Util.getTranslate("Setting", "Error"));
						sw.execute();
					}
					if(Util.getTranslate("Setting", "Logout").equals(e.getActionCommand())){
						System.exit(0);
					}
				}
			});
			btnLogIn.setBounds(268, 229, 91, 30);
			frame.getContentPane().add(btnLogIn);
	
			ButtonGroup bgroup1 = new ButtonGroup();
			rdbtnNewRadioButton = new JRadioButton(Util.getTranslate("Setting","Production"));
			rdbtnNewRadioButton.setBounds(192, 160, 113, 21);
			rdbtnNewRadioButton.setSelected(true);				
			frame.getContentPane().add(rdbtnNewRadioButton);
			bgroup1.add(rdbtnNewRadioButton);
	
			rdbtnNewRadioButton_1 = new JRadioButton(Util.getTranslate("Setting","Sandbox"));
			rdbtnNewRadioButton_1.setBounds(378, 160, 125, 21);				
			frame.getContentPane().add(rdbtnNewRadioButton_1);
			bgroup1.add(rdbtnNewRadioButton_1);
			if(propertiesout.getProperty("sfdc.endpoint").contains("test.")){
				rdbtnNewRadioButton_1.setSelected(true);
			}
			
			errorLabel = new JLabel("");
			errorLabel.setBounds(53, 120, 519, 25);
			frame.getContentPane().add(errorLabel);
			Icon image;
			if(isBatch){
			    image = new ImageIcon(WSC.class.getResource("/common/icons/process.gif"));
			}else{
				image = new ImageIcon("common/icons/process.gif");
			}
			
			lblProcess = new JLabel(image);
			lblProcess.setVisible(false);
			lblProcess.setBounds(245, 250, 32, 32);
			frame.getContentPane().add(lblProcess);
			
			Vector<Component> order = new Vector<Component>(10);
			order.add(usernameField);
			order.add(passwordField);
			order.add(rdbtnNewRadioButton);
			order.add(rdbtnNewRadioButton_1);
			order.add(btnLogIn);
			newPolicy = new MyOwnFocusTraversalPolicy(order);
			frame.setFocusTraversalPolicy(newPolicy);
			frame.getRootPane().setDefaultButton(btnLogIn);
			Util.logger.info("Frame initialization complete.");
		} catch (Exception e) {
			Util.logger.error("Frame initialization failure.");
			Util.logger.fatal(e.getStackTrace());
		}		
	}
	public void setInputEnable(boolean isEnabled){
		usernameField.setEnabled(isEnabled);
		passwordField.setEnabled(isEnabled);
		rdbtnNewRadioButton.setEnabled(isEnabled);
		rdbtnNewRadioButton_1.setEnabled(isEnabled);
		btnLogIn.setEnabled(isEnabled);	
		usernameField.setFocusable(isEnabled);
		passwordField.setFocusable(isEnabled);
	}
	public void actionPerformed(ActionEvent e) {
		if (e.getSource() == cb) {
			language = languageMap.get(cb.getSelectedItem().toString());
		}
		if (Util.getTranslate("Setting", "Setting").equals(e.getActionCommand())) {
			SettingWin(null);
			settingFrame.setVisible(true);
		}
		if ("Exit".equals(e.getActionCommand())) {
			System.exit(0);
		}
		if(Util.getTranslate("Setting", "About").equals(e.getActionCommand())){
			VersionWin(frame);
			versionFrame.setVisible(true);
		}
	}
    public void VersionWin(JFrame fra){
    	Util.logger.info("VersionWin Start.");
        versionFrame = new JDialog(fra,"Version");
        int newX = fra.getX()+((FRAME_HEIGHT - FRAME_WIDTH)/2);
        int newY = fra.getY()+((FRAME_HEIGHT - FRAME_WIDTH)/2);        
    	versionFrame.setBounds(newX, newY, 300, 200);
    	versionFrame.setResizable(false);
    	versionFrame.getContentPane().setLayout(null); 
		
		JLabel venMessage = new JLabel(Util.getTranslate("Setting", "Version"));
		venMessage.setBounds(0, 0, 300, 200);
		versionFrame.getContentPane().add(venMessage);
		Util.logger.info("VersionWin Complete.");
    }
	public void SettingWin(final SelectMetadata csm) {
		Util.logger.info("SettingWin Start.");

		settingFrame = new JFrame(Util.getTranslate("Setting", "Setting"));
		//left top LOGO image
		settingFrame.setIconImage(new ImageIcon("./common/icons/logo.png").getImage());
		settingFrame.setBounds(200, 100, 550, 500);
		settingFrame.setResizable(false);
		settingFrame.getContentPane().setLayout(null);

		JButton btnSave = new JButton(Util.getTranslate("Setting", "Save"));
		btnSave.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				try {
					Util.LoadProperties(language);
				} catch (IOException e) {
					e.printStackTrace();
				}
				String licensePath = licenseField.getText();
				if(licensePath==null||"".equals(licenseField.getText())){
					
					licenseDate.setText(Util.getTranslate("MESSAGE", "NullLicense"));
			    	licenseDate.setForeground(Color.red);
			    
				}else{ 
					String licenseStr=Util.checkLicense(licensePath);
					System.out.println(licenseStr);
					if(licenseStr==null|| licenseStr.equals("")){
						//isLicensePath = false;
						licenseDate.setText(Util.getTranslate("Setting","ErrorLicense"));
					}
					else{
						Date curDate = new Date();
					    Date licensedDate=curDate;
						String[] licenseInfo = licenseStr.split(",",0);
						
						try {
							licensedDate = new SimpleDateFormat("yyyy/MM/dd").parse(licenseInfo[1]);
						} catch (ParseException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				    
						if(curDate.after(licensedDate)){
				    		licenseDate.setText(Util.getTranslate("Setting","LICENSEDATE")+String.valueOf(licenseInfo[1]));
				    	}else{
				    		licenseDate.setText(Util.getTranslate("Setting","LICENSEDATE")+String.valueOf(licenseInfo[1]));
				    		 settingFrame.setVisible(false);
						     savePro("");
				    	}
					}
					
			}
				if (csm != null) {
					csm.changeLanguage();
				} else {
					changeLanguage();
				}
			}
		});
		btnSave.setBounds(165, 420, 100, 25);
		settingFrame.getContentPane().add(btnSave);

		JButton btnCancel = new JButton(Util.getTranslate("Setting", "Cancel"));
		btnCancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				settingFrame.setVisible(false);
			}
		});
		btnCancel.setBounds(285, 420, 100, 25);//
		settingFrame.getContentPane().add(btnCancel);

		JLabel lblLanguage = new JLabel(Util.getTranslate("Setting", "Language"));
		lblLanguage.setBounds(10, 40, 110, 20);//
		lblLanguage.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblLanguage);

		cb = new JComboBox();
		for (String s : languageMap.keySet()) {
			cb.addItem(s);
		}
		if (propertiesout.getProperty("Language").isEmpty()) {
			cb.setSelectedItem("English");
		} else {
			for (String s : languageMap.keySet()) {
				if (propertiesout.getProperty("Language").equals(
						languageMap.get(s))) {
					cb.setSelectedItem(s);
				}
			}
		}
		language = languageMap.get(cb.getSelectedItem().toString());
		cb.setBounds(140, 40, 90, 25);
		cb.addActionListener(this);
		settingFrame.getContentPane().add(cb);
		
		JLabel lblProxyHost = new JLabel(Util.getTranslate("Setting","Proxy_host"));
		System.out.println(lblProxyHost);
		lblProxyHost.setBounds(10, 83, 110, 20);
		lblProxyHost.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblProxyHost);

		hostField = new JTextField(propertiesout.getProperty("sfdc.proxyHost"));
		hostField.setBounds(140, 80, 330, 25);
		settingFrame.getContentPane().add(hostField);
		hostField.setColumns(10);

		JLabel lblProxyPort = new JLabel(Util.getTranslate("Setting","Proxy_port"));
		lblProxyPort.setBounds(10, 123, 110, 20);
		lblProxyPort.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblProxyPort);

		portField = new JTextField(propertiesout.getProperty("sfdc.proxyPort"));
		portField.setBounds(140, 120, 90, 25);
		settingFrame.getContentPane().add(portField);
		portField.setColumns(10);

		JLabel lblProxyUser = new JLabel(Util.getTranslate("Setting","PROXY_USERNAME"));
		lblProxyUser.setBounds(10, 163, 110, 20);
		lblProxyUser.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblProxyUser);

		proxyUsername = new JTextField(propertiesout.getProperty("sfdc.proxyUsername"));
		proxyUsername.setBounds(140, 160, 330, 25);
		settingFrame.getContentPane().add(proxyUsername);
		proxyUsername.setColumns(10);		


		JLabel lblProxyPass = new JLabel(Util.getTranslate("Setting","PROXY_PASSWORK"));
		lblProxyPass.setBounds(10, 203, 110, 20);
		lblProxyPass.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblProxyPass);

		proxyPassword = new JPasswordField(Util.decrypt(propertiesout.getProperty("sfdc.proxyPassword")));
		proxyPassword.setBounds(140, 200, 330, 25);
		settingFrame.getContentPane().add(proxyPassword);
		proxyPassword.setColumns(10);			
		
		JLabel lblOutputDiretary = new JLabel(Util.getTranslate("Setting","Location"));
		lblOutputDiretary.setBounds(10, 243, 110, 20);
		lblOutputDiretary.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblOutputDiretary);		
		
	    cbModified= new JCheckBox(Util.getTranslate("Setting", "MODIFIEDFLAG"));
		if("true".equals(propertiesout.getProperty("setting.modifiedFlag"))){
			cbModified.setSelected(true);
		}
		cbModified.setBounds(250, 40, 150, 25);
		settingFrame.getContentPane().add(cbModified);
				
		if (propertiesout.getProperty("process.outputSuccess").isEmpty()) {
			localField = new JTextField(System.getProperty("user.dir")
					+ "\\Downloads");
		} else {
			localField = new JTextField(
					propertiesout.getProperty("process.outputSuccess"));
		}

		localField.setColumns(10);
		localField.setBounds(140, 240, 330, 25);
		settingFrame.getContentPane().add(localField);
		
		JButton btnSelect = new JButton(Util.getTranslate("Setting", "Select"));
		btnSelect.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if (jfc.showOpenDialog(settingFrame) == JFileChooser.APPROVE_OPTION) {
					localField.setText(jfc.getSelectedFile().getAbsolutePath());
				}
			}
		});
		btnSelect.setBounds(140, 270, 90, 25);
		settingFrame.getContentPane().add(btnSelect);
		
		//by wangyuxin
		settingSubfolder = new JCheckBox(Util.getTranslate("Setting", "Subfolder"));
		if("true".equals(propertiesout.getProperty("setting.Subfolder"))){
			settingSubfolder.setSelected(true);
		}
		settingSubfolder.setBounds(250, 270, 220, 25);
		settingFrame.getContentPane().add(settingSubfolder);
		
		lblLicense = new JLabel(Util.getTranslate("Setting","License"));
		lblLicense.setBounds(10, 310, 110, 20);
		lblLicense.setHorizontalAlignment(SwingConstants.RIGHT);
		settingFrame.getContentPane().add(lblLicense);
		
		licenseField = new JTextField();
		licenseField.setText(propertiesout.getProperty("LicensePath"));
		licenseField.setBounds(140, 310, 330, 25);
		settingFrame.getContentPane().add(licenseField);
		
		btnLicense = new JButton(Util.getTranslate("Setting", "Select"));
		btnLicense.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				if(chooser.showOpenDialog(settingFrame) == JFileChooser.APPROVE_OPTION){
					File file = chooser.getSelectedFile();
					licenseField.setText(file.getAbsolutePath());
				   try {
						String licensePath = file.getAbsolutePath();
						if(licensePath==null){
							licenseDate.setText(Util.getTranslate("Message", "licenseFileNotExist"));
							
						}else{
							String licenseStr=Util.checkLicense(licensePath);
								if(licenseStr==null){
									licenseDate.setText(Util.getTranslate("Setting", "ErrorLicense"));
									
								}else{
									Date curDate = new Date();
								    Date licensedDate=curDate;
									String[] licenseInfo = licenseStr.split(",",0);
									
									licensedDate = new SimpleDateFormat("yyyy/MM/dd").parse(licenseInfo[1]);

							    	//if(curDate.after(licensedDate)){
							    		licenseDate.setText(Util.getTranslate("Setting","LICENSEDATE")+String.valueOf(licenseInfo[1]));
							    	//}else{
							    	//	licenseDate.setText(Util.getTranslate("Setting", "LICENSEDATE")+String.valueOf(licenseInfo[1]));
							    	//}
								}
							
						}
						
				    }catch(Exception ec){
				    	System.out.println(ec);
				    	licenseDate.setText(Util.getTranslate("Setting", "ErrorLicense"));
				    	licenseDate.setForeground(Color.red);
				    }	
				    
				}
			}
		});
		btnLicense.setBounds(140, 340, 90, 25);
		settingFrame.getContentPane().add(btnLicense);
		
	    licenseDate = new JLabel(Util.getTranslate("Setting", "LICENSEDATE")+licenseExpireDate);
		licenseDate.setBounds(250, 340, 200, 25);
		licenseDate.setForeground(Color.red);
		settingFrame.getContentPane().add(licenseDate);			
		Util.logger.info("SettingWin complete.");		
	}
	
	public boolean checkOutputPathValid(String outputPath){
		boolean isValid = true;
		File file = null;
		if(outputPath == null || outputPath.trim().equals("")){
			return false;
		}
		try {
			file = new File(outputPath + "\\test.txt");
			OutputStream out = new FileOutputStream(file);
			out.write("test".getBytes());
			out.flush();
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		if(isValid && file != null){
			file.delete();
		}
		return isValid;
	}
	// for tab use
	public static class MyOwnFocusTraversalPolicy extends FocusTraversalPolicy {
		Vector<Component> order;

		public MyOwnFocusTraversalPolicy(Vector<Component> order) {
			this.order = new Vector<Component>(order.size());
			this.order.addAll(order);
		}

		public Component getComponentAfter(Container focusCycleRoot,
				Component aComponent) {
			int idx = (order.indexOf(aComponent) + 1) % order.size();
			return order.get(idx);
		}

		public Component getComponentBefore(Container focusCycleRoot,
				Component aComponent) {
			int idx = order.indexOf(aComponent) - 1;
			if (idx < 0) {
				idx = order.size() - 1;
			}
			return order.get(idx);
		}

		public Component getDefaultComponent(Container focusCycleRoot) {
			return order.get(0);
		}

		public Component getLastComponent(Container focusCycleRoot) {
			return order.lastElement();
		}

		public Component getFirstComponent(Container focusCycleRoot) {
			return order.get(0);
		}
	}
}
