package wsc;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GridLayout;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.nio.channels.FileChannel;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.Enumeration;
import java.util.EventObject;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.SortedMap;
import java.util.TreeMap;

import javax.swing.BorderFactory;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.JTree;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.event.CellEditorListener;
import javax.swing.event.ChangeEvent;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.TreeModelEvent;
import javax.swing.event.TreeModelListener;
import javax.swing.text.BadLocationException;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeCellRenderer;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreeCellEditor;
import javax.swing.tree.TreeCellRenderer;
import javax.swing.tree.TreeModel;
import javax.swing.tree.TreePath;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.sforce.soap.metadata.ApprovalProcess;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.DescribeMetadataObject;
import com.sforce.soap.metadata.DescribeMetadataResult;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.Folder;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.metadata.Network;
import com.sforce.soap.metadata.SharingRules;
import com.sforce.soap.metadata.Workflow;
import com.sforce.soap.partner.fault.ExceptionCode;
import com.sforce.soap.partner.fault.UnexpectedErrorFault;
import com.sforce.ws.ConnectionException;

import source.Util;
import source.UtilConnectionInfc;
import source.UtilExportInfc;


public class SelectMetadata implements ActionListener {

	public JFrame frame;
	private JTextField textField;
	private JTree tree;
	private MetadataConnection metadataConnection;
	public String lastRunTime;
	private JPanel panel;
	private JCheckBox chckbxShowOnlySelected;
	private JCheckBox chckbxOpenDir;
	private Map<String, List<String>> allMetadataMap;
	private Map<String, List<String>> onlySelecteMetadataMap;
	private Map<String, List<String>> selecteMetadataMap;
	private Map<String, String> metaDataLabelMap;
	private String text = "";
	private UtilConnectionInfc uti;
	private Util ut;
	private JMenu menuSetting;
	private JMenuItem mItemSet;
	private JMenu menuHelp;
	private JMenuItem mAbout;
	
	private JMenuItem mExit;
	private Map<String, String> languageMap;
	private JLabel timeLabel;
	private JLabel searchLabel;
	private JButton expandButton;
	private JButton collapseButton;
	private JButton refButton;
	private JButton btnOk;
	private JLabel lblProcess;
	WSC window; 
	/**
	 * Create the application.
	 * 構造方法
	 */
	public SelectMetadata(MetadataConnection metadataConnection,
			UtilConnectionInfc uti,WSC window) {
		this.uti = uti;
		this.metadataConnection = metadataConnection;
		
		metaDataLabelMap = new HashMap<String, String>();
		
		initialize();
		this.window = window; 
	}
	/**
	 * use thread update UI
	 * 画面ＵＩの翻訳を更新
	 */
	public void changeLanguage() {
			new Thread(new Runnable() {
				@Override
				public void run() {
					frame.setTitle(Util.getTranslate("Select", "Title"));
					chckbxShowOnlySelected.setText(Util.getTranslate("Select", "Show"));
					chckbxOpenDir.setText(Util.getTranslate("Select", "Open"));
					timeLabel.setText(Util.getTranslate("Select", "Time"));
					searchLabel.setText(Util.getTranslate("Select", "Search"));
					menuSetting.setText(Util.getTranslate("Setting", "Setting"));
					mItemSet.setText(Util.getTranslate("Setting", "Setting"));
					menuHelp.setText(Util.getTranslate("Setting", "Help"));
					mAbout.setText(Util.getTranslate("Setting", "About"));
					expandButton.setToolTipText(Util.getTranslate("Select", "Expand"));
					collapseButton.setToolTipText(Util.getTranslate("Select", "Collapse"));
					refButton.setToolTipText(Util.getTranslate("Select", "Refresh"));
					btnOk.setText(Util.getTranslate("Select", "OK"));
					languageMap = new HashMap<String, String>();					
					languageMap.put(Util.getTranslate("language", "EN"), "EN");
					languageMap.put(Util.getTranslate("language", "JP"), "JP");
					languageMap.put(Util.getTranslate("language", "CN"), "CN");
				}
			}).start();
	}
	/**
	 * Initialize the contents of the frame.
	 * 初期化方法
	 */
	private void initialize() {
		Util.logger.info("-------------SelectMetadata initialize started");
		try {
		     UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());//change the look
//		     
		  } catch (ClassNotFoundException|InstantiationException|IllegalAccessException|UnsupportedLookAndFeelException ex) {
		      ex.printStackTrace();
		}
		ut = new Util();
		try {
			ut.LoadProperties(uti.language);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		
		allMetadataMap=getMetadataFromServer();	
		Object[] unsort_key = allMetadataMap.keySet().toArray();

		frame = new JFrame(Util.getTranslate("Select", "Title"));
		frame.setIconImage(new ImageIcon("./common/icons/logo.png").getImage());
		frame.setBounds(680, 10, 600, 660);
		//overwrite close event
		//画面閉じる方法
        frame.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
    			frame.setVisible(false);
            	MetadataLoginUtil.userLogout();    			
    			window.errorLabel.setText("");							
    			window.usernameField.setEnabled(true);
    			window.passwordField.setEnabled(true);
    			window.rdbtnNewRadioButton.setEnabled(true);
    			window.rdbtnNewRadioButton_1.setEnabled(true);
    			window.btnLogIn.setEnabled(true);
				window.usernameField.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
				window.usernameField.setFocusable(true);
				window.passwordField.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
				window.passwordField.setFocusable(true);    			
    			window.frame.setVisible(true);
    			           	
            }
        });		
		frame.setResizable(false);
		frame.getContentPane().setLayout(null);
		//Menubar
		JMenuBar mb = new JMenuBar();
		menuSetting = new JMenu(Util.getTranslate("Setting", "Setting"));
		mItemSet = new JMenuItem(Util.getTranslate("Setting", "Setting"));
		mExit = new JMenuItem("Exit");
		mExit.addActionListener(this);
		//mExit = new JMenuItem(Util.getTranslate("Setting", "About"));
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
		
		//検索用テキストボックス
	    searchLabel = new JLabel(Util.getTranslate("Select", "Search"));
	    searchLabel.setBounds(12, 42, 106, 13);
		frame.getContentPane().add(searchLabel);	
		
		textField = new JTextField();
		textField.setBounds(100, 39, 470, 19);
		frame.getContentPane().add(textField);
		textField.setColumns(10);
		//検索内容更新
		textField.getDocument().addDocumentListener(new DocumentListener() {            
            @Override
            public void removeUpdate(DocumentEvent e) {
            	//selecteMetadataMap=getSelectedNode(tree,false);               
                try {
                    text = e.getDocument().getText(e.getDocument().getStartPosition().getOffset(), e.getDocument().getLength());
                } catch (BadLocationException ex) {
                    ex.printStackTrace();
                }
                
                if(!text.isEmpty()){
                	if(!chckbxShowOnlySelected.isSelected()){
	                    showAllTree(tree,getSearchedMetadata(allMetadataMap,text));
	                    expandAll(tree);
                	}else{
	                    showAllTree(tree,getSearchedMetadata(selecteMetadataMap,text));
	                    expandAll(tree);
                	}
                }else{
                	if(!chckbxShowOnlySelected.isSelected()){
                		showAllTree(tree,allMetadataMap);
                    	collapseAll(tree);
                	}else{
                		showAllTree(tree,getSearchedMetadata(selecteMetadataMap,text));
	                    expandAll(tree);
                	}
                	
                }

            }
             
            @Override
            public void insertUpdate(DocumentEvent e) {
            	//selecteMetadataMap=getSelectedNode(tree,false);
                try {
                    text = e.getDocument().getText(e.getDocument().getStartPosition().getOffset(), e.getDocument().getLength());
                } catch (BadLocationException ex) {
                    ex.printStackTrace();
                }
                if(!chckbxShowOnlySelected.isSelected()){
                	showAllTree(tree,getSearchedMetadata(allMetadataMap,text));
                    expandAll(tree);
                }else{
                	showAllTree(tree,getSearchedMetadata(selecteMetadataMap,text));
                    expandAll(tree);
                }
                
            }             
            @Override
            public void changedUpdate(DocumentEvent e) {
            }
        });	
		
		chckbxShowOnlySelected = new JCheckBox(Util.getTranslate("Select", "Show"));
		chckbxShowOnlySelected.setBounds(8, 67, 183, 21);
		//チェック内容だけ表示更新
		ItemListener itemListener = new ItemListener(){       	 
            public void itemStateChanged(ItemEvent e) {                 
                Object obj = e.getItem();
                if(obj.equals(chckbxShowOnlySelected)) {
                    if(chckbxShowOnlySelected.isSelected()) {
                    	showSelectedTree(tree);    
                    	onlySelecteMetadataMap=getSelectedNode(tree,false);
                    	expandAll(tree);
                    }else{
                    	//selecteMetadataMap=getSelectedNode(tree,false);
                    	selecteMetadataMap=onlySelecteMetadataMap;
                    	System.out.println("--------------------"+selecteMetadataMap);
                    	if(text.isEmpty()){
                    		showAllTree(tree,allMetadataMap);
                    		collapseAll(tree);
                    	}else{
                    		showAllTree(tree,getSearchedMetadata(allMetadataMap,text));
                    		expandAll(tree);
                    	}
                    	
                    }
                }     
            }
       };

        chckbxShowOnlySelected.setSelected(false);
        chckbxShowOnlySelected.addItemListener(itemListener);
		frame.getContentPane().add(chckbxShowOnlySelected);
		if(WSC.isBatch){
			expandButton = new JButton(new ImageIcon(SelectMetadata.class.getResource("/common/icons/Expand.png")));
		}else{
			expandButton = new JButton(new ImageIcon("common/icons/Expand.png"));
		}

		
		expandButton.setToolTipText(Util.getTranslate("Select", "Expand"));
		//全部展開操作
		expandButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				expandAll(tree);
			}
		});
		expandButton .setPreferredSize(new Dimension(26,26));
		expandButton.setBounds(12, 10, 26, 26);
		frame.getContentPane().add(expandButton);
		if(WSC.isBatch){
			collapseButton = new JButton(new ImageIcon(SelectMetadata.class.getResource("/common/icons/Collapse.png")));
		}else{
			collapseButton = new JButton(new ImageIcon("common/icons/Collapse.png"));
		}	    

		
		collapseButton.setToolTipText(Util.getTranslate("Select", "Collapse"));
		//全部閉め操作
		collapseButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				collapseAll(tree);
			}
		});
		collapseButton.setPreferredSize(new Dimension(26,26));
		collapseButton.setBounds(40, 10, 26, 26);
		frame.getContentPane().add(collapseButton);
		if(WSC.isBatch){
			refButton = new JButton(new ImageIcon(SelectMetadata.class.getResource("/common/icons/Refresh.png")));
		}else{
			refButton = new JButton(new ImageIcon("common/icons/Refresh.png"));
		}	    

		refButton.setToolTipText(Util.getTranslate("Select", "Refresh"));
		//表示内容再更新操作
		refButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				frame.setCursor(Cursor.WAIT_CURSOR);
				frame.remove(panel);
				allMetadataMap = getMetadataFromServer();
				panel=listMetadata();
				panel.setBounds(12, 90, 564, 469);
				frame.add(panel);
				frame.invalidate();
				frame.validate();
				frame.setCursor(Cursor.DEFAULT_CURSOR);

			}
		});
		refButton.setPreferredSize(new Dimension(26,26));
		refButton.setBounds(68, 10, 26, 26);
		frame.getContentPane().add(refButton);
		
	    timeLabel = new JLabel(Util.getTranslate("Select", "Time"));
		timeLabel.setBounds(365, 10, 106, 13);
		frame.getContentPane().add(timeLabel);	
		Icon image;
		if(WSC.isBatch){
			image = new ImageIcon(SelectMetadata.class.getResource("/common/icons/folder.png"));
		}else{
			image = new ImageIcon("common/icons/bprocess.gif");
		}			
		lblProcess = new JLabel(image);
		lblProcess.setBounds(250,255, 124, 124);
		lblProcess.setVisible(false);
		frame.getContentPane().add(lblProcess);
		
		btnOk = new JButton(Util.getTranslate("Select", "OK"));
		//エクスポート開始操作
		btnOk.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent arg0) {	
				Date now = new Date(); 
				SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
				String systemTime = dateFormat.format(now)+".log";
				File tempFile;
				String c;
				if(WSC.isBatch){
					tempFile=new File(".././logs/normal_app.log");
					String d=tempFile.getParent();
					c=d+File.separator+"logs";
				}else{
					tempFile=new File("logs/normal_app.log");
					c=tempFile.getParent();
				}	  
		        File mm=new File(c,systemTime);  
		        fileChannelCopy(tempFile,mm);
		        try{
			        FileWriter fw = new FileWriter(tempFile);
			        BufferedWriter bw1 = new BufferedWriter(fw);
			        bw1.write("");
			        bw1.close();
		        }catch(Exception e){
		        }
				new Thread(new Runnable() {
					@Override
					public void run() {
						btnOk.setEnabled(false);
						lblProcess.setVisible(true);
						frame.setCursor(Cursor.WAIT_CURSOR);
						SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
						SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");//Set the date format
						Date d = new Date();
						String path = window.propertiesout.getProperty("process.outputSuccess");
						if("true".equals(window.propertiesout.getProperty("setting.Subfolder"))){
							path += "\\"+sdf.format(d);
							File outfile = new File(path);
							outfile.mkdir();
							uti.setDownloadPath(path);
						}else{
							uti.setDownloadPath(path);
						}
						try {
							d = df.parse(lastRunTime);
						} catch (ParseException e) {
							e.printStackTrace();
						}
						uti.setLastUpdateTime(d.getTime());
						// get new current time
						Date date = new Date();
						lastRunTime = df.format(date);
						Map<String, List<String>> selectedMap = getSelectedNode(tree,
								true);
						uti.setExportMap(selectedMap);
						// Modify By YU Start
						UtilExportInfc pubExport = new UtilExportInfc();
						try {
							pubExport.ExportInfc();
							lblProcess.setVisible(false);
							if (chckbxOpenDir.isSelected()){
								Runtime rt = Runtime.getRuntime();
							    String cmd = "explorer " + path;
							    rt.exec(cmd);
							}
							JOptionPane.showMessageDialog(null,Util.getTranslate("Message","success"), "Success",JOptionPane.INFORMATION_MESSAGE);
						} catch (UnexpectedErrorFault e) {
							lblProcess.setVisible(false);
							if(e.getExceptionCode()==ExceptionCode.INVALID_SESSION_ID){
								JOptionPane.showMessageDialog(null, Util.getTranslate("Message","INVALID_SESSION_ID"),"INVALID_SESSION_ID", JOptionPane.ERROR_MESSAGE);
				    			frame.setVisible(false);  			
				    			window.errorLabel.setText("");							
				    			window.usernameField.setEnabled(true);
				    			window.passwordField.setEnabled(true);
				    			window.rdbtnNewRadioButton.setEnabled(true);
				    			window.rdbtnNewRadioButton_1.setEnabled(true);
				    			window.btnLogIn.setEnabled(true);
								window.usernameField.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
								window.usernameField.setFocusable(true);
								window.passwordField.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
								window.passwordField.setFocusable(true);    			
				    			window.frame.setVisible(true);						
							}
							e.printStackTrace();
						} catch(Exception e){
							lblProcess.setVisible(false);
							JOptionPane.showMessageDialog(null, Util.getTranslate("Message","failure"),"Failure", JOptionPane.ERROR_MESSAGE);
							e.printStackTrace();							
						}
						
						// Modify By YU End
						btnOk.setEnabled(true);
						frame.setCursor(Cursor.DEFAULT_CURSOR);
					}
				}).start();
				
			}
			
			//ファイルコピー（tempFile　ＴＯ　mm）
			private void fileChannelCopy(File tempFile, File mm) {
				FileInputStream fi = null;
		        FileOutputStream fo = null;
		        FileChannel in = null;
		        FileChannel out = null;
		        try {
		            fi = new FileInputStream(tempFile);
		            fo = new FileOutputStream(mm);
		            in = fi.getChannel();
		            out = fo.getChannel();
		            in.transferTo(0, in.size(), out);
		        } catch (IOException e) {
		           e.printStackTrace();
		        } finally {
		            try {
		                fi.close();
		                in.close();
		                fo.close();
		                out.close();
		            } catch (IOException e) {
		                e.printStackTrace();
		            }
		        }
			}
		});
		btnOk.setBounds(485, 570, 91, 21);
		frame.getContentPane().add(btnOk);
		
		
		chckbxOpenDir = new JCheckBox(Util.getTranslate("Select", "Open"));
		chckbxOpenDir.setBounds(12, 570, 183, 21);
		chckbxOpenDir.setSelected(true);
		frame.getContentPane().add(chckbxOpenDir);
		
		
		panel=listMetadata();
		panel.setBounds(12, 90, 564, 469);
		frame.getContentPane().add(panel);
		
		//前回実行時刻
		JLabel lastRunTimeLabel = new JLabel(lastRunTime);
		lastRunTimeLabel.setBounds(468, 10, 119, 13);
		frame.getContentPane().add(lastRunTimeLabel);
		selecteMetadataMap=getSelectedNode(tree,false);
	}
	
	//設定メニュー、閉める操作、ヘルプ操作
	public void actionPerformed(ActionEvent e) {
		if (Util.getTranslate("Setting","Setting").equals(e.getActionCommand())) {
			window.SettingWin(this);
			window.settingFrame.setVisible(true);
		}
		//if (Util.getTranslate("Setting","Setting").equals(e.getActionCommand())) {
		if ("Exit".equals(e.getActionCommand())) {
			frame.setVisible(false);  			
			window.errorLabel.setText("");							
			window.usernameField.setEnabled(true);
			window.passwordField.setEnabled(true);
			window.rdbtnNewRadioButton.setEnabled(true);
			window.rdbtnNewRadioButton_1.setEnabled(true);
			window.btnLogIn.setEnabled(true);
			window.usernameField.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
			window.usernameField.setFocusable(true);
			window.passwordField.setCursor(new Cursor(Cursor.DEFAULT_CURSOR));
			window.passwordField.setFocusable(true);    			
			window.frame.setVisible(true);
			MetadataLoginUtil.userLogout();
			
		}	
		if(Util.getTranslate("Setting", "About").equals(e.getActionCommand())){
			window.VersionWin(frame);
			window.versionFrame.setVisible(true);
		}		
	}
	
	//エクスポート操作（旧方法みたい？）
	//added 2014/09/18
	public void exportToExcel(final Map<String,List<String>> map){
		
		new Thread(new Runnable(){
			public void run() {
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");//Set the date format
				Date d = new Date();
				String path = window.propertiesout.getProperty("process.outputSuccess");
				if("true".equals(window.propertiesout.getProperty("setting.Subfolder"))){
					path += "\\"+sdf.format(d);
					File outfile = new File(path);
					outfile.mkdir();
					uti.setDownloadPath(path);
				}else{
					uti.setDownloadPath(path);
				}
				try {
					d = df.parse(lastRunTime);
				} catch (ParseException e) {
					e.printStackTrace();
				}
				uti.setLastUpdateTime(d.getTime());
				// get new current time
				Date date = new Date();
				lastRunTime = df.format(date);
				//Map<String, List<String>> selectedMap = getSelectedNode(tree,true);
				uti.setExportMap(map);
				// Modify By YU Start
				UtilExportInfc pubExport = new UtilExportInfc();
				try {
					pubExport.ExportInfc();
					System.out.println("output completed");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}			
		}).start();
	}
	
	//選択リスト展開方法
	public void expandAll(JTree tree) {
		int row = 0;
		while (row < tree.getRowCount()) {
			tree.expandRow(row);
			row++;
		}
	}
	
	//選択リスト閉める方法
	public void collapseAll(JTree tree) {
		int row = tree.getRowCount() - 1;
		while (row >= 1) {
			tree.collapseRow(row);
			row--;
		}
	}

	//検索したデータ結果を取得方法
	public Map<String, List<String>> getSearchedMetadata(
			Map<String, List<String>> map, String text) {
		Map<String, List<String>> searchMap = new HashMap<String, List<String>>();
		for (String str : map.keySet()) {
			List<String> l = new ArrayList<String>();
			for (String s : map.get(str)) {
				if (!text.isEmpty()) {
					if (s.toUpperCase().contains(text.toUpperCase())) {
						l.add(s);
					}
				} else {
					l.add(s);
				}
			}
			Collections.sort(l);
			if (l.size() > 0 || str.toUpperCase().contains(text.toUpperCase())) {

				searchMap.put(str, l);
			}
		}
		if (!searchMap.isEmpty()) {
			searchMap = mapSortByKey(searchMap);
		}
		return searchMap;
	}

	//選択した選択値だけを表示する方法
	public void showSelectedTree(JTree tree) {
		TreeModel oldmodel = tree.getModel();
		DefaultMutableTreeNode oldroot = (DefaultMutableTreeNode) oldmodel
				.getRoot();
		CheckBoxNode rootcheck = (CheckBoxNode) oldroot.getUserObject();
		Enumeration breadth = oldroot.children();
		DefaultMutableTreeNode root = new DefaultMutableTreeNode();
		root.setUserObject(new CheckBoxNode(rootcheck.label, rootcheck.status));
		while (breadth.hasMoreElements()) {
			DefaultMutableTreeNode oldtype = (DefaultMutableTreeNode) breadth
					.nextElement();
			CheckBoxNode typecheck = (CheckBoxNode) oldtype.getUserObject();
			if (typecheck.status.equals(Status.SELECTED)
					|| typecheck.status.equals(Status.INDETERMINATE)) {
				DefaultMutableTreeNode type = new DefaultMutableTreeNode();
				type.setUserObject(new CheckBoxNode(typecheck.label,
						typecheck.status));
				Enumeration e = oldtype.children();
				while (e.hasMoreElements()) {
					DefaultMutableTreeNode oldnode = (DefaultMutableTreeNode) e
							.nextElement();
					CheckBoxNode nodecheck = (CheckBoxNode) oldnode
							.getUserObject();
					if (nodecheck.status.equals(Status.SELECTED)) {
						DefaultMutableTreeNode node = new DefaultMutableTreeNode();
						node.setUserObject(new CheckBoxNode(nodecheck.label,
								nodecheck.status));
						type.add(node);
					}
				}
				root.add(type);
			}
		}
		DefaultTreeModel model = new DefaultTreeModel(root);
		model.addTreeModelListener(new CheckBoxStatusUpdateListener());
		tree.setModel(model);
		tree.updateUI();
	}

    //全部選択値表示する方法
	public void showAllTree(JTree tree, Map<String, List<String>> metadataMap) {
		DefaultMutableTreeNode root = new DefaultMutableTreeNode(Util.getTranslate("MetaDataType","COMPONENTS"));
		Object obj = root.getUserObject();
		root.setUserObject(new CheckBoxNode((String) obj, Status.DESELECTED));
		DefaultTreeModel model = new DefaultTreeModel(root);
		model.addTreeModelListener(new CheckBoxStatusUpdateListener());
		CheckBoxStatusUpdateListener cl = new CheckBoxStatusUpdateListener();
		Map<String, List<String>> typeMap = selecteMetadataMap;
		for (String typeStr : metadataMap.keySet()) {
			List<String> typeList = new ArrayList<String>();
			if (typeMap.keySet().contains(typeStr)) {
				typeList = typeMap.get(typeStr);
			}

			DefaultMutableTreeNode type = new DefaultMutableTreeNode(typeStr);
			Object ob = type.getUserObject();
			if (typeMap.keySet().contains(typeStr)
					&& (typeMap.get(typeStr).size() == 0)) {
				type.setUserObject(new CheckBoxNode((String) ob,
						Status.SELECTED));
			} else {
				type.setUserObject(new CheckBoxNode((String) ob,
						Status.DESELECTED));
			}

			List<String> allFile = metadataMap.get(typeStr);
			for (String n : allFile) {
				DefaultMutableTreeNode node = new DefaultMutableTreeNode(n);
				Object o = node.getUserObject();
				if (typeList.contains(o)) {
					node.setUserObject(new CheckBoxNode((String) o,
							Status.SELECTED));
				} else {
					node.setUserObject(new CheckBoxNode((String) o,
							Status.DESELECTED));
				}
				type.add(node);
				cl.updateParentUserObject(type);
			}
			root.add(type);
			cl.updateParentUserObject(root);
		}
		tree.setModel(model);
		tree.updateUI();
	}

	//Metadataを取得（全部うオブジェクト名とラベル翻訳）
	public Map<String, List<String>> getMetadataFromServer() {
		Util.logger.info("-------------SelectMetadata getMetadataFromServer started");
		Map<String, List<String>> metadataMap = new HashMap<String, List<String>>();
		try {

			DescribeMetadataResult d = metadataConnection.describeMetadata(Util.API_VERSION);
			for (DescribeMetadataObject m : d.getMetadataObjects()) {
				System.out.println("m.getXmlName()=="+m.getXmlName());
			}
			for (DescribeMetadataObject m : d.getMetadataObjects()) {
				//avoid repeat
				//Util.logger.info("m========="+m);
				if(isUseable(m.getXmlName())){	
					ListMetadataQuery query = new ListMetadataQuery();
					query.setType(m.getXmlName());
					FileProperties[] lmr = metadataConnection.listMetadata(
							new ListMetadataQuery[] { query }, Util.API_VERSION);
					if (lmr != null) {
						List<String> allFile = new ArrayList<String>();
						for (FileProperties n : lmr) {
							allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));
							
						}
						Collections.sort(allFile);
						metadataMap.put(m.getXmlName(), allFile);
						//Util.logger.info("allFile="+allFile);
						Util.logger.info("getMetadataFromServer m.getXmlName()="+m.getXmlName());
						Util.logger.info("getMetadataFromServer Util.getTranslate="+Util.getTranslate("MetaDataType",m.getXmlName()));
						metaDataLabelMap.put(Util.getTranslate("MetaDataType",m.getXmlName()),m.getXmlName());
					}
				}
			}
		
			
			/** HomePage, User Group**/
			List<String> homeList = new ArrayList<String>();
			homeList.add(Util.getTranslate("MetaDataType","HomePageComponent"));
			homeList.add(Util.getTranslate("MetaDataType","HomePageLayout"));
			homeList.add(Util.getTranslate("MetaDataType","CustomPageWebLink"));
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","CustomPageWebLink"),"CustomPageWebLink");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","HomePageComponent"),"HomePageComponent");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","HomePageLayout"),"HomePageLayout");
			metadataMap.put("HomePage", homeList);
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","HomePage"),"HomePage");
			List<String> groupList = new ArrayList<String>();
			groupList.add(Util.getTranslate("MetaDataType","Group"));
			groupList.add(Util.getTranslate("MetaDataType","Queue"));
			groupList.add(Util.getTranslate("MetaDataType","User"));
			groupList.add(Util.getTranslate("MetaDataType","Role"));
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","Group"),"Group");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","Queue"),"Queue");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","User"),"User");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","Role"),"Role");			
			metadataMap.put("UserGroup", groupList);
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","UserGroup"),"UserGroup");			
			/** Document,email template, report, Dashboard **/
			metadataMap.put("Report", FolderHelper("Report"));
			metadataMap.put("Dashboard", FolderHelper("Dashboard"));	
			//metadataMap.put("EmailTemplate", FolderHelper("Email"));	
			metadataMap.put("Document", FolderHelper("Document"));	
			//metadataMap.put("CustomLabels", FolderHelper("CustomLabels"));	
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","Report"),"Report");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","Dashboard"),"Dashboard");
			//metaDataLabelMap.put(Util.getTranslate("MetaDataType","EmailTemplate"),"EmailTemplate");
			metaDataLabelMap.put(Util.getTranslate("MetaDataType","Document"),"Document");			
			//
		} catch (ConnectionException|UnsupportedEncodingException ce) {
			ce.getStackTrace();
		}
		// seperate standard object from custom object
		List<String> stand = new ArrayList<String>();
		List<String> custom = new ArrayList<String>();
		List<String> kav = new ArrayList<String>();
		List<String> external = new ArrayList<String>();
		List<String> customSetting = new ArrayList<String>();
		List<String> customMetadata = new ArrayList<String>();
		List<String> workFlow = new ArrayList<String>();
		
		List<String> cObjlist = new ArrayList<String>();
		for (String s : metadataMap.get("CustomObject")) {
			//s=nameSpace+s;
			if (s.contains("__c")) {
				cObjlist.add(s);
			} else if (s.contains("__kav")) {
				kav.add(s);
			} else if(s.contains("__x")){
				external.add(s);
			}else if(s.contains("__mdt")){
				customMetadata.add(s);
			}else{
				stand.add(s);
			}
		}
		List<Metadata> mds1 = ut.readMateData("CustomObject",cObjlist);
		if(mds1.size()>0){	
			for(int i=0;i<mds1.size();i++){
				CustomObject object = null;
				object = (CustomObject)mds1.get(i);
				if( object.getCustomSettingsType() == null){
					// when getCustomSettingsType is null ，s is CustomObject
					custom.add(object.getFullName());
				}else{
					// when getCustomSettingsType is not null，s is CustomSetting
					customSetting.add(object.getFullName());
				}
			}	
		}
		//check out workflow if they have data to export and show in the selection screen if they do
		List<String> list = new ArrayList<String>(); 
		for(String w:metadataMap.get("Workflow")){
				list.add(w);
		}
		List<Metadata> workflowmds = ut.readMateData("Workflow",list);
		if(workflowmds.size()>0){
			for(Metadata md:workflowmds){
			  Workflow wf = (Workflow)md;
			  if(wf.getRules().length>0
					||wf.getFlowActions().length>0){
						workFlow.add(md.getFullName());
			  } 
		    }
		}
		//get sharing rules
		List<String> share = new ArrayList<String>();
		List<Metadata> md = ut.readMateData("SharingRules",metadataMap.get("SharingRules"));
		share.add("Organization-Wide Defaults");
		if (md != null) {
			for(int i=0;i<md.size();i++){
				SharingRules sr = (SharingRules)md.get(i);
				boolean hasRule=false;
				if(sr.getSharingOwnerRules().length>0){
					hasRule=true;
				}
				if(sr.getSharingCriteriaRules().length>0){
					hasRule=true;
				}
				if(sr.getSharingTerritoryRules().length>0){
					hasRule=true;
				}
				if(hasRule){
					share.add(sr.getFullName());
				}
			}	
		}
		metadataMap.put("SharingSetting", share);
		//set SharingRules to empty since SharingSetting is SharingRules
		metadataMap.put("SharingRules",null);
		Collections.sort(workFlow);
		metadataMap.put("Workflow", workFlow);
		Collections.sort(stand);
		metadataMap.put("StandardObject", stand);
		Collections.sort(custom);
		metadataMap.put("CustomObject", custom);
		Collections.sort(customSetting);
		metadataMap.put("CustomSetting", customSetting);
		Collections.sort(customMetadata);
		metadataMap.put("CustomMetadata", customMetadata);		
		Collections.sort(external);
		metadataMap.put("EXTERNALOBJECT", external);
		
		//Append Other Derived MetaDataType
		List<String> derivedMetaDataType = Arrays.asList("CustomObject","StandardObject","SharingSetting","CustomSetting","ArticleType","EXTERNALOBJECT");
		for(Integer i=0;i<derivedMetaDataType.size();i++){
			metaDataLabelMap.put(Util.getTranslate("MetaDataType",derivedMetaDataType.get(i)),derivedMetaDataType.get(i));
		}
		for(String key: metadataMap.keySet()){
			if(translateListType.contains(key)){
				List<String> trans = new ArrayList<String>();
				List<String> valueList=metadataMap.get(key);
				for(String s:valueList){
					if(s.equals("unfiled$public")){
						valueList.remove(s);
						trans.add(Util.getTranslate("FOLDERNAME","UNFILEDPUBLIC"));
						metaDataLabelMap.put(Util.getTranslate("FOLDERNAME","UNFILEDPUBLIC"), s);
						break;
					}
				}
				List<Metadata> mdInfos = ut.readMateData(key+"Folder", metadataMap.get(key));
				for (Metadata md2 : mdInfos) {
					Folder fd=(Folder)md2;
					if(fd.getName()!=null){
						trans.add(fd.getName());
						metaDataLabelMap.put(fd.getName(), fd.getFullName());
					}else{
						trans.add(fd.getFullName());
						metaDataLabelMap.put(fd.getFullName(), fd.getFullName());
					}
				}
				if(mdInfos!=null&&mdInfos.size()>0){
					metadataMap.put(key,trans);
				}
			}
		}
		
		List<String> valueList=metadataMap.get("ApprovalProcess");
		if(valueList!=null&&valueList.size()>0){
			List<String> trans = new ArrayList<String>();
			List<Metadata> mdInfos = ut.readMateData("ApprovalProcess", valueList);
			for (Metadata md2 : mdInfos) {
				if (md2 != null) {
					ApprovalProcess obj = (ApprovalProcess) md2;
					if(obj.getLabel()!=null){
						trans.add(obj.getLabel());
						metaDataLabelMap.put(obj.getLabel(),obj.getFullName());
					}else{
						trans.add(obj.getFullName());
						metaDataLabelMap.put(obj.getFullName(),obj.getFullName());
					}
				}
			}
			metadataMap.put("ApprovalProcess",trans);
		}
		
		List<String> valueList2=metadataMap.get("CustomObjectTranslation");
		if(valueList2!=null&&valueList2.size()>0){
			List<String> trans = new ArrayList<String>();
			for (String str : valueList2) {
				if (str != null) {
					String[] strList=str.split("-",0);
					if(strList.length==2){
						String s=ut.getLabelforAll(strList[0])+"-"+Util.getTranslate("TRANSLATIONLANGUAGE", strList[1]);
						trans.add(s);
						metaDataLabelMap.put(s,str);
					}else{
						trans.add(str);
						metaDataLabelMap.put(str,str);
					}
				}
			}
			metadataMap.put("CustomObjectTranslation",trans);
		}
						
		List<String> valueList4=metadataMap.get("Network");
		if(valueList4!=null&&valueList4.size()>0){
			List<String> trans = new ArrayList<String>();
			List<Metadata> mdInfos = ut.readMateData("Network", valueList4);
			for (Metadata md2 : mdInfos) {
				if (md2 != null) {
					Network obj = (Network) md2;
					trans.add(obj.getFullName());
					metaDataLabelMap.put(obj.getFullName(),obj.getFullName());					
				}
			}
			metadataMap.put("Network",trans);
		}
		//add by dan 2017-9-14 start
		/*
		List<String> brandingList=metadataMap.get("NetworkBranding");
		System.out.println("brandingList-------"+brandingList);
		List<String> brandingListNew = new ArrayList<String>();
		if(brandingList!=null&&brandingList.size()>0){
			for(String brand:brandingList){
				brand = brand.substring(2);
				if(brand.contains("_")){
					brand = brand.replace('_', ' ');
				}
				brand = brand.toUpperCase();
				brandingListNew.add(brand);
			}
			System.out.println("brandingListNew-------"+brandingListNew);
			List<Metadata> mdInfosbrand = ut.readMateData("NetworkBranding", brandingListNew);
			System.out.println("mdInfosbrand-------"+mdInfosbrand);
		}*/
		//add by dan 2017-9-14 end
		for(String key: metadataMap.keySet()){
			if(translatableType.contains(key)){
				List<String> trans = new ArrayList<String>();
				for(String value : metadataMap.get(key)){
					String s=ut.getLabelApi(value);
					trans.add(s);
					metaDataLabelMap.put(s, value);
				}
				metadataMap.put(key,trans);
			}
		}
		metadataMap = mapSortByKey(metadataMap);
		return metadataMap;

	}
	List<String> translateListType = Arrays.asList("Report","Dashboard","Document");
	List<String> translateLabelType = Arrays.asList("ApprovalProcess");
	//フォルダを全部取得方法
	List<String> translatableType = Arrays.asList("Workflow","CustomObject","StandardObject","SharingSetting","CustomSetting","ArticleType","EXTERNALOBJECT","CustomTab");
	public Map<String,String> translateMap=new HashMap<String,String>();
	private List<String> FolderHelper(String string){
		List<String> allFolders  = new ArrayList<String>();
		try{			
			ListMetadataQuery queryReportF = new ListMetadataQuery();
			queryReportF.setType(string+"Folder");
			FileProperties[] rfp = metadataConnection.listMetadata(
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
		return allFolders;
	}
	
	//取得したMetadataをソートする方法
	private SortedMap<String, List<String>> mapSortByKey(
			Map<String, List<String>> unsort_map) {
		TreeMap<String, List<String>> result = new TreeMap<String, List<String>>();
		Object[] unsort_key = unsort_map.keySet().toArray();
		Arrays.sort(unsort_key);
		for (int i = 0; i < unsort_key.length; i++) {		
			//result.put(unsort_key[i].toString(), unsort_map.get(unsort_key[i]));
			String metaDataLabel=Util.getTranslate("MetaDataType",unsort_key[i].toString());
			if(metaDataLabel==null){
				metaDataLabel=unsort_key[i].toString();
			}

			result.put(metaDataLabel, unsort_map.get(unsort_key[i]));
			
		}
		unsort_key = result.keySet().toArray();				
		return result.tailMap(result.firstKey());
	}
	
	//取得したMetadataを画面パネルに表示する方法
	public JPanel listMetadata() {
		JPanel panel = new JPanel();// create panel
		panel.setLayout(new GridLayout(0, 1));//set gridlayout manager
		panel.setBackground(Color.WHITE);
		panel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
		panel.setPreferredSize(new Dimension(320, 240));
		DefaultMutableTreeNode root = new DefaultMutableTreeNode(Util.getTranslate("MetaDataType","COMPONENTS"));
		Object obj = root.getUserObject();
		root.setUserObject(new CheckBoxNode((String) obj, Status.DESELECTED));
		DefaultTreeModel model = new DefaultTreeModel(root);
		model.addTreeModelListener(new CheckBoxStatusUpdateListener());
		CheckBoxStatusUpdateListener cl = new CheckBoxStatusUpdateListener();
		try {
			Map<String, List<String>> typeMap = getCustomSetting();
			Map<String, List<String>> metadataMap = allMetadataMap;
			for (String typeStr : metadataMap.keySet()) {
				List<String> typeList = new ArrayList<String>();
				if (typeMap.keySet().contains(typeStr)) {
					typeList = typeMap.get(typeStr);
				}
				DefaultMutableTreeNode type = new DefaultMutableTreeNode(
						typeStr);
				Object ob = type.getUserObject();
				if (typeMap.keySet().contains(typeStr)
						&& typeMap.get(typeStr).size() == 0) {
					type.setUserObject(new CheckBoxNode((String) ob,
							Status.SELECTED));
				} else {
					type.setUserObject(new CheckBoxNode((String) ob,
							Status.DESELECTED));
				}
				List<String> allFile = metadataMap.get(typeStr);
				for (String n : allFile) {
					String labelN=mapLabelBack(n);
					DefaultMutableTreeNode node = new DefaultMutableTreeNode(n);
					Object o = node.getUserObject();
					if (typeList.contains(labelN)) {
						node.setUserObject(new CheckBoxNode((String) o,
								Status.SELECTED));
					} else {
						node.setUserObject(new CheckBoxNode((String) o,
								Status.DESELECTED));
					}
					type.add(node);
					cl.updateParentUserObject(type);
				}
				root.add(type);
				cl.updateParentUserObject(root);
			}

			tree = new JTree(model) {
				@Override
				public void updateUI() {
					setCellRenderer(null);
					setCellEditor(null);
					super.updateUI();
					setCellRenderer(new CheckBoxNodeRenderer());
					setCellEditor(new CheckBoxNodeEditor());
				}
			};

			tree.setEditable(true);
			tree.setBorder(BorderFactory.createEmptyBorder(4, 4, 4, 4));
			tree.expandRow(0);
			panel.add(tree);
			panel.add(new JScrollPane(tree));
		} catch (Exception e) {
			e.getStackTrace();
		}
		return panel;
	}

	//選択したMetadataを取得（Ｍａｐに出す）
	private Map<String, List<String>> getSelectedNode(JTree tree, boolean isOk) {
		HashMap<String, List<String>> typeMap = new LinkedHashMap<String, List<String>>();
		Map<String, List<String>> typekeyMap = new HashMap<String, List<String>>();
		TreeModel model = tree.getModel();
		DefaultMutableTreeNode root = (DefaultMutableTreeNode) model.getRoot();
		Enumeration e = root.children();
		while (e.hasMoreElements()) {
			List<String> members = new ArrayList<String>();
			DefaultMutableTreeNode type = (DefaultMutableTreeNode) e
					.nextElement();
			CheckBoxNode typecheck = (CheckBoxNode) type.getUserObject();
			Enumeration en = type.children();
			while (en.hasMoreElements()) {
				DefaultMutableTreeNode node = (DefaultMutableTreeNode) en.nextElement();
				CheckBoxNode check = (CheckBoxNode) node.getUserObject();
				if (check.status == Status.SELECTED) {
					members.add(mapLabelBack(check.label));
				}
			}
			if (typecheck.status == Status.SELECTED || members.size() > 0) {

				typekeyMap.put(((CheckBoxNode) type.getUserObject()).label,
						members);
				}
		}
		if(!typekeyMap.isEmpty()){
		typekeyMap = mapSortByKey(typekeyMap);
		Object[] keylist=typekeyMap.keySet().toArray();
		for(int i=0;i<keylist.length;i++){
			String key=keylist[i].toString();
			typeMap.put(mapLabelBack(key),typekeyMap.get(key));
		}
		}
		if (isOk) {
			try {
				WriteXML wx = new WriteXML(typeMap, Util.API_VERSION, lastRunTime);
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
		return typeMap;
	}

	//前回選択した選択値を再設定用方法
	public Map<String, List<String>> getCustomSetting() throws Exception {

		// Edit the path, if necessary, if your package.xml file is located
		// elsewhere
		File file;
		if(WSC.isBatch){
			file = new File(".././conf/package.xml");
		}else{
			file = new File("conf/package.xml");
		}
		Map<String, List<String>> typeMap = new HashMap<String, List<String>>();
		if (!file.exists() || !file.isFile()) {
			SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			lastRunTime = df.format(new Date());
			return typeMap;
		}
		DocumentBuilder db = DocumentBuilderFactory.newInstance()
				.newDocumentBuilder();
		InputStream inputStream = new FileInputStream(file);
		Element d = db.parse(inputStream).getDocumentElement();
		for (Node c = d.getFirstChild(); c != null; c = c.getNextSibling()) {
			if (c instanceof Element) {
				if (c.getNodeName() == "lastRunTime") {
					getLastRunTime(c.getTextContent());

				}
				Element ce = (Element) c;
				NodeList nodeList = ce.getElementsByTagName("name");
				if (nodeList.getLength() == 0) {
					continue;
				}
				String name = nodeList.item(0).getTextContent();
				NodeList m = ce.getElementsByTagName("members");
				List<String> members = new ArrayList<String>();
				for (int i = 0; i < m.getLength(); i++) {
					Node mm = m.item(i);
					if(name.equals("StandardObject")){
						members.add(mm.getTextContent());
					}
					else{
						if(WSC.isBatch){
							members.add(mm.getTextContent());
						}else{
							members.add(mm.getTextContent());
							//members.add(Util.getTranslate("MetaDataType",mm.getTextContent()));
						}
					}
				}				
				if(WSC.isBatch){
					typeMap.put(name,members);
				}else{
					typeMap.put(Util.getTranslate("MetaDataType",name),members);
				}
			}
		}
		return typeMap;
	}
	//metadata to display
	private String[] useableMetadata = {"AnalyticSnapshot",
										"ApexClass",
										"ApexComponent",
										"ApexPage",
										"ApexTrigger",
										"ApprovalProcess",
										"AssignmentRules",
										"AuraDefinitionBundle",
										"AutoResponseRules",
										"CustomLabels",
										"CustomObject",
										"CustomObjectTranslation",
										"CustomTab",
										"Dashboard",
										"Document",
										"EscalationRules",
										"Layout",
										"Profile",
										"Report",
										"ReportType",
										"PermissionSet",
										"Settings",
										"StaticResource",
										"Workflow",
										"SharingRules",
										"ExternalDataSource",
										"EntitlementProcess",
										"Network"
										
	};
	
	//表示必要のmetadataかないかの判断用方法
	public boolean isUseable(String xmlName){
		for (String name : useableMetadata) {
			if(name.equals(xmlName)){
				return true;
			}
		}
		return false;		
	}
	//前回のエクスポート時間を返す方法
	private void getLastRunTime(String time) {
		lastRunTime = time;
	}
	//ラベルからmetadataオブジェクトを返す方法
	private String mapLabelBack(String label) {
		String retStr;
		retStr = (String) metaDataLabelMap.get(label);
		if(retStr==null){
			retStr=label;
		}
		return retStr;
	}		
}

class TriStateCheckBox extends JCheckBox {
	private Icon currentIcon;

	@Override
	public void updateUI() {
		currentIcon = getIcon();
		setIcon(null);
		super.updateUI();
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				if (currentIcon != null) {
					setIcon(new IndeterminateIcon());
				}
				setOpaque(false);
			}
		});
	}
}

class IndeterminateIcon implements Icon {
	private static final Color FOREGROUND = new Color(50, 20, 255, 200); // TEST:
																			// UIManager.getColor("CheckBox.foreground");
	private static final int SIDE_MARGIN = 4;
	private static final int HEIGHT = 4;
	private final Icon icon = UIManager.getIcon("CheckBox.icon");

	@Override
	public void paintIcon(Component c, Graphics g, int x, int y) {
		icon.paintIcon(c, g, x, y);
		int w = getIconWidth();
		int h = getIconHeight();
		Graphics2D g2 = (Graphics2D) g.create();
		g2.setPaint(FOREGROUND);
		g2.translate(x, y);
		g2.fillRect(SIDE_MARGIN, HEIGHT, w - SIDE_MARGIN - SIDE_MARGIN, HEIGHT);
		g2.dispose();
	}

	@Override
	public int getIconWidth() {
		return icon.getIconWidth();
	}

	@Override
	public int getIconHeight() {
		return icon.getIconHeight();
	}
}

enum Status {
	SELECTED, DESELECTED, INDETERMINATE
}

class CheckBoxNode {
	public final String label;
	public final Status status;

	public CheckBoxNode(String label) {
		this.label = label;
		status = Status.INDETERMINATE;
	}

	public CheckBoxNode(String label, Status status) {
		this.label = label;
		this.status = status;
	}

	@Override
	public String toString() {
		return label;
	}
}

class CheckBoxStatusUpdateListener implements TreeModelListener {
	private boolean adjusting;

	@Override
	public void treeNodesChanged(TreeModelEvent e) {
		if (adjusting) {
			return;
		}
		adjusting = true;
		TreePath parent = e.getTreePath();
		Object[] children = e.getChildren();
		DefaultTreeModel model = (DefaultTreeModel) e.getSource();
		DefaultMutableTreeNode node;
		CheckBoxNode c;
		if (children != null && children.length == 1) {
			node = (DefaultMutableTreeNode) children[0];
			c = (CheckBoxNode) node.getUserObject();
			DefaultMutableTreeNode n = (DefaultMutableTreeNode) parent
					.getLastPathComponent();
			while (n != null) {
				updateParentUserObject(n);
				DefaultMutableTreeNode tmp = (DefaultMutableTreeNode) n
						.getParent();
				if (tmp == null) {
					break;
				} else {
					n = tmp;
				}
			}
			model.nodeChanged(n);
		} else {
			node = (DefaultMutableTreeNode) model.getRoot();
			c = (CheckBoxNode) node.getUserObject();
		}
		updateAllChildrenUserObject(node, c.status);
		model.nodeChanged(node);
		adjusting = false;
	}

	public void updateParentUserObject(DefaultMutableTreeNode parent) {
		String label = ((CheckBoxNode) parent.getUserObject()).label;
		int selectedCount = 0;
		int indeterminateCount = 0;
		Enumeration children = parent.children();
		while (children.hasMoreElements()) {
			DefaultMutableTreeNode node = (DefaultMutableTreeNode) children
					.nextElement();
			CheckBoxNode check = (CheckBoxNode) node.getUserObject();
			if (check.status == Status.INDETERMINATE) {
				indeterminateCount++;
				break;
			}
			if (check.status == Status.SELECTED) {
				selectedCount++;
			}
		}
		if (indeterminateCount > 0) {
			parent.setUserObject(new CheckBoxNode(label));
		} else if (selectedCount == 0) {
			parent.setUserObject(new CheckBoxNode(label, Status.DESELECTED));
		} else if (selectedCount == parent.getChildCount()) {
			parent.setUserObject(new CheckBoxNode(label, Status.SELECTED));
		} else {
			parent.setUserObject(new CheckBoxNode(label));
		}
	}

	public void updateAllChildrenUserObject(DefaultMutableTreeNode root,
			Status status) {
		Enumeration breadth = root.breadthFirstEnumeration();
		while (breadth.hasMoreElements()) {
			DefaultMutableTreeNode node = (DefaultMutableTreeNode) breadth
					.nextElement();
			if (Objects.equals(root, node)) {
				continue;
			}
			CheckBoxNode check = (CheckBoxNode) node.getUserObject();
			node.setUserObject(new CheckBoxNode(check.label, status));
		}
	}

	@Override
	public void treeNodesInserted(TreeModelEvent e) { /* not needed */
	}

	@Override
	public void treeNodesRemoved(TreeModelEvent e) { /* not needed */
	}

	@Override
	public void treeStructureChanged(TreeModelEvent e) { /* not needed */
	}

}

// extends JCheckBox TreeCellRenderer Editor version
class CheckBoxNodeRenderer extends TriStateCheckBox implements TreeCellRenderer {
	private final DefaultTreeCellRenderer renderer = new DefaultTreeCellRenderer();
	private final JPanel panel = new JPanel(new BorderLayout());

	public CheckBoxNodeRenderer() {
		super();
		String uiName = getUI().getClass().getName();
		if (uiName.contains("Synth")
				&& System.getProperty("java.version").startsWith("1.7.0")) {
			renderer.setBackgroundSelectionColor(new Color(0, 0, 0, 0));
		}
		panel.setFocusable(false);
		panel.setRequestFocusEnabled(false);
		panel.setOpaque(false);
		panel.add(this, BorderLayout.WEST);
		this.setOpaque(false);
	}

	@Override
	public Component getTreeCellRendererComponent(JTree tree, Object value,
			boolean selected, boolean expanded, boolean leaf, int row,
			boolean hasFocus) {
		JLabel l = (JLabel) renderer.getTreeCellRendererComponent(tree, value,
				selected, expanded, leaf, row, hasFocus);
		l.setFont(tree.getFont());
		if (value instanceof DefaultMutableTreeNode) {
			this.setEnabled(tree.isEnabled());
			this.setFont(tree.getFont());
			DefaultMutableTreeNode no = (DefaultMutableTreeNode) value;
			if (no.getLevel() == 1 || no.getLevel() == 0) {
				if(WSC.isBatch){
					renderer.setIcon(new
							ImageIcon(SelectMetadata.class.getResource("/common/icons/folder.png")));					
				}else{
					renderer.setIcon(new ImageIcon("common/icons/folder.png"));
				}

			} else {
				if(WSC.isBatch){
					renderer.setIcon(new
							ImageIcon(SelectMetadata.class.getResource("/common/icons/file.png")));					
				}else{
					renderer.setIcon(new ImageIcon("common/icons/file.png"));
				}				
			}

			Object userObject = no.getUserObject();
			if (userObject instanceof CheckBoxNode) {
				CheckBoxNode node = (CheckBoxNode) userObject;
				if (node.status == Status.INDETERMINATE) {
					setIcon(new IndeterminateIcon());
				} else {
					setIcon(null);
				}
				l.setText(node.label);
				setSelected(node.status == Status.SELECTED);
			}
			panel.add(l);
			return panel;
		}
		return l;
	}

	@Override
	public void updateUI() {
		super.updateUI();
		if (panel != null) {
			panel.updateUI();
		}
		setName("Tree.cellRenderer");
	}
}

class CheckBoxNodeEditor extends TriStateCheckBox implements TreeCellEditor {
	private final DefaultTreeCellRenderer renderer = new DefaultTreeCellRenderer();
	private final JPanel panel = new JPanel(new BorderLayout());
	private String str;

	public CheckBoxNodeEditor() {
		super();
		this.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				stopCellEditing();
			}
		});
		panel.setFocusable(false);
		panel.setRequestFocusEnabled(false);
		panel.setOpaque(false);
		panel.add(this, BorderLayout.WEST);
		this.setOpaque(false);
	}

	@Override
	public Component getTreeCellEditorComponent(JTree tree, Object value,
			boolean isSelected, boolean expanded, boolean leaf, int row) {
		JLabel l = (JLabel) renderer.getTreeCellRendererComponent(tree, value,
				true, expanded, leaf, row, true);
		l.setFont(tree.getFont());

		if (value instanceof DefaultMutableTreeNode) {
			this.setEnabled(tree.isEnabled());
			this.setFont(tree.getFont());
			DefaultMutableTreeNode no = (DefaultMutableTreeNode) value;
			if (no.getLevel() == 1 || no.getLevel() == 0) {
				if(WSC.isBatch){
					 renderer.setIcon(new
								ImageIcon(SelectMetadata.class.getResource("/common/icons/folder.png")));					
				}else{
					renderer.setIcon(new ImageIcon("common/icons/folder.png"));
				}

			} else {
				if(WSC.isBatch){
					renderer.setIcon(new
							ImageIcon(SelectMetadata.class.getResource("/common/icons/file.png")));				
				}else{
					renderer.setIcon(new ImageIcon("common/icons/file.png"));
				}				
			}
			Object userObject = no.getUserObject();
			if (userObject instanceof CheckBoxNode) {
				CheckBoxNode node = (CheckBoxNode) userObject;
				if (node.status == Status.INDETERMINATE) {
					setIcon(new IndeterminateIcon());
				} else {
					setIcon(null);
				}
				l.setText(node.label);
				setSelected(node.status == Status.SELECTED);
				str = node.label;
			}
			// panel.add(this, BorderLayout.WEST);
			panel.add(l);
			return panel;
		}
		return l;
	}

	@Override
	public Object getCellEditorValue() {
		return new CheckBoxNode(str, isSelected() ? Status.SELECTED
				: Status.DESELECTED);
	}

	@Override
	public boolean isCellEditable(EventObject e) {
		if (e instanceof MouseEvent && e.getSource() instanceof JTree) {
			MouseEvent me = (MouseEvent) e;
			JTree tree = (JTree) e.getSource();
			TreePath path = tree.getPathForLocation(me.getX(), me.getY());
			Rectangle r = tree.getPathBounds(path);
			if (r == null) {
				return false;
			}
			Dimension d = getPreferredSize();
			r.setSize(new Dimension(d.width, r.height));
			if (r.contains(me.getX(), me.getY())) {
				if (str == null
						&& System.getProperty("java.version").startsWith(
								"1.7.0")) {
					setBounds(new Rectangle(0, 0, d.width, r.height));
				}
				return true;
			}
		}
		return false;
	}

	@Override
	public void updateUI() {
		super.updateUI();
		setName("Tree.cellEditor");
		if (panel != null) {
			panel.updateUI();
		}
	}

	@Override
	public boolean shouldSelectCell(EventObject anEvent) {
		return true;
	}

	@Override
	public boolean stopCellEditing() {
		fireEditingStopped();
		return true;
	}

	@Override
	public void cancelCellEditing() {
		fireEditingCanceled();
	}

	@Override
	public void addCellEditorListener(CellEditorListener l) {
		listenerList.add(CellEditorListener.class, l);
	}

	@Override
	public void removeCellEditorListener(CellEditorListener l) {
		listenerList.remove(CellEditorListener.class, l);
	}

	public CellEditorListener[] getCellEditorListeners() {
		return listenerList.getListeners(CellEditorListener.class);
	}

	protected void fireEditingStopped() {
		// Guaranteed to return a non-null array
		Object[] listeners = listenerList.getListenerList();
		// Process the listeners last to first, notifying
		// those that are interested in this event
		for (int i = listeners.length - 2; i >= 0; i -= 2) {
			if (listeners[i] == CellEditorListener.class) {
				// Lazily create the event:
				if (changeEvent == null) {
					changeEvent = new ChangeEvent(this);
				}
				((CellEditorListener) listeners[i + 1])
						.editingStopped(changeEvent);
			}
		}
	}

	protected void fireEditingCanceled() {
		// Guaranteed to return a non-null array
		Object[] listeners = listenerList.getListenerList();
		// Process the listeners last to first, notifying
		// those that are interested in this event
		for (int i = listeners.length - 2; i >= 0; i -= 2) {
			if (listeners[i] == CellEditorListener.class) {
				// Lazily create the event:
				if (changeEvent == null) {
					changeEvent = new ChangeEvent(this);
				}
				((CellEditorListener) listeners[i + 1])
						.editingCanceled(changeEvent);
			}
		}
	}


}