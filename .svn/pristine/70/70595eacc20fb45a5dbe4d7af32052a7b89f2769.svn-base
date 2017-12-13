package source;

import java.io.IOException;
import java.net.URLDecoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;










import com.sforce.ws.bind.XmlObject;










import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import wsc.MetadataLoginUtil;

import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Queue;
import com.sforce.soap.metadata.QueueSobject;
import com.sforce.soap.metadata.Role;
import com.sforce.soap.metadata.CustomField;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class ReadGroupRoleQueueSync {

	private XSSFWorkbook workBook;
	public void readGroupRoleQueue(String type,List<String> objectsList)throws Exception{
		Util.logger.info("ReadGroupRoleQueueSync started");	
		Util.logger.debug("objectsList="+objectsList);			
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		Util.nameSequence=0;
		Util.sheetSequence=0;
		//create catelog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();	
		/***Create Group Role Queue Sheets***/
		String GroupSheetName ="";
		if(objectsList.contains("Group")){
			GroupSheetName=Util.makeSheetName("Group");
		}		
		XSSFSheet excelGroupSheet;	
		
		String QueueSheetName ="";
		if(objectsList.contains("Queue")){
			QueueSheetName=Util.makeSheetName("Queue");
		}
		XSSFSheet excelQueueSheet;	
		
		String UserSheetName ="";
		if(objectsList.contains("User")){
			UserSheetName=Util.makeSheetName("User");
		}
		XSSFSheet excelSheet;		
		
		String RoleSheetName="";
		if(objectsList.contains("Role")){
			RoleSheetName=Util.makeSheetName("Role");
		}
		XSSFSheet excelRoleSheet;

		
		Util ut = new Util();
		//make link map
		Map<String,String> roleNameMap = new HashMap<String,String>();
		//make user name map for hyperlink use
		Map<String,String> userNameMap = new HashMap<String,String>();
		//make group name map for hyperlink use
		Map<String,String> groupNameMap = new HashMap<String,String>();
		
		List<String> nameList = new ArrayList<String>();
		List<String> objList=new ArrayList<String>();
		objList.add("User");
		List<Metadata> mInfo = ut.readMateData("CustomObject", objList);
		for (Metadata md : mInfo) {
			CustomObject co=(CustomObject)md;
			if(co.getFields().length>0){					
				for( Integer i=0; i<co.getFields().length; i++ ){
					CustomField cf = (CustomField)co.getFields()[i];
					if(cf.getFullName().contains("__c"))
					nameList.add(cf.getFullName());
				}
			}
		}
		if(nameList.size()>0){
		nameList.remove(0);
		}
		String userSoql ="Select Id,Username,Name,UserType,IsActive,LastName,Manager.Name,FirstName,LastModifiedDate,UserRole.DeveloperName,Profile.Name,LocaleSidKey,LanguageLocaleKey,Phone,FAX,ForecastEnabled,Department,Title,Email,ManagerId,";
		for(String namestr:nameList){
			userSoql+=namestr+",";
		}
		userSoql+="EmailEncodingKey From User Order by Id";
		SObject [] userObjects= ut.apiQuery(userSoql);
		for(SObject obj : userObjects){
			String userCellName= Util.makeNameValue(obj.getId());
			userNameMap.put(obj.getId(),userCellName);
		}
		
		for(String str : objectsList){
			if(str.equals("Group")){	
				Util.logger.info("Group started");	
				//Create Group Sheet
				excelGroupSheet= excelTemplate.createSheet(Util.cutSheetName(GroupSheetName));
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelGroupSheet,Util.cutSheetName(GroupSheetName),GroupSheetName);	
				ListMetadataQuery query = new ListMetadataQuery();
				query.setType(str);
				List<String> allFile = new ArrayList<String>();
				FileProperties[] lmr = MetadataLoginUtil.metadataConnection.listMetadata(
						new ListMetadataQuery[] { query }, Util.API_VERSION);
				if (lmr != null) {
					for (FileProperties n : lmr) {
						allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));
					}
				}				
				Map<String,String> resultMap = ut.getComparedResult(str,UtilConnectionInfc.getLastUpdateTime());
				

				if(allFile.size()>0){
					String names = ut.getObjectNames(allFile);	
					Util.logger.debug("names="+names);	
					//Create Group query String
					String sql = "Select Name, Id, DeveloperName, DoesIncludeBosses From Group WHERE DeveloperName in ("+ names +") Order By Id";				
					SObject [] groupObjects= ut.apiQuery(sql);

					for(SObject obj : groupObjects){
						groupNameMap.put(Util.nullFilter(obj.getField("Name")),Util.makeNameValue(Util.nullFilter(obj.getField("Name"))));
					}
					//Create Group Table
					excelTemplate.createTableHeaders(excelGroupSheet,"Group",excelGroupSheet.getLastRowNum()+Util.RowIntervalNum);	
					
					for(SObject obj : groupObjects){	
						Util.logger.debug("groupObject="+obj);	
						if(obj.getField("Name")!=null){
							Util.logger.debug("groupObject id="+obj.getField("Id"));	
							//Create groupRow
							XSSFRow groupRow = excelGroupSheet.createRow(excelGroupSheet.getLastRowNum()+1);	
							
							//Make sure hyper link name start with a letter or underscore and not contain spaces						
							if(groupNameMap.get(Util.nullFilter(obj.getField("Name")))!=null){
								excelTemplate.createCellName(groupNameMap.get(Util.nullFilter(obj.getField("Name"))),GroupSheetName,excelGroupSheet.getLastRowNum()+1);	
							}
							Integer cellNum = 1;
							if(UtilConnectionInfc.modifiedFlag){
								//変更あり
								excelTemplate.createCell(groupRow,cellNum++,ut.getUpdateFlag(resultMap,"Group."+obj.getField("DeveloperName")));
														
							}
							//グループ名
							excelTemplate.createCell(groupRow,cellNum++,Util.nullFilter(obj.getField("DeveloperName")));
							//表示ラベル
							excelTemplate.createCell(groupRow,cellNum++,Util.nullFilter(obj.getField("Name")));
							//階層を使用したアクセス許可
							excelTemplate.createCell(groupRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getField("DoesIncludeBosses"))));					
							
							String sqlmember = "Select UserOrGroupId From GroupMember Where GroupId='"+Util.nullFilter(obj.getField("Id"))+"'Order By Id";					
							SObject [] memberObjects= ut.apiQuery(sqlmember);	
	
							List<String> glist = new ArrayList<String>();
							List<String> ulist = new ArrayList<String>();
	
							for(SObject mobj : memberObjects ){	
								String str1 = Util.nullFilter(mobj.getField("UserOrGroupId"));
								String str2 =str1.substring(0, 3);
								if(str2.equals("005")){
									ulist.add(str1);
								}
								else if(str2.equals("00G")){
									glist.add(str1);
								}
							}					
							if(ulist.size()>0){
								String uId ="";
								for(String Id : ulist){
									uId += "'"+Id +"'"+",";
								}
								uId = uId.substring(0,uId.length()-1);	
								String usersql = "Select id,Username,LastName,FirstName From User Where Id in ("+ uId +") Order By Id";
								SObject [] userObj= ut.apiQuery(usersql);
								Integer rowNu = excelGroupSheet.getLastRowNum();	
								for(SObject uobj : userObj){	
									String hyperVal = userNameMap.get(Util.nullFilter(uobj.getId()));
									String displayVal = Util.getTranslate("GroupMemberType","User")+":"+Util.nullFilter(uobj.getField("LastName"))+" "+Util.nullFilter(uobj.getField("FirstName"));											
									if(excelGroupSheet.getRow(rowNu)!=null){
										//メンバ�E
										excelTemplate.createCellValue(excelGroupSheet,rowNu, cellNum ,hyperVal,displayVal);
									}else{
										XSSFRow itemRow = excelGroupSheet.createRow(rowNu);	
										excelTemplate.createCell(itemRow,1,"");
										excelTemplate.createCell(itemRow,2,"");
										excelTemplate.createCell(itemRow,3,"");
										if(UtilConnectionInfc.modifiedFlag){
											excelTemplate.createCell(itemRow,4,"");											
										}
										excelTemplate.createCellValue(excelGroupSheet,rowNu, cellNum ,hyperVal,displayVal);
									}																			
									rowNu=rowNu+1;
								}			
							}
							if(glist.size()>0){
								String gId ="";
								for(String Id : glist){
									gId += "'"+Id +"'" + ",";
								}		
								gId = gId.substring(0,gId.length()-1);						
								String grosql = "Select Type,Name,DeveloperName From Group Where Id in ("+ gId +") Order By Id";
								SObject [] groupObj= ut.apiQuery(grosql);
								Integer rowNo = excelGroupSheet.getLastRowNum()+1;
								for(SObject gobj : groupObj){
									if(gobj.getField("Type").equals("Regular")){	
										Util.logger.info("it is a Regular group");

										String hyperVal = groupNameMap.get(Util.nullFilter(gobj.getField("Name")));
										String displayVal = Util.getTranslate("GroupMemberType","GROUP")+":"+Util.nullFilter(gobj.getField("Name"));											
										if(excelGroupSheet.getRow(rowNo)!=null){
											excelTemplate.createCellValue(excelGroupSheet,rowNo, cellNum ,hyperVal,displayVal);
										}else{
											excelGroupSheet.createRow(rowNo);
											//for cell style
											XSSFRow itemRow = excelGroupSheet.getRow(rowNo);
											excelTemplate.createCell(itemRow,1,"");
											excelTemplate.createCell(itemRow,2,"");
											excelTemplate.createCell(itemRow,3,"");
											if(UtilConnectionInfc.modifiedFlag){
												excelTemplate.createCell(itemRow,4,"");											
											}
											excelTemplate.createCellValue(excelGroupSheet,rowNo, cellNum,hyperVal,displayVal);
										}
									}
									else if(gobj.getField("Type").equals("Role") || gobj.getField("Type").equals("RoleAndSubordinates")){									
										Util.logger.debug("it is a Role or RoleAndSubordinates.");
										String hyperVal ="";
										int leng = Util.nullFilter(gobj.getField("DeveloperName")).length();
										String Name = Util.nullFilter(gobj.getField("DeveloperName")).substring(0,leng-1);
										String cellName = Util.makeNameValue(gobj.getField("DeveloperName").toString());
										if(roleNameMap.get(Name)==null){	
											int len = cellName.length();
										    hyperVal = cellName.substring(0,len-1);
										}else{
											hyperVal = roleNameMap.get(Name);
										}	
										roleNameMap.put(Name, hyperVal);
										String displayVal="";
										if(gobj.getField("Type").equals("Role")){
											displayVal=Util.getTranslate("GroupMemberType","ROLE")+":"+Name;	
										}else{
											displayVal=Util.getTranslate("GroupMemberType","RoleAndSubordinates")+":"+Name;	
										}
																		
										if(excelGroupSheet.getRow(rowNo)!=null){
											excelTemplate.createCellValue(excelGroupSheet,rowNo, cellNum ,hyperVal,displayVal);
										}else{											
											excelGroupSheet.createRow(rowNo);
											//for cell style
											XSSFRow itemRow = excelGroupSheet.getRow(rowNo);
											excelTemplate.createCell(itemRow,1,"");
											excelTemplate.createCell(itemRow,2,"");
											excelTemplate.createCell(itemRow,3,"");
											if(UtilConnectionInfc.modifiedFlag){
												excelTemplate.createCell(itemRow,4,"");											
											}
											excelTemplate.createCellValue(excelGroupSheet,rowNo, cellNum ,hyperVal,displayVal);
										}
									}							
									rowNo=rowNo+1;
								}
	
							}
							if(glist.size()==0&&ulist.size()==0){
								excelTemplate.createCell(groupRow,cellNum++,"");
							}
						}					
					}
				}else{
					//Create Group Table
					excelTemplate.createTableHeaders(excelGroupSheet,"Group",excelGroupSheet.getLastRowNum()+Util.RowIntervalNum);	
				}
				excelTemplate.adjustColumnWidth(excelGroupSheet);
				Util.logger.info("Group completed");					
			}	
			if(str.equals("User")){
				//Create Sheets	
				excelSheet= excelTemplate.createSheet(Util.cutSheetName(UserSheetName));
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelSheet,Util.cutSheetName(UserSheetName),UserSheetName);

				//Create Catalog menu
				excelTemplate.createTableHeaders(excelSheet,"User",excelSheet.getLastRowNum()+Util.RowIntervalNum);
				XSSFRow headerRow = excelSheet.getRow(excelSheet.getLastRowNum());
				for(int j=0;j<nameList.size();j++){
					Integer cellNum = Integer.valueOf(headerRow.getLastCellNum());
					XSSFCell cell = headerRow.createCell(cellNum);
					cell.setCellValue(ut.getLabelApi("User."+nameList.get(j)));
					cell.setCellStyle(excelTemplate.createCHeaderStyle());
				}	

				for(SObject obj : userObjects){
					//Create columnRow
					Integer colNo=1;
					XSSFRow userRow = excelSheet.createRow(excelSheet.getLastRowNum()+1);

					if(UtilConnectionInfc.modifiedFlag){
						SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
						Date d = new Date();
						try {
							d = df.parse(ut.getLocalTime(Util.nullFilter(obj.getField("LastModifiedDate"))));
						} catch (ParseException e) {
							e.printStackTrace();
						}
						//変更あり
						if (UtilConnectionInfc.getLastUpdateTime()>(d.getTime())){
							excelTemplate.createCell(userRow,colNo++,Util.getTranslate("IsChanged","false"));
						}else{
							excelTemplate.createCell(userRow,colNo++,Util.getTranslate("IsChanged","true"));
						}			
					}
					excelTemplate.createCellName(userNameMap.get(obj.getId()),UserSheetName,excelSheet.getLastRowNum()+1);	
					//ユーザ吁E
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(obj.getField("Username")));
					//ユーザ種別
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(obj.getField("UserType")));
					//有効
					excelTemplate.createCell(userRow,colNo++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getField("IsActive"))));
					//姁E
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(obj.getField("LastName")));
					//吁E
					excelTemplate.createCell(userRow,colNo++,Util.stringFilter(Util.nullFilter(obj.getField("FirstName"))));
					if(obj.getField("UserRole")!=null){
						XmlObject xmlobjrole =(XmlObject) obj.getField("UserRole");
						Integer rowNo = excelSheet.getLastRowNum();
						String hyperVal ="";
						if(roleNameMap.get(xmlobjrole.getField("DeveloperName").toString())==null){
							String cellName = Util.makeNameValue(xmlobjrole.getField("DeveloperName").toString());
							roleNameMap.put(xmlobjrole.getField("DeveloperName").toString(), cellName);	
						    hyperVal = RoleSheetName+"!"+cellName;
						}else{
							hyperVal = RoleSheetName+"!"+roleNameMap.get(xmlobjrole.getField("DeveloperName").toString());
						}							
						String displayVal =Util.nullFilter(xmlobjrole.getField("DeveloperName"));
						//ロール
						excelTemplate.createCellValue(excelSheet,rowNo,colNo++,hyperVal,displayVal);
					}
					else{
						excelTemplate.createCell(userRow,colNo++,"");						
					}
					XmlObject xmlobjpro =(XmlObject) obj.getField("Profile");
					//プロファイル
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(xmlobjpro.getField("Name")));
					//地埁E
					excelTemplate.createCell(userRow,colNo++,Util.getTranslate("LOCALESIDKEY",Util.nullFilter(obj.getField("LocaleSidKey"))));
					//言誁E
					excelTemplate.createCell(userRow,colNo++,Util.getTranslate("TRANSLATIONLANGUAGE",Util.nullFilter(obj.getField("LanguageLocaleKey"))));
					String phone=" ";
					if(obj.getField("Phone")!=null){
						phone=Util.nullFilter(obj.getField("Phone"));
					}
					//電話
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(phone));
					String fax=" ";
					if(obj.getField("Fax")!=null){
						fax=Util.nullFilter(obj.getField("Fax"));
					}
					//FAX
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(fax));
					//売上予測を許可
					excelTemplate.createCell(userRow,colNo++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getField("ForecastEnabled"))));
					String depart=" ";
					if(obj.getField("Department")!=null){
						depart=Util.nullFilter(obj.getField("Department"));
					}
					//部署
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(depart));
					String title=" ";
					if(obj.getField("Title")!=null){
						title=Util.nullFilter(obj.getField("Title"));
					}
					//役職
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(title));
					//メール
					excelTemplate.createCell(userRow,colNo++,Util.nullFilter(obj.getField("Email")));
					//メールの斁E��コーチE
					excelTemplate.createCell(userRow,colNo++,Util.getTranslate("EMAILENCODINGKEY",Util.nullFilter(obj.getField("EmailEncodingKey"))));
					//System.out.println("Util.nullFilter(obj.getField('EmailEncodingKey'))="+Util.nullFilter(obj.getField("EmailEncodingKey")));
					if(obj.getField("ManagerId")!=null){
						//XmlObject xmlobjMan =(XmlObject) obj.getField("Manager");
						Integer rowNo = excelSheet.getLastRowNum();
						
						String hyperVal =userNameMap.get(obj.getField("ManagerId").toString());
						String displayVal =Util.nullFilter(obj.getField("ManagerId"));
						//マネージャ
						excelTemplate.createCellValue(excelSheet,rowNo,colNo++,hyperVal,displayVal);
					}
					else{						
						excelTemplate.createCell(userRow,colNo++,"");
					}
					if(nameList.size()>0){
						for(int j=0;j<nameList.size();j++){
							
							excelTemplate.createCell(userRow,colNo++,Util.nullFilter(obj.getField(nameList.get(j))));
						}
					}
				}
				excelTemplate.adjustColumnWidth(excelSheet);
			}
			if(str.equals("Role")){
				Util.logger.info("Role started");
				//Create Role Sheet
				excelRoleSheet= excelTemplate.createSheet(Util.cutSheetName(RoleSheetName));	
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelRoleSheet,Util.cutSheetName(RoleSheetName),RoleSheetName);	
				ListMetadataQuery query = new ListMetadataQuery();
				query.setType(str);
				List<String> allFile = new ArrayList<String>();
				FileProperties[] lmr = MetadataLoginUtil.metadataConnection.listMetadata(
						new ListMetadataQuery[] { query }, Util.API_VERSION);
				if (lmr != null) {
					for (FileProperties n : lmr) {
						allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));
					}
				}				
				Map<String,String> resultMap = ut.getComparedResult(str,UtilConnectionInfc.getLastUpdateTime());
				List<Metadata> mdInfos = ut.readMateData(str, allFile);
				//Create Role Table
				excelTemplate.createTableHeaders(excelRoleSheet,"Role",excelRoleSheet.getLastRowNum()+Util.RowIntervalNum);
				for(Metadata md : mdInfos){
					Role obj = (Role) md;
					if(roleNameMap.get(obj.getFullName())==null){
						String cellName = Util.makeNameValue(obj.getFullName());
						roleNameMap.put(obj.getFullName(), cellName);
					}					
				}
				for (Metadata md : mdInfos) {																											
					Role obj = (Role) md;	
					//Create roleRow
					XSSFRow roleRow = excelRoleSheet.createRow(excelRoleSheet.getLastRowNum()+1);
					//Create HyperLink name			
					excelTemplate.createCellName(roleNameMap.get(obj.getFullName()),RoleSheetName,excelRoleSheet.getLastRowNum()+1);					
					//excelTemplate.createCellName("ROLE:"+obj.getFullName(),RoleSheetName,excelRoleSheet.getLastRowNum()+1);
					//String objUpdateFlag=resultMap.get("Role." + obj.getFullName());
					Integer cellNum = 1;
					if(UtilConnectionInfc.modifiedFlag){
						//変更あり
						excelTemplate.createCell(roleRow,cellNum++,ut.getUpdateFlag(resultMap,"Role." + obj.getFullName()));
												
					}
					//ロール吁E
					excelTemplate.createCell(roleRow,cellNum++,Util.nullFilter(obj.getFullName()));
					//表示ラベル
					excelTemplate.createCell(roleRow,cellNum++,Util.nullFilter(obj.getName()));	
					//ParentRole
					if(obj.getParentRole()!= null){		
						Integer rowNo = excelRoleSheet.getLastRowNum();
						String hyperVal = roleNameMap.get(obj.getParentRole());
						String displayVal =obj.getParentRole();
						//こ�Eロールの上位ロール
						excelTemplate.createCellValue(excelRoleSheet,rowNo, cellNum++ ,hyperVal,displayVal);						
					} else{
						excelTemplate.createCell(roleRow,cellNum++,"");
					}
					//レポ�Eトに表示するロール吁E
					excelTemplate.createCell(roleRow,cellNum++,Util.nullFilter(obj.getDescription()));
					
					//取引�E責任老E��クセスレベル
					excelTemplate.createCell(roleRow,cellNum++,Util.getTranslate("RoleRelatedAccess", Util.nullFilter(obj.getContactAccessLevel())));
					//啁E��E��クセスレベル
					excelTemplate.createCell(roleRow,cellNum++,Util.getTranslate("RoleRelatedAccess", Util.nullFilter(obj.getOpportunityAccessLevel())));
					//ケースアクセスレベル
					excelTemplate.createCell(roleRow,cellNum++,Util.getTranslate("RoleRelatedAccess", Util.nullFilter(obj.getCaseAccessLevel())));
				}
				excelTemplate.adjustColumnWidth(excelRoleSheet);
			}
			if(str.equals("Queue")){
				Util.logger.info("Queue started");
				//Create Queue sheet
				excelQueueSheet= excelTemplate.createSheet(Util.cutSheetName(QueueSheetName));	
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelQueueSheet,Util.cutSheetName(QueueSheetName),QueueSheetName);
				ListMetadataQuery query = new ListMetadataQuery();
				query.setType(str);
				List<String> allFile = new ArrayList<String>();
				FileProperties[] lmr = MetadataLoginUtil.metadataConnection.listMetadata(
						new ListMetadataQuery[] { query }, Util.API_VERSION);
				if (lmr != null) {
					for (FileProperties n : lmr) {
						allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));
					}
				}				
				List<Metadata> mdInfos = ut.readMateData(str, allFile);
				Map<String,String> resultMap = ut.getComparedResult(str,UtilConnectionInfc.getLastUpdateTime());	
				//Create Queue Table
				excelTemplate.createTableHeaders(excelQueueSheet,"Queue",excelQueueSheet.getLastRowNum()+Util.RowIntervalNum);
				if(allFile.size()>0){
					String names = ut.getObjectNames(allFile);	
					Util.logger.debug("names="+names);
				    for (Metadata md : mdInfos){																											
	
						Queue obj = (Queue) md;	
						//Create queueRow
						XSSFRow queueRow = excelQueueSheet.createRow(excelQueueSheet.getLastRowNum()+1);
						//String objUpdateFlag=resultMap.get("Queue."+obj.getFullName());
						Integer cellNum = 1;
						if(UtilConnectionInfc.modifiedFlag){
							//変更あり
							excelTemplate.createCell(queueRow,cellNum++,ut.getUpdateFlag(resultMap,"Queue."+obj.getFullName()));
							
						}
						//キュー吁E
						excelTemplate.createCell(queueRow,cellNum++,Util.nullFilter(obj.getFullName()));
						//表示ラベル
						excelTemplate.createCell(queueRow,cellNum++,Util.nullFilter(obj.getName()));
						//メール
						excelTemplate.createCell(queueRow,cellNum++,Util.nullFilter(obj.getEmail()));
						//メンバ�Eへのメールの送信
						excelTemplate.createCell(queueRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(obj.getDoesSendEmailToMembers())));					    						
						String QueueSobject ="";
						for(QueueSobject q : obj.getQueueSobject()){
							QueueSobject += ut.getLabelApi(q.getSobjectType()) +"\n";	
						}
						//サポ�EトされるオブジェクチE
						excelTemplate.createCell(queueRow,cellNum++,Util.nullFilter(QueueSobject));
						
						//add zhang 2015-11-19 
						String sqlmember = "Select UserOrGroupId From GroupMember Where Group.Name='"+Util.nullFilter(obj.getName())+"'Order By Id";					
						SObject [] memberObjects= ut.apiQuery(sqlmember);	
	
						List<String> glist = new ArrayList<String>();
						List<String> ulist = new ArrayList<String>();
	
						for(SObject mobj : memberObjects ){		
							String str1 = Util.nullFilter(mobj.getField("UserOrGroupId"));
							String str2 =str1.substring(0, 3);
							if(str2.equals("005")){
								ulist.add(str1);
							}
							else if(str2.equals("00G")){
								glist.add(str1);
							}
						}					
						if(ulist.size()>0){
							String uId ="";
							for(String Id : ulist){
								uId += "'"+Id +"'"+",";
							}
							uId = uId.substring(0,uId.length()-1);	
							String usersql = "Select id,Username,LastName,FirstName From User Where Id in ("+ uId +") Order By Id";
							SObject [] userObj= ut.apiQuery(usersql);
							Integer rowNu = excelQueueSheet.getLastRowNum();	
							for(SObject uobj : userObj){
								String hyperVal = userNameMap.get(Util.nullFilter(uobj.getId()));
								String displayVal = Util.getTranslate("GroupMemberType","User")+":"+Util.nullFilter(uobj.getField("LastName"))+" "+Util.nullFilter(uobj.getField("FirstName"));											
								if(excelQueueSheet.getRow(rowNu)!=null){
									//メンバ�E
									excelTemplate.createCellValue(excelQueueSheet,rowNu, cellNum ,hyperVal,displayVal);
								}else{
									XSSFRow itemRow = excelQueueSheet.createRow(rowNu);	
									excelTemplate.createCell(itemRow,1,"");
									excelTemplate.createCell(itemRow,2,"");
									excelTemplate.createCell(itemRow,3,"");
									excelTemplate.createCell(itemRow,4,"");		
									excelTemplate.createCell(itemRow,5,"");	
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(itemRow,6,"");											
									}
									excelTemplate.createCellValue(excelQueueSheet,rowNu, cellNum ,hyperVal,displayVal);
								}																			
								rowNu=rowNu+1;
							}			
						}
						if(glist.size()>0){
							String gId ="";
							for(String Id : glist){
								gId += "'"+Id +"'" + ",";
							}		
							gId = gId.substring(0,gId.length()-1);						
							String grosql = "Select Type,Name,DeveloperName From Group Where Id in ("+ gId +") Order By Id";
							SObject [] groupObj= ut.apiQuery(grosql);
							Integer rowNo = excelQueueSheet.getLastRowNum()+1;
							
							for(SObject gobj : groupObj){
								if(gobj.getField("Type").equals("Regular")){	
									Util.logger.info("it is a Regular group");
									String hyperVal = groupNameMap.get(Util.nullFilter(gobj.getField("Name")));
									String displayVal = Util.getTranslate("GroupMemberType","GROUP")+":"+Util.nullFilter(gobj.getField("Name"));
									
									if(excelQueueSheet.getRow(rowNo)!=null){
										excelTemplate.createCellValue(excelQueueSheet,rowNo, cellNum ,hyperVal,displayVal);
									}else{
										excelQueueSheet.createRow(rowNo);
										//for cell style
										XSSFRow itemRow = excelQueueSheet.getRow(rowNo);
										excelTemplate.createCell(itemRow,1,"");
										excelTemplate.createCell(itemRow,2,"");
										excelTemplate.createCell(itemRow,3,"");
										excelTemplate.createCell(itemRow,4,"");	
										excelTemplate.createCell(itemRow,5,"");	
										if(UtilConnectionInfc.modifiedFlag){
											excelTemplate.createCell(itemRow,6,"");											
										}
										excelTemplate.createCellValue(excelQueueSheet,rowNo, cellNum,hyperVal,displayVal);
									}
								}
								else if(gobj.getField("Type").equals("Role") || gobj.getField("Type").equals("RoleAndSubordinates")){									
									Util.logger.debug("it is a Role or RoleAndSubordinates.");
									String hyperVal ="";
									int leng = Util.nullFilter(gobj.getField("DeveloperName")).length();
									String Name = Util.nullFilter(gobj.getField("DeveloperName")).substring(0,leng-1);
									String cellName = Util.makeNameValue(gobj.getField("DeveloperName").toString());
									if(roleNameMap.get(Name)==null){	
										int len = cellName.length();
									    hyperVal = cellName.substring(0,len-1);
									}else{
										hyperVal = roleNameMap.get(Name);
									}	
									roleNameMap.put(Name, hyperVal);
									String displayVal="";
									if(gobj.getField("Type").equals("Role")){
										displayVal=Util.getTranslate("GroupMemberType","ROLE")+":"+Name;	
									}else{
										displayVal=Util.getTranslate("GroupMemberType","RoleAndSubordinates")+":"+Name;	
									}
																	
									if(excelQueueSheet.getRow(rowNo)!=null){
										excelTemplate.createCellValue(excelQueueSheet,rowNo, cellNum ,hyperVal,displayVal);
									}else{											
										excelQueueSheet.createRow(rowNo);
										//for cell style
										XSSFRow itemRow = excelQueueSheet.getRow(rowNo);
										excelTemplate.createCell(itemRow,1,"");
										excelTemplate.createCell(itemRow,2,"");
										excelTemplate.createCell(itemRow,3,"");
										excelTemplate.createCell(itemRow,4,"");		
										excelTemplate.createCell(itemRow,5,"");	
										if(UtilConnectionInfc.modifiedFlag){
											excelTemplate.createCell(itemRow,6,"");											
										}
	
										excelTemplate.createCellValue(excelQueueSheet,rowNo, cellNum ,hyperVal,displayVal);
									}
								}							
								rowNo=rowNo+1;
							}
	
						}
						if(glist.size()==0&&ulist.size()==0){
							excelTemplate.createCell(queueRow,cellNum++,"");
						}
					}
			    }else{
					Integer cellNum = 1;
					XSSFRow queueRow = excelQueueSheet.createRow(excelQueueSheet.getLastRowNum()+1);	
					excelTemplate.createCell(queueRow,cellNum++,"NO Queue DATA.");
			    }
				excelTemplate.adjustColumnWidth(excelQueueSheet);
				Util.logger.info("Queue completed");
			}
	
		}	
				
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel("UserGroup","");
		}else{
			Util.logger.warn("no result to export!!!");
		}
		Util.logger.info("ReadGroupRoleQueueSync End");		
	}	
}
