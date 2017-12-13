package source;

import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;




import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.FilterItem;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Workflow;
import com.sforce.soap.metadata.WorkflowActionReference;
import com.sforce.soap.metadata.WorkflowAlert;
import com.sforce.soap.metadata.WorkflowEmailRecipient;
import com.sforce.soap.metadata.WorkflowFieldUpdate;
import com.sforce.soap.metadata.WorkflowOutboundMessage;
import com.sforce.soap.metadata.WorkflowRule;
import com.sforce.soap.metadata.WorkflowTask;
import com.sforce.soap.metadata.WorkflowTimeTrigger;
import com.sforce.ws.ConnectionException;

public class ReadWorkFlowSync {
	
	private XSSFWorkbook workBook;
	
	public void readWorkFlow(String type,List<String> objectsList) throws Exception{
		Util.logger.info("ReadWorkFlowSync Started.");
		Util.nameSequence=0;
		Util.sheetSequence=0;
		Util ut = new Util();
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//Create catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		Map<String,String> workflowMap = new HashMap<String,String>();
		/*** Loop MetaData results ***/
		for (Metadata md : mdInfos) {
			if (md != null) {
				// Create WorkFlow object
				Workflow obj = (Workflow) md;
				Util.logger.info("Workflow="+obj +" started.");				
				/*** WorkFlow Rules ***/
				if(obj.getRules().length>0){
					
					Map<String,WorkflowTimeTrigger> triggersMap= new LinkedHashMap<String,WorkflowTimeTrigger>();
					
					//Create sheet
					String rulesSheetName = Util.makeSheetName(obj.getFullName()+"_Rule");
					
					XSSFSheet excelRuleSheet= excelTemplate.createSheet(Util.cutSheetName(rulesSheetName));
					//Create Catalog Menu
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelRuleSheet,Util.cutSheetName(rulesSheetName),rulesSheetName);
					
					//Create Table
					//ルール
					Util.logger.info("WorkFlow Rule started.");	
					excelTemplate.createTableHeaders(excelRuleSheet,"WorkFlow Rule",excelRuleSheet.getLastRowNum()+Util.RowIntervalNum);
					
					for( Integer wi=0; wi<obj.getRules().length; wi++ ){
						
						Integer itemNum = excelRuleSheet.getLastRowNum()+1;
						Integer rowStart=itemNum;
						
						//Create columnRow
						XSSFRow columnRow = excelRuleSheet.createRow(excelRuleSheet.getLastRowNum()+1);
						//Rule
						Integer cellNum = 1;
						WorkflowRule tempRule=(WorkflowRule)obj.getRules()[wi];
						if(UtilConnectionInfc.modifiedFlag){
							excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"WorkflowRule."+obj.getFullName()+"."+tempRule.getFullName()));
													
						}
						//ルール名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.translateSpecialChar(String.valueOf(tempRule.getFullName()))));
						//説明
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(tempRule.getDescription()));
						//有効
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(tempRule.getActive())));
						//トリガ種別
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("WorkflowTriggerType",Util.nullFilter(tempRule.getTriggerType())));
						//数式
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(tempRule.getFormula()));
						//Rule CriteriaItems
						String criteriaItem = "";
						for(Integer items=0; items<tempRule.getCriteriaItems().length; items++){
							FilterItem  item = (FilterItem)tempRule.getCriteriaItems()[items];
							criteriaItem += "["+item.getField()+"]"+Util.getTranslate("FILTEROPERATION",String.valueOf(item.getOperation()))+"["+item.getValue()+"]\n";
						}
						//ルール条件
						excelTemplate.createCell(columnRow,cellNum++,ut.getFilterItem(obj.getFullName(), tempRule.getCriteriaItems()));
						//検索条件
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(tempRule.getBooleanFilter()));						
						//Rule Actions
						//アクション種別と名前
						for(Integer actions=0; actions<tempRule.getActions().length; actions++){							
							WorkflowActionReference action = (WorkflowActionReference)tempRule.getActions()[actions];
							//Create columnRow
							Integer rowNo = itemNum+actions;
							String hyperVal ="";
							if(workflowMap.get(action.getType()+action.getName())==null){
								hyperVal=Util.makeNameValue(action.getType()+action.getName());
								workflowMap.put(action.getType()+action.getName(), hyperVal);
							}
							else{
								hyperVal=workflowMap.get(action.getType()+action.getName());
							}
							String displayVal = Util.getTranslate("WorkflowActionType",String.valueOf(action.getType()))+"."+Util.nullFilter(action.getName());
							excelTemplate.createCellValue(excelRuleSheet,rowNo,cellNum,hyperVal,displayVal);
						}
						cellNum++;
						//Rule WorkflowTimeTriggers
						//時間ベースのアクション
						for(Integer triggers=0; triggers<tempRule.getWorkflowTimeTriggers().length; triggers++){
							WorkflowTimeTrigger timeTrigger = (WorkflowTimeTrigger )tempRule.getWorkflowTimeTriggers()[triggers];
							//Create columnRow
							Integer rowNo = itemNum+triggers;
							String hyperVal = Util.makeNameValue(tempRule.getFullName()+".timetrigger_"+(triggers+1));
							String displayVal = tempRule.getFullName()+".timetrigger_"+(triggers+1);
							excelTemplate.createCellValue(excelRuleSheet,rowNo,cellNum,hyperVal,displayVal);
							triggersMap.put(hyperVal, timeTrigger);
						}
						cellNum++;
                        
						Integer rowEnd=excelRuleSheet.getLastRowNum();
						//To make sure the table is fully drawed 
						for(int rowNo=rowStart;rowNo<=rowEnd;rowNo++){
							//Create columnRow
							columnRow = excelRuleSheet.getRow(rowNo);
							for(Integer colNo=1; colNo<cellNum; colNo++){
								if(columnRow.getCell(colNo)==null){
									excelTemplate.createCell(columnRow,colNo,"");
								}
							}

						}
					}
					
					
					Util.logger.info("WorkFlow Rule completed.");						
					//Create Table
					//時間ベースのアクション
					Util.logger.info("WorkFlow Time Triggers started.");
					excelTemplate.createTableHeaders(excelRuleSheet,"WorkFlow Time Triggers",excelRuleSheet.getLastRowNum()+Util.RowIntervalNum);
					
					Set<Map.Entry<String, WorkflowTimeTrigger>> headerSet=triggersMap.entrySet();
					for (Map.Entry<String, WorkflowTimeTrigger> header : headerSet) {
						XSSFRow itemRow = excelRuleSheet.createRow(excelRuleSheet.getLastRowNum()+1);
						//Create Name, used for HyperLink 
						excelTemplate.createCellName(header.getKey(),rulesSheetName,excelRuleSheet.getLastRowNum()+1);
						Integer cellNum = 1;
						//アクションの名前
						excelTemplate.createCell(itemRow,cellNum++,Util.nullFilter(header.getKey()));
						//基準日
						if(header.getValue().getOffsetFromField()!=null){
							excelTemplate.createCell(itemRow,cellNum++,ut.getLabelforAll(Util.nullFilter(header.getValue().getOffsetFromField())));
						}else{
							excelTemplate.createCell(itemRow,cellNum++,Util.getTranslate("WorkFlowTimeTrigger","DefaultDate"));
						}
						//時間の長さ
						excelTemplate.createCell(itemRow,cellNum++,Util.nullFilter(header.getValue().getTimeLength()));
						//時間の単位
						excelTemplate.createCell(itemRow,cellNum++,Util.getTranslate("WorkflowTimeUnit",Util.nullFilter(header.getValue().getWorkflowTimeTriggerUnit())));
						//アクション種別と名前
						for(Integer i=0; i<header.getValue().getActions().length;i++){
							WorkflowActionReference action = (WorkflowActionReference)header.getValue().getActions()[i];
							//Create columnRow
							Integer rowNo = excelRuleSheet.getLastRowNum();
							String hyperVal = " ";
							hyperVal=Util.makeNameValue(action.getType()+action.getName());
							workflowMap.put(action.getType()+action.getName(), hyperVal);							
							String displayVal =Util.getTranslate("WorkflowActionType",String.valueOf(action.getType()) + "." + action.getName());
							//Create HyperLink
							
							excelTemplate.createCellValue(excelRuleSheet,rowNo,cellNum,hyperVal,displayVal);
						}
						
						if(header.getValue().getActions().length==0){
							excelTemplate.createCell(itemRow,cellNum++,"");
						}
						cellNum++;
						
					}
					excelTemplate.adjustColumnWidth(excelRuleSheet);
				}
				Util.logger.info("WorkFlow Time Triggers completed.");				
				/*** WorkFlow Actions ***/
				Util.logger.info("WorkFlow Task started.");	
				if( obj.getTasks().length>0 || 
					obj.getAlerts().length>0 || 
					obj.getFieldUpdates().length>0 || 
					obj.getOutboundMessages().length>0 || 
					obj.getFlowActions().length>0 ||
					obj.getKnowledgePublishes().length>0 )
				{
					//ToDo
					String actionsSheetName = Util.makeSheetName(obj.getFullName()+"_Action");
					XSSFSheet excelActionSheet= excelTemplate.createSheet(Util.cutSheetName(actionsSheetName));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelActionSheet,Util.cutSheetName(actionsSheetName),actionsSheetName);
					excelTemplate.createTableHeaders(excelActionSheet,"WorkFlow Task",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
				
					if(obj.getTasks().length>0){
						for( Integer t=0; t<obj.getTasks().length; t++ ){
							
							//Create columnRow
							XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
							
							WorkflowTask task=(WorkflowTask)obj.getTasks()[t];
							//Create Name, used for HyperLink
							if(workflowMap.get("Task"+task.getFullName())!=null){
								excelTemplate.createCellName(workflowMap.get("Task"+task.getFullName()),actionsSheetName,excelActionSheet.getLastRowNum()+1);
							}
							int cellNum=1;
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"WorkflowTask."+obj.getFullName()+"."+task.getFullName()));
								
							}
							//一意の名前
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getFullName()));
							//任命先種別
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ActionTaskAssignedToType",Util.nullFilter(task.getAssignedToType())));
							//任命先
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getAssignedTo()));
							//件名
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getSubject()));
							//期日
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getDueDateOffset()));
							//ワークフロータイムトリガ
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getOffsetFromField()));
							//任命先へ通知
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(task.getNotifyAssignee())));
							//保護コンポーネント
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(task.getProtected())));
							//状況
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getStatus()));
							//優先度
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getPriority()));
							//詳細情報
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(task.getDescription()));
						}
					}
					Util.logger.info("WorkFlow Task completed.");						
					//Create Table
					//メールアラート
					Util.logger.info("WorkFlow Alert completed.");	
					excelTemplate.createTableHeaders(excelActionSheet,"WorkFlow Alert",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
					
					if(obj.getAlerts().length>0){
						for( Integer a=0; a<obj.getAlerts().length; a++ ){
							//Create columnRow
							XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
							WorkflowAlert alert = (WorkflowAlert)obj.getAlerts()[a];
							//Create Name, used for HyperLink
							int cellNum=1;
							if(workflowMap.get("Alert"+alert.getFullName())!=null){
								excelTemplate.createCellName(workflowMap.get("Alert"+alert.getFullName()),actionsSheetName,excelActionSheet.getLastRowNum()+1);
							}
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"WorkflowAlert."+obj.getFullName()+"."+alert.getFullName()));
								
							}
							//一意の名前
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(alert.getFullName()));
							//説明
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(alert.getDescription()));
							//メールテンプレート
							excelTemplate.createCell(columnRow,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(alert.getTemplate())));
							//保護コンポーネント
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(alert.getProtected())));
							//Alerts WorkflowEmailRecipient
							String wfEmailRecipient = "";
							for(Integer items=0; items<alert.getRecipients().length; items++){
								WorkflowEmailRecipient wfer = alert.getRecipients()[items];
								wfEmailRecipient += Util.getTranslate("ActionEmailRecipientType",String.valueOf(wfer.getType()));
								if(wfer.getField()!=null){
									wfEmailRecipient += ":"+ wfer.getField();
								}
								if(wfer.getRecipient()!=null){
									wfEmailRecipient += ":"+ ut.getUserLabel("Username", wfer.getRecipient());
								}
								wfEmailRecipient += "\r\n";
							}
							//メール受信者
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfEmailRecipient));
							//Alerts CcEmails
							String ccEmails = "";
							for(Integer items=0; items<alert.getCcEmails().length; items++){
								ccEmails += alert.getCcEmails()[items]+"\r\n";
							}
							//追加のメール
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ccEmails));
							//差出人種別
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ActionEmailSenderType",Util.nullFilter(alert.getSenderType())));
							//差出人メールアドレス
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(alert.getSenderAddress()));
						}
					}
					Util.logger.info("WorkFlow Alert completed.");						
					
					//Create Table
					//項目自動更新
					Util.logger.info("WorkFlow Field Update start.");						
					excelTemplate.createTableHeaders(excelActionSheet,"WorkFlow Field Update",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getFieldUpdates().length>0){
						for( Integer t=0; t<obj.getFieldUpdates().length; t++ ){
							//Create columnRow
							XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
							WorkflowFieldUpdate wffu=(WorkflowFieldUpdate)obj.getFieldUpdates()[t];
							//Create Name, used for HyperLink
							int cellNum=1;
							String hyperVal="";
							if(workflowMap.get("FieldUpdate"+wffu.getName())==null){
								hyperVal=Util.makeNameValue("FieldUpdate"+wffu.getName());
								workflowMap.put("FieldUpdate"+wffu.getName(), hyperVal);
							}
							else{
								hyperVal=workflowMap.get("FieldUpdate"+wffu.getName());
							}
							if(workflowMap.get("FieldUpdate"+wffu.getFullName())!=null){
								excelTemplate.createCellName(workflowMap.get("FieldUpdate"+wffu.getFullName()),actionsSheetName,excelActionSheet.getLastRowNum()+1);
							}
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"WorkflowFieldUpdate."+obj.getFullName()+"."+wffu.getFullName()));
								
							}
							//一意の名前
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wffu.getFullName()));
							//名前
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wffu.getName()));
							//説明
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wffu.getDescription()));
							if(wffu.getTargetObject()==null){
							
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLabelApi(obj.getFullName())));
								//更新する項目
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLabelApi(obj.getFullName()+"."+wffu.getField())));
							}else{
							
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLabelApi(wffu.getTargetObject())));
								//オブジェクト
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ut.getLabelApi(wffu.getTargetObject()+"."+wffu.getField())));
							}
							//項目変更後にワークフロールールを再評価する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getReevaluateOnChange())));
							//保護コンポーネント
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getProtected())));
						    //条件
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("FieldUpdateOperation",String.valueOf(wffu.getOperation()))));
							//リテラル値
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wffu.getLiteralValue()));
							//項目データの種別
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("LookupValueType",Util.nullFilter(wffu.getLookupValueType())));
							//項目値
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wffu.getLookupValue()));
							//数式
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wffu.getFormula()));
						//任命先へ通知
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getNotifyAssignee())));
						}
					}
					Util.logger.info("WorkFlow Field Update completed.");						
					//Create Table
					//アウトバウンドメッセージ
					Util.logger.info("WorkFlow OutboundMessages started.");	
					excelTemplate.createTableHeaders(excelActionSheet,"WorkFlow OutboundMessage",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
					if(obj.getOutboundMessages().length>0){
						for( Integer t=0; t<obj.getOutboundMessages().length; t++ ){
							//Create columnRow
							XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
							WorkflowOutboundMessage wfom=(WorkflowOutboundMessage)obj.getOutboundMessages()[t];
							//Create Name, used for HyperLink
							int cellNum=1;
							if(workflowMap.get("OutboundMessages"+wfom.getFullName())!=null){
								excelTemplate.createCellName(workflowMap.get("OutboundMessages"+wfom.getFullName()),actionsSheetName,excelActionSheet.getLastRowNum()+1);
							}
							if(UtilConnectionInfc.modifiedFlag){
								excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"WorkflowOutboundMessage."+obj.getFullName()+"."+wfom.getFullName()));
								
							}
							//一意の名前
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfom.getFullName()));
						//名前
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfom.getName()));
							//説明
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfom.getDescription()));
						//エンドポイント URL
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfom.getEndpointUrl()));
							//送信ユーザ
							excelTemplate.createCell(columnRow,cellNum++,ut.getUserLabel("Username", Util.nullFilter(wfom.getIntegrationUser())));
							//保護コンポーネント
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wfom.getProtected())));
							//送信セッション ID
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wfom.getIncludeSessionId())));
							//配信不能メッセージキュー
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wfom.getUseDeadLetterQueue())));
							//送信する項目
							//excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfom.getFields()));
							String[] strs=wfom.getFields();
							if(strs!=null&&strs.length>0){
								String showStr="";
								for(String filed: strs){
									showStr+=filed+"\r\n";
								}
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(showStr));
							}else{
								excelTemplate.createCell(columnRow,cellNum++,"");
							}
							//API バージョン
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wfom.getApiVersion()));
						}
					}
					Util.logger.info("WorkFlow OutboundMessages completed.");	
					excelTemplate.adjustColumnWidth(excelActionSheet);
				}
				Util.logger.info("Workflow="+obj +" completed.");	
			}
			else {
				Util.logger.warn("Empty metadata.");	
			}
		}
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//Auto width  does not work on link value
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.warn("no result to export!!!");	
		}
		Util.logger.info("ReadWorkFlowSync End.");
	}

}
