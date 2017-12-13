package source;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;







import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;






import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Workflow;
import com.sforce.soap.metadata.WorkflowActionReference;
import com.sforce.soap.metadata.WorkflowAlert;
import com.sforce.soap.metadata.WorkflowEmailRecipient;
import com.sforce.soap.metadata.WorkflowFieldUpdate;
import com.sforce.soap.metadata.EntitlementProcess;
import com.sforce.soap.metadata.EntitlementProcessMilestoneItem;
import com.sforce.soap.metadata.EntitlementProcessMilestoneTimeTrigger;
import com.sforce.soap.metadata.WorkflowOutboundMessage;
import com.sforce.soap.metadata.WorkflowTask;
import com.sforce.ws.ConnectionException;

public class ReadEntitlementProcessesSync {
	
	private XSSFWorkbook workBook;
	
	public void readEntitlementProcesses(String type,List<String> objectsList) throws Exception{
		Util.logger.info("ReadEntitlementProcessesSync Started.");
		Util.logger.debug("objectsList="+objectsList);
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
		Integer lastIndex=0;
		for (Metadata md : mdInfos) {
			lastIndex+=1;
			if (md != null) {
				// Create EntitlementProcess object
				EntitlementProcess obj = (EntitlementProcess) md;

				//Create EntitlementProcess sheet				
				String apDislayName = obj.getFullName();
				String apSheetName = Util.makeSheetName(apDislayName);
				String actionDisplayName=obj.getFullName()+".Action";
				String actionsSheetName = Util.makeSheetName(actionDisplayName);				
				XSSFSheet excelApSheet= excelTemplate.createSheet(Util.cutSheetName(apSheetName));				
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelApSheet,Util.cutSheetName(apSheetName),apSheetName);
				Map<String,String> actionNameMap = new HashMap<String,String>();

				//Create EntitlementProcess Table
				Integer rowNum = 0;
				excelTemplate.createTableHeaders(excelApSheet,"Entitlement Process",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				XSSFRow columnRowOne = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
				Integer cellNum = 1;
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(columnRowOne,cellNum++,ut.getUpdateFlag(resultMap,"EntitlementProcess."+obj.getFullName()));
										
				}
				
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getFullName()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getIsVersionDefault())));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getDescription()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getActive())));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getEntryStartDateField()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getExitCriteriaFormula()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getExitCriteriaBooleanFilter()));
				excelTemplate.createCell(columnRowOne,cellNum++,ut.getFilterItem(obj.getFullName(), obj.getExitCriteriaFilterItems()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getBusinessHours()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getVersionNotes()));
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getVersionNumber()));							
				
				//Create Approval Step Table
				rowNum = excelApSheet.getLastRowNum()+Util.RowIntervalNum;
				List<String> actionlist = new ArrayList();
				//get action and make name map for hyper link
				List<String> workflowList = new ArrayList<String>();
				workflowList.add("Case");
				List<Metadata> workflowAction = ut.readMateData("Workflow", workflowList);
				List<EntitlementProcessMilestoneTimeTrigger> timeTriggerList= new ArrayList<EntitlementProcessMilestoneTimeTrigger>();
				excelTemplate.createTableHeaders(excelApSheet,"Entitlement Process MilestoneItem",rowNum);
				for(Integer i=0;i<obj.getMilestones().length;i++){
					Integer itemNum = excelApSheet.getLastRowNum()+1;
					//Create columnRow
					XSSFRow columnRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
					Integer cellNum2 = 1;
					
					EntitlementProcessMilestoneItem tempAp=(EntitlementProcessMilestoneItem)obj.getMilestones()[i];		
					String str = tempAp.getMilestoneName();
					workflowList.add(str);  
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(i+1));
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getMilestoneName()));
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getBusinessHours()));
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getMinutesCustomClass()));
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getMinutesToComplete()));
					//excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getTimeTriggers()));//NeedToFix
					if(tempAp.getTimeTriggers().length>0){
						for(EntitlementProcessMilestoneTimeTrigger t:tempAp.getTimeTriggers()){
							timeTriggerList.add(t);
						}
					}
					//excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getSuccessActions()));//NeedToFix
					excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(tempAp.getUseCriteriaStartTime())));
					//Seting EntryCriteria
					if(tempAp.getMilestoneCriteriaFilterItems()!=null){
						excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getMilestoneCriteriaFormula()));
						excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getCriteriaBooleanFilter()));
						excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getFilterItem(obj.getFullName(), tempAp.getMilestoneCriteriaFilterItems())));
					}else{
						excelTemplate.createCell(columnRow,cellNum2++,"");
						excelTemplate.createCell(columnRow,cellNum2++,"");
						excelTemplate.createCell(columnRow,cellNum2++,"");
					}						

					if(tempAp.getSuccessActions()!=null){
						WorkflowActionReference[] wa= tempAp.getSuccessActions();
						for(Integer i3=0; i3<wa.length; i3++){
							WorkflowActionReference action = (WorkflowActionReference)wa[i3];
							//Create columnRow
							Integer rowNo = itemNum+i3;
							String cellName ="";
							if(actionNameMap.get(obj.getFullName()+action.getName())==null){
								cellName = Util.makeNameValue(obj.getFullName()+action.getName());
								actionNameMap.put(obj.getFullName()+action.getName(), cellName);
							}else{
								cellName = actionNameMap.get(obj.getFullName()+action.getName());
							}
							String hyperVal = actionsSheetName+"!"+cellName;								
							//String hyperVal = Util.makeNameValue(obj.getFullName().substring(obj.getFullName().length()-3,obj.getFullName().length())+action.getType())+action.getName();
							String displayVal = Util.getTranslate("StepApprovalAction",String.valueOf(action.getType())) + "." + action.getName();
							excelTemplate.createCellValue(excelApSheet,rowNo,cellNum2,hyperVal,displayVal);
							actionlist.add(obj.getFullName()+action.getName());
						}
						cellNum2++;
					}else{
						excelTemplate.createCell(columnRow,cellNum2++,"");
					}
				}
				excelTemplate.createTableHeaders(excelApSheet,"WorkFlow Time Triggers",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				
				for (EntitlementProcessMilestoneTimeTrigger tt:timeTriggerList) {
					Util.logger.debug("EntitlementProcessMilestoneTimeTrigger="+tt);
					XSSFRow itemRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
					//Create Name, used for HyperLink 
					//excelTemplate.createCellName(obj.getFullName(),apSheetName,excelApSheet.getLastRowNum()+1);
					Integer cellNums = 1;
					excelTemplate.createCell(itemRow,cellNums++,Util.nullFilter(obj.getFullName()));
					excelTemplate.createCell(itemRow,cellNums++,Util.nullFilter(tt.getTimeLength()));
					excelTemplate.createCell(itemRow,cellNums++,Util.getTranslate("WorkflowTimeUnit",Util.nullFilter(tt.getWorkflowTimeTriggerUnit())));
					
					for(Integer i=0; i<tt.getActions().length;i++){
						WorkflowActionReference action = (WorkflowActionReference)tt.getActions()[i];
						Util.logger.debug("WorkflowActionReference="+action);
						//Create columnRow
						Integer rowNo = excelApSheet.getLastRowNum();
						String hyperVal = " ";
						String cellName ="";
						if(actionNameMap.get(obj.getFullName()+action.getName())==null){
							cellName = Util.makeNameValue(obj.getFullName()+action.getName());
							actionNameMap.put(obj.getFullName()+action.getName(), cellName);
						}else{
							cellName = actionNameMap.get(obj.getFullName()+action.getName());
						}
						hyperVal=actionsSheetName+"!"+Util.makeNameValue(action.getType()+action.getName());
						workflowMap.put(action.getType()+action.getName(), hyperVal);							
						Util.logger.debug("hyperVal=="+hyperVal);
						String displayVal =Util.getTranslate("WorkflowActionType",String.valueOf(action.getType() + "." + action.getName()));
						//Create HyperLink
						excelTemplate.createCellValue(excelApSheet,rowNo,cellNums,hyperVal,displayVal);
						actionlist.add(obj.getFullName()+action.getName());
					}
					if(tt.getActions().length==0){
						excelTemplate.createCell(itemRow,cellNums++,"");
					}
					cellNums++;
					
				}

				excelTemplate.adjustColumnWidth(excelApSheet);
				
				//Action sheet *****************************************************************/
				Map<String, String> resultMapWork = null;
				try {
					resultMapWork = ut.getComparedResult("Workflow",UtilConnectionInfc.getLastUpdateTime());
				} catch (ConnectionException e1) {
					e1.printStackTrace();
				}
				XSSFSheet excelActionSheet= excelTemplate.createSheet(Util.cutSheetName(actionsSheetName));	
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelActionSheet,Util.cutSheetName(actionsSheetName),actionsSheetName);
				if(workflowAction!=null){
					for(Metadata meta: workflowAction){
						Workflow objw = (Workflow) meta;
						//Create Task Table
						excelTemplate.createTableHeaders(excelActionSheet,"Workflow Task",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
						if(objw.getTasks().length>0){
							for(WorkflowTask task:objw.getTasks()){								
								if(actionlist.contains(obj.getFullName()+task.getFullName())){
									//Create columnRow
									XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
									Integer cellNum2=1;
									String cellName = actionNameMap.get(obj.getFullName()+task.getFullName());							
									//Create HyperLink name 
									excelTemplate.createCellName(cellName,actionsSheetName,excelActionSheet.getLastRowNum()+1);
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRowOne,cellNum++,ut.getUpdateFlag(resultMapWork,"WorkflowTask."+objw.getFullName()+"."+task.getFullName()));
										
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getFullName()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("ActionTaskAssignedToType",Util.nullFilter(task.getAssignedToType())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getAssignedTo()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getSubject()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getDueDateOffset()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(task.getNotifyAssignee())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(task.getProtected())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getStatus()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getPriority()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getDescription()));
									}
							}
						}
						//Create Alerts Table
						excelTemplate.createTableHeaders(excelActionSheet,"Workflow Alert",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);							
						if(objw.getAlerts().length>0){
							for(WorkflowAlert alert:objw.getAlerts()){
								if(actionlist.contains(obj.getFullName()+alert.getFullName())){
									//Create columnRow
									XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
									//Create HyperLink name
									String cellName = actionNameMap.get(obj.getFullName()+alert.getFullName());
									
									excelTemplate.createCellName(cellName,actionsSheetName,excelActionSheet.getLastRowNum()+1);
									Integer cellNum2=1;
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRowOne,cellNum++,ut.getUpdateFlag(resultMapWork,"WorkflowAlert."+objw.getFullName()+"."+alert.getFullName()));
										
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getFullName()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getDescription()));
									excelTemplate.createCell(columnRow,cellNum2++,ut.getEmailTemplateLabel(Util.nullFilter(alert.getTemplate())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(alert.getProtected())));
	
									//Alerts WorkflowEmailRecipient
									String wfEmailRecipient = "";
									for(Integer items=0; items<alert.getRecipients().length; items++){
										WorkflowEmailRecipient wfer = alert.getRecipients()[items];
										wfEmailRecipient += Util.getTranslate("ActionEmailRecipientType",String.valueOf(wfer.getType()));
										if(wfer.getField()!=null){
											wfEmailRecipient += ":"+ wfer.getField();
										}
										if(wfer.getRecipient()!=null){
											wfEmailRecipient += ":"+ wfer.getRecipient();
										}
										wfEmailRecipient += "\r\n";
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfEmailRecipient));
									//Alerts CcEmails
									String ccEmails = "";
									for(Integer items=0; items<alert.getCcEmails().length; items++){
										ccEmails += alert.getCcEmails()[items]+"\r\n";
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ccEmails));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("ActionEmailSenderType",Util.nullFilter(alert.getSenderType())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getSenderAddress()));
								}
							}
						}
						//Create FieldUpdates Table
						excelTemplate.createTableHeaders(excelActionSheet,"Workflow FieldUpdate",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);						
						if(objw.getFieldUpdates().length>0){
							for(WorkflowFieldUpdate wffu:objw.getFieldUpdates()){
								if(actionlist.contains(obj.getFullName()+wffu.getFullName())){
									//Create columnRow
									XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
									
									String cellName = actionNameMap.get(obj.getFullName()+wffu.getFullName());								
									//Create HyperLink name
									excelTemplate.createCellName(cellName,actionsSheetName,excelActionSheet.getLastRowNum()+1);	
									Integer cellNum2=1;
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRowOne,cellNum++,ut.getUpdateFlag(resultMapWork,"WorkflowFieldUpdate."+objw.getFullName()+"."+wffu.getFullName()));
										
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getName()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getFullName()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getDescription()));
									if(wffu.getTargetObject()==null){
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(objw.getFullName())));
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(objw.getFullName()+"."+wffu.getField())));
									}else{
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(wffu.getTargetObject())));
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(wffu.getTargetObject()+"."+wffu.getField())));
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getReevaluateOnChange())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getProtected())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("FieldUpdateOperation",Util.nullFilter(wffu.getOperation())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getLiteralValue()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getLookupValue()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("LookupValueType",Util.nullFilter(wffu.getLookupValueType())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getFormula()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getNotifyAssignee())));
								}
							}
						}
						//Create OutboundMessage Table
						excelTemplate.createTableHeaders(excelActionSheet,"Workflow OutboundMessage",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
						if(objw.getOutboundMessages().length>0){
							for(WorkflowOutboundMessage wfom:objw.getOutboundMessages()){
								if(actionlist.contains(obj.getFullName()+wfom.getFullName())){
									//Create columnRow
									XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
									String cellName = actionNameMap.get(obj.getFullName()+wfom.getFullName());
																		
									//Create HyperLink name
									excelTemplate.createCellName(cellName,actionsSheetName,excelActionSheet.getLastRowNum()+1);
									Integer cellNum2 = 1;
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRowOne,cellNum++,ut.getUpdateFlag(resultMapWork,"WorkflowOutboundMessage."+objw.getFullName()+"."+wfom.getFullName()));
										
									}
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getName()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getFullName()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getDescription()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getEndpointUrl()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getIntegrationUser()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wfom.getProtected())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wfom.getIncludeSessionId())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wfom.getUseDeadLetterQueue())));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getFields()));
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getApiVersion()));
								}
							}
						}
					}						
				}
				excelTemplate.adjustColumnWidth(excelActionSheet);		
			}						
			if(ut.createExcel(workBook, excelTemplate, type, objectsList.size(), lastIndex)){
				excelTemplate.CreateWorkBook(type);
				workBook = excelTemplate.workBook;
			}
		}
		Util.logger.info("readApprovalProcess End.");
	}
}