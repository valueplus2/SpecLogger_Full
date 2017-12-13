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

import com.sforce.soap.metadata.ApprovalProcess;
import com.sforce.soap.metadata.ApprovalStep;
import com.sforce.soap.metadata.ApprovalStepApprover;
import com.sforce.soap.metadata.ApprovalSubmitter;
import com.sforce.soap.metadata.Approver;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Workflow;
import com.sforce.soap.metadata.WorkflowActionReference;
import com.sforce.soap.metadata.WorkflowAlert;
import com.sforce.soap.metadata.WorkflowEmailRecipient;
import com.sforce.soap.metadata.WorkflowFieldUpdate;
import com.sforce.soap.metadata.WorkflowFlowAction;
import com.sforce.soap.metadata.WorkflowKnowledgePublish;
import com.sforce.soap.metadata.WorkflowOutboundMessage;
import com.sforce.soap.metadata.WorkflowTask;
import com.sforce.ws.ConnectionException;

public class ReadApprovalProcessSync {	
	private XSSFWorkbook workBook;
	
	public void readApprovalProcess(String type,List<String> objectsList) throws Exception {
		Util.logger.info("readApprovalProcess Start.");
		Util ut = new Util();
	
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String, String> resultMap = null;
		try {
			resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		} catch (ConnectionException e1) {
			e1.printStackTrace();
		}
		Util.nameSequence=0;
		Util.sheetSequence=0;
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		
		//Create catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		Integer lastIndex=0;
		for (Metadata md : mdInfos) {
			lastIndex+=1;
			if (md != null) {

				ApprovalProcess obj = (ApprovalProcess) md;    //Creat ApprovalProcess Object
				//Create ApprovalProcessName sheet				
				String apDislayName = obj.getFullName();
				String apSheetName = Util.makeSheetName(apDislayName);
				String actionDisplayName=obj.getFullName()+".Action";
				
				String actionsSheetName = Util.makeSheetName(actionDisplayName);				
				XSSFSheet excelApSheet= excelTemplate.createSheet(Util.cutSheetName(apSheetName));				
				//excelTemplate.createCatalogMenu(catalogSheet,excelApSheet,apSheetName,apDislayName);
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelApSheet,Util.cutSheetName(apSheetName),apSheetName);
				//get workflow action and make name map for hyper link
				String str = obj.getFullName().substring(0, obj.getFullName().indexOf('.'));
				List<String> workflowList = new ArrayList<String>();
				workflowList.add(str);  
				List<Metadata> workflowAction = ut.readMateData("Workflow", workflowList);
				Map<String,String> actionNameMap = new HashMap<String,String>();
				Map<String,String> actionLabelMap = new HashMap<String,String>();
				if(workflowAction!=null){
					for(Metadata meta: workflowAction){
						Workflow objw = (Workflow) meta;
						if(objw.getTasks().length>0){
							for(WorkflowTask task:objw.getTasks()){
								String cellName = obj.getFullName()+task.getFullName();
								actionNameMap.put(cellName, Util.makeNameValue(cellName));
								actionLabelMap.put(cellName,task.getFullName());
							}
						}
						if(objw.getAlerts().length>0){
							for(WorkflowAlert alert:objw.getAlerts()){
								String cellName = obj.getFullName()+alert.getFullName();
								actionNameMap.put(cellName, Util.makeNameValue(cellName));
								actionLabelMap.put(cellName,alert.getDescription());
							}
						}
						if(objw.getFieldUpdates().length>0){
							for(WorkflowFieldUpdate wffu:objw.getFieldUpdates()){
								String cellName = obj.getFullName()+wffu.getFullName();
								actionNameMap.put(cellName, Util.makeNameValue(cellName));
								actionLabelMap.put(cellName,wffu.getName());
							}
						}
						if(objw.getOutboundMessages().length>0){
							for(WorkflowOutboundMessage wfom:objw.getOutboundMessages()){
								String cellName = obj.getFullName()+wfom.getFullName();
								actionNameMap.put(cellName, Util.makeNameValue(cellName));
								actionLabelMap.put(cellName,wfom.getName());
							}
						}
						if(objw.getFlowActions().length>0){
							for(WorkflowFlowAction wfa:objw.getFlowActions()){
								String cellName = obj.getFullName()+wfa.getFullName();
								actionNameMap.put(cellName, Util.makeNameValue(cellName));
								actionLabelMap.put(cellName,wfa.getLabel());
							}
						}
						if(objw.getKnowledgePublishes().length>0){
							for(WorkflowKnowledgePublish wfkp:objw.getKnowledgePublishes()){
								String cellName = obj.getFullName()+wfkp.getFullName();
								actionNameMap.put(cellName, Util.makeNameValue(cellName));
								actionLabelMap.put(cellName,wfkp.getLabel());
							}		
						}
					}
				}
				//Create ApprovalProcessName Table(承認�Eロセス)
				Integer rowNum = 0;
				excelTemplate.createTableHeaders(excelApSheet,"Approval Process",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				XSSFRow columnRowOne = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
				Integer cellNum = 1;
					//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(columnRowOne,cellNum++,ut.getUpdateFlag(resultMap,"ApprovalProcess."+obj.getFullName()));
											
				}
				//一意�E名前
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getFullName()));
				//プロセス吁E
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getLabel()));
				//説昁E
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getDescription()));
				//有効
				excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(obj.getActive())));
				if(obj.getEntryCriteria()!=null){
					//条件
					excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ut.getFilterItem(obj.getFullName().substring(0, obj.getFullName().indexOf('.')), obj.getEntryCriteria().getCriteriaItems())));
					//数弁E
					excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getEntryCriteria().getFormula()));
					//条件ロジチE��
					excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getEntryCriteria().getBooleanFilter()));
				}else{
					excelTemplate.createCell(columnRowOne,cellNum++,"");
					excelTemplate.createCell(columnRowOne,cellNum++,"");
					excelTemplate.createCell(columnRowOne,cellNum++,"");
				}
				//レコード�E編雁E��限老E
				excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("RECORDEDITABILITY", Util.nullFilter(obj.getRecordEditability())));
				//承認割り当てメールチE��プレーチE
				excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(obj.getEmailTemplate())));
				//ポストテンプレーチE
				excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(obj.getPostTemplate())));
				//申請老E��承認申請�E取り消しを許可
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getAllowRecall()));
				//最終承認時レコード�EロチE��
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getFinalApprovalRecordLock()));
				//最終却下時レコード�EロチE��
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getFinalRejectionRecordLock()));
				//承認履歴を表示
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getShowApprovalHistory()));
				//モバイルチE��イスからアクセスの許可
				excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getEnableMobileDeviceAccess()));

				//hyperlink
				Integer itNum = excelApSheet.getLastRowNum();				
				List<String> actionlist = new ArrayList();
				if(obj.getNextAutomatedApprover()!=null){	
					//所有老E�E承認老E��E��を使用
					excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getNextAutomatedApprover().getUseApproverFieldOfRecordOwner()));
					//割り当て先として使用するユーザ頁E��
					if(obj.getNextAutomatedApprover().getUserHierarchyField()!=null){
						excelTemplate.createCell(columnRowOne,cellNum++,ut.getLabelforAll("User."+Util.nullFilter(obj.getNextAutomatedApprover().getUserHierarchyField())));
					}else{
						excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(obj.getNextAutomatedApprover().getUserHierarchyField()));
					}
				}else{
					excelTemplate.createCell(columnRowOne,cellNum++,"");
					excelTemplate.createCell(columnRowOne,cellNum++,"");
				}
				
				//Create Approval Submitter Table(承認申請老E
				rowNum = excelApSheet.getLastRowNum()+Util.RowIntervalNum;
				excelTemplate.createTableHeaders(excelApSheet,"Approval Submitter",rowNum);
				for(Integer i=0;i<obj.getAllowedSubmitters().length;i++){				
					//Create columnRow
					XSSFRow columnRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
					ApprovalSubmitter tempAp=(ApprovalSubmitter)obj.getAllowedSubmitters()[i];
					Integer cellNum2 = 1;
					//タイチE
					excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("SUBMITTERTYPE",Util.nullFilter(tempAp.getType())));
					//申請老E
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getSubmitter()));
				}	
				
				//Create Approval PageField Table(承認�Eージ頁E��)
				rowNum = excelApSheet.getLastRowNum()+Util.RowIntervalNum;
				excelTemplate.createTableHeaders(excelApSheet,"Approval PageField",rowNum);
				//Create columnRow
				XSSFRow columnRow1 = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
				Integer cellNum3 = 1;
				String ApprovalPageFieldStr = "";
				ApprovalPageFieldStr+=ut.getLabelforAll(Util.nullFilter(obj.getFullName().substring(0,obj.getFullName().indexOf('.')+1)+"name"))+"\n";
				for(int k = 1;k<obj.getApprovalPageFields().getField().length;k++){
					ApprovalPageFieldStr+=ut.getLabelforAll(Util.nullFilter(obj.getFullName().substring(0,obj.getFullName().indexOf('.')+1)+obj.getApprovalPageFields().getField()[k]))+"\n";
					//System.out.println(obj.getFullName().substring(0,obj.getFullName().indexOf('.')+1)+obj.getApprovalPageFields().getField()[k]+"\n");
				}	
				//頁E��
				excelTemplate.createCell(columnRow1,cellNum3++,ApprovalPageFieldStr);
				
				//create initial Submission Actions table(申請時のアクション)
				excelTemplate.createTableHeaders(excelApSheet,"initial Submission Actions",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				if(obj.getInitialSubmissionActions()!=null){
					Integer cellNumTem = 1;					
					for(WorkflowActionReference action :obj.getInitialSubmissionActions().getAction()){
						XSSFRow actionRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
						String cellName = actionNameMap.get(obj.getFullName()+action.getName());
						String hyperVal = actionsSheetName+"!"+cellName;
						String displayVal = actionLabelMap.get(obj.getFullName()+action.getName());
						//String displayVal = action.getName();
						//種別
						excelTemplate.createCell(actionRow,cellNumTem++,Util.getTranslate("ACTIONTYPE",Util.nullFilter(action.getType())));
						//名前
						excelTemplate.createCellValue(excelApSheet,actionRow.getRowNum(),cellNumTem++,hyperVal,displayVal);
						cellNumTem=1;
						actionlist.add(obj.getFullName()+action.getName());
					}
				}								
				
				//Create Approval Step Table(承認スチE��チE
				rowNum = excelApSheet.getLastRowNum()+Util.RowIntervalNum;
				excelTemplate.createTableHeaders(excelApSheet,"Approval Step",rowNum);
				for(Integer i=0;i<obj.getApprovalStep().length;i++){
					Integer itemNum = excelApSheet.getLastRowNum()+1;
					//Create columnRow
					XSSFRow columnRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
					Integer cellNum2 = 1;
					
					ApprovalStep tempAp=(ApprovalStep)obj.getApprovalStep()[i];
					//スチE��プ番号
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(i+1));
					//名前
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getName()));
					//ラベル
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getLabel()));
					//説昁E
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getDescription()));
					//Seting EntryCriteria
					if(tempAp.getEntryCriteria()!=null){
						//頁E��
						excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getFilterItem(obj.getFullName().substring(0, obj.getFullName().indexOf('.')), tempAp.getEntryCriteria().getCriteriaItems())));
						//数弁E
						excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getEntryCriteria().getFormula()));
						//絞り込み条件
						excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(tempAp.getEntryCriteria().getBooleanFilter()));
					}else{
						excelTemplate.createCell(columnRow,cellNum2++,"");
						excelTemplate.createCell(columnRow,cellNum2++,"");
						excelTemplate.createCell(columnRow,cellNum2++,"");
					}						

					String assignedApprover ="";
					ApprovalStepApprover asa = (ApprovalStepApprover)tempAp.getAssignedApprover();
					Approver[] ap = asa.getApprover();
					for(Integer i2=0;i2<ap.length;i2++){
						if(Util.nullFilter(ap[i2].getType()).equals("adhoc")){
							assignedApprover += Util.getTranslate("SUBMITTERTYPE",Util.nullFilter(ap[i2].getType()))+"\n";
						}else{
							String Usertype=Util.getTranslate("SUBMITTERTYPE",Util.nullFilter(ap[i2].getType()));
							String Username=ut.getUserLabel("Username", ap[i2].getName());
							assignedApprover += Usertype+":"+Username+"\n";
							//assignedApprover += ap[i2].getType()+":"+ap[i2].getName()+"\n";
						}
					}
					//承認老E
					excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(assignedApprover));
					//褁E��承認老E�E場吁E
					excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("WHENMULTIPLEAPPROVERS", Util.nullFilter(asa.getWhenMultipleApprovers())));

					if(tempAp.getApprovalActions()!=null){
						WorkflowActionReference[] wa= tempAp.getApprovalActions().getAction();
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
							//承認時のアクション
							excelTemplate.createCellValue(excelApSheet,rowNo,cellNum2++,hyperVal,displayVal);
						}
					}else{
						excelTemplate.createCell(columnRow,cellNum2++,"");
					}
					if(tempAp.getRejectionActions()!=null){
						for(Integer i4=0; i4<tempAp.getRejectionActions().getAction().length; i4++){
							WorkflowActionReference action = (WorkflowActionReference)tempAp.getRejectionActions().getAction()[i4];
							//Create columnRow
							Integer rowNo = itemNum+i4;
							String cellName ="";
							if(actionNameMap.get(obj.getFullName()+action.getName())==null){
								cellName = Util.makeNameValue(obj.getFullName()+action.getName());
								actionNameMap.put(obj.getFullName()+action.getName(), cellName);
							}else{
								cellName = actionNameMap.get(obj.getFullName()+action.getName());
							}
							String hyperVal = actionsSheetName+"!"+cellName;								
							String displayVal = Util.getTranslate("RejectionActions",String.valueOf(action.getType())) + "." + action.getName();
							//却下時のアクション
							excelTemplate.createCellValue(excelApSheet,rowNo,cellNum2++,hyperVal,displayVal);
						}
					}else{
						excelTemplate.createCell(columnRow,cellNum2++,"");
					}								
				}
				//create final Approval Actions table(最終承認時のアクション)
				excelTemplate.createTableHeaders(excelApSheet,"final Approval Actions",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				if(obj.getFinalApprovalActions()!=null){
					Integer cellNumTem = 1;					
					for(WorkflowActionReference action :obj.getFinalApprovalActions().getAction()){
						XSSFRow actionRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
						String cellName = actionNameMap.get(obj.getFullName()+action.getName());
						String hyperVal = actionsSheetName+"!"+cellName;
						String displayVal = actionLabelMap.get(obj.getFullName()+action.getName());
						//String displayVal = action.getName();
						//種別
						excelTemplate.createCell(actionRow,cellNumTem++,Util.getTranslate("ACTIONTYPE",Util.nullFilter(action.getType())));
						//名前
						excelTemplate.createCellValue(excelApSheet,actionRow.getRowNum(),cellNumTem++,hyperVal,displayVal);
						cellNumTem=1;
						actionlist.add(obj.getFullName()+action.getName());
					}
				}
				//create final Rejection Actions table(最終却下時のアクション)
				excelTemplate.createTableHeaders(excelApSheet,"final Rejection Actions",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				if(obj.getFinalRejectionActions()!=null){
					Integer cellNumTem = 1;					
					for(WorkflowActionReference action :obj.getFinalRejectionActions().getAction()){
						XSSFRow actionRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
						String cellName = actionNameMap.get(obj.getFullName()+action.getName());
						String hyperVal = actionsSheetName+"!"+cellName;
						String displayVal = actionLabelMap.get(obj.getFullName()+action.getName());
						//String displayVal = action.getName();
						//種別
						excelTemplate.createCell(actionRow,cellNumTem++,Util.getTranslate("ACTIONTYPE",Util.nullFilter(action.getType())));
						//名前
						excelTemplate.createCellValue(excelApSheet,actionRow.getRowNum(),cellNumTem++,hyperVal,displayVal);
						cellNumTem=1;
						actionlist.add(obj.getFullName()+action.getName());
					}
				}
				//create recall Actions table(取り消しアクション)
				excelTemplate.createTableHeaders(excelApSheet,"recall Actions",excelApSheet.getLastRowNum()+Util.RowIntervalNum);
				if(obj.getRecallActions()!=null){
					Integer cellNumTem = 1;					
					for(WorkflowActionReference action :obj.getRecallActions().getAction()){
						XSSFRow actionRow = excelApSheet.createRow(excelApSheet.getLastRowNum()+1);
						String cellName = actionNameMap.get(obj.getFullName()+action.getName());
						String hyperVal = actionsSheetName+"!"+cellName;
						String displayVal = actionLabelMap.get(obj.getFullName()+action.getName());
						//String displayVal = action.getName();
						//種別
						excelTemplate.createCell(actionRow,cellNumTem++,Util.getTranslate("ACTIONTYPE",Util.nullFilter(action.getType())));
						//名前
						excelTemplate.createCellValue(excelApSheet,actionRow.getRowNum(),cellNumTem++,hyperVal,displayVal);
						cellNumTem=1;
						actionlist.add(obj.getFullName()+action.getName());
					}
					
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
				//excelTemplate.createCatalogMenu(catalogSheet,excelActionSheet,actionsSheetName,actionDisplayName);
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelActionSheet,Util.cutSheetName(actionsSheetName),actionsSheetName);//dan add
				if(workflowAction!=null){
					for(Metadata meta: workflowAction){
						Workflow objw = (Workflow) meta;
						//Create Task Table(タスク)
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
									//変更あり
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("IsChanged",Util.nullFilter(resultMapWork.get("WorkflowTask."+objw.getFullName()+"."+task.getFullName()))));
									}
									//一意�E名前
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getFullName()));
									//任命允E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getAssignedTo()));
									//承認老E�E選抁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("ActionTaskAssignedToType",Util.nullFilter(task.getAssignedToType())));
									//件吁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getSubject()));
									//期日
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getDueDateOffset()));
									//任命先へ通知
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getNotifyAssignee()));
									//保護コンポ�EネンチE
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getProtected()));
									//状況E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getStatus()));
									//優先度
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getPriority()));
									//説昁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(task.getDescription()));
									}
							}
						}
						//Create Alerts Table(メールアラーチE
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
									//変更あり
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("IsChanged",Util.nullFilter(resultMapWork.get("WorkflowAlert."+objw.getFullName()+"."+alert.getFullName()))));
									}
									//一意�E名前
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getFullName()));
									//説昁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getDescription()));
									//メールチE��プレーチE
									excelTemplate.createCell(columnRow,cellNum2++,ut.getEmailTemplateLabel(Util.nullFilter(alert.getTemplate())));
									//保護コンポ�EネンチE
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getProtected()));
	
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
									//メール受信老E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfEmailRecipient));
									//Alerts CcEmails
									String ccEmails = "";
									for(Integer items=0; items<alert.getCcEmails().length; items++){
										ccEmails += alert.getCcEmails()[items]+"\r\n";
									}
									//追加のメール
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ccEmails));
									//差出人類型
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("ActionEmailSenderType",Util.nullFilter(alert.getSenderType())));
									//差出人メールアドレス
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(alert.getSenderAddress()));
								}
							}
						}
						//Create FieldUpdates Table(頁E��自動更新)
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
									//変更あり
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("IsChanged",Util.nullFilter(resultMapWork.get("WorkflowFieldUpdate."+objw.getFullName()+"."+wffu.getFullName()))));
									}
									//一意�E名前
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getFullName()));
									//名前
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getName()));
									//説昁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getDescription()));
									if(wffu.getTargetObject()==null){
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(apDislayName.substring(0, apDislayName.indexOf('.')))));
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(apDislayName.substring(0, apDislayName.indexOf('.'))+"."+wffu.getField())));
									}else{
										//オブジェクチE
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(wffu.getTargetObject())));
										//更新する頁E��
										excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(ut.getLabelApi(wffu.getTargetObject()+"."+wffu.getField())));
									}
									//頁E��変更後にワークフロールールを�E評価する
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getReevaluateOnChange())));
									//保護コンポ�EネンチE
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getProtected())));
									//条件
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(Util.getTranslate("FieldUpdateOperation",String.valueOf(wffu.getOperation()))));
									//リチE��ル値
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getLiteralValue()));
									//頁E��チE�Eタの種別
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("LookupValueType",Util.nullFilter(wffu.getLookupValueType())));
									//頁E��値
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wffu.getLookupValue()));
									//数弁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(Util.nullFilter(wffu.getFormula())));
									//任命先へ通知
									excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wffu.getNotifyAssignee())));
								}
							}
						}
						//Create OutboundMessage Table(アウトバウンドメチE��ージ)
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
									//変更あり
									if(UtilConnectionInfc.modifiedFlag){
										excelTemplate.createCell(columnRow,cellNum2++,ut.getTranslate("IsChanged", Util.nullFilter(resultMapWork.get("WorkflowOutboundMessage."+objw.getFullName()+"."+wfom.getFullName()))));
									}
									//一意�E名前
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getFullName()));
									//名前
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getName()));
									//説昁E
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getDescription()));
									//エンド�EインチEURL
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getEndpointUrl()));
									//送信ユーザ
									excelTemplate.createCell(columnRow,cellNum2++,ut.getUserLabel("Username", Util.nullFilter(wfom.getIntegrationUser())));
									//保護コンポ�EネンチE
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getProtected()));
									//送信セチE��ョン ID
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getIncludeSessionId()));
									//配信不�EメチE��ージキュー
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getUseDeadLetterQueue()));
									//excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getFields()));//送信するオブジェクトテスト頁E��
									String[] strs=wfom.getFields();
									if(strs!=null&&strs.length>0){
										String showStr="";
										for(String filed: strs){
											showStr+=filed+"\r\n";
										}
										//送信するオブジェクトテスト頁E��
										excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(showStr));
									}else{
										excelTemplate.createCell(columnRow,cellNum++,"");
									}
									//API バ�Eジョン
									excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfom.getApiVersion()));
								}
							}
						}
						//Create KnowledgePublishes Table(ナレチE��アクション)
						excelTemplate.createTableHeaders(excelActionSheet,"Workflow KnowledgePublish",excelActionSheet.getLastRowNum()+Util.RowIntervalNum);
						if(objw.getKnowledgePublishes().length>0){
							for(WorkflowKnowledgePublish wfkp:objw.getKnowledgePublishes()){
								//Create columnRow
								XSSFRow columnRow = excelActionSheet.createRow(excelActionSheet.getLastRowNum()+1);
								String cellName = actionNameMap.get(obj.getFullName()+wfkp.getFullName());																	
								//Create HyperLink name
								excelTemplate.createCellName(cellName,actionsSheetName,excelActionSheet.getLastRowNum()+1);
								Integer cellNum2 = 1;
								//変更あり
								if(UtilConnectionInfc.modifiedFlag){
									excelTemplate.createCell(columnRow,cellNum2++,ut.getTranslate("IsChanged",Util.nullFilter( resultMapWork.get("WorkflowKnowledgePublish."+obj.getFullName()+"."+wfkp.getFullName()))));									
								}
								//名前
								excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfkp.getFullName()));
								//アクション
								excelTemplate.createCell(columnRow,cellNum2++,Util.getTranslate("KnowledgeWorkflowAction",Util.nullFilter(wfkp.getAction())));
								//説昁E
								excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfkp.getDescription()));
								//ラベル
								excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfkp.getLabel()));
								//言誁E
								excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfkp.getLanguage()));
								//保護コンポ�EネンチE
								excelTemplate.createCell(columnRow,cellNum2++,Util.nullFilter(wfkp.getProtected()));
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
