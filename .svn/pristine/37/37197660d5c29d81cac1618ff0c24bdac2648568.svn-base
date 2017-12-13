package source;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.EscalationAction;
import com.sforce.soap.metadata.EscalationRule;
import com.sforce.soap.metadata.EscalationRules;
import com.sforce.soap.metadata.FilterItem;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.RuleEntry;
import com.sforce.ws.ConnectionException;

public class ReadEscalationRuleSync {

	private XSSFWorkbook workBook;

	public void readEscalationRule(String type,List<String> objectsList) throws Exception{
		Util.logger.info("ReadEscalationRule Start.");
		Util.nameSequence=0;
		Util.sheetSequence=0;
		Util ut = new Util();
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//Create Catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();

		/*** Loop MetaData results ***/
		for (Metadata md : mdInfos) {
			if (md != null) {
				EscalationRules obj =(EscalationRules)md;
				EscalationRule[] etarr = obj.getEscalationRule();
				if(etarr.length>0){
					for(EscalationRule et :etarr){
						/*** Escalation Rule ***/
						String objectName =Util.makeSheetName(obj.getFullName()+"."+et.getFullName());
						XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
						excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
						int cellNum=1;
						//エスカレーションルール
						excelTemplate.createTableHeaders(excelObjectSheet,"Escalation Rule",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						//Create columnRow
						XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						//変更あり
						if(UtilConnectionInfc.modifiedFlag){
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ISCHANGED",Util.nullFilter(resultMap.get("EscalationRule."+obj.getFullName()+"."+et.getFullName()))));
						}
						//ルール名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(et.getFullName()));
						//有効
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE", Util.nullFilter(et.getActive())));
						
						/*** Rule Entry ***/
						//エントリ
					    excelTemplate.createTableHeaders(excelObjectSheet,"Rule Entry",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
					    if(et.getRuleEntry().length>0){
							for(RuleEntry re : (RuleEntry[])et.getRuleEntry()){
								//Create columnRow
								cellNum=1;
								XSSFRow columnRowOne = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//検索条件ロジック
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getBooleanFilter()));
								/*if(re.getCriteriaItems().length>0){
									String entryCriteriaObj ="";
									for(int i=0;i<re.getCriteriaItems().length;i++){
										FilterItem aec = re.getCriteriaItems()[i];
										entryCriteriaObj += aec.getField()+","
												+Util.getTranslate("FilterOperation",String.valueOf(aec.getOperation()))+","
												+aec.getValue()+","+aec.getValueField()+"\r\n";
									}
									excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(entryCriteriaObj));//割り当て条件
								}else{
									excelTemplate.createCell(columnRowOne,cellNum++,"");
								}*/
								//ルール条件
								excelTemplate.createCell(columnRowOne,cellNum++,ut.getFilterItem("Case",re.getCriteriaItems()));
								//入力規則数式
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getFormula()));
								//営業時間指定
								excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("EscalationRuleBusinessHoursSource", Util.nullFilter(re.getBusinessHoursSource())));
								//営業時間
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getBusinessHours()));
								//ケースを初めて変更した後に無効化
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getDisableEscalationWhenModified()));
								//エスカレーション時刻設定方法
								excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("EscalationRuleEscalationStartTime", Util.nullFilter(re.getEscalationStartTime())));
								int num=cellNum;
								EscalationAction[] escalationAction = re.getEscalationAction();
								if(escalationAction.length>0){
									for(int i =0;i<escalationAction.length;i++){
										EscalationAction ea = escalationAction[i];
										if(i==0){
											//割り当て先
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getAssignedTo()));
											//再割り当て先種別
											excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("EscalationRuleAssignedToType", Util.nullFilter(ea.getAssignedToType())));
											String str = "";
											for(int j =0;j< ea.getNotifyEmail().length;j++){
												if(j!=0){
													str +=";";
												}
												str += ea.getNotifyEmail()[j].toString();
											}
											//追加のメール
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(str));
											//経過時間（秒）
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getMinutesToEscalation()));
											//ケース所有者に通知
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getNotifyCaseOwner()));
											//通知メールテンプレート
											excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(ea.getNotifyToTemplate())));
											System.out.println("Util.nullFilter(ea.getNotifyToTemplate())=="+Util.nullFilter(ea.getNotifyToTemplate()));
											System.out.println("ut.getEmailTemplateLabel(Util.nullFilter(ea.getNotifyToTemplate()))="+ut.getEmailTemplateLabel(Util.nullFilter(ea.getNotifyToTemplate())));
											//通知先
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getNotifyTo()));
											//割り当て通知メールテンプレート
											excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(ea.getAssignedToTemplate())));
											cellNum=num;
										}else{
											columnRowOne=excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
											//割り当て先
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getAssignedTo()));
											//再割り当て先種別
											excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("EscalationRuleAssignedToType", Util.nullFilter(ea.getAssignedToType())));
											String str = "";
											for(int j =0;j< ea.getNotifyEmail().length;j++){
												if(j!=0){
													str +=";";
												}
												str += ea.getNotifyEmail()[j].toString();
											}
											//追加のメール
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(str));
											//経過時間(秒)
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getMinutesToEscalation()));
											//ケース所有者に通知
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getNotifyCaseOwner()));
											//通知メールテンプレート
											excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(ea.getNotifyToTemplate())));
											//通知先
											excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(ea.getNotifyTo()));
											//割り当て通知メールテンプレート
											excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(ea.getAssignedToTemplate())));
											cellNum=num;
										}																														
									}
								}else{
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");
									excelTemplate.createCell(columnRowOne,cellNum++,"");									
								}
							}
						}
						excelTemplate.adjustColumnWidth(excelObjectSheet);
					}					
				}
			}
		}
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.debug("***no result to export!!!");
		}
		Util.logger.info("ReadEscalationRule End.");
	}
}
