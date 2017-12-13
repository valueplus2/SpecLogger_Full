package source;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.AssignmentRule;
import com.sforce.soap.metadata.AssignmentRules;


import com.sforce.soap.metadata.FilterItem;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.RuleEntry;
import com.sforce.ws.ConnectionException;

public class ReadAssignmentRuleSync {

	private XSSFWorkbook workBook;

	public void readAssignmentRule(String type,List<String> objectsList) throws Exception{
		Util.logger.info("readAssignmentRule Start.");
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
		Integer lastIndex=0;
		for (Metadata md : mdInfos) {
			lastIndex+=1;
			if (md != null) {
				AssignmentRules obj =(AssignmentRules)md;
				AssignmentRule[] agarr = obj.getAssignmentRule();
				if(agarr.length>0){
					for(AssignmentRule ag :agarr){
						/*** Assignment Rule ***/
						//割り当てルール
						String objectName =Util.makeSheetName(obj.getFullName()+"."+ag.getFullName());
						XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
						excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
						int cellNum=1;
						excelTemplate.createTableHeaders(excelObjectSheet,"Assignment Rule",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						//Create columnRow
						XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						if(UtilConnectionInfc.modifiedFlag){
							excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"AssignmentRule."+obj.getFullName()+"."+ag.getFullName()));
												
						}
						//ルール名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ag.getFullName()));
						//有効
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(ag.getActive()));
						/*** Rule Entry ***/
						//ルールエントリ
						if(ag.getRuleEntry().length>0){
							excelTemplate.createTableHeaders(excelObjectSheet,"Rule Entry",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
							for(RuleEntry re : (RuleEntry[])ag.getRuleEntry()){
								//Create columnRow
								cellNum=1;
								XSSFRow columnRowOne = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//検索条件ロジック
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getBooleanFilter()));
								
								if(re.getCriteriaItems().length>0){
									String entryCriteriaObj ="";
									for(int i=0;i<re.getCriteriaItems().length;i++){
										FilterItem aec = re.getCriteriaItems()[i];
										entryCriteriaObj += aec.getField()+","+
										Util.getTranslate("FILTEROPERATION", String.valueOf(aec.getOperation()))+","+aec.getValue()+","+aec.getValueField()+"\r\n";
									}
									////検索条件リスト
									excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(entryCriteriaObj));
								}else{
									excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(""));
								}
								//数式
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getFormula()));
								//任命先タイプ
								excelTemplate.createCell(columnRowOne,cellNum++,Util.getTranslate("AssignToLookupValueType", Util.nullFilter(re.getAssignedToType())));
								//任命先
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getAssignedTo()));
								//メールテンプレート
								excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(re.getTemplate())));
								String strTeam ="";
								if(re.getTeam().length>0){
									for(int i=0;i<re.getTeam().length;i++){
										strTeam +=re.getTeam()[i]+"\r\n";
									}
								}
								//ケースチーム
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(strTeam));
								//割り当て完了後ケースチームをリセットするか
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getOverrideExistingTeams()));
								//所有者を再割り当てしない
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getNotifyCcRecipients()));
							}
						}
						excelTemplate.adjustColumnWidth(excelObjectSheet);
					}
				}
				
			}	
			if(ut.createExcel(workBook, excelTemplate, type, objectsList.size(), lastIndex)){
				excelTemplate.CreateWorkBook(type);
				workBook = excelTemplate.workBook;
			}				
		}
		Util.logger.info("readAssignmentRule End.");
	}
}
