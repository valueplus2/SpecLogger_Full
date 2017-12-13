package source;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.AutoResponseRule;
import com.sforce.soap.metadata.AutoResponseRules;
import com.sforce.soap.metadata.FilterItem;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.RuleEntry;
import com.sforce.ws.ConnectionException;

public class ReadAutoResponseRuleSync {
	
	private XSSFWorkbook workBook;
	
    public void readAutoResponseRule(String type,List<String> objectsList) throws Exception{
    	Util.logger.info("readAutoResponseRule Start."); 
    	Util ut = new Util();
    	Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData("AutoResponseRules",objectsList);
		Util.logger.debug("type----------="+mdInfos);
		Map<String,String> resultMap = ut.getComparedResult("AutoResponseRules",UtilConnectionInfc.getLastUpdateTime());
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//Create Catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();

		/*** Loop MetaData results ***/
		//自動レスポンスルール
		for (Metadata md : mdInfos) {
			if (md != null) {
				AutoResponseRules obj =(AutoResponseRules)md;
				AutoResponseRule[] autorr = obj.getAutoResponseRule();
				String objectName =Util.makeSheetName(obj.getFullName());
				XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
				excelTemplate.createTableHeaders(excelObjectSheet,"AutoResponse Rule",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
				if(obj.getFullName()!=null){
					for(AutoResponseRule arr :autorr){
						/*** AutoResponse Rule ***/						
						int cellNum=1;
						//Create columnRow
						XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						if(UtilConnectionInfc.modifiedFlag){
							excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"AutoResponseRule."+obj.getFullName()+"."+arr.getFullName()));
												
						}
						//ルール名
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(arr.getFullName()));
						//有効
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(arr.getActive()));
						/*** Rule Entry ***/
						if(arr.getRuleEntry().length>0){
							excelTemplate.createTableHeaders(excelObjectSheet,"Rule Entry",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
							for(RuleEntry re : (RuleEntry[])arr.getRuleEntry()){
								//Create columnRow
								cellNum=1;
								XSSFRow columnRowOne = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//検索条件ロジック
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getBooleanFilter()));
								if(re.getCriteriaItems().length>0){
									String entryCriteriaObj ="";
									for(int i=0;i<re.getCriteriaItems().length;i++){
										FilterItem aec = re.getCriteriaItems()[i];
										entryCriteriaObj += aec.getField()+","+aec.getOperation()+","+aec.getValue()+","+aec.getValueField()+"\r\n";
									}
									//検索条件リスト
									excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(entryCriteriaObj));
								}else{
									excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(""));
								}
								//数式
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getFormula()));
								//返信先メールアドレス
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getReplyToEmail()));
								//差出人メールアドレス
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getSenderEmail()));
								//差出人名前
								excelTemplate.createCell(columnRowOne,cellNum++,Util.nullFilter(re.getSenderName()));
								//メールテンプレート
								excelTemplate.createCell(columnRowOne,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(re.getTemplate())));
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
			Util.logger.info("***no result to export!!!");
		}
		Util.logger.info("readAutoResponseRule End.");
    }
}
