package source;

import java.util.List;
import java.util.Map;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.AnalyticSnapshot;
import com.sforce.soap.metadata.AnalyticSnapshotMapping;
import com.sforce.soap.metadata.ReportSummaryType;


public class ReadAnalyticSnapshotSync {
	private XSSFWorkbook workBook;

	public void readAnalyticSnapshot(String type, List<String> objectsList) throws Exception {
		Util.logger.info("readAnalyticSnapshot Start.");
		Util ut = new Util();
		List<Metadata> mdInfos = ut.readMateData(type,objectsList);
		Map<String, String> resultMap = ut.getComparedResult(type,
				UtilConnectionInfc.getLastUpdateTime());
		Util.nameSequence=0;
		Util.sheetSequence=0;
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();	 
		
		/*** Loop MetaData results ***/
		Integer lastIndex=0;
		for (Metadata md : mdInfos) {
			if (md != null) {
				// Create WorkFlow object
				AnalyticSnapshot obj = (AnalyticSnapshot) md;
				// create analytic snapshot's sheet
				String mapsSheetName = Util.makeSheetName(obj.getFullName());
				XSSFSheet excelMapSheet = excelTemplate.createSheet(Util.cutSheetName(mapsSheetName));
				//create catalog menu
				 excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelMapSheet,Util.cutSheetName(mapsSheetName),mapsSheetName);
				/***Analytic Snapshot attribute****/
				// create analytic snaoshot attribute's table
				//レポート作成スナップショット
				excelTemplate.createTableHeaders(excelMapSheet,"Analytic Snapshot Attribute",excelMapSheet.getLastRowNum()+Util.RowIntervalNum );
				// create ColumnRow 
				XSSFRow attRow = excelMapSheet.createRow(excelMapSheet.getLastRowNum()+1);
				Integer cellNum = 1;
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(attRow,cellNum++,ut.getUpdateFlag(resultMap,"AnalyticSnapshot."+obj.getFullName()));					
				}				
				//レポート作成スナップショット名
				excelTemplate.createCell(attRow,cellNum++,Util.nullFilter(obj.getName()));
				//レポート作成スナップショットの一意の名前
				excelTemplate.createCell(attRow,cellNum++,Util.nullFilter(obj.getFullName()));
				//説明
				excelTemplate.createCell(attRow,cellNum++,Util.nullFilter(obj.getDescription()));
				//グルーピング列
				if(Util.nullFilter(obj.getGroupColumn()).equals("GRAND_SUMMARY")){
					excelTemplate.createCell(attRow,cellNum++,Util.getTranslate("GroupColumn",Util.nullFilter(obj.getGroupColumn())));
				}else{
					excelTemplate.createCell(attRow,cellNum++,ut.getLabelforAll(Util.nullFilter(obj.getGroupColumn())));
				}
				//実行ユーザ
				excelTemplate.createCell(attRow,cellNum++,ut.getUserLabel("Username", Util.nullFilter(obj.getRunningUser())));
				//ソースレポート
				excelTemplate.createCell(attRow,cellNum++,Util.nullFilter(obj.getSourceReport()));
				//対象オブジェクト
				excelTemplate.createCell(attRow,cellNum++,Util.nullFilter(ut.getLabelApi(obj.getTargetObject())));
				
				/*** Analytic Snapshot Mappings ***/
				if (obj.getMappings().length > 0) {
					// create Table of map
					Integer num = excelMapSheet.getLastRowNum()+Util.RowIntervalNum;
					//スナップショット項目の対応付け
					excelTemplate.createTableHeaders(excelMapSheet,"Analytic Snapshot Mapping", num);

					for (Integer wi = 0; wi < obj.getMappings().length; wi++) {
						// Create columnRow
						XSSFRow columnRow = excelMapSheet
								.createRow(excelMapSheet.getLastRowNum() + 1);
						
						Integer cellNum2 = 1;
						// Mapping
						AnalyticSnapshotMapping tempMap = (AnalyticSnapshotMapping) obj.getMappings()[wi];
						ReportSummaryType tempRst =tempMap.getAggregateType();
						//集計タイプ
						if(tempRst!=null){
							excelTemplate.createCell(columnRow,cellNum2++,
									Util.getTranslate("ReportSummaryType",Util.nullFilter(tempRst)));
						}else{
							excelTemplate.createCell(columnRow,cellNum2++,
									Util.getTranslate("ReportSummaryType",Util.nullFilter("None")));
						}
						//ソースレポート列
						excelTemplate.createCell(columnRow,cellNum2++,
								Util.nullFilter(ut.getLabelApi(tempMap.getSourceField())));
						//ソースレポートタイプ
						excelTemplate.createCell(columnRow,cellNum2++,
								Util.getTranslate("REPORTJOBSOURCETYPES",Util.nullFilter(tempMap.getSourceType().name())));
						//対象オブジェクト項目	
						excelTemplate.createCell(columnRow,cellNum2++,
								Util.nullFilter(ut.getLabelApi(tempMap.getTargetField())));					
					}
				}
				excelTemplate.adjustColumnWidth(excelMapSheet);
			}
			if(ut.createExcel(workBook, excelTemplate, type, objectsList.size(), lastIndex)){
				excelTemplate.CreateWorkBook(type);
				workBook = excelTemplate.workBook;
			}				
		}
		Util.logger.info("readAnalyticSnapshot End.");
	}
}
