package source;

import java.io.IOException;
import java.util.List;
//import java.util.Map;




import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.CustomField;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.partner.sobject.SObject;
import com.sforce.ws.ConnectionException;

public class ReadCustomSettingSync {

	private XSSFWorkbook workBook;

	public void readCustomSetting(String type,List<String> objectsList) throws Exception{
		try{
			Util.logger.info("readCustomSetting Start.");
			Util ut = new Util();
			List<Metadata> mdInfos = ut.readMateData("CustomSetting", objectsList);
			//Map<String,String> resultMap = ut.getComparedResult("CustomSetting",UtilConnectionInfc.getLastUpdateTime());
			Util.nameSequence=0;
			Util.sheetSequence=0;
			/*** Get Excel template and create workBook(Common) ***/
			CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
			workBook = excelTemplate.workBook;
			//Create catalog sheet
			//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
			/*** Loop MetaData results ***/
			Integer lastIndex=0;
			for (Metadata md : mdInfos) {
				lastIndex+=1;
				if (md != null) {
					// Create CustomSetting object
					CustomObject obj = (CustomObject) md;
					if(obj.getCustomSettingsType()!=null){
						/**---create Custom Setting Attribute table --*/
						String objectName = Util.makeSheetName(obj.getFullName());
						XSSFSheet excelObjectSheet= excelTemplate.createSheet(Util.cutSheetName(objectName));

						excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelObjectSheet,Util.cutSheetName(objectName),objectName);
						int cellNum = 1;
						//カスタム設定属性
						excelTemplate.createTableHeaders(excelObjectSheet,"Custom Setting Attribute",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						//Create columnRow
						XSSFRow columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
						//表示ラベル
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getLabel()));
						//表示
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("CUSTOMSETTINGSVISIBILITY", Util.nullFilter(obj.getVisibility())));
						//設定種別
						excelTemplate.createCell(columnRow,cellNum++,
								Util.getTranslate("CustomSettingsType", Util.nullFilter(obj.getCustomSettingsType())));
						//説明
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getDescription()));
						/**---------------------------------------------------------*/
						CustomField cfList[] = obj.getFields();
						String querySql = "select Id,";
						if(cfList!=null){
							for(CustomField cf:cfList){
								querySql += (cf.getFullName()+",");
							}
						}
						querySql += ("Name from " + obj.getFullName());
						//added by cheng 15-11-20 start
						querySql += " order by id ";
						//added by cheng 15-11-20 end
						//System.out.println(querySql);
						SObject[] sobjArr = ut.apiQuery(querySql);
						//レコード数
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(sobjArr.length));
						/** -----create CustomField table --*/
						//カスタム項目
						excelTemplate.createTableHeaders(excelObjectSheet,"Custom Field",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						if(cfList!=null){
							for(CustomField cf:cfList){
								cellNum = 1;
								columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//表示ラベル
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getLabel()));
								//項目名
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getFullName()));
								//データ型
								excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("FIELDTYPE",Util.nullFilter(cf.getType())));
								//大文字と小文字を区別する
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE", Util.nullFilter(cf.getCaseSensitive()))));
								//デフォルト値
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getDefaultValue()));
								//説明
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getDescription()));
								//インラインヘルプテキスト
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getInlineHelpText()));
								//暗号化
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getEncrypted()))));
								//数式
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getFormula()));
								//数式内の空白の処理
								excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("TreatBlanksAs", Util.nullFilter(cf.getFormulaTreatBlanksAs())));
								//表示形式
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getDisplayFormat()));
								//外部 ID
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getExternalId()))));
								//文字数
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getLength()));
								//小数点の位置
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getScale()));
								//精度
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getPrecision()));
								//必須項目
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getRequired()))));
								//ユニーク
								excelTemplate.createCell(columnRow,cellNum++,(Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cf.getUnique()))));
								//表示される行
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getVisibleLines()));
								//開始番号
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cf.getStartingNumber()));
							}
						}
						/** -----create Manage table -----*/
						//管理
						excelTemplate.createTableHeaders(excelObjectSheet,"Manage",excelObjectSheet.getLastRowNum()+Util.RowIntervalNum);
						XSSFRow headerRow = excelObjectSheet.getRow(excelObjectSheet.getLastRowNum());
						// write manage table header to excel
						if(cfList!=null){
							for(int i=0;i<cfList.length;i++){
								XSSFCell cell = headerRow.createCell(i+2);
								cell.setCellValue(cfList[i].getLabel());
								cell.setCellStyle(excelTemplate.createCHeaderStyle());
							}
						}
						//write custom setting records to excel
						if(sobjArr!=null){
							for(SObject sobj:sobjArr){
								cellNum = 1;
								columnRow = excelObjectSheet.createRow(excelObjectSheet.getLastRowNum()+1);
								//オブジェクト名
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(sobj.getField("Name")));
								if(cfList!=null){
									for(CustomField cf:cfList){
										excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(sobj.getField(cf.getFullName())));
									}
								}
							}
						}
						/**----------------------------------------------------------*/
						excelTemplate.adjustColumnWidth(excelObjectSheet);
					} 

				}
				if(ut.createExcel(workBook, excelTemplate, type, objectsList.size(), lastIndex)){
					excelTemplate.CreateWorkBook(type);
					workBook = excelTemplate.workBook;
				}					
				
			}
		}catch(Exception e){

		}
		Util.logger.info("readCustomSetting End.");
	}
}
