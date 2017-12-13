package source;

import java.io.FileNotFoundException;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.CustomTab;
import com.sforce.soap.metadata.Metadata;
import com.sforce.ws.ConnectionException;

public class ReadCustomTabSync {
	private XSSFWorkbook workBook;
	private CreateExcelTemplate excelTemplate;
	Util ut;
	public void readCustomTab(String type , List<String> objectsList) throws Exception{
		Util.logger.info("readCustomTab Start.");
		ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String, String> resultMap = ut.getComparedResult(type, UtilConnectionInfc.getLastUpdateTime());
		excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;
		//XSSFSheet catalog = excelTemplate.createCatalogSheet();
		////Create Sheets
		String sheetname=Util.makeSheetName("CustomTab");
		XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(sheetname));
		
		//Create TableHeaders(カスタムタブ)
		excelTemplate.createTableHeaders(sheet,"CustomTab",sheet.getLastRowNum()+ Util.RowIntervalNum);
		//Create Catalog menu
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet,Util.cutSheetName(sheetname), sheetname);
		for (Metadata m : mdInfos) {
			if(m!=null){
				CustomTab tab = (CustomTab)m;
	
				int cellNum = 1;
				XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(row,cellNum++,ut.getUpdateFlag(resultMap,type+"."+tab.getFullName()));
					
				}
				//タブ名
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getFullName()));
				//excelTemplate.createCell(row,cellNum++,tab.getLabel());
				String cuslabel = "";
				String typeValue = "";
				//2015-6-8 added by duchuanchuan start
				if(tab.getAuraComponent()!=null){
					typeValue = Util.getTranslate("CustomTabType", "auraComponent");
				}
				//2015-6-8 added by duchuanchuan end
				if(tab.getCustomObject()){
					typeValue += Util.getTranslate("CustomTabType", "customObject");
					List<Metadata> mdList = ut.readMateData("CustomObject",Arrays.asList(tab.getFullName()));
					if(!mdList.isEmpty()){
						CustomObject cusObj = (CustomObject)mdList.get(0);
						cuslabel = cusObj.getLabel();
					}
				}
				if(tab.getFlexiPage() != null){
					typeValue += " , " + Util.getTranslate("CustomTabType", "flexiPage");
				}
				if(tab.getPage() != null){
					typeValue += " , " + Util.getTranslate("CustomTabType", "page");
				}
				if(tab.getScontrol() != null){
					typeValue += " , " + Util.getTranslate("CustomTabType", "scontrol");
				}
				if(tab.getUrl()!=null){
					typeValue += " , " + Util.getTranslate("CustomTabType", "url");
				}
				if(typeValue.indexOf(",") == 1){
					typeValue = typeValue.substring(3 ,typeValue.length());
				}
				if(cuslabel==""){
					//表示ラベル
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getLabel()));
				}else{
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(cuslabel));
				}
				//タイプ
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(typeValue));
				String sourceValue = "";
				if(tab.getCustomObject()){
					sourceValue += tab.getFullName();
				}
				if(tab.getFlexiPage() != null){
					sourceValue += " , " + tab.getFlexiPage();
				}
				if(tab.getPage() != null){
					sourceValue += " , " + tab.getPage();
				}
				if(tab.getScontrol() != null){
					sourceValue += " , " + tab.getScontrol();
				}
				if(tab.getUrl()!=null){
					sourceValue += " , " + tab.getUrl();
				}
				if(sourceValue.indexOf(",") == 1){
					sourceValue = sourceValue.substring(3 ,sourceValue.length());
				}
				//コンテンツ
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(sourceValue));
				//説明
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getDescription() != null 
						? URLDecoder.decode(tab.getDescription(),"UTF-8") : null));
				//コンテンツフレームの高さ(ピクセル)
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getFrameHeight()));
				//サイドバー付き
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(tab.getHasSidebar())));
				//this field can not be used in salesforce 31
				//アイコン
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getIcon()));
				//モバイル利用可能
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(tab.getMobileReady())));
				String str1 = tab.getMotif().substring(0,tab.getMotif().indexOf(":")+1);
				String str2 = tab.getMotif().substring(tab.getMotif().indexOf(":")+1);
				str2 = str2.replaceAll(" ","");
				//タブスタイル
				excelTemplate.createCell(row,cellNum++,str1+Util.getTranslate("TABSTYLE",Util.nullFilter(str2)));
				//スプラッシュページのカスタムリンク
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getSplashPageLink()));
				//文字コード
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(tab.getUrlEncodingKey() != null 
						? tab.getUrlEncodingKey().toString():null));
			}
			excelTemplate.adjustColumnWidth(sheet);
			if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
				//excelTemplate.adjustColumnWidth(catalog);
				excelTemplate.exportExcel(type,"");
			}else{
				Util.logger.warn("***no result to export!!!");
			}
			Util.logger.info("readCustomTab End.");
		}
	}
}
