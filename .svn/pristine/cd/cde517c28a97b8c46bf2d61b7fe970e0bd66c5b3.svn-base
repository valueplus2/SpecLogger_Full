package source;


import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

import com.sforce.soap.metadata.CustomLabelTranslation;
import com.sforce.soap.metadata.CustomLabels;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.CustomLabel;
import com.sforce.soap.metadata.Translations;
import com.sforce.ws.ConnectionException;

public class ReadCustomLabelSync {

	/**
	 * @param args
	 */
	private XSSFWorkbook workBook;
	
	public void readCustomLabel(String type,List<String> objectList)throws Exception{
		Util.logger.info("ReadCustomLabelSync started");	
		Util.logger.debug("objectsList="+objectList);			
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData(type, objectList);		
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		workBook = excelTemplate.workBook;

		//out = excelTemplate.out;

		//创建目录sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		List<String> objectListTran = Arrays.asList(new String[]{"de","es","fr","it","ja","sv","ko",
				                                                  "zh_TW","zh_CN","pt_BR","nl_NL","da",
				                                                  "th","fi","ru","es_MX","hu","pl","cs",
				                                                  "tr","in","ro","vi","uk","iw","el","bg",
				                                                  "en_GB","ar","no","fr_CA","ka","sr","sh",
				                                                  "sk","en_AU","en_MY","en_IN","en_PH","en_CA",
				                                                  "sl","ro_MD","hr","bs","mk","lv","lt","et",
				                                                  "sq","sh_ME","mt","ga","eu","cy","is","pt_PT",
				                                                  "ms","tl","lb","rm","hy","hi","ur","eo","en_US"});
		for (Metadata md : mdInfos) {
			if (md != null) {
				
				CustomLabels obj = (CustomLabels) md;
				Util.logger.debug("CustomLabels="+obj);					
				String CustomLabelSheetName =Util.makeSheetName( obj.getFullName());
				XSSFSheet excelCustomLabelSheet= excelTemplate.createSheet(Util.cutSheetName(CustomLabelSheetName));
				//创建目录信息
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelCustomLabelSheet,Util.cutSheetName(CustomLabelSheetName),CustomLabelSheetName);
				
				//excelTemplate.createTableHeaders(excelCustomLabelSheet,"Custom Label",0);
				//カスタムラベル
				excelTemplate.createTableHeaders(excelCustomLabelSheet,"Custom Label",excelCustomLabelSheet.getLastRowNum()+Util.RowIntervalNum);
				XSSFRow headerRow = excelCustomLabelSheet.getRow(excelCustomLabelSheet.getLastRowNum());

			    //all translations on the platform
			    List<Metadata> mdInfosTran = ut.readMateData("Translations", objectListTran);
			    //the name of language list
				List<String> nameList = new ArrayList<String>();
				Map<String,String> transMap = new LinkedHashMap<String,String>();
				
				for (Metadata metadata : mdInfosTran) {
					Translations translations = (Translations)metadata;
					Util.logger.debug("Translations="+translations);
					if(translations != null){
						//其他语言的name
						nameList.add(translations.getFullName());
						//所有的label
						CustomLabelTranslation[] labelTranslations =  translations.getCustomLabels();
						if(labelTranslations != null){
							for (CustomLabelTranslation customLabelTranslation : labelTranslations) {
								Util.logger.debug("CustomLabelTranslation="+customLabelTranslation);
								if(customLabelTranslation != null ){									
									transMap.put(translations.getFullName()+"_"+customLabelTranslation.getName(), customLabelTranslation.getLabel());
								}
							}
						}
					}
				}
			    			    
				for(int j=0;j<nameList.size();j++){				
					XSSFCell cell = headerRow.createCell(7+j);
					//set cell value(language name)
					cell.setCellValue(Util.getTranslate("translationLanguage", nameList.get(j))); 
					//set cell style
					cell.setCellStyle(excelTemplate.createCHeaderStyle());
				}								    
				if(obj.getLabels().length>0){
					for( Integer t=0; t<obj.getLabels().length; t++ ){
						//create row
						XSSFRow columnRow = excelCustomLabelSheet.createRow(excelCustomLabelSheet.getLastRowNum()+1);
						CustomLabel tempLabel=(CustomLabel)obj.getLabels()[t];
						Util.logger.debug("CustomLabel="+tempLabel);	
						Integer cellNum = 1;
						//excelTemplate.createCellName("TempLabel."+tempLabel.getFullName(),CustomLabelSheetName,excelCustomLabelSheet.getLastRowNum()+1);
						//名前
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(tempLabel.getFullName()));
						//簡単な説明
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(tempLabel.getShortDescription()));
						//言語
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("translationLanguage",Util.nullFilter(tempLabel.getLanguage())));
						//保護コンポーネント
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BooleanValue",Util.nullFilter(tempLabel.getProtected())));
						//カテゴリ
						excelTemplate.createCell(columnRow,cellNum++,tempLabel.getCategories()!=null?Util.nullFilter(tempLabel.getCategories()):"");
						//値
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(tempLabel.getValue()));
						for(int j=0;j<nameList.size();j++){
							if(transMap.get(nameList.get(j)+"_"+String.valueOf(tempLabel.getFullName()))!=null){
								excelTemplate.createCell(columnRow,cellNum++,transMap.get(nameList.get(j)+"_"+Util.nullFilter(tempLabel.getFullName())));
							}
						}	

					}
				}
				
				excelTemplate.adjustColumnWidth(excelCustomLabelSheet);
			}else {
				Util.logger.warn("Empty metadata.");
			}
		}
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.warn("no result to export!!!.");
		}
		Util.logger.info("ReadCustomLabelSync End.");
	}
	public boolean isHad(List<String> list,String str){
		boolean has = false; 
		for(String temp:list){
			if(temp.equals(str)){
				has = true;
				break;
			}
		}
		return has;
	}
	
}
