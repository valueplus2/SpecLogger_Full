package source;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.StaticResource;
import com.sforce.ws.ConnectionException;

public class ReadStaticResourceSync {
	
	private static XSSFWorkbook workBook;
	private static final int BUFF_SIZE = 2048;
	public void ReadStaticResource(String type,List<String> objectsList) throws Exception{
		Util.logger.info("ReadStaticResource Start.");
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
		
		//Create sheet
		String rulesSheetName =Util.makeSheetName("StaticResource");
		XSSFSheet excelRuleSheet= excelTemplate.createSheet(Util.cutSheetName(rulesSheetName));
		//Create catalog menu
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelRuleSheet,Util.cutSheetName(rulesSheetName),rulesSheetName);
		//Create Table
		excelTemplate.createTableHeaders(excelRuleSheet,"StaticResource",excelRuleSheet.getLastRowNum()+Util.RowIntervalNum);

		/*** Loop MetaData results ***/
		for (Metadata md : mdInfos) {
			if(md != null){
				StaticResource obj= (StaticResource)md; 
				//Create columnRow
				int cellNum=1;
				XSSFRow columnRow = excelRuleSheet.createRow(excelRuleSheet.getLastRowNum()+1);
				if(UtilConnectionInfc.modifiedFlag){
					
					excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"StaticResource."+obj.getFullName()));
					
				}
				//名前
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getFullName()));
				//MIME タイプ
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getContentType()));
				//キャッシュコントロール
				excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("STATICRESOURCECACHECONTROL",Util.nullFilter(obj.getCacheControl())));
				//説明
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(obj.getDescription()));
				//Export File
				ExportStaticResource(obj.getFullName(),obj.getContent());
			}
		}
		if(excelRuleSheet!=null){
			excelTemplate.adjustColumnWidth(excelRuleSheet);
		}
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			Util.logger.warn("***no result to export!!!");
		}
		Util.logger.info("ReadStaticResource End.");
	}
	
	private void ExportStaticResource(String fileName,byte [] source){
		String filePath = UtilConnectionInfc.getDownloadPath()+"\\"+"StaticResource";
		File file = new File(filePath);
		if(!file.exists()){
			file.mkdir();
		}
		//Export File
		FileOutputStream outf;
		try {
			outf = new FileOutputStream(filePath+"\\"+fileName);
			BufferedOutputStream bufferout= new BufferedOutputStream(outf);
			int len = source.length;
			if(len<BUFF_SIZE){
				bufferout.write(source);
			}else{
				int off = 0;
				while(off<len){
					if((len-off)>=BUFF_SIZE)
						bufferout.write(source,off,BUFF_SIZE);
					else
						bufferout.write(source,off,len-off);
					off += BUFF_SIZE;
				}
			}
			bufferout.flush();
			bufferout.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
