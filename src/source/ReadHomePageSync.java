package source;

import java.io.IOException;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.CustomPageWebLink;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.HomePageComponent;
import com.sforce.soap.metadata.HomePageLayout;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.ws.ConnectionException;

public class ReadHomePageSync {

	private XSSFWorkbook workBook;
	private UtilConnectionInfc connecthelp;
	public void readHomePage(String homeType,List<String> homeList) throws Exception{
		Util.logger.info("ReadHomePage Start.");
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(homeType);
		workBook = excelTemplate.workBook;
		
		Map<String,List<String>> multiMap = new HashMap<String,List<String>>();
		
		//创建目录sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		String linksheetname=Util.makeSheetName("Custom Page WebLink");
		String pagesheetname=Util.makeSheetName("Homepage Component");
		String layoutsheetname=Util.makeSheetName("Homepage Layout");
		//Create Custom Page WebLink sheet
		XSSFSheet linkSheet= excelTemplate.createSheet(Util.cutSheetName(linksheetname));
		//Create Table Headers(カスタムリンク)
		excelTemplate.createTableHeaders(linkSheet,"Custom Page WebLink",linkSheet.getLastRowNum()+Util.RowIntervalNum);
		//Create Homepage Component sheet
		XSSFSheet pageSheet= excelTemplate.createSheet(Util.cutSheetName(pagesheetname));
		//Create Table Headers(コンポーネント)
		excelTemplate.createTableHeaders(pageSheet,"Homepage Component",pageSheet.getLastRowNum()+Util.RowIntervalNum);
		//Create Homepage Layout sheet
		XSSFSheet layoutSheet= excelTemplate.createSheet(Util.cutSheetName(layoutsheetname));
		//Create Table Headers(ページレイアウト)
		excelTemplate.createTableHeaders(layoutSheet,"Homepage Layout",layoutSheet.getLastRowNum()+Util.RowIntervalNum);
		
		//创建目录信息
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,linkSheet,Util.cutSheetName(linksheetname),linksheetname);
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,pageSheet,Util.cutSheetName(pagesheetname),pagesheetname);
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,layoutSheet,Util.cutSheetName(layoutsheetname),layoutsheetname);		
		Map<String,String> linkNameMap = new HashMap<String,String>();
		Map<String,String> pageNameMap = new HashMap<String,String>();		
		for(String s:homeList){
			ListMetadataQuery query = new ListMetadataQuery();
			query.setType(s);
			FileProperties[] lmr = connecthelp.getMetadataConnection().listMetadata(
					new ListMetadataQuery[] { query }, Util.API_VERSION);
			if (lmr != null) {
				List<String> allFile = new ArrayList<String>();
				for (FileProperties n : lmr) {
					if(n.getNamespacePrefix()!=null&&n.getManageableState()!=null&&(n.getManageableState().toString().equals("unmanaged")||n.getManageableState().toString().equals("released"))){
						allFile.add(n.getNamespacePrefix()+"__"+URLDecoder.decode(n.getFullName(),"utf-8"));
					}else{
						allFile.add(URLDecoder.decode(n.getFullName(),"utf-8"));
					}
				}
				Collections.sort(allFile);
				multiMap.put(s, allFile);
			}			
		}
		Util.logger.debug(multiMap);
		//Loop Types list循环导出类型
		Set<Map.Entry<String, List<String>>> entryseSet = multiMap.entrySet();
		for (Map.Entry<String, List<String>> entry : entryseSet) {
			if(entry.getValue().size()>0){
				
				String type = entry.getKey();
				List<String> objectsList = entry.getValue();
				List<Metadata> mdInfos = ut.readMateData(type, objectsList);
				Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
				for (Metadata md : mdInfos) {
					/*** Custom Page Web Link ***/
					if( type.equals("CustomPageWebLink")){
						// Create CustomObject object
						CustomPageWebLink obj = (CustomPageWebLink) md;						
						Integer cellNum = 1;			
						//Create columnRow
						XSSFRow linkRow = linkSheet.createRow(linkSheet.getLastRowNum()+1);
						if(UtilConnectionInfc.modifiedFlag){
							//変更あり
							excelTemplate.createCell(linkRow,cellNum++,ut.getUpdateFlag(resultMap,"CustomPageWebLink."+obj.getFullName()));
							
						}
						String cellName = "";
						if(linkNameMap.get(obj.getFullName())==null){
							cellName = Util.makeNameValue(obj.getFullName());
							linkNameMap.put(obj.getFullName(), cellName);
						}else{
							cellName = linkNameMap.get(obj.getFullName());
						}																
						excelTemplate.createCellName(cellName,linksheetname,linkRow.getRowNum()+1);
						//名前
						excelTemplate.createCell(linkRow,cellNum++,Util.nullFilter(obj.getFullName()));
						//表示ラベル
						excelTemplate.createCell(linkRow,cellNum++,Util.nullFilter(obj.getMasterLabel()));
						//内容のソース
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("WebLinkType", Util.nullFilter(obj.getLinkType())));
						//内容
						if(String.valueOf(obj.getLinkType())=="page"){
							excelTemplate.createCell(linkRow,cellNum,Util.nullFilter(obj.getPage()));
						}else
						if(String.valueOf(obj.getLinkType())=="sControl"){
							excelTemplate.createCell(linkRow,cellNum,Util.nullFilter(obj.getScontrol()));
						}else{
							excelTemplate.createCell(linkRow,cellNum,Util.nullFilter(obj.getUrl()));
						}
						cellNum++;
						//使用形態
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("WebLinkAvailability", Util.nullFilter(obj.getAvailability())));
						//説明
						excelTemplate.createCell(linkRow,cellNum++,Util.nullFilter(obj.getDescription()));
						//表示方法
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("WebLinkDisplayType", Util.nullFilter(obj.getDisplayType())));
						//リンクのエンコード
						excelTemplate.createCell(linkRow,cellNum++,Util.nullFilter(obj.getEncodingKey()));
						//メニューバーの表示
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getHasMenubar())));
						//スクロールバーの表示
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getHasScrollbars())));
						//ツールバーの表示
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getHasToolbar())));
						//サイズ変更可能
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getIsResizable())));
						//行選択必須
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getRequireRowSelection())));
						//アドレスバーの表示
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowsLocation())));
						//ステータスバーの表示
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowsStatus())));
						//ウィンドウの位置
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("WebLinkPosition", Util.nullFilter(obj.getPosition())));
						//高さ (ピクセル単位)
						excelTemplate.createCell(linkRow,cellNum++,Util.nullFilter(obj.getHeight()));
						//幅(ピクセル)
						excelTemplate.createCell(linkRow,cellNum++,Util.nullFilter(obj.getWidth()));
						//動作
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("WebLinkWindowType", Util.nullFilter(obj.getOpenType())));
						//保護コンポーネント
						excelTemplate.createCell(linkRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getProtected())));
					}
					excelTemplate.adjustColumnWidth(linkSheet);
					/*** Home Page Component ***/
					if( type.equals("HomePageComponent")){
						// Create CustomObject object
						HomePageComponent obj = (HomePageComponent) md;										
						//Create columnRow
						int cellNum = 1;
						XSSFRow pageRow = pageSheet.createRow(pageSheet.getLastRowNum()+1);
						if(UtilConnectionInfc.modifiedFlag){
							//変更あり
							excelTemplate.createCell(pageRow,cellNum++,Util.getTranslate("ISCHANGED",Util.nullFilter(resultMap.get("HomePageComponent."+obj.getFullName()))));
						}
						String cellName = "";
						if(pageNameMap.get(obj.getFullName())==null){
							cellName = Util.makeNameValue(obj.getFullName());	
							pageNameMap.put(obj.getFullName(), cellName);
						}else{
							cellName =pageNameMap.get(obj.getFullName());
						}						
						excelTemplate.createCellName(cellName,pagesheetname,pageRow.getRowNum()+1);
						//名前
						excelTemplate.createCell(pageRow,cellNum++,Util.nullFilter(obj.getFullName()));
						//種別
						excelTemplate.createCell(pageRow,cellNum++,Util.getTranslate("PageComponentType", Util.nullFilter(obj.getPageComponentType())));
						if(obj.getPage()!=null){
							excelTemplate.createCell(pageRow,cellNum++,Util.nullFilter(obj.getPage()));
						}
						if(obj.getBody()!=null){
							excelTemplate.createCell(pageRow,cellNum++,Util.nullFilter(obj.getBody()));
						}
						if(obj.getLinks().length>0){
							//参数 ： 当前sheet，行数，列数，hyperName,Cell表现值
							Integer cellTem = cellNum;
							String hyperVal0 = linksheetname+"!"+linkNameMap.get(String.valueOf(obj.getLinks()[0]));	
							excelTemplate.createCellValue(pageSheet,pageRow.getRowNum(),cellNum++,hyperVal0,obj.getLinks()[0]);
							//pageRow.createCell(3).setCellValue();
							for(int i=1;i<obj.getLinks().length;i++){
								XSSFRow srow = pageSheet.createRow(pageSheet.getLastRowNum()+1);
								String hyperVal = linksheetname+"!"+linkNameMap.get(obj.getLinks()[i]);						
								excelTemplate.createCellValue(pageSheet,srow.getRowNum(),cellTem,hyperVal,obj.getLinks()[i]);
							}						
						}
						//高さ (ピクセル単位)
						excelTemplate.createCell(pageRow,cellNum++,Util.nullFilter(obj.getHeight()));
						//幅(広/狭)
						excelTemplate.createCell(pageRow,cellNum++,Util.getTranslate("PICTURE",Util.nullFilter(obj.getWidth())));
						//スクロールバーを表示
						excelTemplate.createCell(pageRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowScrollbars())));
						//ラベルを表示
						excelTemplate.createCell(pageRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowLabel())));						

					}
					excelTemplate.adjustColumnWidth(pageSheet);
					/*** Home Page Layout ***/
					if( type.equals("HomePageLayout")){
						// Create CustomObject object
						HomePageLayout obj = (HomePageLayout) md;						
						int cellNum = 1;
						//Create columnRow
						XSSFRow layoutRow = layoutSheet.createRow(layoutSheet.getLastRowNum()+1);
						if(UtilConnectionInfc.modifiedFlag){
							//変更あり
							excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(resultMap.get("HomePageLayout."+obj.getFullName())));
						}
						//名前
						excelTemplate.createCell(layoutRow,cellNum++,obj.getFullName());
						Integer maxRow = obj.getNarrowComponents().length>obj.getWideComponents().length?obj.getNarrowComponents().length:obj.getWideComponents().length;
						for(int i=1;i<maxRow;i++){
							XSSFRow row = layoutSheet.createRow(layoutSheet.getLastRowNum()+1);
							excelTemplate.createCell(row,1,"");
							excelTemplate.createCell(row,2,"");
							if(obj.getNarrowComponents().length>0&&obj.getNarrowComponents().length>i){
								if(obj.getNarrowComponents()[0].contains("standard-")){
									//狭い (左) 列
									excelTemplate.createCell(layoutRow,cellNum,Util.getTranslate("standardComponent", Util.nullFilter(obj.getNarrowComponents()[0])));
								}else{
									String cellName = "";	
									if(pageNameMap.get(obj.getNarrowComponents()[0])==null){
										cellName = Util.makeNameValue(obj.getNarrowComponents()[0]);	
										pageNameMap.put(obj.getNarrowComponents()[0], cellName);
									}else{
										cellName =pageNameMap.get(obj.getNarrowComponents()[0]);
									}										
									String hyperVal= pagesheetname+"!"+cellName;									
									excelTemplate.createCellValue(layoutSheet,layoutRow.getRowNum(),cellNum,hyperVal,obj.getNarrowComponents()[0]);
								}
								if(obj.getNarrowComponents()[i].contains("standard-")){
									//狭い (左) 列
									excelTemplate.createCell(row,cellNum,Util.getTranslate("standardComponent",Util.nullFilter( obj.getNarrowComponents()[i])));
									
								}else{
									String cellName = "";	
									if(pageNameMap.get(obj.getNarrowComponents()[i])==null){
										cellName = Util.makeNameValue(obj.getNarrowComponents()[i]);	
										pageNameMap.put(obj.getNarrowComponents()[i], cellName);
									}else{
										cellName =pageNameMap.get(obj.getNarrowComponents()[i]);
									}										
									String hyperVal= pagesheetname+"!"+pageNameMap.get(obj.getNarrowComponents()[i]);
									excelTemplate.createCellValue(layoutSheet,row.getRowNum(),cellNum,hyperVal,obj.getNarrowComponents()[i]);
								}
							}else{
								excelTemplate.createCell(row,cellNum,"");

							}
							if(obj.getWideComponents().length>0&&obj.getWideComponents().length>i){
								
								if(obj.getWideComponents()[0].contains("standard-")){
									//広い (右) 列
									excelTemplate.createCell(layoutRow,cellNum+1,Util.getTranslate("standardComponent", Util.nullFilter(obj.getWideComponents()[0])));
								}else{
									String cellName = "";	
									if(pageNameMap.get(obj.getWideComponents()[0])==null){
										cellName = Util.makeNameValue(obj.getWideComponents()[0]);	
										pageNameMap.put(obj.getWideComponents()[0], cellName);
									}else{
										cellName =pageNameMap.get(obj.getWideComponents()[0]);
									}										
									String hyperVal= pagesheetname+"!"+cellName;
									excelTemplate.createCellValue(layoutSheet,layoutRow.getRowNum(),cellNum+1,hyperVal,obj.getWideComponents()[0]);
								}
								if(obj.getWideComponents()[i].contains("standard-")){
									//広い (右) 列
									excelTemplate.createCell(row,cellNum+1,Util.getTranslate("standardComponent", Util.nullFilter(obj.getWideComponents()[i])));
								}else{
									String cellName = "";	
									if(pageNameMap.get(obj.getWideComponents()[i])==null){
										cellName = Util.makeNameValue(obj.getWideComponents()[i]);	
										pageNameMap.put(obj.getWideComponents()[i], cellName);
									}else{
										cellName =pageNameMap.get(obj.getWideComponents()[i]);
									}										
									String hyperVal= pagesheetname+"!"+cellName;									
									excelTemplate.createCellValue(layoutSheet,row.getRowNum(),cellNum+1,hyperVal,obj.getWideComponents()[i]);
								}
							}else{
								excelTemplate.createCell(row,cellNum+1,"");
							}
							
						}	
					}
					excelTemplate.adjustColumnWidth(layoutSheet);
				}
			}
		}	
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel("HomePage","");
		}else{
			Util.logger.warn("***no result to export!!!");
		}
		Util.logger.info("ReadHomePage End.");
	}
}
