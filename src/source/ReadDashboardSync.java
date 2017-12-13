package source;

import java.io.FileNotFoundException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import wsc.MetadataLoginUtil;

import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.FolderShare;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Dashboard;
import com.sforce.soap.metadata.DashboardFolder;
import com.sforce.soap.metadata.DashboardComponent;
import com.sforce.soap.metadata.ChartSummary;
import com.sforce.soap.metadata.DashboardComponentSection;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.metadata.DashboardFilter;
import com.sforce.soap.metadata.DashboardFilterOption;
import com.sforce.soap.metadata.DashboardTableColumn;
import com.sforce.soap.metadata.DashboardFilterColumn;
import com.sforce.ws.ConnectionException;

public class ReadDashboardSync {
	
	private XSSFWorkbook workbook;
	private CreateExcelTemplate excelTemplate;
	private MetadataConnection metadataConnection;
	private List<String> list = new ArrayList<String>();
	private Map<String, String> resultMap;
	private Util ut;
	//private XSSFSheet catalogSheet;
	
	public void readDashboardFloder(String type,List<String> objectsList)
			throws Exception {
		Util.logger.info("readDashboardFloder Start."); 
		excelTemplate = new CreateExcelTemplate(type);
		ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		workbook = this.excelTemplate.workBook;
		//catalogSheet = excelTemplate.createCatalogSheet();
		metadataConnection = MetadataLoginUtil.metadataConnection;
		metadataConnection = UtilConnectionInfc.getMetadataConnection();
		// deal reportFolder
		List<Metadata> mdInfos = ut.readMateData("DashboardFolder", objectsList);
		Map<String, String> floderMap = this.getCompare("DashboardFolder",UtilConnectionInfc.getLastUpdateTime());
		String sheetName = Util.makeSheetName("DashboardFolder");
		XSSFSheet dashboardFolderSheet = excelTemplate.createSheet(Util.cutSheetName(sheetName));
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, dashboardFolderSheet,Util.cutSheetName(sheetName),sheetName);
		//Create DashboardFolder Table
		excelTemplate.createTableHeaders(dashboardFolderSheet,
				"DashboardFolder",
				dashboardFolderSheet.getLastRowNum() + Util.RowIntervalNum);
		for (Metadata md : mdInfos) {
			/*** Loop MetaData results ***/
			if (md != null) {
				DashboardFolder dashboardFolder = (DashboardFolder) md;
				/******** Report Folder *********/
				Integer rowNum = dashboardFolderSheet.getLastRowNum() + 1;
				// create a new row to write
				XSSFRow newRow = dashboardFolderSheet.createRow(rowNum);
				int cellNum = 1;
				// create cell and write in data
				if(UtilConnectionInfc.modifiedFlag){
					//変更あり
					excelTemplate.createCell(newRow,cellNum++,ut.getUpdateFlag(floderMap,"DashboardFolder."+dashboardFolder.getFullName()));
					
				}
				//名前
				excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(dashboardFolder.getFullName()));
				//表示ラベル
				excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(dashboardFolder.getName()));
				FolderShare[] fs =  dashboardFolder.getFolderShares();
				for (Integer i = 0; i < fs.length; i++){
					if (i > 0){
						 rowNum = dashboardFolderSheet.getLastRowNum() + 1;
						// create a new row to write
						 newRow = dashboardFolderSheet.createRow(rowNum);
						 cellNum = 1;
						 excelTemplate.createCell(newRow,cellNum++,"");
						 excelTemplate.createCell(newRow,cellNum++,"");
						 excelTemplate.createCell(newRow,cellNum++,"");
					}
                    //共有先種類
					excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("SharedToType",Util.nullFilter(fs[i].getSharedToType())));
					//共有先
					excelTemplate.createCell(newRow,cellNum++,ut.getUserLabel("Username", Util.nullFilter(fs[i].getSharedTo())));
					//アクセス
					excelTemplate.createCell(newRow,cellNum++,Util.getTranslate("AccessLevel", Util.nullFilter(fs[i].getAccessLevel())));
				}
				if(fs.length==0){
					 excelTemplate.createCell(newRow,cellNum++,"");
					 excelTemplate.createCell(newRow,cellNum++,"");
					 excelTemplate.createCell(newRow,cellNum++,"");
				}

			}
		}
		// deal file
		for (String str : objectsList) {
			ListMetadataQuery queries = new ListMetadataQuery();
			queries.setType("Dashboard");
			queries.setFolder(str);
			FileProperties[] fileProperties = metadataConnection.listMetadata(new ListMetadataQuery[] { queries }, 31.0);
			for (FileProperties f : fileProperties) {
				list.add(f.getFullName());
			}
			resultMap = ut.getComparedResult(type,str, UtilConnectionInfc.getLastUpdateTime());
		}	
		this.readDashboard(type, list);
		excelTemplate.adjustColumnWidth(dashboardFolderSheet);
		if (workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null) {
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		} else {
			System.out.println("***no result to export!!!");
		}
		Util.logger.info("readDashboardFloder End.");
	}
	public void readDashboard(String type, List<String> list)
			throws FileNotFoundException {
		Util.logger.info("readDashboard Start.");
		List<Metadata> mdInfc = ut.readMateData("Dashboard", list);	
		for (Metadata m : mdInfc) {
			Dashboard db=(Dashboard)m;
			Map<String,String> dashboardMap = new HashMap<String,String>();
			String dashSheet=Util.makeSheetName(db.getFullName());
			XSSFSheet dashboardSheet = excelTemplate.createSheet(Util.cutSheetName(dashSheet));
			if (m!= null){
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, dashboardSheet,Util.cutSheetName(dashSheet),dashSheet);
				//Create Dash board Table
				excelTemplate.createTableHeaders(dashboardSheet, "Dash board",dashboardSheet.getLastRowNum() + Util.RowIntervalNum);
				XSSFRow columnRow = dashboardSheet.createRow(dashboardSheet.getLastRowNum() + 1);
				int cellNum = 1;
				if(UtilConnectionInfc.modifiedFlag){
					//変更あり
					excelTemplate.createCell(columnRow,cellNum++,ut.getUpdateFlag(resultMap,"Dashboard."+db.getFullName()));
					
				}
				//名前
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getFullName()));
				//タイトル
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getTitle()));
				//タイトルの色
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getTitleColor()));
				//タイトルのサイズ
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getTitleSize()));
				//テキストの色
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getTextColor()));
				//背景グラデーションの向き
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("ChartBackgroundDirection",db.getBackgroundFadeDirection().name())));
				//開始の色
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getBackgroundStartColor()));
				//終了の色
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getBackgroundEndColor()));
				System.out.println("------DashboardFilters="+db.getDashboardFilters());
				System.out.println("------DashboardFilters2="+db.getDashboardFilters().length);
				if(db.getDashboardFilters().length>0){								
					Integer rowNo = dashboardSheet.getLastRowNum();
					String hyperVal = Util.makeNameValue(db.getDashboardFilters().toString());
					String displayVal =db.getDashboardFilters().toString();
					dashboardMap.put(db.getDashboardFilters().toString(), hyperVal);
					//フィルター
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,hyperVal,displayVal);
				}else{
					Integer rowNo = dashboardSheet.getLastRowNum();
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,"","");
				}	
				//種類
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("DashboardType",db.getDashboardType().name())));
				//説明
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getDescription()));
				//実行ユーザ
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(db.getRunningUser()));
			}
			//Create Section Table
			excelTemplate.createTableHeaders(dashboardSheet, "Section",
					dashboardSheet.getLastRowNum() + Util.RowIntervalNum);			
			DashboardComponentSection dcs0=(DashboardComponentSection)db.getLeftSection();
			DashboardComponentSection dcs1=(DashboardComponentSection)db.getMiddleSection();
			DashboardComponentSection dcs2=(DashboardComponentSection)db.getRightSection();
			int cellNum = 1;
			XSSFRow Row = dashboardSheet.createRow(dashboardSheet.getLastRowNum() + 1);
			if (m!= null){
				//左セクション
				if(dcs0!=null){
					excelTemplate.createCell(Row,cellNum++,Util.getTranslate("DashboardComponentSize",Util.nullFilter(dcs0.getColumnSize())));
				}else{
					excelTemplate.createCell(Row,cellNum++,"");					
				}
				//中央セクション
				if(dcs1!=null){
					excelTemplate.createCell(Row,cellNum++,Util.getTranslate("DashboardComponentSize",Util.nullFilter(dcs1.getColumnSize())));
				}else{
					excelTemplate.createCell(Row,cellNum++,"");					
				}
				//右セクション
				if(dcs2!=null){	
					excelTemplate.createCell(Row,cellNum++,Util.getTranslate("DashboardComponentSize",Util.nullFilter(dcs2.getColumnSize())));
				}else{
					excelTemplate.createCell(Row,cellNum++,"");					
				}	
			}
			cellNum = 1;
			//2016/01/28 Modified by WYX START 
			//if((dcs0.getComponents().length!=0) || (dcs1.getComponents().length!=0) || (dcs2.getComponents().length!=0)){
			if((dcs0!=null && dcs0.getComponents().length!=0) || (dcs1!=null && dcs1.getComponents().length!=0) || (dcs2!=null && dcs2.getComponents().length!=0)){
			//2016/01/28 Modified by WYX END
				if(dcs0!=null&&dcs0.getComponents().length!=0){	
					Integer rowNo = dashboardSheet.getLastRowNum()+1;
					String hyperVal =Util.makeNameValue(dcs0.getComponents().toString());
					String displayVal =dcs0.getComponents().toString();
					dashboardMap.put(dcs0.getComponents().toString(), hyperVal);
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,hyperVal,displayVal);
				}else{
					Integer rowNo = dashboardSheet.getLastRowNum()+1;
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,"","");
				}
				if(dcs1!=null&&dcs1.getComponents().length!=0){	
					Integer rowNo = dashboardSheet.getLastRowNum();
					String hyperVal =Util.makeNameValue(dcs1.getComponents().toString());
					String displayVal =dcs1.getComponents().toString();
					dashboardMap.put(dcs1.getComponents().toString(), hyperVal);
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,hyperVal,displayVal);
				}
				else{
					Integer rowNo = dashboardSheet.getLastRowNum();
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,"","");
				}
				if(dcs2!=null&&dcs2.getComponents().length!=0){		
					Integer rowNo = dashboardSheet.getLastRowNum();
					String hyperVal =Util.makeNameValue(dcs2.getComponents().toString());
					String displayVal =dcs2.getComponents().toString();
					dashboardMap.put(dcs2.getComponents().toString(), hyperVal);
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,hyperVal,displayVal);
				}else{
					Integer rowNo = dashboardSheet.getLastRowNum();
					excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,"","");
				}
			}
			//Create DashboardComponent Table
			excelTemplate.createTableHeaders(dashboardSheet, "DashboardComponent",dashboardSheet.getLastRowNum() + Util.RowIntervalNum);
			
			if(dcs0!=null&&dcs0.getComponents().length>0){
				excelTemplate.createCellName(dashboardMap.get(dcs0.getComponents().toString()),dashSheet,dashboardSheet.getLastRowNum()+2);
				outPutComponent(dcs0.getComponents(),dashboardSheet,dashboardMap);
			}	
			if(dcs1!=null&&dcs1.getComponents().length>0){
				excelTemplate.createCellName(dashboardMap.get(dcs1.getComponents().toString()),dashSheet,dashboardSheet.getLastRowNum()+2);
				outPutComponent(dcs1.getComponents(),dashboardSheet,dashboardMap);
			}
			if(dcs2!=null&&dcs2.getComponents().length>0){
				excelTemplate.createCellName(dashboardMap.get(dcs2.getComponents().toString()),dashSheet,dashboardSheet.getLastRowNum()+2);
				outPutComponent(dcs2.getComponents(),dashboardSheet,dashboardMap);
			}
			//Create DashboardFilterOptions Table
			excelTemplate.createTableHeaders(dashboardSheet, "DashboardFilterOptions",dashboardSheet.getLastRowNum() + Util.RowIntervalNum);
			if(db.getDashboardFilters().length>0){
				excelTemplate.createCellName(dashboardMap.get(db.getDashboardFilters().toString()),dashSheet,dashboardSheet.getLastRowNum()+1);
				for( Integer t=0; t<db.getDashboardFilters().length; t++ ){
					DashboardFilter df=(DashboardFilter)db.getDashboardFilters()[t];
					if(df.getDashboardFilterOptions()!=null){
						for( Integer i=0; i<df.getDashboardFilterOptions().length; i++ ){						
							DashboardFilterOption dfo=(DashboardFilterOption)df.getDashboardFilterOptions()[i];
							for (Integer j=0; j< dfo.getValues().length; j++){
								XSSFRow columnRow = dashboardSheet.createRow(dashboardSheet.getLastRowNum() + 1);
								cellNum=1;
								//表示ラベル
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(df.getName()));	
								//演算子
								excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("DashboardFilterOperation",Util.nullFilter(dfo.getOperator())));
								//値
								excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dfo.getValues()[j]));
								
							}
						}
					}
				}
			}
			//Create DashboardTableColumn Table
			excelTemplate.createTableHeaders(dashboardSheet, "DashboardTableColumn",dashboardSheet.getLastRowNum() + Util.RowIntervalNum);
			//2016/01/28 Modified by WYX START 
			//outPutDashboardTableColumn(db.getLeftSection().getComponents(),dashSheet,dashboardSheet,dashboardMap);
			if(db.getLeftSection()!=null){
				outPutDashboardTableColumn(db.getLeftSection().getComponents(),dashSheet,dashboardSheet,dashboardMap);
			}
			//2016/01/28 Modified by WYX END 
			if(db.getMiddleSection()!=null){
				outPutDashboardTableColumn(db.getMiddleSection().getComponents(),dashSheet,dashboardSheet,dashboardMap);
			}
			//2016/01/28 Modified by WYX START 
			//outPutDashboardTableColumn(db.getRightSection().getComponents(),dashSheet,dashboardSheet,dashboardMap);
			if(db.getRightSection()!=null){
				outPutDashboardTableColumn(db.getRightSection().getComponents(),dashSheet,dashboardSheet,dashboardMap);
			}
			//2016/01/28 Modified by WYX END
			excelTemplate.adjustColumnWidth(dashboardSheet);
		}
		Util.logger.info("readDashboard End.");
	}
	public Map<String,String> getCompare(String type,Long lastUpdateTime) throws ConnectionException {
		ListMetadataQuery query = new ListMetadataQuery();
		query.setType(type);
		Map<String, String> map = new LinkedHashMap<String, String>();
		FileProperties[] filePro = metadataConnection.listMetadata(
				new ListMetadataQuery[] { query }, Util.API_VERSION);
		for (FileProperties fPro : filePro) {
			map.put(type + "." + fPro.getFullName(),
					fPro.getLastModifiedDate().getTimeInMillis() > lastUpdateTime ? "TRUE" : "FALSE");
		}
		return map;
	}
	private void outPutComponent(DashboardComponent dcList[],XSSFSheet dashboardSheet,Map<String,String> dashboardMap){
		int cellNum = 1;
		for( Integer t=0; t<dcList.length; t++ ){
			DashboardComponent dc=dcList[t];
			DashboardFilterColumn dfc=new DashboardFilterColumn();
			cellNum=1;
			XSSFRow columnRow = dashboardSheet.createRow(dashboardSheet.getLastRowNum() + 1);
			//タイトル
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getTitle()));
			//ヘッダー
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getHeader()));
			//フッター
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getFooter()));
			//最大軸範囲
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getChartAxisRangeMax()));
			//最小軸範囲
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getChartAxisRangeMin()));
			Integer intChartSummary = dc.getChartSummary().length;
			if(intChartSummary>0){
				ChartSummary cs=(ChartSummary)dc.getChartSummary()[0];
				//グラフの列名
				excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ChartAxis",Util.nullFilter(cs.getColumn())));
				//グラフの集計
				excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ReportSummaryType",Util.nullFilter(cs.getAggregate())));
				//グラフの軸線
				excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cs.getAxisBinding()));
			}else{
				excelTemplate.createCell(columnRow,cellNum++,"");
				excelTemplate.createCell(columnRow,cellNum++,"");
				excelTemplate.createCell(columnRow,cellNum++,"");
			}	
			//種別
			excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("DashboardComponentType",Util.nullFilter(dc.getComponentType())));
			//検索条件列
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dfc.getColumn()));
			if(dc.getDashboardTableColumn().length> 0){								
				Integer rowNo = dashboardSheet.getLastRowNum();
				String hyperVal = Util.makeNameValue(dc.getDashboardTableColumn().toString());
				String displayVal =dc.getDashboardTableColumn().toString();
				dashboardMap.put(dc.getDashboardTableColumn().toString(), hyperVal);
				//テーブルの列
				excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,hyperVal,displayVal);
			}else{
				Integer rowNo = dashboardSheet.getLastRowNum();
				excelTemplate.createCellValue(dashboardSheet,rowNo,cellNum++,"","");
				}
			//表示単位
			excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ChartUnits",Util.nullFilter(dc.getDisplayUnits())));
			//ドリルダウン先URL
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getDrillDownUrl()));
			//ドリル可能フラグ
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getDrillEnabled()));
			//レコード詳細ページ
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getDrillToDetailEnabled()));
			//小グループを「その他」グループと組み合わせる
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(!dc.getExpandOthers()));
			//値の表示
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getShowValues()));
			//% を表示
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getShowPercentage()));
			//詳細のフロート表示
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getEnableHover()));
			//合計の表示
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getShowTotal()));
			String[] gc = dc.getGroupingColumn();
			String strGc = "";
			if (gc.length==0){
			}else if (gc.length==1){
				strGc =  gc[0];
			}else{
				for (Integer i = 0; i < gc.length - 1; i++){
					strGc = strGc + gc[i] + "\n";
				}
				strGc = strGc + gc[gc.length - 1];
			}
			//グループ化列
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(strGc));
			//最小
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getGaugeMin()));
			//ローレンジの色
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getIndicatorLowColor()));
			//ブレークポイント 1
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getIndicatorBreakpoint1()));
			//ミドルレンジの色
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getIndicatorMiddleColor()));
			//ブレークポイント 2
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getIndicatorBreakpoint2()));
			//ハイレンジの色
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getIndicatorHighColor()));
			//最大
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getGaugeMax()));
			//凡例の表示位置
			excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ChartLegendPosition",Util.nullFilter(dc.getLegendPosition())));
			//表示する最大件数
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getMaxValuesDisplayed()));
			//指標ラベル
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getMetricLabel()));
			//Visualforceページ
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getPage()));
			//Visualforceページの高さ (ピクセル単位)
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getPageHeightInPixels()));
			//レポート
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getReport()));
			//カスタム Sコントロール
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getScontrol()));
			//Sコントロールの高さ (ピクセル単位)
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getScontrolHeightInPixels()));
			//Chartsの写真表示フラグ
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getShowPicturesOnCharts()));
			//テーブルの写真表示フラグ
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getShowPicturesOnTables()));
			//行を並び替え
			excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("DashboardComponentFilter",Util.nullFilter(dc.getSortBy())));
			//軸範囲
			excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ChartRangeType",Util.nullFilter(dc.getChartAxisRange())));
			//useReportChart
			excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dc.getUseReportChart()));
			if(intChartSummary>1){
				for( Integer i=1; i<dc.getChartSummary().length; i++ ){
					cellNum=1;
					columnRow = dashboardSheet.createRow(dashboardSheet.getLastRowNum() + 1);
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");

					ChartSummary cs=(ChartSummary)dc.getChartSummary()[i];
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ChartAxis",Util.nullFilter(cs.getColumn())));
					excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ReportSummaryType",Util.nullFilter(cs.getAggregate())));
					excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cs.getAxisBinding()));

					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
					excelTemplate.createCell(columnRow,cellNum++,"");
				}
			}
		}	
	}
	
	private void outPutDashboardTableColumn(DashboardComponent dcList[],String dashSheet,XSSFSheet dashboardSheet,Map<String,String> dashboardMap){
		int cellNum = 1;
		if(dcList.toString()!=null){
			for( Integer t=0; t<dcList.length; t++ ){
				DashboardComponent dc=(DashboardComponent)dcList[t];
				if(dc.getDashboardTableColumn().length>0){
					excelTemplate.createCellName(dashboardMap.get(dc.getDashboardTableColumn().toString()),dashSheet,dashboardSheet.getLastRowNum()+2);
					for( Integer i=0; i<dc.getDashboardTableColumn().length; i++ ){
						
						DashboardTableColumn dtc=(DashboardTableColumn)dc.getDashboardTableColumn()[i];					
						XSSFRow columnRow = dashboardSheet.createRow(dashboardSheet.getLastRowNum() + 1);
						cellNum=1;
						//集計タイプ
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("ReportSummaryType",Util.nullFilter(dtc.getAggregateType())));
						//テーブルの列
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dtc.getColumn()));
						//合計を表示
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(dtc.getShowTotal()));
						//ソート
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("DashboardComponentFilter",Util.nullFilter(dtc.getSortBy())));

					}
				}
			}
		}
	}
}
