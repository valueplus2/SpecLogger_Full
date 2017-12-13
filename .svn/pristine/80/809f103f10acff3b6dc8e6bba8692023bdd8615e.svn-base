package source;

import java.io.FileNotFoundException;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import wsc.MetadataLoginUtil;

import com.sforce.soap.metadata.ChartSummary;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.FolderShare;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.metadata.Report;
import com.sforce.soap.metadata.ReportAggregate;
import com.sforce.soap.metadata.ReportAggregateReference;
import com.sforce.soap.metadata.ReportBlockInfo;
import com.sforce.soap.metadata.ReportBucketField;
import com.sforce.soap.metadata.ReportBucketFieldValue;
import com.sforce.soap.metadata.ReportChart;
import com.sforce.soap.metadata.ReportColumn;
import com.sforce.soap.metadata.ReportColorRange;
import com.sforce.soap.metadata.ReportCrossFilter;
import com.sforce.soap.metadata.ReportFilter;
import com.sforce.soap.metadata.ReportFilterItem;
import com.sforce.soap.metadata.ReportFolder;
import com.sforce.soap.metadata.ReportSummaryType;
import com.sforce.soap.metadata.ReportTimeFrameFilter;
import com.sforce.ws.ConnectionException;

public class ReadReportSync {
	private XSSFWorkbook workbook;
	private CreateExcelTemplate excelTemplate;
	private MetadataConnection metadataConnection;
	private List<String> list = new ArrayList<String>();
	private Map<String, String> resultMap;
	private Util ut;
	//private XSSFSheet catalogSheet;

	public void readReportFloder(String type, List<String> objectsList)
			throws Exception,
			UnsupportedEncodingException {
		Util.logger.info("readReportFloder started");	
		//Initialization
		excelTemplate = new CreateExcelTemplate(type);
		ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		workbook = this.excelTemplate.workBook;
		//catalogSheet = excelTemplate.createCatalogSheet();
		metadataConnection = MetadataLoginUtil.metadataConnection;
		//remove unfiled$public
		List<String> folderlist = new ArrayList<String>();
		for(String s:objectsList){
			if(s.equals("unfiled$public"))
				continue;
			folderlist.add(s);
		}
		// deal reportFolder
		//レポートフォルダ
		List<Metadata> mdInfos = ut.readMateData("ReportFolder",folderlist);
		System.out.println(mdInfos);
		Map<String, String> floderMap = this.getCompare("ReportFolder",	UtilConnectionInfc.getLastUpdateTime());
		String sheetName =Util.makeSheetName("ReportFolder");
		XSSFSheet reportFolderSheet = excelTemplate.createSheet(Util.cutSheetName(sheetName));
		excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, reportFolderSheet,Util.cutSheetName(reportFolderSheet.getSheetName()),sheetName);
		excelTemplate.createTableHeaders(reportFolderSheet,
				"ReportFolder",
				reportFolderSheet.getLastRowNum() + Util.RowIntervalNum);
		for (Metadata md : mdInfos) {
			/*** Loop MetaData results ***/
			if (md != null) {
				
				ReportFolder reportFolder = (ReportFolder) md;
				/******** Report Folder *********/
				Integer rowNum = reportFolderSheet.getLastRowNum() + 1;
				// create a new row to write
				XSSFRow newRow = reportFolderSheet.createRow(rowNum);
				// create cell and write in data
				int cellNum = 1;
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(newRow,cellNum++,ut.getUpdateFlag(floderMap,"ReportFolder."+reportFolder.getFullName()));
					
				}
				//API名
				excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(reportFolder.getFullName()));
				//表示ラベル
				excelTemplate.createCell(newRow,cellNum++,Util.nullFilter(reportFolder.getName()));				
				FolderShare[] fs =  reportFolder.getFolderShares();
				for (Integer i = 0; i < fs.length; i++){
					if (i > 0){
						 rowNum = reportFolderSheet.getLastRowNum() + 1;
						// create a new row to write
						 newRow = reportFolderSheet.createRow(rowNum);
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
				
				excelTemplate.adjustColumnWidth(reportFolderSheet);
			}
			
		}
		
		// deal file
		for (String str : objectsList) {
			ListMetadataQuery queries = new ListMetadataQuery();
			queries.setType("Report");
			queries.setFolder(str);
			FileProperties[] fileProperties = metadataConnection.listMetadata(
					new ListMetadataQuery[] { queries }, Util.API_VERSION);
			for (FileProperties f : fileProperties) {
				list.add(f.getFullName());
			}
			resultMap = ut.getComparedResult(type,
					str, UtilConnectionInfc.getLastUpdateTime());
			this.readReport(type, list);
		}
		if (workbook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null) {
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		} else {
			Util.logger.warn("no result to export!!!");
		}
		
		Util.logger.info("readReportFloder End");
	}

	public void readReport(String type, List<String> list)
			throws FileNotFoundException {
		Util.logger.info("readReport started");
		List<Metadata> mdInfc = ut.readMateData("Report", list);
//レポートのプロパティ
		for (Metadata m : mdInfc) {
			Report report = (Report) m;
			if(report.getFullName()!=null){
				Util.logger.debug("report="+report);
				String reportSheetName=Util.makeSheetName(report.getFullName());
				XSSFSheet reportSheet = excelTemplate.createSheet(Util.cutSheetName(reportSheetName));
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, reportSheet,Util.cutSheetName(reportSheetName),reportSheetName);
				excelTemplate.createTableHeaders(reportSheet, "ReportAttribute",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				int cellNum = 1;
				XSSFRow row = reportSheet
						.createRow(reportSheet.getLastRowNum() + 1);
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(row,cellNum++,ut.getUpdateFlag(resultMap,type + "." + report.getFullName()));
					
				}
				//レポートの一意の名前
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getFullName()));
				//レポート名
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getName()));
				//形式
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("ReportFormat", Util.nullFilter(report.getFormat())));
				//説明
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getDescription()));
				//レポートタイプ
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getReportType()));
				//ドリルダウンのロール名
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getRoleHierarchyFilter()));
				//最大行数
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getRowLimit()));
				//トレンドレポート現在日表示
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(report.getShowCurrentDate())));
				//詳細表示
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(report.getShowDetails())));
				//ソート列
				excelTemplate.createCell(row,cellNum++,ut.getLabelforAll(Util.nullFilter(report.getSortColumn())));
				//ソート順
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("SORTORDER", Util.nullFilter(report.getSortOrder())));
				//ドリルダウンのテリトリー名
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getTerritoryHierarchyFilter()));
				//ドリルダウンのユーザ名
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getUserFilter()));
				//通貨
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getCurrency()));
				//ディビジョンの使用
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(report.getDivision()));
	
				// ReportAggregate
				//集計項目定義
				excelTemplate.createTableHeaders(reportSheet, "ReportAggregate",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getAggregates().length > 0) {
					ReportAggregate[] aggregate = report.getAggregates();
					for (ReportAggregate a : aggregate) {
						XSSFRow arow = reportSheet.createRow(reportSheet
								.getLastRowNum() + 1);
						cellNum = 1;
						//API参照名
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getDeveloperName()));
						//説明
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getDescription()));
						//データ型
						excelTemplate.createCell(arow,cellNum++, Util.getTranslate("ReportAggregateDatatype", Util.nullFilter(a.getDatatype().toString())));// 必填字段
						//有効
						excelTemplate.createCell(arow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(a.getIsActive())));
						//表示ラベル
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getMasterLabel()));
						//行グループ化レベル
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getAcrossGroupingContext()));
						//カスタム集計項目
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getCalculatedFormula()));
						//列グループ化レベル
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getDownGroupingContext()));
						//少数点以下桁数
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getScale()));
						//レコードタイプ
						excelTemplate.createCell(arow,cellNum++,Util.nullFilter(a.getReportType()));
					}
				}
				// ReportColumn  項目 (列)定義
				excelTemplate.createTableHeaders(reportSheet, "ReportColumn",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getColumns().length > 0) {
					for (ReportColumn c : report.getColumns()) {
						XSSFRow crow = reportSheet.createRow(reportSheet
								.getLastRowNum() + 1);
						cellNum = 1;
						//項目名
						String fieldStr=Util.nullFilter(c.getField());
						String fieldOutput=ut.getLabelforAll(fieldStr);
						if(fieldOutput.equals(fieldStr)){
							if (report.getBuckets().length > 0) {
								for (ReportBucketField rbf : report.getBuckets()) {
									if(rbf.getDeveloperName().equals(fieldStr)){
										fieldOutput=rbf.getMasterLabel();
										break;
									}
								}
							}
							if (report.getAggregates().length > 0) {
								for (ReportAggregate a : report.getAggregates()) {
									if(a.getDeveloperName().equals(fieldStr)){
										fieldOutput=a.getMasterLabel();
										break;
									}
								}
							}
						}
						excelTemplate.createCell(crow,cellNum++,fieldOutput);
						ReportSummaryType[] reportSummaryTypes = c.getAggregateTypes();
						List<String> summaryTypeList = new ArrayList<String>();
						for (int i = 0; reportSummaryTypes != null && i < reportSummaryTypes.length; i++) {
							summaryTypeList.add(Util.getTranslate("ReportSummaryType", reportSummaryTypes[i].toString()));
						}
						String str = summaryTypeList.toString();
						str = str.substring(1, str.length() - 1);
						//集計方法
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(str));
						//色置換え
						excelTemplate.createCell(crow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(c.getReverseColors())));
					    //差異列表示
						excelTemplate.createCell(crow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(c.getShowChanges())));
					}
				}
				//2015-6-8 added by duchuancuan start
				//crossFilters  クロス条件
				excelTemplate.createTableHeaders(reportSheet, "crossFilters",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if(report.getCrossFilters().length>0){
					for(ReportCrossFilter rcf :report.getCrossFilters()){
						XSSFRow crow = reportSheet.createRow(reportSheet
								.getLastRowNum() + 1);
						cellNum = 1;
						String itemsValue = "";
						for (ReportFilterItem filterItem : rcf.getCriteriaItems()) {
							String opstr =filterItem.getColumn() +"\t"+ Util.getTranslate("FilterOperation", String.valueOf(filterItem.getOperator())) + "\t"+filterItem.getValue();
							itemsValue += (opstr+"\n");
							/*for(int i=0;i<filterList.size();i++){
								excelTemplate.createCell(crow,cellNum,
										filterList.get(i));
							}*/
						}
						itemsValue = itemsValue.trim();
						//サブ条件
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(itemsValue));
						//operation
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(rcf.getOperation().toString()));
						//親オブジェクト
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(rcf.getPrimaryTableColumn()));
						//子オブジェクト
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(rcf.getRelatedTable()));
						//結合項目
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(rcf.getRelatedTableJoinColumn()));
					}
				}
				//2015-6-8 added by duchuancuan end
				// ReportFilter 条件付き強調表示
				excelTemplate.createTableHeaders(reportSheet, "ReportColorRange",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getColorRanges().length > 0) {
					for (ReportColorRange c : report.getColorRanges()) {
						XSSFRow crow = reportSheet.createRow(reportSheet
								.getLastRowNum() + 1);
						cellNum = 1;
						//項目名
						String fieldStr=Util.nullFilter(c.getColumnName());
						String fieldOutput=ut.getLabelforAll(fieldStr);
						if(fieldOutput.equals(fieldStr)){
							if (report.getBuckets().length > 0) {
								for (ReportBucketField rbf : report.getBuckets()) {
									if(rbf.getDeveloperName().equals(fieldStr)){
										fieldOutput=rbf.getMasterLabel();
										break;
									}
								}
							}
							if (report.getAggregates().length > 0) {
								for (ReportAggregate a : report.getAggregates()) {
									if(a.getDeveloperName().equals(fieldStr)){
										fieldOutput=a.getMasterLabel();
										break;
									}
								}
							}
						}
						excelTemplate.createCell(crow,cellNum++,fieldOutput);
						//集計方法
						excelTemplate.createCell(crow,cellNum++,Util.getTranslate("ReportSummaryType", Util.nullFilter(c.getAggregate())));
						//ハイ分割値
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(c.getHighBreakpoint()));
						//ハイレンジ色
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(c.getHighColor()));
						//ロー分割値
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(c.getLowBreakpoint()));
						//ローレンジ色
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(c.getLowColor()));
						//ミドルレンジ色
						excelTemplate.createCell(crow,cellNum++,Util.nullFilter(c.getMidColor()));
					}
				}
				
				excelTemplate.createTableHeaders(reportSheet, "ReportFilter",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getFilter() != null) {
					ReportFilter reportFilter = report.getFilter();
					XSSFRow frow = reportSheet.createRow(reportSheet
							.getLastRowNum() + 1);
					cellNum = 1;
					excelTemplate.createCell(frow,cellNum++,Util.nullFilter(reportFilter.getBooleanFilter()));
					excelTemplate.createCell(frow,cellNum++,
							reportFilter.getLanguage() != null ? Util.getTranslate("translationLanguage",reportFilter
									.getLanguage().toString()) : null);
					//List<String> filterList = new ArrayList<String>();
					String itemsValue = "";
					for (ReportFilterItem filterItem : reportFilter.getCriteriaItems()) {
						//filterList.add(Util.getTranslate("FilterOperation", filterItem.getOperator().toString()));
						/*for(int i=0;i<filterList.size();i++){
							excelTemplate.createCell(frow,cellNum,
									filterList.get(i));
						}*/
						String opstr =filterItem.getColumn() +"\t"+ Util.getTranslate("FilterOperation", filterItem.getOperator().toString()) + "\t"+filterItem.getValue();
						itemsValue += (opstr+"\n");
					}
					itemsValue = itemsValue.trim();
					excelTemplate.createCell(frow,cellNum++,Util.nullFilter(itemsValue));
				}
				// ReportTimeFrameFilter 期間指定
				excelTemplate.createTableHeaders(reportSheet,"ReportTimeFrameFilter", reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getTimeFrameFilter() != null) {
					ReportTimeFrameFilter frameFilter = report.getTimeFrameFilter();
					cellNum = 1;
					XSSFRow frow = reportSheet.createRow(reportSheet.getLastRowNum() + 1);
					//対象日付項目
					excelTemplate.createCell(frow,cellNum++,Util.nullFilter(frameFilter.getDateColumn()));
					//期間
					excelTemplate.createCell(frow,cellNum++,Util.nullFilter(Util.getTranslate("UserDateInterval", String.valueOf(frameFilter.getInterval()))));
					SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
					String startDate = frameFilter.getStartDate() != null ? dateFormat.format(frameFilter.getStartDate().getTime()) : "";
					//開始日
					excelTemplate.createCell(frow,cellNum++,Util.nullFilter(startDate));
					String endDate = frameFilter.getEndDate() != null ? dateFormat.format(frameFilter.getEndDate().getTime()) : "";
					//終了日
					excelTemplate.createCell(frow,cellNum++,Util.nullFilter(endDate));
				}
				// Joinedreport
				// System.out.println("Joinedreport:"+report.);
				// excelTemplate.createTableHeaders(reportSheet,
				// "ReportTimeFrameFilter", reportSheet.getLastRowNum()+2);
	
				// ReportBlockInfo 結合レポートブロック定義
				excelTemplate.createTableHeaders(reportSheet, "ReportBlockInfo",reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getBlockInfo() != null) {
					ReportBlockInfo blockInfo = report.getBlockInfo();
					cellNum = 1;
					XSSFRow brow = reportSheet.createRow(reportSheet.getLastRowNum() + 1);
					//ブロックID
					excelTemplate.createCell(brow,cellNum++,Util.nullFilter(blockInfo.getBlockId()));
				//結合エンティティ
					excelTemplate.createCell(brow,cellNum++,Util.nullFilter(blockInfo.getJoinTable()));
					ReportAggregateReference[] aggregateReferences = blockInfo.getAggregateReferences();
					List<String> referenceList = new ArrayList<String>();
					if (aggregateReferences.length > 0) {
						for (ReportAggregateReference reference : aggregateReferences) {
							referenceList.add(reference.getAggregate());
						}
					}
					//カスタム集計項目
					excelTemplate.createCell(brow,cellNum++,
							Util.nullFilter(referenceList.toString().substring(1,
									referenceList.toString().length() - 1)));
				}
				// ReportBucketField バケット項目定義
				excelTemplate.createTableHeaders(reportSheet, "ReportBucketField",
						reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getBuckets().length > 0) {
					for (ReportBucketField rbf : report.getBuckets()) {
						XSSFRow brow = reportSheet.createRow(reportSheet.getLastRowNum() + 1);
						cellNum = 1;
						//API参照名
						excelTemplate.createCell(brow,cellNum++,Util.nullFilter(rbf.getDeveloperName()));
						//表示ラベル
						excelTemplate.createCell(brow,cellNum++,Util.nullFilter(rbf.getMasterLabel()));
						//otherBucketLabel
						excelTemplate.createCell(brow,cellNum++,Util.nullFilter(rbf.getOtherBucketLabel()));
						//ソース項目
						excelTemplate.createCell(brow,cellNum++,Util.nullFilter(rbf.getSourceColumnName()));
						//バケット種別
						excelTemplate.createCell(brow,cellNum++,Util.getTranslate("ReportBucketFieldType",Util.nullFilter( rbf.getBucketType().toString())));
						//空値をゼロ処理
						excelTemplate.createCell(brow,cellNum++,Util.getTranslate("ReportFormulaNullTreatment",Util.getTranslate("BooleanValue", Util.nullFilter(rbf.getNullTreatment()))));
						ReportBucketFieldValue[] bucketFields = rbf.getValues();
						if(bucketFields.length>0){
							String strSourceValue = "";
							for (Integer i = 0; i < bucketFields[0].getSourceValues().length; i++){
								if (strSourceValue != ""){
									strSourceValue = strSourceValue + "\n";
								}
								if (bucketFields[0].getSourceValues()[i].getSourceValue()==null){
									strSourceValue = strSourceValue + Util.nullFilter(bucketFields[0].getSourceValues()[i].getFrom()) + " - " + Util.nullFilter(bucketFields[0].getSourceValues()[i].getTo());
								}else{
									strSourceValue = strSourceValue + Util.nullFilter(bucketFields[0].getSourceValues()[i].getSourceValue());
								}
							}
							//sourceValues
							excelTemplate.createCell(brow,cellNum++,Util.nullFilter(strSourceValue));
						
							excelTemplate.createCell(brow,cellNum++,Util.nullFilter(bucketFields[0].getValue()));
						}else{
							excelTemplate.createCell(brow,cellNum++,"");
							excelTemplate.createCell(brow,cellNum++,"");
						}
						
						if(bucketFields.length>1){
							
							for (Integer i = 1; i < bucketFields.length; i++){
								brow = reportSheet.createRow(reportSheet.getLastRowNum() + 1);
								cellNum = 1;
								excelTemplate.createCell(brow,cellNum++,"");
								excelTemplate.createCell(brow,cellNum++,"");
								excelTemplate.createCell(brow,cellNum++,"");
								excelTemplate.createCell(brow,cellNum++,"");
								excelTemplate.createCell(brow,cellNum++,"");
								excelTemplate.createCell(brow,cellNum++,"");
								String strSourceValue = "";
								for (Integer j = 0; j < bucketFields[i].getSourceValues().length; j++){
									if (strSourceValue != ""){
										strSourceValue = strSourceValue + "\n";
									}
									if (bucketFields[i].getSourceValues()[j].getSourceValue()==null){
										strSourceValue = strSourceValue + Util.nullFilter(bucketFields[i].getSourceValues()[j].getFrom()) + " - " + Util.nullFilter(bucketFields[i].getSourceValues()[j].getTo());
									}else{
										strSourceValue = strSourceValue + Util.nullFilter(bucketFields[i].getSourceValues()[j].getSourceValue());
									}
								}
								excelTemplate.createCell(brow,cellNum++,Util.nullFilter(strSourceValue));
								excelTemplate.createCell(brow,cellNum++,Util.nullFilter(bucketFields[i].getValue()));
							}
						}
						
						
						// String value = "";
						// String sourceVlaue = "";
						// if(rbf.getValues() != null && rbf.getValues().length >
						// 0){
						// ReportBucketFieldValue[] bucketFields = rbf.getValues();
						// for (ReportBucketFieldValue fieldValue : bucketFields) {
						// System.out.println("********************");
						// System.out.println(fieldValue.getValue());
						// System.out.println(Arrays.asList(fieldValue.getSourceValues()).toString());
						// System.out.println("********************");
						// }
						// }
						// brow.createCell(7).setCellValue(Arrays.asList().toString());
					}
				}
				// ReportChart グラフ定義
				excelTemplate.createTableHeaders(reportSheet, "ReportChart",reportSheet.getLastRowNum() + Util.RowIntervalNum);
				if (report.getChart() != null) {
					ReportChart reChart = report.getChart();
					// 创建Report Chart的Table
					cellNum = 1;
					XSSFRow crow = reportSheet.createRow(reportSheet.getLastRowNum() + 1);
					//種別
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(Util.getTranslate("ChartType", reChart.getChartType().toString())));
					//グループ化単位
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getGroupingColumn()));
					//背景色FROM
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getBackgroundColor1()));
					//背景色TO
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getBackgroundColor2()));
					
					//グラデーション
					excelTemplate.createCell(crow,cellNum++,
									Util.getTranslate("ChartBackgroundDirection",Util.nullFilter(reChart.getBackgroundFadeDir())));
					//フロート表示を有効化
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getEnableHoverLabels()));
					//小グループを「その他」グループと組み合わせる
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getExpandOthers()));
					//凡例の表示位置
					excelTemplate.createCell(crow,cellNum++,
							reChart.getLegendPosition() != null ? Util.getTranslate("legendPosition", reChart.getLegendPosition().toString()) : "");
					//グラフの位置
					excelTemplate.createCell(crow,cellNum++,
							reChart.getLocation() != null ? Util.getTranslate("ChartPosition", reChart.getLocation().toString()) : "");
					//グループ化単位
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getSecondaryGroupingColumn()));
					//軸ラベルを表示する
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getShowAxisLabels()));
					//合計の表示
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getShowTotal()));
					//値の表示
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getShowValues()));
					//グラフサイズ
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(Util.getTranslate("ChartPosition",String.valueOf(reChart.getSize()))));
					// summaryAggregate集計方法(~17)
					excelTemplate.createCell(crow,cellNum++,"");
					ChartSummary[] summaryTypes = reChart.getChartSummaries();
					String summarStr = "";
					for (int i = 0; summaryTypes != null && i < summaryTypes.length; i++) {
						if (summarStr != ""){
							summarStr = summarStr + "\n";
						}								
						if (summaryTypes[i].getAggregate() != null) {
							summarStr = summarStr + "aggregate:" + Util.getTranslate("ReportSummaryType", summaryTypes[i].getAggregate().toString());
						}else{
							summarStr = summarStr + "aggregate:" + Util.getTranslate("ReportSummaryType", "NONE");
						}
						summarStr = summarStr + ", axisBinding:" + Util.nullFilter(summaryTypes[i].getAxisBinding());
						summarStr = summarStr + ", column:" + Util.nullFilter(summaryTypes[i].getColumn());
					}
					//集計方法
					excelTemplate.createCell(crow,cellNum++,Util.nullFilter(summarStr));
					//Y軸の範囲TO
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getSummaryAxisManualRangeEnd()));
					//Y軸の範囲From
					excelTemplate.createCell(crow,cellNum++,
									Util.nullFilter(reChart.getSummaryAxisManualRangeStart()));
					//Y軸の範囲
					excelTemplate.createCell(crow,cellNum++,
							Util.getTranslate("ChartRangeType", Util.nullFilter(String.valueOf(reChart.getSummaryAxisRange()))));
					// summaryColumn
					//Y軸
					excelTemplate.createCell(crow,cellNum++,"");
					//テキスト色
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getTextColor()));
					//テキストサイズ
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getTextSize()));
					//グラフタイトル
					excelTemplate.createCell(crow,cellNum++,reChart.getTitle());
					//タイトル色
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getTitleColor()));
					//タイトルサイズ
					excelTemplate.createCell(crow,cellNum++,
							Util.nullFilter(reChart.getTitleSize()));
	
				}
				excelTemplate.adjustColumnWidth(reportSheet);
			}
			list.clear();
		}
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
}
