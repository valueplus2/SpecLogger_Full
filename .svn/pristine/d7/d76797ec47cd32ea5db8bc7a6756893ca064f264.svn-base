package source;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.AnalyticsCloudComponentLayoutItem;
import com.sforce.soap.metadata.LayoutColumn;
import com.sforce.soap.metadata.LayoutHeader;
import com.sforce.soap.metadata.LayoutItem;
import com.sforce.soap.metadata.LayoutSection;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Layout;
import com.sforce.soap.metadata.RelatedListItem;
import com.sforce.soap.metadata.ReportChartComponentLayoutItem;
import com.sforce.ws.ConnectionException;

public class ReadLayoutSync {
	
	private XSSFWorkbook workBook;

	
	public void readLayout(String type,List<String> objectsList) throws Exception{
				
		/*** Get Excel template and create workBook(Common) ***/
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		Util ut = new Util();
		Util.sheetSequence=0;
		Util.nameSequence=0;
		workBook = excelTemplate.workBook;
		//Create Catalog sheet
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String,String> resultMap = ut.getComparedResult(type,UtilConnectionInfc.getLastUpdateTime());
		Map<String,String> layoutdMap = new HashMap<String,String>();
		for (Metadata md : mdInfos) {
			if (md != null) {
				// Create Layout object
				Layout obj = (Layout) md;
				//Create Layout sheet
				String LayoutSheetName = Util.makeSheetName(obj.getFullName());
				XSSFSheet excelLayoutSheet = excelTemplate.createSheet(Util.cutSheetName(LayoutSheetName));					
				//目次を作成
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelLayoutSheet,Util.cutSheetName(LayoutSheetName),LayoutSheetName);
				int cellNum=1;
				//Create Layout Table(ページレイアウト)
				excelTemplate.createTableHeaders(excelLayoutSheet,"Layout",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);
				//Create layoutRow
				XSSFRow layoutRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
				//変更あり
				if(UtilConnectionInfc.modifiedFlag){
					excelTemplate.createCell(layoutRow,cellNum++,ut.getUpdateFlag(resultMap,"Layout."+obj.getFullName()));
											
				}
				String[] strs=obj.getFullName().split("-");
				if(strs.length==2){
					//API 参照名
					excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(strs[0]));
					//名前	
					excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(strs[1]));
				}else if(strs.length==1){
					excelTemplate.createCell(layoutRow,cellNum++,"");	
					excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(strs[1]));	
				}else{
					excelTemplate.createCell(layoutRow,cellNum++,"");	
					excelTemplate.createCell(layoutRow,cellNum++,"");	
				}
				String button1 = "";
				for(String s : obj.getCustomButtons()){
					button1 += s + "\n";
				}
				//カスタムボタン
				excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(button1));
				String button2 = "";
				for(String s : obj.getExcludeButtons()){
					button2 += s + "\n";
				}		
				//除外ボタン	
				excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(button2));			
				String layoutHeader = "" ;				
				 for(LayoutHeader lay : obj.getHeaders()){					 
					 layoutHeader += Util.getTranslate("layoutHeader",String.valueOf(lay));				 
				 }			
				//レイアウトヘッダー
				excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(layoutHeader));
				String mul = "";
				for(String st : obj.getMultilineLayoutFields()){
					mul += st;
				}		
				//商談チームレイアウト
				excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(mul));
				
				String rel = "";
				for(String st : obj.getRelatedObjects()){
					rel += st;
				}
				//関連オブジェクト
				excelTemplate.createCell(layoutRow,cellNum++,Util.nullFilter(rel));	
			    //強調表示パネル
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowHighlightsPanel())));
				//相互関係ログ
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowInteractionLogPanel())));
				//ナレッジサイドバー
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowKnowledgeComponent())));
				//割り当てルール
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowRunAssignmentRulesCheckbox())));
				//割り当てルール自動設定
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getRunAssignmentRulesDefault())));
				//ソリューション情報セクション
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowSolutionSection())));
				//[登録とファイルを添付]ボタンを表示
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowSubmitAndAttachButton())));
				//「メール通知」 チェックボックス
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getShowEmailCheckbox())));
				//「メール通知」 デフォルトで選択
				excelTemplate.createCell(layoutRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(obj.getEmailDefault())));
			    
				//Create LayoutSections Table(レイアウト　セクション)
				excelTemplate.createTableHeaders(excelLayoutSheet,"LayoutSections",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);		
				if(obj.getLayoutSections().length>0){				
					for( Integer t=0; t<obj.getLayoutSections().length; t++ ){					
						//Create sectionRow
						cellNum=1;
						XSSFRow sectionRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);						
						LayoutSection section=(LayoutSection)obj.getLayoutSections()[t];	
						//セクション名
						excelTemplate.createCell(sectionRow,cellNum++,Util.getTranslate("layoutsectionName",Util.nullFilter(section.getLabel())));
						//スタイル
						excelTemplate.createCell(sectionRow,cellNum++,Util.getTranslate("LayoutSection",Util.nullFilter(section.getStyle())));
						//カスタムラベル
						excelTemplate.createCell(sectionRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(section.getCustomLabel())));
						//ヘッダー（詳細ページ）
						excelTemplate.createCell(sectionRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(section.getDetailHeading())));
						//ヘッダー（編集ページ）	
						excelTemplate.createCell(sectionRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(section.getEditHeading())));				
						
					}						
				}
				//create LayoutItem Table(レイアウト項目)
				excelTemplate.createTableHeaders(excelLayoutSheet,"LayoutItem",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);					
			    if(obj.getLayoutSections().length>0){		 
					XSSFRow itemRow;
			    	for( Integer t=0; t<obj.getLayoutSections().length; t++ ){		    		
			    		LayoutSection section=(LayoutSection)obj.getLayoutSections()[t];		    		
						//LayoutSections LayoutColumns
			    		cellNum=1;
			    		if(section.getLabel()!=null&&!(section.getLayoutColumns().length==3)){
			    		itemRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);	
			    		itemRow.createCell(1).setCellValue(section.getLabel());		
			    		itemRow.getCell(1).setCellStyle(excelTemplate.createCHeaderStyle());

			    		Integer beginRow = excelLayoutSheet.getLastRowNum();
			    		Integer maxColumnRow = 0;
						for(Integer columns=0; columns<section.getLayoutColumns().length; columns++){							
							LayoutColumn  col = (LayoutColumn)section.getLayoutColumns()[columns];
							//LayoutColumns LayoutItems
							Integer rowCount=0;
							for(Integer items=0; items<col.getLayoutItems().length; items++){	
								rowCount++;

								if(rowCount>maxColumnRow){
									//Create itemRow
									itemRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);	
									maxColumnRow++;
								}else{
									itemRow = excelLayoutSheet.getRow(beginRow+rowCount);
								}
								if(beginRow==0){
									beginRow = excelLayoutSheet.getLastRowNum();
								}
								LayoutItem  item = (LayoutItem)col.getLayoutItems()[items];
								Integer rowNo = beginRow+rowCount;
								if(item.getField() != null){				
									//itemRow.createCell(columns*2+1).setCellValue(item.getField());										
									//itemRow.createCell(columns*2+2).setCellValue(Util.getTranslate("UiBehavior",String.valueOf(item.getBehavior())));
									//項目
									excelTemplate.createCell(itemRow,columns*2+1,Util.nullFilter(item.getField()));
									//プロパティ
									excelTemplate.createCell(itemRow,columns*2+2,Util.getTranslate("UiBehavior",Util.nullFilter(item.getBehavior())));
								}else if(item.getReportChartComponent()!= null){								
									String hyperVal = Util.makeNameValue(item.getReportChartComponent().getReportName());
									String displayVal =String.valueOf(item.getReportChartComponent().getReportName());
									layoutdMap.put(item.getReportChartComponent().getReportName(), hyperVal);
									excelTemplate.createCellValue(excelLayoutSheet,rowNo,columns*2+1,hyperVal,displayVal);
									excelTemplate.createCell(itemRow,columns*2+2,"");
								}else if(item.getPage() != null){								
									String hyperVal = Util.makeNameValue(item.getPage());
									String displayVal =String.valueOf(item.getPage());
									layoutdMap.put(item.getPage(), hyperVal);
									excelTemplate.createCellValue(excelLayoutSheet,rowNo,columns*2+1,hyperVal,displayVal);
									excelTemplate.createCell(itemRow,columns*2+2,"");
								}else if(item.getScontrol() != null){									
									String hyperVal = Util.makeNameValue(item.getScontrol());
									String displayVal =String.valueOf(item.getScontrol());
									layoutdMap.put(item.getScontrol(), hyperVal);
									excelTemplate.createCellValue(excelLayoutSheet,rowNo,columns*2+1,hyperVal,displayVal);	
									excelTemplate.createCell(itemRow,columns*2+2,"");
								}else if(item.getEmptySpace()){								
									//itemRow.createCell(columns*2+1).setCellValue("Blank");	
									excelTemplate.createCell(itemRow,columns*2+1,"Blank");
									excelTemplate.createCell(itemRow,columns*2+2,"");
								}else if(item.getAnalyticsCloudComponent()!=null){
									String hyperVal = Util.makeNameValue(item.getAnalyticsCloudComponent().getDevName());
									String displayVal = item.getAnalyticsCloudComponent().getDevName();
									layoutdMap.put(item.getAnalyticsCloudComponent().getDevName(), hyperVal);
									excelTemplate.createCellValue(excelLayoutSheet,rowNo,columns*2+1,hyperVal,displayVal);
									excelTemplate.createCell(itemRow,columns*2+2,"");
								}
							}							
						}		    		
			    	}	
			    	}
			    }										
			    //Create Custom Link Table(Custom Link)
			    excelTemplate.createTableHeaders(excelLayoutSheet,"CustomLink",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);
			    Integer rowNum = excelLayoutSheet.getLastRowNum()+1;
			    Integer maxColumnRow2 = 0;	    
			    if(obj.getLayoutSections().length>0){	
			    	XSSFRow itemRow2 = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
			    	boolean isEmpty=true;
			    	for( Integer t=0; t<obj.getLayoutSections().length; t++ ){		    		
			    		LayoutSection section=(LayoutSection)obj.getLayoutSections()[t];		    		
						//LayoutSections LayoutColumns
			    		cellNum=1;	
						for(Integer columns=0; columns<section.getLayoutColumns().length; columns++){	
							Integer temRowNum=rowNum;
							if(section.getLayoutColumns().length==3){							
							LayoutColumn  col = (LayoutColumn)section.getLayoutColumns()[columns];									
								//LayoutColumns LayoutItems
								for(Integer items=0; items<col.getLayoutItems().length; items++){
									isEmpty=false;
									if(items>maxColumnRow2){
										//Create itemRow
										itemRow2 = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);	
										maxColumnRow2++;
									}else{
										itemRow2 = excelLayoutSheet.getRow(items+rowNum);
									}
									LayoutItem  item = (LayoutItem)col.getLayoutItems()[items];
									if(item.getCustomLink() != null){	
										String hyperVal = Util.makeNameValue(item.getCustomLink());
										String displayVal =String.valueOf(item.getCustomLink());
										layoutdMap.put(item.getCustomLink(), hyperVal);
										//項目
										excelTemplate.createCellValue(excelLayoutSheet,itemRow2.getRowNum(),columns+1,hyperVal,displayVal);
									}else if(item.getEmptySpace()){								
										//itemRow.createCell(columns*2+1).setCellValue("Blank");	
										excelTemplate.createCell(itemRow2,columns+1,"Blank");
									}
								}
							}
						}
			    	}
			    	if(isEmpty)
			    		excelLayoutSheet.removeRow(itemRow2);
			    }
			    //Create ItemAdditionalInformation Table(項目追加情報)
				excelTemplate.createTableHeaders(excelLayoutSheet,"ItemAdditionalInformation",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);			
			    if(obj.getLayoutSections().length>0){			    	
			    	for( Integer t=0; t<obj.getLayoutSections().length; t++ ){			    		
			    		LayoutSection section=(LayoutSection)obj.getLayoutSections()[t];
						//LayoutSections LayoutColumns
						for(Integer columns=0; columns<section.getLayoutColumns().length; columns++){							
							LayoutColumn  col = (LayoutColumn)section.getLayoutColumns()[columns];														
							//LayoutColumns LayoutItems
							for(Integer items=0; items<col.getLayoutItems().length; items++){							
								//Create itemAddRow
								XSSFRow itemAddRow;								
								LayoutItem  layoutitem = (LayoutItem)col.getLayoutItems()[items];								
								cellNum=1;
								if(layoutitem.getPage() != null){
									itemAddRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
									//Create HyperLink　Name
									excelTemplate.createCellName(layoutdMap.get(layoutitem.getPage()),LayoutSheetName,excelLayoutSheet.getLastRowNum()+1);									
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getPage()));
									//高さ	
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getHeight()));	
									//幅
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getWidth()));	
									//ラベル表示
									excelTemplate.createCell(itemAddRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(layoutitem.getShowLabel())));	
									//スクロールバー表示	
									excelTemplate.createCell(itemAddRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(layoutitem.getShowScrollbars())));									
									//レポートグラフ設定
									if(layoutitem.getReportChartComponent()!= null){								
										String hyperVal = Util.makeNameValue(layoutitem.getReportChartComponent().getReportName());
										String displayVal =String.valueOf(layoutitem.getReportChartComponent().getReportName());
										layoutdMap.put(layoutitem.getReportChartComponent().getReportName(), hyperVal);
										excelTemplate.createCellValue(excelLayoutSheet,itemAddRow.getRowNum(),cellNum++,hyperVal,displayVal);
									}else{
										excelTemplate.createCell(itemAddRow,cellNum++,"");
									}
									//Analytics Cloud ダッシュボード
									if(layoutitem.getAnalyticsCloudComponent()!=null){
										String hyperVal = Util.makeNameValue(layoutitem.getAnalyticsCloudComponent().getDevName());
										String displayVal =String.valueOf(layoutitem.getAnalyticsCloudComponent().getDevName());
										layoutdMap.put(layoutitem.getAnalyticsCloudComponent().getDevName(), hyperVal);
										excelTemplate.createCellValue(excelLayoutSheet,itemAddRow.getRowNum(),cellNum++,hyperVal,displayVal);
									}else{
										excelTemplate.createCell(itemAddRow,cellNum++,"");
									}
								}
								else if(layoutitem.getCustomLink() != null){
									itemAddRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
									//Create HyperLink　Name
									
									excelTemplate.createCellName(layoutdMap.get(layoutitem.getCustomLink()),LayoutSheetName,excelLayoutSheet.getLastRowNum()+1);
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getCustomLink()));
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getHeight()));								
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getWidth()));								
									excelTemplate.createCell(itemAddRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(layoutitem.getShowLabel())));								
									excelTemplate.createCell(itemAddRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(layoutitem.getShowScrollbars())));		
									if(layoutitem.getReportChartComponent()!= null){								
										String hyperVal = Util.makeNameValue(layoutitem.getReportChartComponent().getReportName());
										String displayVal =String.valueOf(layoutitem.getReportChartComponent().getReportName());
										layoutdMap.put(layoutitem.getReportChartComponent().getReportName(), hyperVal);
										excelTemplate.createCellValue(excelLayoutSheet,itemAddRow.getRowNum(),cellNum++,hyperVal,displayVal);
									}else{
										excelTemplate.createCell(itemAddRow,cellNum++,"");
									}
									if(layoutitem.getAnalyticsCloudComponent()!=null){
										String hyperVal = Util.makeNameValue(layoutitem.getAnalyticsCloudComponent().getDevName());
										String displayVal =String.valueOf(layoutitem.getAnalyticsCloudComponent().getDevName());
										layoutdMap.put(layoutitem.getAnalyticsCloudComponent().getDevName(), hyperVal);
										excelTemplate.createCellValue(excelLayoutSheet,itemAddRow.getRowNum(),cellNum++,hyperVal,displayVal);
									}else{
										excelTemplate.createCell(itemAddRow,cellNum++,"");
									}									
								}else if(layoutitem.getScontrol() != null){
									itemAddRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
									//Create HyperLink　Name
									excelTemplate.createCellName(layoutdMap.get(layoutitem.getScontrol()),LayoutSheetName,excelLayoutSheet.getLastRowNum()+1);
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getScontrol()));
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getHeight()));								
									excelTemplate.createCell(itemAddRow,cellNum++,Util.nullFilter(layoutitem.getWidth()));								
									excelTemplate.createCell(itemAddRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(layoutitem.getShowLabel())));								
									excelTemplate.createCell(itemAddRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(layoutitem.getShowScrollbars())));																			
									if(layoutitem.getReportChartComponent()!= null){								
										String hyperVal = Util.makeNameValue(layoutitem.getReportChartComponent().getReportName());
										String displayVal =String.valueOf(layoutitem.getReportChartComponent().getReportName());
										layoutdMap.put(layoutitem.getReportChartComponent().getReportName(), hyperVal);
										excelTemplate.createCellValue(excelLayoutSheet,itemAddRow.getRowNum(),cellNum++,hyperVal,displayVal);
									}else{
										excelTemplate.createCell(itemAddRow,cellNum++,"");
									}
									if(layoutitem.getAnalyticsCloudComponent()!=null){
										String hyperVal = Util.makeNameValue(layoutitem.getAnalyticsCloudComponent().getDevName());
										String displayVal =String.valueOf(layoutitem.getAnalyticsCloudComponent().getDevName());
										layoutdMap.put(layoutitem.getAnalyticsCloudComponent().getDevName(), hyperVal);
										excelTemplate.createCellValue(excelLayoutSheet,itemAddRow.getRowNum(),cellNum++,hyperVal,displayVal);
									}else{
										excelTemplate.createCell(itemAddRow,cellNum++,"");
									}																								
								}							
							}		    		
				    	}	
			    	}
			    }				

			    //Create ReportChartComponentDetail Table(レポートグラフ設定)
				excelTemplate.createTableHeaders(excelLayoutSheet,"ReportChartComponentDetail",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);
			    if(obj.getLayoutSections().length>0){			    	
			    	for( Integer t=0; t<obj.getLayoutSections().length; t++ ){		    		
			    		LayoutSection section=(LayoutSection)obj.getLayoutSections()[t];			    		
						//LayoutSections LayoutColumns
						for(Integer columns=0; columns<section.getLayoutColumns().length; columns++){							
							LayoutColumn  col = (LayoutColumn)section.getLayoutColumns()[columns];
							//LayoutColumns LayoutItems
							for(Integer items=0; items<col.getLayoutItems().length; items++){							
								LayoutItem  item = (LayoutItem)col.getLayoutItems()[items];												
								//LayoutItems ReportChartComponentLayoutItem
								if(item.getReportChartComponent() != null){		 
									//Create chartRow
									XSSFRow chartRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
									//Create HyperLink　Name
									cellNum=1;
									excelTemplate.createCellName(layoutdMap.get(item.getReportChartComponent().getReportName()),LayoutSheetName,excelLayoutSheet.getLastRowNum()+1);									
									ReportChartComponentLayoutItem  repchart = (ReportChartComponentLayoutItem)item.getReportChartComponent();
									//レポートAPI名
									excelTemplate.createCell(chartRow,cellNum++,Util.nullFilter(repchart.getReportName()));		
									//ユーザがページを開くたびに更新
									excelTemplate.createCell(chartRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(!repchart.getCacheData())));				
									//絞込み項目
									excelTemplate.createCell(chartRow,cellNum++,Util.nullFilter(repchart.getContextFilterableField()));
									//エラー文字列
									excelTemplate.createCell(chartRow,cellNum++,Util.nullFilter(repchart.getError()));
									//エラーのあるグラフを非表示
									excelTemplate.createCell(chartRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(repchart.getHideOnError())));
									//絞り込み条件
									excelTemplate.createCell(chartRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(repchart.getIncludeContext())));
									//タイトル表示
									excelTemplate.createCell(chartRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(repchart.getShowTitle())));
									//サイズ
									excelTemplate.createCell(chartRow,cellNum++,Util.getTranslate("ReportChartComponentSize",Util.nullFilter(repchart.getSize())));
								}							
							}							
						}		    		
			    	}			    	
			    }
			    //create analyticsCloudComponentDetail table(Analytics Cloud ダッシュボードの詳細)
			    excelTemplate.createTableHeaders(excelLayoutSheet,"analyticsCloudComponentDetail",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);
			    LayoutSection[] lss = obj.getLayoutSections();
			    if(lss.length>0){
			    	for(Integer t=0;t<lss.length;t++){
			    		LayoutSection section = lss[t];
			    		LayoutColumn[] lcs = section.getLayoutColumns();
			    		for(LayoutColumn col:lcs){
			    			for(LayoutItem item:col.getLayoutItems()){
			    				if(item.getAnalyticsCloudComponent()!=null){
			    					XSSFRow anaRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
			    					//create link name
			    					cellNum = 1;
			    					excelTemplate.createCellName(layoutdMap.get(item.getAnalyticsCloudComponent().getDevName()), LayoutSheetName, excelLayoutSheet.getLastRowNum()+1);
			    					AnalyticsCloudComponentLayoutItem anaItem = item.getAnalyticsCloudComponent();
			    					//タイプ
			    					excelTemplate.createCell(anaRow, cellNum++, Util.nullFilter(anaItem.getAssetType()));
			    					//名前
			    					excelTemplate.createCell(anaRow, cellNum++, Util.nullFilter(anaItem.getDevName()));
			    					//エラー
			    					excelTemplate.createCell(anaRow, cellNum++, Util.nullFilter(anaItem.getError()));
			    					//フィルタ
			    					excelTemplate.createCell(anaRow, cellNum++, Util.nullFilter(anaItem.getFilter()));
			    					//高さ
			    					excelTemplate.createCell(anaRow, cellNum++, Util.nullFilter(anaItem.getHeight()));
			    					//エラー非表示
			    					excelTemplate.createCell(anaRow, cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(anaItem.getHideOnError())));
			    					//タイトル表示
			    					excelTemplate.createCell(anaRow, cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(anaItem.getShowTitle())));
			    					//幅
			    					excelTemplate.createCell(anaRow, cellNum++, Util.nullFilter(anaItem.getWidth()));
			    				}
			    			}
			    		}
			    	}
			    }
				//Create RelatedListItem Table(関連リスト)
				excelTemplate.createTableHeaders(excelLayoutSheet,"RelatedListItem",excelLayoutSheet.getLastRowNum()+Util.RowIntervalNum);
				if(obj.getRelatedLists().length>0){					
					for( Integer t=0; t<obj.getRelatedLists().length; t++ ){					
						//Create listItemRow
						cellNum=1;
						XSSFRow listItemRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);						
						RelatedListItem list=(RelatedListItem)obj.getRelatedLists()[t];	
						//名前
						excelTemplate.createCell(listItemRow,cellNum++,Util.nullFilter(list.getRelatedList()));
						String button3 = "";
						for(String s : list.getCustomButtons()){
							button3 += s + "\n";
						}
						//カスタムボタン
						excelTemplate.createCell(listItemRow,cellNum++,Util.nullFilter(button3));
						String button4 = "";
						for(String s : list.getExcludeButtons()){
							button4 += s + "\n";
						}		
						//除外したボタン
						excelTemplate.createCell(listItemRow,cellNum++,Util.nullFilter(button4));			
						for(String s : list.getFields()){
							if(s.equals(list.getSortField())){
								s+= "(" + Util.getTranslate("SortOrder", String.valueOf(list.getSortOrder()))+")" ;							
							}
							//項目
							excelTemplate.createCell(listItemRow,cellNum,ut.getLabelforAll(Util.nullFilter(s)));
							XSSFRow tempRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
							for(int i=1;i<=cellNum;i++)
								excelTemplate.createCell(tempRow,i,"");
							listItemRow=tempRow;
						}											
						if(list.getFields().length==0){
							excelTemplate.createCell(listItemRow,cellNum,"");
							XSSFRow tempRow = excelLayoutSheet.createRow(excelLayoutSheet.getLastRowNum()+1);
							for(int i=1;i<cellNum;i++)
								excelTemplate.createCell(tempRow,i,"");
							listItemRow=tempRow;							
						}
						
					}						
				}	
				if(excelLayoutSheet!=null){
					excelTemplate.adjustColumnWidth(excelLayoutSheet);
				}
			}
			else {
				System.out.println("Empty metadata.");
			}				
		}
		
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}else{
			System.out.println("***no result to export!!!");
		}			
	}
}

