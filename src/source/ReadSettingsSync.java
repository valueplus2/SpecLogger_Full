package source;

import java.io.IOException;
import java.net.URLDecoder;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.metadata.AccountSettings;
import com.sforce.soap.metadata.ActivitiesSettings;
import com.sforce.soap.metadata.CaseSettings;
import com.sforce.soap.metadata.EmailToCaseSettings;
import com.sforce.soap.metadata.EmailToCaseRoutingAddress;
import com.sforce.soap.metadata.CompanySettings;
import com.sforce.soap.metadata.ContractSettings;
import com.sforce.soap.metadata.EntitlementSettings;
import com.sforce.soap.metadata.ForecastingSettings;
import com.sforce.soap.metadata.ForecastingTypeSettings;
import com.sforce.soap.metadata.ChatterAnswersSettings;
import com.sforce.soap.metadata.WebToCaseSettings;
import com.sforce.soap.metadata.MobileSettings;
import com.sforce.soap.metadata.ChatterMobileSettings;
import com.sforce.soap.metadata.DashboardMobileSettings;
import com.sforce.soap.metadata.SFDCMobileSettings;
import com.sforce.soap.metadata.LiveAgentSettings;
import com.sforce.soap.metadata.AddressSettings;
import com.sforce.soap.metadata.BusinessHoursEntry;
import com.sforce.soap.metadata.BusinessHoursSettings;
import com.sforce.soap.metadata.CountriesAndStates;
import com.sforce.soap.metadata.Country;
import com.sforce.soap.metadata.FindSimilarOppFilter;
import com.sforce.soap.metadata.Holiday;
import com.sforce.soap.metadata.IdeasSettings;
import com.sforce.soap.metadata.IpRange;
import com.sforce.soap.metadata.KnowledgeAnswerSettings;
import com.sforce.soap.metadata.KnowledgeCaseSettings;
import com.sforce.soap.metadata.KnowledgeLanguage;
import com.sforce.soap.metadata.KnowledgeLanguageSettings;
import com.sforce.soap.metadata.KnowledgeSettings;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.NetworkAccess;
import com.sforce.soap.metadata.OpportunitySettings;
import com.sforce.soap.metadata.OrderSettings;
import com.sforce.soap.metadata.PasswordPolicies;
import com.sforce.soap.metadata.ProductSettings;
import com.sforce.soap.metadata.QuoteSettings;
import com.sforce.soap.metadata.SecuritySettings;
import com.sforce.soap.metadata.SessionSettings;
import com.sforce.soap.metadata.State;

import java.util.LinkedHashMap;
import java.util.Map;

import com.sforce.ws.ConnectionException;

public class ReadSettingsSync {	
	Util ut = new Util();
	private XSSFWorkbook workBook;	
	public void readSettings(String type,List<String> objectList)throws Exception{	
		Util.logger.info("ReadSetting Start.");
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		//excelTemplate.CreateWorkBook(type);
		Util.nameSequence=0;
		Util.sheetSequence=0;
		//XSSFSheet catalogSheet = excelTemplate.createCatalogSheet();
		workBook = excelTemplate.workBook;		
		for(int k=0;k<objectList.size();k++){
			String types = Util.makeSheetName(objectList.get(k) + "Settings");
			String readTypes = objectList.get(k) + "Settings";
			List<Metadata> mdInfos = ut.readMateData(readTypes,objectList.subList(k, k+1));
			for (Metadata md : mdInfos) {
				//取引先設定
				if("Account".equals(md.getFullName())){
					XSSFSheet excelAccountSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelAccountSettingsSheet,Util.cutSheetName(types),types);
					AccountSettings ats = (AccountSettings)md;
					if(ats!=null){
						int cellNum = 1;
						excelTemplate.createTableHeaders(excelAccountSettingsSheet, "AccountSettings",excelAccountSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
						XSSFRow columnRow = excelAccountSettingsSheet.createRow(excelAccountSettingsSheet.getLastRowNum()+1);
						//[階層の表示] リンクを表示
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ats.getShowViewHierarchyLink())));
						//取引先チームの有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ats.getEnableAccountTeams())));
						//すべてのユーザが取引先所有者レポートを使用する
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ats.getEnableAccountOwnerReport())));
					}
					excelTemplate.adjustColumnWidth(excelAccountSettingsSheet);
				}			
				//活動設定と、カレンダー用のユーザインターフェース設定
				if("Activities".equals(md.getFullName())){
					XSSFSheet excelActivitiesSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelActivitiesSettingsSheet,Util.cutSheetName(types),types);
					ActivitiesSettings ass = (ActivitiesSettings)md;
					if(ass!=null){
						int cellNum = 1;
						excelTemplate.createTableHeaders(excelActivitiesSettingsSheet, "ActivitiesSettings",excelActivitiesSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
						XSSFRow columnRow = excelActivitiesSettingsSheet.createRow(excelActivitiesSettingsSheet.getLastRowNum()+1);
						//グループ ToDo を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableGroupTasks())));
						//サイドバーカレンダーショートカットを有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableSidebarCalendarShortcut())));
						//定期的な行動の作成の有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableRecurringEvents())));
						//定期的な ToDo の作成の有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableRecurringTasks())));
						//活動アラームの有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableActivityReminders())));
						//メール追跡を有効にする
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableEmailTracking())));
						//マルチユーザカレンダービューに行動の詳細を表示
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getShowEventDetailsMultiUserCalendar())));
						//複数日の行動の有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableMultidayEvents())));
						//[ホーム] タブの [カレンダー] セクションの [要請済みミーティングを表示]
						//excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getShowRequestedMeetingsOnHomePage())));
						//ミーティングのお願いにカスタムロゴを表示
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getShowCustomLogoMeetingRequests())));
						//ミーティングのお願いロゴ
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.UrlFilter(ass.getMeetingRequestsLogo())));
						//ユーザが複数取引先責任者をToDoと行動に関連付けられるようにする
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getAllowUsersToRelateMultipleContactsToTasksAndEvents())));
						//カレンダービューでのクリック作成行動を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableClickCreateEvents())));
						//カレンダービューでのドラッグアンドドロップ編集を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableDragAndDropScheduling())));
						//リストビューでのドラッグアンドドロップスケジュール設定を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getEnableListViewScheduling())));
						//ホームページの行動のフロート表示を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getShowHomePageHoverLinksForEvents())));
						//私の ToDo 一覧のリンクのフロート表示を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ass.getShowMyTasksHoverLinks())));
					}
					excelTemplate.adjustColumnWidth(excelActivitiesSettingsSheet);
				}
				
				if("ChatterAnswers".equals(md.getFullName())){
					XSSFSheet excelChatterAnswersSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelChatterAnswersSettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelChatterAnswersSettingsSheet, "ChatterAnswersSettings",excelChatterAnswersSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						ChatterAnswersSettings cas = (ChatterAnswersSettings)m;
						if(cas!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelChatterAnswersSettingsSheet.createRow(excelChatterAnswersSettingsSheet.getLastRowNum()+1);
							//Chatterアンサーを有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEnableChatterAnswers())));
							//Chatterアンサーをポータルに表示
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getShowInPortals())));
							//リッチテキストエディタを有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEnableRichTextEditor())));
							//評価を有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEnableReputation())));
							//メールによる回答の投稿を許可する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEnableAnswerViaEmail())));
							//Facebookのシングルサインオンを有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEnableFacebookSSO())));
							//Facebook認証プロバイダ
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cas.getFacebookAuthProvider()));
							//フォローする質問の最良の回答を選択する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEmailFollowersOnBestAnswer())));
							//フォローする質問に返信する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEmailFollowersOnReply())));
							//それらの質問に非公開返信を送信する (サポートデスク)
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEmailOwnerOnPrivateReply())));
							//所有する質問に返信する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEmailOwnerOnReply())));
							//検索/質問パブリッシャーをインラインで表示する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cas.getEnableInlinePublisher())));
						}
					}
					excelTemplate.adjustColumnWidth(excelChatterAnswersSettingsSheet);
				}
				if("Company".equals(md.getFullName())){
					XSSFSheet excelCompanySettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelCompanySettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelCompanySettingsSheet, "CompanySettings",excelCompanySettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						CompanySettings cs = (CompanySettings)m;
						if(cs!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelCompanySettingsSheet.createRow(excelCompanySettingsSheet.getLastRowNum()+1);
							//会計年度の表記
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("FISCALYEARNAME",Util.nullFilter(cs.getFiscalYear().getFiscalYearNameBasedOn())));
							//会計年度期首月
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("STARTMONTH",Util.nullFilter(cs.getFiscalYear().getStartMonth())));
						}
					}
					excelTemplate.adjustColumnWidth(excelCompanySettingsSheet);
				}
				if("Contract".equals(md.getFullName())){
					XSSFSheet excelContractSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelContractSettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelContractSettingsSheet, "ContractSettings",excelContractSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						ContractSettings cs = (ContractSettings)m;
						if(cs!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelContractSettingsSheet.createRow(excelContractSettingsSheet.getLastRowNum()+1);
							//契約終了日の自動計算
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cs.getAutoCalculateEndDate())));
							//契約終了通知メールを取引先と契約所有者に送信する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cs.getNotifyOwnersOnContractExpiration())));
							//すべての状況の履歴管理
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cs.getEnableContractHistoryTracking())));
							//期限切れ契約
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cs.getAutoExpireContracts())));
							//自動期限延長
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cs.getAutoExpirationDelay()));
							//自動期限受信
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cs.getAutoExpirationRecipient()));		
						}
					}
					excelTemplate.adjustColumnWidth(excelContractSettingsSheet);
				}
				if("Entitlement".equals(md.getFullName())){
					XSSFSheet excelEntitlementSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelEntitlementSettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelEntitlementSettingsSheet, "EntitlementSettings",excelEntitlementSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						EntitlementSettings es = (EntitlementSettings)m;
						if(es!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelEntitlementSettingsSheet.createRow(excelEntitlementSettingsSheet.getLastRowNum()+1);
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getEnableEntitlements())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getEnableEntitlementVersioning())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getEntitlementLookupLimitedToActiveStatus())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getEntitlementLookupLimitedToSameAccount())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getEntitlementLookupLimitedToSameAsset())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getEntitlementLookupLimitedToSameContact())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getAssetLookupLimitedToActiveEntitlementsOnAccount())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getAssetLookupLimitedToActiveEntitlementsOnContact())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getAssetLookupLimitedToSameAccount())));
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(es.getAssetLookupLimitedToSameContact())));
						}
					}
					excelTemplate.adjustColumnWidth(excelEntitlementSettingsSheet);
				}
				//サポート設定
				if("Case".equals(md.getFullName())){
					XSSFSheet excelCaseSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelCaseSettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelCaseSettingsSheet, "CaseSettings",excelCaseSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						CaseSettings cst = (CaseSettings)m;
						if(cst!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelCaseSettingsSheet.createRow(excelCaseSettingsSheet.getLastRowNum()+1);
							//デフォルトのケース所有者種別
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cst.getDefaultCaseOwnerType()));
							//デフォルトのケース所有者
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cst.getDefaultCaseOwner()));
							//デフォルトのケース所有者に通知する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getNotifyDefaultCaseOwner())));
							//自動ケース更新ユーザ
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(cst.getDefaultCaseUser()));
							//ケース作成時のテンプレート
							excelTemplate.createCell(columnRow,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(cst.getCaseCreateNotificationTemplate())));
							//ケース割り当て時のテンプレート
							excelTemplate.createCell(columnRow,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(cst.getCaseAssignNotificationTemplate())));
							//ケースクローズ時のテンプレート
							excelTemplate.createCell(columnRow,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(cst.getCaseCloseNotificationTemplate())));
							//取引先責任者へのケースコメント通知を有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getNotifyContactOnCaseComment())));
							//ケースコメントのテンプレート
							excelTemplate.createCell(columnRow,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(cst.getCaseCommentNotificationTemplate())));
							//ケース所有者に新規ケースコメントを通知する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getNotifyOwnerOnCaseComment())));
							//アーリートリガの有効
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getEnableEarlyEscalationRuleTriggers())));
							//推奨ソリューションを有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getEnableSuggestedSolutions())));
							//ケース通知をシステムアドレスから送信
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getUseSystemEmailAddress())));
						    //ケース所有権の変更時にケース所有者に通知
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getNotifyOwnerOnCaseOwnerChange())));
							//クローズケースの状況項目を表示
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getCloseCaseThroughStatusChange())));
							//ケースフィード有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getEnableCaseFeed())));
							//デフォルトメールテンプレートまたはメールアクションのデフォルトハンドラを有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cst.getEnableEmailActionDefaultsHandler())));
						}
					}
					// caseFeedItemSettings
					/*excelTemplate.createTableHeaders(excelCaseSettingsSheet, "FeedItemSettings", excelCaseSettingsSheet.getLastRowNum()+1);
					for(Metadata m:mdInfos){
						CaseSettings cst = (CaseSettings)m;
						Util.logger.debug("======"+cst.getCaseFeedItemSettings().length);
						FeedItemSettings[] fiss = cst.getCaseFeedItemSettings();
						if(fiss.length>0){
							for(int i = 0;i<fiss.length;i++){
								FeedItemSettings fis = fiss[i];
								if(fis!=null){
									int cellNum = 1;
									XSSFRow columnRow = excelCaseSettingsSheet.createRow(excelCaseSettingsSheet.getLastRowNum()+1);
									excelTemplate.createCell(columnRow, cellNum, String.valueOf(fis.getCharacterLimit()));
									excelTemplate.createCell(columnRow, cellNum, String.valueOf(fis.getCollapseThread()));
									excelTemplate.createCell(columnRow, cellNum, Util.getTranslate("FeedItemDisplayFormat", Util.nullFilter(fis.getDisplayFormat())));
									excelTemplate.createCell(columnRow,cellNum,Util.getTranslate("FeedItemType", Util.nullFilter(fis.getFeedItemType())));
								}
							}
						}
					}*/
					//メール-to-ケースの設定
					excelTemplate.createTableHeaders(excelCaseSettingsSheet, "EmailToCaseSettings",excelCaseSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						CaseSettings cst= (CaseSettings)m;
						EmailToCaseSettings etcs  = cst.getEmailToCase();
						if(etcs!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelCaseSettingsSheet.createRow(excelCaseSettingsSheet.getLastRowNum()+1);
							//メール-to-ケースの有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getEnableEmailToCase())));
							//ケース所有者に新規メールを通知する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getNotifyOwnerOnNewCaseEmail())));
							//HTML メールの有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getEnableHtmlEmail())));
							//メール件名にスレッド ID を挿入する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getEnableThreadIDInSubject())));
							//メール内容にスレッド ID を挿入する
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getEnableThreadIDInBody())));
							//オンデマンドサービスの有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getEnableOnDemandEmailToCase())));
							//上限を超えたメールメッセージを受信した時のアクション
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("EmailToCaseOnFailureActionType",Util.nullFilter(etcs.getOverEmailLimitAction())));
							//メールスレッドの前にユーザの署名を配置
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcs.getPreQuoteSignature())));
							//送信者が許可されていない時のアクション
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("EmailToCaseOnFailureActionType",Util.nullFilter(etcs.getUnauthorizedSenderAction())));
						}
					}
					//ルーティングアドレス
					excelTemplate.createTableHeaders(excelCaseSettingsSheet, "EmailToCaseRoutingAddress",excelCaseSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						CaseSettings cst= (CaseSettings)m;					
						EmailToCaseSettings etcs  = cst.getEmailToCase();
//						etcs.getRoutingAddresses().
						if(etcs.getRoutingAddresses().length>0){
							for( Integer a=0; a<etcs.getRoutingAddresses().length; a++ ){
								EmailToCaseRoutingAddress etcra=(EmailToCaseRoutingAddress)etcs.getRoutingAddresses()[a];
								if(etcra!=null){
									int cellNum = 1;
									XSSFRow columnRow = excelCaseSettingsSheet.createRow(excelCaseSettingsSheet.getLastRowNum()+1);
									//アドレスタイプ
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getAddressType()));					
									//ソース
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getAddressType()));
									//ルーティング名
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getRoutingName()));					
									//メールサービスアドレス
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getEmailAddress()));					
									//メールヘッダーの保存
									excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcra.getSaveEmailHeaders())));			
									//許可する送信元
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getAuthorizedSenders()));					
									//メールからの ToDo の作成
									excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(etcra.getCreateTask())));					
									//ToDoの状況
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getTaskStatus()));
									//デフォルトのケース所有者種類
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getCaseOwnerType()));
									//デフォルトのケース所有者
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getCaseOwner()));
									//ケース優先度
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getCasePriority()));
									//ケース発生源
									excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(etcra.getCaseOrigin()));			
								}
							}
						}
					}
					//Web-to-ケース設定
					excelTemplate.createTableHeaders(excelCaseSettingsSheet, "WebToCaseSettings",excelCaseSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						CaseSettings cst= (CaseSettings)m;
						WebToCaseSettings wtcs  = cst.getWebToCase();	
						if(cst!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelCaseSettingsSheet.createRow(excelCaseSettingsSheet.getLastRowNum()+1);
							//Web-to-ケースの有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(wtcs.getEnableWebToCase())));
							//デフォルトのケース 発生源
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(wtcs.getCaseOrigin()));
							//デフォルトのレスポンス用テンプレート
							excelTemplate.createCell(columnRow,cellNum++,ut.getEmailTemplateLabel(Util.nullFilter(wtcs.getDefaultResponseTemplate())));
						}
					}
					excelTemplate.adjustColumnWidth(excelCaseSettingsSheet);
				}	
				//売上予測
			if("Forecasting".equals(md.getFullName())){
				XSSFSheet excelForecastingSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
				excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelForecastingSettingsSheet,Util.cutSheetName(types),types);
				excelTemplate.createTableHeaders(excelForecastingSettingsSheet, "ForecastingSettings",excelForecastingSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
				for (Metadata m : mdInfos) {
					ForecastingSettings fs = (ForecastingSettings)m;
					if(fs!=null){
						int cellNum = 1;
						XSSFRow columnRow = excelForecastingSettingsSheet.createRow(excelForecastingSettingsSheet.getLastRowNum()+1);						
						//売上予測を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(fs.getEnableForecasts())));
						//デフォルト通貨
						excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(fs.getDisplayCurrency()));
					}
					
				}
				//売上予測種別
				excelTemplate.createTableHeaders(excelForecastingSettingsSheet, "ForecastingTypeSettings",excelForecastingSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
				for (Metadata m : mdInfos) {
					ForecastingSettings fs = (ForecastingSettings)m;
					if(fs.getForecastingTypeSettings().length>0){						
						for( Integer a=0; a<fs.getForecastingTypeSettings().length; a++ ){
							int cellNum = 1;
							ForecastingTypeSettings fts=(ForecastingTypeSettings)fs.getForecastingTypeSettings()[a];
							XSSFRow columnRow = excelForecastingSettingsSheet.createRow(excelForecastingSettingsSheet.getLastRowNum()+1);							
							//売上予測種別名
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(Util.getTranslate("ForecastingType", fts.getName())));
							//有効
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(fts.getActive())));
							//期間
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(fts.getForecastRangeSettings()!=null ? Util.nullFilter(Util.getTranslate("ForecastRangePeriodType", fts.getForecastRangeSettings().getPeriodType().name())):""));
							//開始日
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(fts.getForecastRangeSettings()!=null ? Util.nullFilter(Util.getTranslate("ForecastRangeBeginning", fts.getForecastRangeSettings().getBeginning()+"")):""));
							//表示
							excelTemplate.createCell(columnRow,cellNum++,Util.nullFilter(fts.getForecastRangeSettings()!=null ? Util.nullFilter(Util.getTranslate("ForecastRangeDisplaying", fts.getForecastRangeSettings().getDisplaying()+"")):""));
							//調整を有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(fts.getAdjustmentsSettings().getEnableAdjustments())));
							//目標を表示
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(fts.getQuotasSettings().getShowQuotas())));
							//表示する項目
							if(fts.getOpportunityListFieldsSelectedSettings().getField().length>0){
								String entryCriteriaObj ="";
								for(int i=0;i<fts.getOpportunityListFieldsSelectedSettings().getField().length;i++){
									String aec = fts.getOpportunityListFieldsSelectedSettings().getField()[i];
									if(i!=0){
										entryCriteriaObj +="\n";
									}
									entryCriteriaObj += ut.getLabelforAll(aec);
								}
								excelTemplate.createCell(columnRow,cellNum++,entryCriteriaObj);
							}else{
								excelTemplate.createCell(columnRow,cellNum++,"");
							}
						}
					}
				}
				excelTemplate.adjustColumnWidth(excelForecastingSettingsSheet);
			}
			//chatterモバイル通知の設定
				if("Mobile".equals(md.getFullName())){
					XSSFSheet excelMobileSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelMobileSettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelMobileSettingsSheet, "ChatterMobileSettings",excelMobileSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						int cellNum = 1;
						MobileSettings ms = (MobileSettings)m;
						ChatterMobileSettings cms=ms.getChatterMobile();
						XSSFRow columnRow = excelMobileSettingsSheet.createRow(excelMobileSettingsSheet.getLastRowNum()+1);
						//アプリケーション内通知を有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cms.getEnablePushNotifications())));		
					}
					//Salesforce Classicの設定
					excelTemplate.createTableHeaders(excelMobileSettingsSheet, "SFDCMobileSettings",excelMobileSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						MobileSettings ms = (MobileSettings)m;
						SFDCMobileSettings sms=ms.getSalesforceMobile();
						XSSFRow columnRow = excelMobileSettingsSheet.createRow(excelMobileSettingsSheet.getLastRowNum()+1);
						int cellNum = 1;
						//Salesforce Classic Lite の有効化
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(sms.getEnableMobileLite())));
						//モバイルデバイスにユーザを恒久的にリンクする
						excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(sms.getEnableUserToDeviceLinking())));
					}
					//モバイルダッシュボード
					excelTemplate.createTableHeaders(excelMobileSettingsSheet, "DashboardMobileSettings",excelMobileSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						int cellNum = 1;
						MobileSettings ms = (MobileSettings)m;
						if(ms.getDashboardMobile()!=null){
							DashboardMobileSettings dms=ms.getDashboardMobile();
							XSSFRow columnRow = excelMobileSettingsSheet.createRow(excelMobileSettingsSheet.getLastRowNum()+1);
							//すべてのユーザでモバイルダッシュボード iPad アプリケーションを有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(dms.getEnableDashboardIPadApp())));
						}
					}
					excelTemplate.adjustColumnWidth(excelMobileSettingsSheet);
				}
				//Live Agent の設定
				if("LiveAgent".equals(md.getFullName())){
					XSSFSheet excelLiveAgentSettingsSheet= excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet,excelLiveAgentSettingsSheet,Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(excelLiveAgentSettingsSheet, "LiveAgentSettings",excelLiveAgentSettingsSheet.getLastRowNum()+Util.RowIntervalNum);
					for (Metadata m : mdInfos) {
						LiveAgentSettings las = (LiveAgentSettings)m;
						if(las!=null){
							int cellNum = 1;
							XSSFRow columnRow = excelLiveAgentSettingsSheet.createRow(excelLiveAgentSettingsSheet.getLastRowNum()+1);
							//Live Agent の有効化
							excelTemplate.createCell(columnRow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(las.getEnableLiveAgent())));	
						}
					}
					excelTemplate.adjustColumnWidth(excelLiveAgentSettingsSheet);
				}
				//国情報
				if("Address".equals(md.getFullName())){
					XSSFSheet Countriessheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, Countriessheet, Util.cutSheetName(types),types);
					AddressSettings as = (AddressSettings) md;
					excelTemplate.createTableHeaders(Countriessheet, "Countries",Countriessheet.getLastRowNum() + Util.RowIntervalNum);
					CountriesAndStates countriesAndStates = as.getCountriesAndStates();
					Map<String, List<State>> stateMap = new LinkedHashMap<String, List<State>>();// 存放state
					Country[] countries = countriesAndStates.getCountries();
					if(countries.length>0){
						for (int i = 0; i < countries.length; i++) {
							int cellNum = 1;
							XSSFRow row = Countriessheet.createRow(Countriessheet.getLastRowNum() + 1);
							stateMap.put(countries[i].getLabel(),Arrays.asList(countries[i].getStates()));
							//有効
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(countries[i].getActive())));
							//参照可能
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(countries[i].getVisible())));
							//国名
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(countries[i].getLabel()));
							//国コード
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(countries[i].getIsoCode()));
							//インテグレーション値
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(countries[i].getIntegrationValue()));
							//デフォルトの国
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(countries[i].getOrgDefault())));
							//標準
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(countries[i].getStandard())));
						}
					}
					//州情報
					excelTemplate.createTableHeaders(Countriessheet, "states",Countriessheet.getLastRowNum() + Util.RowIntervalNum);
					for (String state : stateMap.keySet()) {
						for (State s : stateMap.get(state)) {
							int cellNum = 1;
							XSSFRow row = Countriessheet.createRow(Countriessheet.getLastRowNum() + 1);
							//国
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(state));
							//州名
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(s.getLabel()));
							//州コード
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(s.getIsoCode()));
							//インテグレーション値
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(s.getIntegrationValue()));
							//有効
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(s.getActive())));
							//参照可能
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(s.getVisible())));
							//標準
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(s.getStandard())));
						}
					}
					excelTemplate.adjustColumnWidth(Countriessheet);
				}
				//組織の営業時間
				if("BusinessHours".equals(md.getFullName())){
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet, Util.cutSheetName(types),types);
					BusinessHoursSettings businessHoursSettings = (BusinessHoursSettings)md;
					excelTemplate.createTableHeaders(sheet, "BusinessHoursEntry", sheet.getLastRowNum()+Util.RowIntervalNum);
					BusinessHoursEntry[] businessHoursEntries = businessHoursSettings.getBusinessHours();
					for (BusinessHoursEntry be : businessHoursEntries) {
						int cellNum = 1;
						XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
						//営業時間名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(be.getName()));
						//デフォルトとして使用する
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(be.getDefault())));
						//有効
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(be.getActive())));
						//タイムゾーン
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("TIMEZONE",Util.nullFilter(be.getTimeZoneId())));
						//曜日
						excelTemplate.createCell(row,cellNum,Util.getTranslate("Weekday","Monday"));
						//営業時間
						if(be.getTuesdayStartTime() != null || be.getTuesdayStartTime() != null){
							excelTemplate.createCell(row,cellNum+1,Util.nullFilter(getTime(be.getMondayStartTime(),be.getMondayEndTime())));
						}
						XSSFRow row1 = sheet.createRow(sheet.getLastRowNum()+1);
						//cell style
						for(int i=1;i<cellNum;i++){
							excelTemplate.createCell(row1,i,"");
						}
						excelTemplate.createCell(row1,cellNum,Util.getTranslate("Weekday","tuesday"));
						if(be.getTuesdayStartTime() != null || be.getTuesdayStartTime() != null){
							excelTemplate.createCell(row1,cellNum+1,Util.nullFilter(getTime(be.getTuesdayStartTime(),be.getTuesdayStartTime())));
						}
						XSSFRow row2 = sheet.createRow(sheet.getLastRowNum()+1);
						//cell style
						for(int i=1;i<cellNum;i++){
							excelTemplate.createCell(row2,i,"");
						}
						excelTemplate.createCell(row2,cellNum,Util.getTranslate("Weekday","wednesday"));
						if(be.getWednesdayStartTime() != null || be.getWednesdayEndTime() != null){
							excelTemplate.createCell(row2,cellNum+1,Util.nullFilter(getTime(be.getWednesdayStartTime(),be.getWednesdayEndTime())));
						}
						XSSFRow row3 = sheet.createRow(sheet.getLastRowNum()+1);
						//cell style
						for(int i=1;i<cellNum;i++){
							excelTemplate.createCell(row3,i,"");
						}
						excelTemplate.createCell(row3,cellNum,Util.getTranslate("Weekday","thursday"));
						if(be.getThursdayStartTime() != null || be.getThursdayEndTime() != null){
							excelTemplate.createCell(row3,cellNum+1,Util.nullFilter(getTime(be.getThursdayStartTime(),be.getThursdayEndTime())));
						}
						XSSFRow row4 = sheet.createRow(sheet.getLastRowNum()+1);
						//cell style
						for(int i=1;i<cellNum;i++){
							excelTemplate.createCell(row4,i,"");
						}
						excelTemplate.createCell(row4,cellNum,Util.getTranslate("Weekday","friday"));
						if(be.getFridayStartTime() != null || be.getFridayEndTime() != null){
							excelTemplate.createCell(row4,cellNum+1,Util.nullFilter(getTime(be.getFridayStartTime(),be.getFridayEndTime())));
						}
						XSSFRow row5 = sheet.createRow(sheet.getLastRowNum()+1);
						//cell style
						for(int i=1;i<cellNum;i++){
							excelTemplate.createCell(row5,i,"");
						}
						excelTemplate.createCell(row5,cellNum,Util.getTranslate("Weekday","saturday"));
						if(be.getSaturdayStartTime() != null || be.getSaturdayEndTime() != null){
							excelTemplate.createCell(row5,cellNum+1,Util.nullFilter(getTime(be.getSaturdayStartTime(),be.getSaturdayEndTime())));
						}
						XSSFRow row6 = sheet.createRow(sheet.getLastRowNum()+1);
						//cell style
						for(int i=1;i<cellNum;i++){
							excelTemplate.createCell(row6,i,"");
						}
						excelTemplate.createCell(row6,cellNum,Util.getTranslate("Weekday","sunday"));
						if(be.getSundayStartTime() != null || be.getSundayEndTime() != null){
							excelTemplate.createCell(row6,cellNum+1,Util.nullFilter(getTime(be.getSundayStartTime(),be.getSundayEndTime())));
						}
					}
					//休日
					excelTemplate.createTableHeaders(sheet, "Holidays", sheet.getLastRowNum()+Util.RowIntervalNum);
					Holiday[] holiday = businessHoursSettings.getHolidays();
					if(holiday.length>0){
						for (Holiday h : holiday) {
							int cellNum = 1;
							XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
							//休日名
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getName()));
						    //説明
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getDescription()));
							SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
							//日付
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getActivityDate() != null 
									? sf.format(h.getActivityDate().getTime())
											: ""));
							//開始時刻
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getStartTime()));
							//終了時刻
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getEndTime()));
							boolean isTrue = false;
							if(h.getStartTime() == null && h.getEndTime() == null){
								isTrue = true;
							}
							//終日
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(isTrue)));
							//繰り返しの休日
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(h.getIsRecurring())));
							//休日の繰り返し種別
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getRecurrenceType()));	
							//休日を繰り返す間隔
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getRecurrenceInterval()));
							//休日を繰り返す週
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("WEEKNUMBER",Util.nullFilter(h.getRecurrenceInstance())));
							//休日を繰り返す曜日
							//excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getRecurrenceDayOfWeekMask()));
							excelTemplate.createCell(row,cellNum++,Util.getTranslate("DAYOFWEEK",Util.nullFilter(h.getRecurrenceDayOfWeek()[0])));							
							//System.out.println("Util.nullFilter(h.getRecurrenceDayOfWeek())[0]=="+Util.nullFilter(h.getRecurrenceDayOfWeek()[0]));
							//System.out.println("Util.nullFilter(h.getRecurrenceDayOfWeek())[0]=="+Util.getTranslate("DAYOFWEEK",Util.nullFilter(h.getRecurrenceDayOfWeek()[0])));
							//休日を繰り返す日付
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getRecurrenceDayOfMonth()));
							//休日を繰り返す月
							excelTemplate.createCell(row,cellNum++,Util.nullFilter(h.getRecurrenceMonthOfYear()));
						}
					}
					excelTemplate.adjustColumnWidth(sheet);
				}
				//アイデアの設定
				if("Ideas".equals(md.getFullName())){
					int cellNum = 1;
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet,Util.cutSheetName(types), types);
					IdeasSettings ideasSettings = (IdeasSettings)md;
					excelTemplate.createTableHeaders(sheet, "IdeasSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
					XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
					//アイデアの有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ideasSettings.getEnableIdeas())));
					//評価を有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ideasSettings.getEnableIdeasReputation())));
					//Chatter ユーザプロファイルを利用
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ideasSettings.getEnableChatterProfile())));
					//カスタムプロファイルページ
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(ideasSettings.getIdeasProfilePage()));
					//アイデアの半減期 (日数)
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(ideasSettings.getHalfLife()));
					//アイデアのテーマを有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ideasSettings.getEnableIdeaThemes())));
					
					excelTemplate.adjustColumnWidth(sheet);
				}
				//ナレッジの設定
				if("Knowledge".equals(md.getFullName())){
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet,Util.cutSheetName(types), types);
					KnowledgeSettings ks = (KnowledgeSettings)md;
					if(ks!=null){
						int cellNum = 1;
						excelTemplate.createTableHeaders(sheet, "KnowledgeSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
						XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
						//ナレッジの有効化
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getEnableKnowledge())));
						//タブからの記事の作成と編集をユーザに許可する
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getEnableCreateEditOnArticlesTab())));
						//Chatterを介したケースのデフレクションの追跡を有効化
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getEnableChatterQuestionKBDeflection())));
						//標準エディタでユーザが外部マルチメディアコンテンツを HTML に追加することを許可します。
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getEnableExternalMediaContent())));
						//記事の概要が内部アプリケーションに表示されます
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getShowArticleSummariesInternalApp())));
						//記事の概要が顧客に表示されます
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getShowArticleSummariesCustomerPortal())));
						//記事の概要がパートナーに表示されます
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getShowArticleSummariesPartnerPortal())));
						//検証状況項目を有効化
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ks.getShowValidationStatusField())));
						//知識ベースのデフォルト言語
						excelTemplate.createCell(row,cellNum++,Util.getTranslate("DEFAULTLANGUAGE",Util.nullFilter(ks.getDefaultLanguage())));
					
						
						//ナレッジ言語の設定
						excelTemplate.createTableHeaders(sheet, "KnowledgeLanguageSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
						KnowledgeLanguageSettings kls = ks.getLanguages();
						if(kls != null){
							KnowledgeLanguage[] kl = kls.getLanguage();
							for (KnowledgeLanguage knowledgeLanguage : kl) {
								cellNum = 1;
								XSSFRow lrow = sheet.createRow(sheet.getLastRowNum()+1);
							    //言語
								excelTemplate.createCell(lrow,cellNum++,Util.getTranslate("DEFAULTLANGUAGE",Util.nullFilter(knowledgeLanguage.getName())));
								//有効
								excelTemplate.createCell(lrow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(knowledgeLanguage.getActive())));
								//デフォルト任命先のタイプ
								excelTemplate.createCell(lrow,cellNum++,Util.nullFilter(knowledgeLanguage.getDefaultAssigneeType() != null 
										? Util.getTranslate("KnowledgeLanguageLookupValueType"
												, knowledgeLanguage.getDefaultAssigneeType().toString()) : ""));
								//デフォルトの任命先
								excelTemplate.createCell(lrow,cellNum++,Util.nullFilter(knowledgeLanguage.getDefaultAssignee()));
								//デフォルト校閲者のタイプ
								excelTemplate.createCell(lrow,cellNum++,Util.nullFilter(knowledgeLanguage.getDefaultReviewerType()!= null 
										? Util.getTranslate("KnowledgeLanguageLookupValueType"
												, knowledgeLanguage.getDefaultReviewerType().toString()) : ""));
								//デフォルトの校閲者
								excelTemplate.createCell(lrow,cellNum++,Util.nullFilter(knowledgeLanguage.getDefaultReviewer()));
							}
						}
						
						//ナレッジケースの設定
						excelTemplate.createTableHeaders(sheet, "KnowledgeCaseSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
						KnowledgeCaseSettings kcs = ks.getCases();
						XSSFRow crow = sheet.createRow(sheet.getLastRowNum()+1);
						if(kcs != null){
							cellNum = 1;
							//ケースからの記事の作成をユーザに許可する
							excelTemplate.createCell(crow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(kcs.getEnableArticleCreation())));
							//エディタの種類
							excelTemplate.createCell(crow,cellNum++,Util.nullFilter(kcs.getEditor()!=null ? Util.getTranslate("KnowledgeCaseEditor"
									, kcs.getEditor().name()) : ""));
							//デフォルトの記事タイプ
							excelTemplate.createCell(crow,cellNum++,Util.nullFilter(kcs.getDefaultContributionArticleType()));
							//新規記事の割り当て先
							excelTemplate.createCell(crow,cellNum++,Util.nullFilter(kcs.getAssignTo()));
							//APEX カスタマイズを使用
							excelTemplate.createCell(crow,cellNum++,Util.nullFilter(kcs.getCustomizationClass()));
							//プロファイルを使用してお客様が利用可能なケースに関する記事の PDF を作成
							excelTemplate.createCell(crow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(kcs.getUseProfileForPDFCreation())));
						//プロファイル
							excelTemplate.createCell(crow,cellNum++,Util.nullFilter(kcs.getArticlePDFCreationProfile()));
						//公開 URL を使用して記事の共有をユーザに許可する
							excelTemplate.createCell(crow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(kcs.getEnableArticlePublicSharingSites())));
							//選択したサイト
							if(kcs.getArticlePublicSharingSites() != null){
								String[] sites = kcs.getArticlePublicSharingSites().getSite();
								if(sites.length == 1){
									excelTemplate.createCell(crow,cellNum++,Util.nullFilter(sites[0]));
								}else{
									for(String s : sites) {
										XSSFRow cr = sheet.createRow(sheet.getLastRowNum()+1);
										excelTemplate.createCell(cr,cellNum++,Util.nullFilter(s));
									}
								}
							}else{
								excelTemplate.createCell(crow,cellNum++,"");
							}
						}
						//ナレッジアンサーの設定
						excelTemplate.createTableHeaders(sheet, "KnowledgeAnswerSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
						XSSFRow arow = sheet.createRow(sheet.getLastRowNum()+1);
						KnowledgeAnswerSettings kas = ks.getAnswers();
						if( kas != null ){
							cellNum = 1;
							//返信からの記事の作成をユーザに許可する
							excelTemplate.createCell(arow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(kas.getEnableArticleCreation())));
							//デフォルトの記事タイプ
							excelTemplate.createCell(arow,cellNum++,Util.nullFilter(kas.getDefaultArticleType()));
							//新規記事の割り当て先
							excelTemplate.createCell(arow,cellNum++,Util.nullFilter(kas.getAssignTo()));
						}
					}
					excelTemplate.adjustColumnWidth(sheet);
				}
				//信頼済み IP 範囲
				if("Security".equals(md.getFullName())){
					int cellNum = 1;
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet,Util.cutSheetName(types), types);
					SecuritySettings securitySettings = (SecuritySettings)md;
					excelTemplate.createTableHeaders(sheet, "NetworkAccess", sheet.getLastRowNum()+Util.RowIntervalNum);
					NetworkAccess networkAccess = securitySettings.getNetworkAccess();
					if(networkAccess != null){
						IpRange[] ips = networkAccess.getIpRanges();
						if(ips.length>0){
							for (IpRange ipRange : ips) {
								XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
								//開始 IP アドレス
								excelTemplate.createCell(row,cellNum++,Util.nullFilter(ipRange.getStart()));
								//終了 IP アドレス
								excelTemplate.createCell(row,cellNum++,Util.nullFilter(ipRange.getEnd()));
								cellNum = 0;
							}
						}
					}
					//パスワードポリシー
					excelTemplate.createTableHeaders(sheet, "PasswordPolicies", sheet.getLastRowNum()+Util.RowIntervalNum);
					PasswordPolicies policies = securitySettings.getPasswordPolicies();
					XSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
					cellNum = 1;
					//パスワードの有効期間
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("Expiration",Util.nullFilter(policies.getExpiration().name())));
					//過去のパスワードの利用制限回数
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(policies.getHistoryRestriction()));
					//最小パスワード長
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(Util.getTranslate("MinPasswordLength",policies.getMinimumPasswordLength())));
					//パスワード文字列の制限
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(Util.getTranslate("Complexity",policies.getComplexity().name())));
					//パスワード質問の制限
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(Util.getTranslate("QuestionRestriction",policies.getQuestionRestriction().name())));
					//ログイン失敗によりロックするまでの回数
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(Util.getTranslate("MaxLoginAttempts",policies.getMaxLoginAttempts().name())));
					//ロックアウトの有効期間
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(Util.getTranslate("LockoutInterval", policies.getLockoutInterval().name())));
					//パスワードのリセットの秘密の回答を非表示にする
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(policies.getObscureSecretAnswer())));
					//パスワードを忘れた場合のメッセージ
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(policies.getPasswordAssistanceMessage()));
					//パスワードを忘れた場合のヘルプリンク
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(policies.getPasswordAssistanceURL()));
					//API 限定ユーザの代替ホームページ
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(policies.getApiOnlyUserHomePageURL()));
					//セッションの設定
					excelTemplate.createTableHeaders(sheet, "SessionSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
					XSSFRow srow = sheet.createRow(sheet.getLastRowNum() + 1);
					SessionSettings session = securitySettings.getSessionSettings();
					if( session != null ){
						cellNum = 1;
						//セッションタイムアウト値
						excelTemplate.createCell(srow,cellNum++,session.getSessionTimeout() != null ? Util.getTranslate("SessionTimeout",session.getSessionTimeout().toString()) : "");
						//ログアウト URL
						excelTemplate.createCell(srow, cellNum, session.getLogoutURL());
						//セッションタイムアウト時の警告ポップアップを無効にする
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getDisableTimeoutWarning())));
						//セッションを最初に使用したドメインにセッションをロックする
						excelTemplate.createCell(srow, cellNum, Util.nullFilter(session.getLockSessionsToDomain()));
						//すべての要求でログイン IP アドレスの制限を適用
						excelTemplate.createCell(srow, cellNum, Util.nullFilter(session.getEnforceIpRangesEveryRequest()));
						//ヘッダーが無効化された Visualforce ページのクリックジャック保護を有効化
						excelTemplate.createCell(srow, cellNum, Util.nullFilter(session.getEnableClickjackNonsetupUserHeaderless()));
						//ログイン時の IP アドレスとセッションをロックする
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getLockSessionsToIp())));
						//ユーザとしてログインしてから再ログインを強制する
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getForceRelogin())));
						//ログインページでキャッシングとオートコンプリート機能を有効にする
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableCacheAndAutocomplete())));
						//SMS による ID 確認を有効にする
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableSMSIdentity())));
						//設定ページのクリックジャック保護を有効化
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableClickjackSetup())));
						//標準ヘッダーがある Visualforce ページのクリックジャック保護を有効化
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableClickjackNonsetupUser())));
						//設定以外の Salesforce ページのクリックジャック保護を有効化
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableClickjackNonsetupSFDC())));
					    //設定ページ以外の GET 要求の CSRF 保護を有効化
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableCSRFOnGet())));
						//設定ページ以外の POST 要求の CSRF 保護を有効化
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnableCSRFOnPost())));
						//クロスドメインセッションで POST 要求を使用
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getEnablePostForSessions())));
						//セッションタイムアウト時に強制的にログアウト
						excelTemplate.createCell(srow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(session.getForceLogoutOnSessionTimeout())));
						
					}
					excelTemplate.adjustColumnWidth(sheet);
				}
				//見積設定
				if("Quote".equals(md.getFullName())){
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet, Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(sheet, "QuoteSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
					QuoteSettings quoteSettings = (QuoteSettings)md;
					if(quoteSettings!=null){
						int cellNum = 1;
						XSSFRow qrow = sheet.createRow(sheet.getLastRowNum()+1);
						//見積の有効化
						excelTemplate.createCell(qrow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(quoteSettings.getEnableQuote())));
					}
					excelTemplate.adjustColumnWidth(sheet);
				}
				//商品設定
				if("Product".equals(md.getFullName())){
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet,Util.cutSheetName(types), types);
					excelTemplate.createTableHeaders(sheet, "ProductSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
					ProductSettings productSettings = (ProductSettings)md;
					XSSFRow prow = sheet.createRow(sheet.getLastRowNum()+1);
					int cellNum = 1;
					//商品レコード上の有効フラグを変更した際に、関連する価格の有効フラグも自動更新する
					excelTemplate.createCell(prow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(productSettings
							.getEnableCascadeActivateToRelatedPrices())));
				//数量によるスケジュールの有効化
					excelTemplate.createCell(prow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(productSettings
							.getEnableQuantitySchedule())));
					//収益によるスケジュールの有効化
					excelTemplate.createCell(prow,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(productSettings
							.getEnableRevenueSchedule())));
					
					excelTemplate.adjustColumnWidth(sheet);
				}
				//注文の設定
				if("Order".equals(md.getFullName())){
					int cellNum = 1;
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet, Util.cutSheetName(types),types);
					excelTemplate.createTableHeaders(sheet, "OrderSettings", sheet.getLastRowNum()+Util.RowIntervalNum);
					OrderSettings orderSettings = (OrderSettings)md;
					XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
					//注文を有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(orderSettings.getEnableOrders())));
					//削減注文を有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(orderSettings.getEnableReductionOrders())));
					//負の数量を有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(orderSettings.getEnableNegativeQuantity())));
					
					excelTemplate.adjustColumnWidth(sheet);
				}
				//OpportunitySettings
				if("Opportunity".equals(md.getFullName())){
					int cellNum = 1;
					XSSFSheet sheet = excelTemplate.createSheet(Util.cutSheetName(types));
					excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, sheet,Util.cutSheetName(types), types);
					excelTemplate.createTableHeaders(sheet, "OpportunitySettings", sheet.getLastRowNum()+Util.RowIntervalNum);
					OpportunitySettings opportunitySettings = (OpportunitySettings)md;
					XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
					//直属の部下を持つユーザに対するリマインダーを自動的に有効にする
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(opportunitySettings.getAutoActivateNewReminders())));
					//組織にアップデートリマインダーを有効にする
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(opportunitySettings.getEnableUpdateReminders())));
				//商品を商談に追加するようユーザに促す
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(opportunitySettings.getPromptToAddProducts())));
					//チームセリング設定の有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(opportunitySettings.getEnableOpportunityTeam())));
					//類似商談の有効化
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(opportunitySettings.getEnableFindSimilarOpportunities())));
					int rownum = sheet.getLastRowNum();
					FindSimilarOppFilter filter = opportunitySettings.getFindSimilarOppFilter();
					
					if(filter != null){
						//一致条件
						String[] fields = filter.getSimilarOpportunitiesMatchFields();
						//類似商談の結果列
						String[] columns = filter.getSimilarOpportunitiesDisplayColumns();
						String resultField="";
						String resultColumn="";
						for(int r = 0; r < fields.length; r++){
							resultField+=ut.getLabelforAll(Util.nullFilter(fields[r]))+"\n";
						}
						for(int r = 0; r < columns.length; r++){
							resultColumn+=ut.getLabelforAll(Util.nullFilter(columns[r]))+"\n";
						}
						if( sheet.getRow(rownum) != null ){
							XSSFRow xssfRow = sheet.getRow(rownum);
							excelTemplate.createCell(xssfRow,cellNum++,Util.nullFilter(resultField));
							excelTemplate.createCell(xssfRow,cellNum++,Util.nullFilter(resultColumn));
						}else{
							XSSFRow xssfRow = sheet.createRow(rownum);
							excelTemplate.createCell(xssfRow,cellNum++,Util.nullFilter(resultField));
							excelTemplate.createCell(xssfRow,cellNum++,Util.nullFilter(resultColumn));
						}
						/*for(int r = 0; r < columns.length || r < fields.length ; r++){
							if( sheet.getRow(rownum + r) != null ){
								XSSFRow xssfRow = sheet.getRow(rownum + r);
								if(r <fields.length){
									excelTemplate.createCell(xssfRow,cellNum,Util.nullFilter(fields[r]));
								}else{
									excelTemplate.createCell(xssfRow,cellNum,"");
								}
								if(r <columns.length){
									excelTemplate.createCell(xssfRow,cellNum + 1,Util.nullFilter(columns[r]));
								}else{
									excelTemplate.createCell(xssfRow,cellNum + 1,"");
								}
							}else{
								XSSFRow xssfRow = sheet.createRow(rownum + r);
								if(r <fields.length){
									excelTemplate.createCell(xssfRow,cellNum,Util.nullFilter(fields[r]));
								}else{
									excelTemplate.createCell(xssfRow,cellNum,"");
								}
								if(r <columns.length){
									excelTemplate.createCell(xssfRow,cellNum + 1,Util.nullFilter(columns[r]));
								}else{
									excelTemplate.createCell(xssfRow,cellNum + 1,"");
								}
							}
						}*/
					}else{
						excelTemplate.createCell(row,cellNum++,"");
						excelTemplate.createCell(row,cellNum++,"");
					}
					excelTemplate.adjustColumnWidth(sheet);
				}
			}
		}			
		
		
		if(workBook.getSheet(Util.getTranslate("Common", "Index")).getRow(1) != null){
			//excelTemplate.adjustColumnWidth(catalogSheet);
			excelTemplate.exportExcel(type,"");
		}
		else{
			Util.logger.warn("***no result to export!!!");
		}	
		Util.logger.info("ReadSetting End.");
	}
	public String getTime(com.sforce.ws.types.Time startTime,com.sforce.ws.types.Time endTime){
		System.out.println(startTime.toString() + " - " +endTime.toString());
		String time = "";
		if(startTime.toString().equals(endTime.toString())){
			time = Util.getTranslate("Weekday","allday");
		}else{
			time = startTime.toString().substring(0, 5) + " - " +endTime.toString().substring(0, 5);
		}
		return time;
	}
}