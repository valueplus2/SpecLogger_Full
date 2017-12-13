package source;

import java.io.IOException;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import wsc.MetadataLoginUtil;

import com.sforce.soap.metadata.CustomApplication;
import com.sforce.soap.metadata.CustomTab;
import com.sforce.soap.metadata.EmailTemplate;
import com.sforce.soap.metadata.FileProperties;
import com.sforce.soap.metadata.ListMetadataQuery;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.Profile;
import com.sforce.soap.metadata.ProfileApexClassAccess;
import com.sforce.soap.metadata.ProfileApexPageAccess;
import com.sforce.soap.metadata.ProfileApplicationVisibility;
import com.sforce.soap.metadata.ProfileCustomPermissions;
import com.sforce.soap.metadata.ProfileFieldLevelSecurity;
import com.sforce.soap.metadata.ProfileLayoutAssignment;
import com.sforce.soap.metadata.ProfileLoginHours;
import com.sforce.soap.metadata.ProfileLoginIpRange;
import com.sforce.soap.metadata.ProfileObjectPermissions;
import com.sforce.soap.metadata.ProfilePasswordPolicy;
import com.sforce.soap.metadata.ProfileRecordTypeVisibility;
import com.sforce.soap.metadata.ProfileTabVisibility;
import com.sforce.soap.metadata.ProfileUserPermission;
import com.sforce.ws.ConnectionException;

public class ReadProfileSync {
	private XSSFWorkbook workbook;
	
	public List<String> getObjectList(ListMetadataQuery queries) throws Exception{
		List<String> list = new ArrayList<String>();			
		FileProperties[] fileProperties = MetadataLoginUtil.metadataConnection.listMetadata(
				new ListMetadataQuery[] { queries }, Util.API_VERSION);
		for (FileProperties f : fileProperties) {
			list.add(f.getFullName());
		}
		return list;
	}
	public void readProfile(String type, List<String> objectsList)
			throws Exception {
		Util.logger.info("ReadProfileSync Started.");
		Util ut = new Util();
		Util.nameSequence=0;
		Util.sheetSequence=0;
		List<Metadata> mdInfos = ut.readMateData(type, objectsList);
		Map<String, String> resultMap = ut.getComparedResult(type,
				UtilConnectionInfc.getLastUpdateTime());
		CreateExcelTemplate excelTemplate = new CreateExcelTemplate(type);
		//XSSFSheet catalog = excelTemplate.createCatalogSheet();
		this.workbook = excelTemplate.workBook;
		
		Map<String,String> appMap = new LinkedHashMap<String,String>();
		Map<String,String> tabMap = new LinkedHashMap<String,String>();
		List<String> list = new ArrayList<String>();			
		ListMetadataQuery queries = new ListMetadataQuery();
		queries.setType("CustomApplication");
		queries.setFolder("applications");
		list=this.getObjectList(queries);
		List<Metadata> mdInfos2 = ut.readMateData("CustomApplication", list);
		for (Metadata md : mdInfos2) {
			if (md != null) {
				CustomApplication ca=(CustomApplication)md;
				if(ca.getLabel()!=null){
					appMap.put(ca.getFullName(),ca.getLabel());
				}else if(!(Util.getTranslate("ApplicationName", ca.getFullName()).equals(ca.getFullName()))){
					appMap.put(ca.getFullName(),Util.getTranslate("ApplicationName", ca.getFullName()));
				}else{
					appMap.put(ca.getFullName(),ca.getFullName());
				}
			}
		}
		queries.setType("CustomTab");
		queries.setFolder("tabs");
		list=this.getObjectList(queries);
		mdInfos2 = ut.readMateData("CustomTab", list);
		for (Metadata md : mdInfos2) {
			if (md != null) {
				CustomTab ct=(CustomTab)md;
				if(ct.getLabel()!=null){
					tabMap.put(ct.getFullName(),ct.getLabel()+"("+ct.getFullName()+")");
				//}else if(!(Util.getTranslate("PorfileTabName", ct.getFullName()).equals(ct.getFullName()))){
				//	tabMap.put(ct.getFullName(),Util.getTranslate("PorfileTabName", ct.getFullName())+"("+ct.getFullName()+")");
				//}else{
				//	tabMap.put(ct.getFullName(),ut.getLabelApi(ct.getFullName()));
				}
			}
		}
		
		//ProfilePasswordPolicyコンポーネントを取得
		Map<String,ProfilePasswordPolicy> profilePassMap = new LinkedHashMap<String,ProfilePasswordPolicy>();
		queries.setType("ProfilePasswordPolicy");
		queries.setFolder("profilePasswordPolicies");
		list=this.getObjectList(queries);
		List<Metadata> mdInfos3 = ut.readMateData("ProfilePasswordPolicy", list);
		for (Metadata md : mdInfos3) {
			if (md != null) {
				ProfilePasswordPolicy pp=(ProfilePasswordPolicy)md;
				if(pp.getFullName()!=null && pp.getProfile()!=null){					
					profilePassMap.put(pp.getProfile().toLowerCase(),pp);
				}
			}
		}
		
		// loop profile detail
		for (int i = 0; i < mdInfos.size(); i++) {
			Profile profile = (Profile) mdInfos.get(i);
			Util.logger.info("profile.getFullName()="+profile.getFullName());
			String sheetName = Util.makeSheetName(profile.getFullName());
			Util.logger.info("profile.getFullName().heetName="+sheetName);
			sheetName = URLDecoder.decode(sheetName, "UTF-8");
			XSSFSheet profileSheet = excelTemplate.createSheet(Util.cutSheetName(sheetName));
			excelTemplate.createCatalogMenu(excelTemplate.catalogSheet, profileSheet,profileSheet.getSheetName(),sheetName);
			//Create Profile Table
			excelTemplate.createTableHeaders(profileSheet, "Profile",
					profileSheet.getLastRowNum() + Util.RowIntervalNum);
			XSSFRow rowProfile = profileSheet.createRow(profileSheet
					.getLastRowNum() + 1);
			int cellNum=1;
			//変更あり
			if(UtilConnectionInfc.modifiedFlag){
				excelTemplate.createCell(rowProfile,cellNum++,ut.getUpdateFlag(resultMap,"Profile." + profile.getFullName()));
				
			}
			//プロファイル名
			excelTemplate.createCell(rowProfile,cellNum++,Util.nullFilter(URLDecoder.decode(profile.getFullName(), "UTF-8")));
			//ユーザライセンス
			excelTemplate.createCell(rowProfile,cellNum++,Util.nullFilter(URLDecoder.decode(profile.getUserLicense(), "UTF-8")));
			//カスタムプロファイル
			excelTemplate.createCell(rowProfile,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(profile.getCustom())));
			//説明
			excelTemplate.createCell(rowProfile,cellNum++,Util.nullFilter(profile.getDescription() != null
					? URLDecoder.decode(profile.getDescription(), "UTF-8") : null));
			//Create Layout Assignments Table(ページレイアウト)
			Util.logger.info("ProfileLayoutAssignments started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileLayoutAssignments",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getLayoutAssignments().length > 0) {
				cellNum=1;
				ProfileLayoutAssignment[] layoutAssignments = profile.getLayoutAssignments();
				for (ProfileLayoutAssignment ls : layoutAssignments) {
					Util.logger.debug("ProfileLayoutAssignment="+ls);

					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					Util.logger.info("layout="+ls.getLayout());
					int last = ls.getLayout().lastIndexOf(" ");
					last = last > 0 ? last : ls.getLayout().length();
					String layout = Util.translateSpecialChar(ls.getLayout());
					//add dan 16/2/17 start
					String objectName = layout.substring(0,layout.indexOf('-'));
					//オブジェクト
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(objectName));
					if(ls.getRecordType()!=null){
						String recordType = ut.getLabelforAll(Util.nullFilter(ls.getRecordType())).substring(ut.getLabelforAll(Util.nullFilter(ls.getRecordType())).indexOf('.')+1);
						//レコードタイプ
						excelTemplate.createCell(row,cellNum++,ut.getLabelforAll(Util.nullFilter(recordType)));
					}else{
						excelTemplate.createCell(row,cellNum++,ut.getLabelforAll(Util.nullFilter(ls.getRecordType())));
					}
					//add dan 16/2/17 end
					//レイアウト
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(layout));
					cellNum=1;
				}
			}
			Util.logger.info("ProfileLayoutAssignments completed.");

			//Create Field Level Security Table(項目レベルセキュリティ)
			Util.logger.info("ProfileFieldLevelSecurity started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileFieldLevelSecurity",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getFieldPermissions().length > 0) {
				ProfileFieldLevelSecurity[] fieldLevelSecurity = profile.getFieldPermissions();
				for (ProfileFieldLevelSecurity fy : fieldLevelSecurity) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					//項目名
					excelTemplate.createCell(row,cellNum++,ut.getLabelforAll(Util.nullFilter(URLDecoder.decode(fy.getField(), "UTF-8"))));
					//参照可能
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(fy.getReadable())));
					//参照のみ
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(fy.getEditable())));
				}
			}
			Util.logger.info("ProfileFieldLevelSecurity completed.");

			//Create ApplicationVisibility Table(カスタムアプリケーション設定)
			Util.logger.info("ProfileApplicationVisibility started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileApplicationVisibility",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getApplicationVisibilities().length > 0) {
				ProfileApplicationVisibility[] applicationVisibilities = profile.getApplicationVisibilities();
				for (ProfileApplicationVisibility ay : applicationVisibilities) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					String apiName=Util.nullFilter(URLDecoder.decode(ay.getApplication(), "UTF-8"));
					if(apiName!=null&&appMap.get(apiName)!=null){
						apiName=appMap.get(apiName);
					}
					//アプリケーション
					excelTemplate.createCell(row,cellNum++,apiName);
					//参照可能
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ay.getVisible())));
					//デフォルト
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ay.getDefault())));
				}
			}
			Util.logger.info("ProfileApplicationVisibility completed.");

			//Create TabVisibility Table(タブの設定)
			Util.logger.info("ProfileTabVisibility started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileTabVisibility",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getTabVisibilities().length > 0) {
				ProfileTabVisibility[] tabVisibilities = profile.getTabVisibilities();
				for (ProfileTabVisibility ty : tabVisibilities) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					String apiName=Util.nullFilter(URLDecoder.decode(ty.getTab(), "UTF-8"));
					if(apiName!=null&&tabMap.get(apiName)!=null){
						apiName=tabMap.get(apiName);
					}else if(apiName.contains("standard-")){
						if(!(Util.getTranslate("PorfileTabName", apiName).equals(apiName))){
							apiName=Util.getTranslate("PorfileTabName", apiName)+"("+apiName+")";
						}else{
							apiName=ut.getLabelApi(apiName.replace("standard-", ""));
						}
					}else{
						apiName=ut.getLabelApi(apiName);
					}
					//タブ
					excelTemplate.createCell(row,cellNum++,apiName);
					//利用可能
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("TabVisibility", Util.nullFilter(ty.getVisibility().toString())));
				}
			}
			Util.logger.info("ProfileTabVisibility completed.");

			Util.logger.info("ProfileRecordTypeVisibility started.");
			//Create RecordTypeVisibility Table(レコードタイプの設定)
			excelTemplate.createTableHeaders(profileSheet,"ProfileRecordTypeVisibility",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getRecordTypeVisibilities().length > 0) {
				ProfileRecordTypeVisibility[] recordTypeVisibilities = profile.getRecordTypeVisibilities();
				for (ProfileRecordTypeVisibility ry : recordTypeVisibilities) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					//レコードタイプ
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(URLDecoder.decode(ry.getRecordType(), "UTF-8")));
					//選択済み
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ry.getVisible())));
					//デフォルト
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ry.getDefault())));
					//個人取引先
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(ry.getPersonAccountDefault())));
				}
			}
			Util.logger.info("ProfileRecordTypeVisibility completed.");

			//Create UserPermissions Table(一般ユーザ権限)
			Util.logger.info("ProfileUserPermission started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileUserPermission",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			
			if (profile.getUserPermissions().length > 0) {
				ProfileUserPermission[] userPermissions = profile.getUserPermissions();
				for (ProfileUserPermission us : userPermissions) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					//権限名
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("USERPERMISSIONNAME",Util.nullFilter(URLDecoder.decode(us.getName(), "UTF-8"))));
				}
			}
			Util.logger.info("ProfileUserPermission completed.");

			//Create ObjectPermissions Table(オブジェクト権限)
			Util.logger.info("ProfileObjectPermissions started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileObjectPermissions",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getObjectPermissions().length > 0) {
				ProfileObjectPermissions[] objectPermissions = profile.getObjectPermissions();
				for (ProfileObjectPermissions os : objectPermissions) {
					//System.out.println("os.getObject()==="+os.getObject());
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					//オブジェクト
					excelTemplate.createCell(row,cellNum++,ut.getLabelApi(Util.nullFilter(URLDecoder.decode(os.getObject(), "UTF-8"))));
					//参照
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(os.getAllowRead())));
					//作成	
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(os.getAllowCreate())));
					//編集
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(os.getAllowEdit())));
					//削除
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(os.getAllowDelete())));
					//すべて表示
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(os.getViewAllRecords())));
					//すべて変更
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(os.isModifyAllRecords())));
				}
			}
			Util.logger.info("ProfileObjectPermissions completed.");

			//Create CustomPermissions Table(カスタム権限)
			Util.logger.info("ProfileCustomPermissions started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileCustomPermissions",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getCustomPermissions().length > 0) {
				ProfileCustomPermissions[] customPermissions = profile.getCustomPermissions();
				for (ProfileCustomPermissions cs : customPermissions) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					//権限名
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(URLDecoder.decode(cs.getName(), "UTF-8")));
					//有効
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(cs.getEnabled())));
				}
			}
			Util.logger.info("ProfileCustomPermissions completed.");

			//Create LoginHours Table(ログイン時間帯の制限)
			Util.logger.info("ProfileLoginHours started.");
			excelTemplate.createTableHeaders(profileSheet,"ProfileLoginHours", profileSheet.getLastRowNum() + Util.RowIntervalNum);
			ProfileLoginHours loginHours = profile.getLoginHours();
			if (profile.getLoginHours() != null) {
				cellNum=1;
				XSSFRow row0 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row0,cellNum++,Util.getTranslate("Weekday", "Monday"));
				excelTemplate.createCell(row0,cellNum++,
						loginHours.getMondayStart() != null ? (Integer
								.parseInt(loginHours.getMondayStart()) / 60)
								+ ":00 -"
								+ Integer.parseInt(loginHours.getMondayEnd())/ 60 + ":00" : Util.getTranslate("Weekday","allday"));
				XSSFRow row1 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				cellNum=1;
				excelTemplate.createCell(row1,cellNum++,Util.getTranslate("Weekday", "Tuesday"));
				excelTemplate.createCell(row1,cellNum++,
						loginHours.getTuesdayStart() != null ? (Integer.parseInt(loginHours.getTuesdayStart()) / 60)
								+ ":00 -"
								+ Integer.parseInt(loginHours.getTuesdayEnd())/ 60 + ":00" : Util.getTranslate("Weekday","allday"));
				cellNum=1;
				XSSFRow row2 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row2,cellNum++,Util.getTranslate("Weekday", "Wednesday"));
				excelTemplate.createCell(row2,cellNum++,
								loginHours.getWednesdayStart() != null ? (Integer.parseInt(loginHours.getWednesdayStart()) / 60)
										+ ":00 -"
										+ Integer.parseInt(loginHours.getWednesdayEnd())/ 60
										+ ":00": Util.getTranslate("Weekday", "allday"));
				cellNum=1;
				XSSFRow row3 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row3,cellNum++,Util.getTranslate("Weekday", "Thursday"));
				excelTemplate.createCell(row3,cellNum++,
						loginHours.getThursdayStart() != null ? (Integer.parseInt(loginHours.getThursdayStart()) / 60)
								+ ":00 -"
								+ Integer.parseInt(loginHours.getThursdayEnd())/ 60 + ":00" : Util.getTranslate("Weekday","allday"));
				cellNum=1;
				XSSFRow row4 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row4,cellNum++,Util.getTranslate("Weekday", "Friday"));
				excelTemplate.createCell(row4,cellNum++,
						loginHours.getFridayStart() != null ? (Integer.parseInt(loginHours.getFridayStart()) / 60)
								+ ":00 -"
								+ Integer.parseInt(loginHours.getFridayEnd())/ 60 + ":00" : Util.getTranslate("Weekday","allday"));
				cellNum=1;
				XSSFRow row5 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row5,cellNum++,Util.getTranslate("Weekday", "Saturday"));
				excelTemplate.createCell(row5,cellNum++,
						loginHours.getSaturdayStart() != null ? (Integer.parseInt(loginHours.getSaturdayStart()) / 60)
								+ ":00 -"
								+ Integer.parseInt(loginHours.getSaturdayEnd())/ 60 + ":00" : Util.getTranslate("Weekday","allday"));
				cellNum=1;
				XSSFRow row6 = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
				excelTemplate.createCell(row6,cellNum++,Util.getTranslate("Weekday", "Sunday"));
				excelTemplate.createCell(row6,cellNum++,
						loginHours.getSundayStart() != null ? (Integer.parseInt(loginHours.getSundayStart()) / 60)
								+ ":00 -"
								+ Integer.parseInt(loginHours.getSundayEnd())/ 60 + ":00" : Util.getTranslate("Weekday","allday"));
			}
			Util.logger.info("ProfileLoginHours completed.");
			
			//Create LoginIpRange Table(ログイン IP アドレスの制限)
			Util.logger.info("ProfileLoginIpRange started.");			
			excelTemplate.createTableHeaders(profileSheet,"ProfileLoginIpRange",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getLoginIpRanges().length > 0) {
				ProfileLoginIpRange[] loginIpRanges = profile.getLoginIpRanges();
				for (ProfileLoginIpRange le : loginIpRanges) {
					cellNum=1;
					XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
					//開始IPアドレス
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(le.getStartAddress()));
					//終了IPアドレス
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(le.getEndAddress()));
					//説明
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(le.getDescription() != null ? URLDecoder.decode(le.getDescription(), "UTF-8") : null));
				}
			}
			Util.logger.info("ProfileLoginIpRange completed.");			

			//Create ProfileApexClassAccess Table(有効なApexクラス)
			Util.logger.info("ProfileApexClassAccess started.");			
			excelTemplate.createTableHeaders(profileSheet,"ProfileApexClassAccess",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getClassAccesses().length > 0) {
				ProfileApexClassAccess[] apexClassAccesses = profile.getClassAccesses();
				for (ProfileApexClassAccess as : apexClassAccesses) {
					if (as.isEnabled()) {
						cellNum=1;
						XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
						//Apexクラス名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(URLDecoder.decode(as.getApexClass(), "UTf-8")));
					}
				}
			}
			Util.logger.info("ProfileApexClassAccess completed.");			
			
			//Create ProfileApexPageAccess Table(有効な Visualforceページ)
			Util.logger.info("ProfileApexPageAccess started.");			
			excelTemplate.createTableHeaders(profileSheet,"ProfileApexPageAccess",profileSheet.getLastRowNum() + Util.RowIntervalNum);
			if (profile.getPageAccesses().length > 0) {
				ProfileApexPageAccess[] apexPageAccesses = profile.getPageAccesses();
				for (ProfileApexPageAccess as : apexPageAccesses) {
					if (as.isEnabled()) {
						cellNum=1;
						XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);
						//Visualforce ページ名
						excelTemplate.createCell(row,cellNum++,Util.nullFilter(URLDecoder.decode(as.getApexPage(), "UTF-8")));
					}
				}
			}
			Util.logger.info("ProfileApexPageAccess completed.");	
			
			//create ProfilePasswordPolicy Table(プロファイルパスワードポリシー)
			Util.logger.info("ProfilePasswordPolicy started.");	
			excelTemplate.createTableHeaders(profileSheet,"ProfilePasswordPolicy",profileSheet.getLastRowNum() + Util.RowIntervalNum);		    
		    ProfilePasswordPolicy passwordPolicy = profilePassMap.get(profile.getFullName().toLowerCase());		    
		    if(passwordPolicy != null){    
				cellNum=1;
				XSSFRow row = profileSheet.createRow(profileSheet.getLastRowNum() + 1);	
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("LockoutInterval",Util.nullFilter(passwordPolicy.getLockoutInterval())));
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("MaxLoginAttempts",Util.nullFilter(passwordPolicy.getMaxLoginAttempts())));
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("MinPasswordLength",Util.nullFilter(passwordPolicy.getMinimumPasswordLength())));				
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(passwordPolicy.getMinimumPasswordLifetime())));
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("BOOLEANVALUE",Util.nullFilter(passwordPolicy.getObscure())));
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("Complexity",Util.nullFilter(passwordPolicy.getPasswordComplexity())));
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("Expiration",Util.nullFilter(passwordPolicy.getPasswordExpiration())));
				if(passwordPolicy.getPasswordHistory() == 0){
					excelTemplate.createCell(row,cellNum++,Util.getTranslate("PASSWORD", "HISTORYZERO"));
				}else{
					excelTemplate.createCell(row,cellNum++,Util.nullFilter(passwordPolicy.getPasswordHistory()+Util.getTranslate("PASSWORD", "HISTORY")));
				}
				excelTemplate.createCell(row,cellNum++,Util.getTranslate("QuestionRestriction",Util.nullFilter(passwordPolicy.getPasswordQuestion())));
				excelTemplate.createCell(row,cellNum++,Util.nullFilter(profile.getFullName()));	
		    }
			Util.logger.info("ProfilePasswordPolicy completed.");
						
			// ProfileExternalDataSourceAccess   the function is abandoned
//			excelTemplate.createTableHeaders(profileSheet,
//					"ProfileExternalDataSourceAccess",
//					profileSheet.getLastRowNum() + 3);
//			if (profile.getExternalDataSourceAccesses().length > 0) {
//				ProfileExternalDataSourceAccess[] externalDataSourceAccesses = profile
//						.getExternalDataSourceAccesses();
//				for (ProfileExternalDataSourceAccess es : externalDataSourceAccesses) {
//					if (es.isEnabled()) {
//						XSSFRow row = profileSheet.createRow(profileSheet
//								.getLastRowNum() + 1);
//						row.createCell(0).setCellValue(
//								es.getExternalDataSource());
//					}
//				}
//			}
			//Need to confirm performance issue
			excelTemplate.adjustColumnWidth(profileSheet);
			
			if(ut.createExcel(workbook, excelTemplate, type, objectsList.size(), i+1)){
				excelTemplate.CreateWorkBook(type);
				workbook = excelTemplate.workBook;
			}				
		}
		Util.logger.info("ReadProfileSync End.");
	}
}
