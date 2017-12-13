package source;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ErrorConstants;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class PoiUtil {
	public static final String EMPTY = "";
    public static final String CRLF = "\n";
	public int getFixedHeight(XSSFWorkbook workBook,XSSFSheet Sheet, XSSFRow row, Cell poiCell, Integer column){
		Object value = getValue(workBook,poiCell,false);
		short fontHeight=workBook.getFontAt(Sheet.getColumnStyle(column).getFontIndex()).getFontHeight();
		double width = getWidth(Sheet, column);
		double fixedHeight = 0;
		// セル値が空の場合は処理をスキップ
		if (value == null || isEmpty(value.toString())) {
			return row.getHeight();
		}
		
		// 改行が存在する場合は行毎の高さを加算
		if (value instanceof String) {
			String values[] = ((String) value).split(CRLF);
			for (String temp : values) {
				// 一時的に行毎の値を設定して表示に必要なセル幅を取得
				setValue(poiCell,temp);
				double fixedWidth = getFixedWidth(Sheet, row.getRowNum(),column);
				
				// Excel100%表示時の改行位置ずれ補正
				fixedWidth /= 1.1769;

				// 表示に必要なセル幅から折り返し文字列を考慮した高さの算出
				fixedHeight += (Math.ceil(fixedWidth / width)) * fontHeight; // 折り返し行分のフォント高サイズ追加
				fixedHeight += (Math.ceil(fixedWidth / width)) * fontHeight * 0.2D; // 折り返し行分のフォントに合わせたExcel自動調整オフセットを追加
			}
		}

		// フォント高サイズ分最終行に余白サイズ追加
		fixedHeight += fontHeight;

		// 元の値を設定
		setValue(poiCell,value);	
		return (int) fixedHeight;		
	}
	public void autoSizeRow(XSSFWorkbook workBook,XSSFSheet Sheet, XSSFRow row, Cell poiCell, Integer column) {
		int height=getFixedHeight(workBook,Sheet,row,poiCell,column);
		if (row.getHeight() > height) {
			return;
		}

		if (poiCell.getCellStyle().getWrapText()) {
			row.setHeight((short)height);
		}
	}	
    public static boolean isEmpty(String value) {
        return value == null || value.length() <= 0;
    }	
	public int getFixedWidth(XSSFSheet Sheet, Integer rowNum, Integer columnIndex) {
		double width = SheetUtil.getColumnWidth(Sheet, columnIndex, false, rowNum,rowNum) * 256D;
		return (int) width;
	}
	public int getWidth(XSSFSheet Sheet, Integer columnIndex) {
		return Sheet.getColumnWidth(columnIndex);
	}	
	/**
	 * セル値をセットします。<br>
	 * @param value セル値
	 */
	public void setValue(Cell poiCell,Object value) {
		if (value == null) {
			poiCell.setCellType(Cell.CELL_TYPE_BLANK);
		} else if (value instanceof Boolean) {
			poiCell.setCellValue((Boolean) value);
		} else if (value instanceof Date) {
			poiCell.setCellValue((Date) value);
		} else if (value instanceof Calendar) {
			poiCell.setCellValue((Calendar) value);
		} else if (value instanceof Short) {
			poiCell.setCellValue((Short) value);
		} else if (value instanceof Integer) {
			poiCell.setCellValue((Integer) value);
		} else if (value instanceof Float) {
			poiCell.setCellValue((Float) value);
		} else if (value instanceof Double) {
			poiCell.setCellValue((Double) value);
		} else if (value instanceof Long) {
			poiCell.setCellValue((Long) value);
		} else if (value instanceof BigDecimal) {
			poiCell.setCellValue(((BigDecimal) value).doubleValue());
		} else if (value instanceof String) {
			if ("".equals(value)) {
				poiCell.setCellType(Cell.CELL_TYPE_BLANK);
			} else {
				poiCell.setCellValue((String) value);
			}
		} else {
			poiCell.setCellValue(value.toString());
		}
	}	
	public Object getValue(XSSFWorkbook workBook,Cell poiCell,boolean calc) {
		if (poiCell.getCellType() == Cell.CELL_TYPE_BLANK) {
			return EMPTY;
		} else if (poiCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			return poiCell.getBooleanCellValue();
		} else if (poiCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			return poiCell.getNumericCellValue();
		} else if (poiCell.getCellType() == Cell.CELL_TYPE_STRING) {
			return poiCell.getStringCellValue();
		} else if (poiCell.getCellType() == Cell.CELL_TYPE_ERROR) {
			return poiCell.getErrorCellValue();
		} else if (poiCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			if (calc) {
				evaluateCell(workBook, poiCell);
				return getValue(workBook,poiCell,false);
			} else {
				return getBaseValue(poiCell);
			}
		} else {
			return EMPTY;
		}
	}
	/**
	 * セル値を取得します。<br>
	 * 計算式が指定されている場合は再計算を行わずに提供します。<br>
	 * @return 文字列として変換したセル値
	 
	public Object getValue() {
		return getValue(false);
	}	*/
    public static void evaluateCell(Workbook workbook, Cell cell) {
        if (workbook == null) {
                return;
        }
        if (cell == null) {
                return;
        }
        if (cell.getCellType() != Cell.CELL_TYPE_FORMULA) {
                return;
        }
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        try {
                evaluator.evaluateFormulaCell(cell);
        } catch (Throwable e) {
                // 非対応自動計算部は無視する
        }
    }	
    public static Object getBaseValue(Cell cell) {
        if (cell == null) {
                return EMPTY;
        }
        try {
                // 真偽値取得
                return cell.getBooleanCellValue();
        } catch (IllegalStateException e1) {
                try {
                        // 数値取得
                        double d = cell.getNumericCellValue();
                        if ((double) ((long) d) == d) {
                                return (long) d;
                        } else {
                                return d;
                        }
                } catch (IllegalStateException e2) {
                        try {
                                // 文字列取得
                                return cell.getStringCellValue();
                        } catch (IllegalStateException e3) {
                                try {
                                        // エラー取得
                                        //return cell.getErrorCellValue();
                                        return ErrorConstants.getText(cell.getErrorCellValue());
                                } catch (IllegalStateException e4) {
                                        return EMPTY;
                                }
                        }
                }
        }
    } 
}
