package hu.ivgraai.gitstats.reader.excel;

import hu.ivgraai.gitstats.reader.Utility;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @author igergo
 * @since Sept 15, 2016
 */
class PoiSheet implements ISheet {

    private final HSSFSheet sheet;
    private final PoiWorkbook.CellSytleManager manager;

    public PoiSheet(HSSFSheet sheet, PoiWorkbook.CellSytleManager manager) {
        this.sheet = sheet;
        this.manager = manager;
    }

    @Override
    public void setRows(List<ExportRow> rows) {
        for (int i = 0; i < rows.size(); ++i) {
            HSSFRow row = sheet.createRow(i);

            ExportRow.ExportCell[] elements = rows.get(i).getElements();
            for (int j = 0; j < elements.length; ++j) {
                HSSFCell cell = row.createCell(j);

                short color, alignment;
                if (IWorkbook.CellColor.Grey == elements[j].backgroundColor) {
                    color = HSSFColor.GREY_25_PERCENT.index;
                } else {
                    color = HSSFColor.WHITE.index;
                }
                switch (elements[j].alignment) {
                case Left:
                    alignment = CellStyle.ALIGN_LEFT;
                    break;
                case Right:
                    alignment = CellStyle.ALIGN_RIGHT;
                    break;
                case Center:
                    alignment = CellStyle.ALIGN_CENTER;
                    break;
                default:
                    alignment = CellStyle.ALIGN_GENERAL;
                }
                cell.setCellStyle(manager.getStyle(color, alignment));

                if (null == elements[j].value) {
                    cell.setCellValue("");
                } else {
                    if ((Integer.class == elements[j].type) || (Long.class == elements[j].type) || (Double.class == elements[j].type)) {
                        try {
                            cell.setCellValue(Double.parseDouble(elements[j].value.toString()));
                        } catch (NumberFormatException e) {
                            cell.setCellValue(elements[j].value.toString());
                        }
                    } else if (Date.class == elements[j].type) {
                        cell.setCellValue(Utility.convertToString((Date) elements[j].value, "yyyy.MM.dd HH:mm:ss"));
                    } else {
                        cell.setCellValue(elements[j].value.toString());
                    }
                }
            }
        }

        autoSizeColumns(rows);
    }

    @Override
    public List<Object[]> getRows() {
        List<Object[]> result = new ArrayList<>();

        /* HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(workbook);
        evaluator.clearAllCachedResultValues();
        evaluator.evaluateAll(); */

        for(int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            List<Object> temp = new ArrayList<>();
            temp.add(i + 1);

            HSSFRow row = sheet.getRow(i);
            for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); ++j) {
                HSSFCell cell = row.getCell(j);

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        temp.add(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        temp.add(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        temp.add(cell.getNumericCellValue());
                        break;
                    default:
                        temp.add(null);
                }
            }
            result.add(temp.toArray(new Object[temp.size()]));
        }

        return result;
    }

    private void autoSizeColumns(List<ExportRow> rows) {
        for (int j = 0; j < rows.get(0).getElements().length; ++j) {
            sheet.autoSizeColumn(j);
        }
    }

}
