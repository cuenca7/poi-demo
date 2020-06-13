package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class ExcelWriter {

    private static List<String> CELL_HEADS;

    static {
        // 类装载时就载入指定好的列头信息，如有需要，可以考虑做成动态生成的列头
        CELL_HEADS = new ArrayList<String>();
        CELL_HEADS.add("组合代码");
        CELL_HEADS.add("数据生成日期");
        CELL_HEADS.add("证券代码");
        CELL_HEADS.add("申赎日期");
        CELL_HEADS.add("申赎数量");
    }

    /**
     * 生成Excel并写入数据信息
     *
     * @param dataList 数据列表
     * @return 写入数据后的工作簿对象
     */
    public static Workbook exportData(List<ExcelDataVO> dataList) {
        // 生成xlsx的Excel
        Workbook workbook = new HSSFWorkbook();

        // 生成Sheet表，写入第一行的列头
        Sheet sheet = buildDataSheet(workbook);
        //构建每行的数据内容
        int rowNum = 2;
        for (ExcelDataVO data : dataList) {
            if (data == null) {
                continue;
            }
            Row row = sheet.createRow(rowNum++);
            convertDataToRow(data, row);
        }
        return workbook;
    }

    private static Sheet buildDataSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        // 设置列头宽度
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            sheet.setColumnWidth(i, 4000);
        }
        // 设置默认行高
        sheet.setDefaultRowHeight((short) 400);
        // 构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
        // 写入第一行合并列数据
        // 基础数据合并
        CellRangeAddress baseRegion = new CellRangeAddress(0, 0, 0, 2);
        sheet.addMergedRegion(baseRegion);
        Row merge = sheet.createRow(0);
        Cell a1 = merge.createCell(0);
        a1.setCellValue("基础数据");

        // 申赎数据合并
        CellRangeAddress prRegion = new CellRangeAddress(0, 0, 3, 4);
        sheet.addMergedRegion(prRegion);
        Cell a3 = merge.createCell(3);
        a3.setCellValue("申赎数据");

        // 写入第二行各列的数据
        Row head = sheet.createRow(1);
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * 设置第一行列头的样式
     *
     * @param workbook 工作簿对象
     * @return 单元格样式对象
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        //style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        //style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        //style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        //style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        //style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(font);
        return style;
    }

    private static void convertDataToRow(ExcelDataVO data, Row row) {
        int cellNum = 0;
        Cell cell;
        // 组合代码
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getFundCode() ? "" : data.getFundCode());
        // 数据生成日期
        cell = row.createCell(cellNum++);
        if (null != data.getRecordCreateDate()) {
            cell.setCellValue(String.valueOf(data.getRecordCreateDate()));
        } else {
            cell.setCellValue("");
        }
        // 证券代码
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getSecurityCode() ? "" : data.getSecurityCode());
        // 申赎日期
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getPurchRedmDate() ? "" : String.valueOf(data.getPurchRedmDate()));
        // 申赎数量
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getPurchRedmQuantity() ? "" : String.valueOf(data.getPurchRedmQuantity()));
    }
}
