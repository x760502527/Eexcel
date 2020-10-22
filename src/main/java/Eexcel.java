
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.MediaType;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

public class Eexcel {
    private Workbook workbook;
    private Sheet sheet;
    private int crow = 0;
    private int ccol = 0;
    private Row row = null;
    private Cell cell = null;

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Cell getCell() {
        return cell;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public int getCrow() {
        return crow;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public Row getRow() {
        return row;
    }

    public void setCrow(int crow) {
        this.crow = crow;
    }

    public int getCcol() {
        return ccol;
    }

    public void setCcol(int ccol) {
        this.ccol = ccol;
    }

    /**
     * @param
     * @return
     * @Date 2017/12/6 13:53
     * @Author dxcr
     * @Description 获得样式
     */
    public CellStyle getCellStyle() {
        CellStyle style = workbook.createCellStyle();
        return style;
    }

    /**
     * @param
     * @return
     * @Date 2017/12/6 13:52
     * @Author dxcr
     * @Description 获得设置后的样式
     */
    public CellStyle getCellStyleByTitle() {
        CellStyle cellStyleHead = workbook.createCellStyle();//创建单元格样式
        cellStyleHead.setVerticalAlignment(VerticalAlignment.CENTER);//垂直对齐方式
        cellStyleHead.setAlignment(HorizontalAlignment.CENTER);//垂直对齐方式
        Font fontStyle = workbook.createFont();
        fontStyle.setFontHeightInPoints((short) 16);
        cellStyleHead.setFont(fontStyle);
        return cellStyleHead;
    }

    public CellStyle getCellStyleByTable() {
        CellStyle cellStyle = workbook.createCellStyle();//创建单元格样式
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置单元格上下左右边框线
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * @param path 文件路径
     * @return 是否成功
     * @Date 2017/12/6 14:17
     * @Author dxcr
     * @Description 生成文件
     */
    public boolean writeFile(String path) {
        FileOutputStream fileOut = null;
        try {
            File file = new File(path);
            if (!file.exists()) {
                file.createNewFile();
            }
            fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
        } catch (Exception ex) {
            ex.printStackTrace();
            return false;
        } finally {
            if (fileOut != null) {
                try {
                    fileOut.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }

    public void writeResponse(String fileName, HttpServletResponse response) {
        try {
            response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
            response.flushBuffer();
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void close() {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (Exception e) {
                throw new RuntimeException("关闭excel连接失败");
            }
        }
    }

    public Eexcel(MultipartFile file) {
        String fileName = file.getOriginalFilename();

        if (fileName.matches("^.+\\.(?i)(xls)$")) {
            workbook = new HSSFWorkbook();//创建excel文件
        } else if (fileName.matches("^.+\\.(?i)(xlsx)$")) {
            workbook = new XSSFWorkbook();//创建excel文件
        } else {
            throw new RuntimeException("Eexcel格式不正确");
        }
    }

    public Eexcel() {
        workbook = new HSSFWorkbook();//创建excel文件
        init();
    }

    public Eexcel(String type) {
        if ("xlsx".equals(type)) {
            workbook = new XSSFWorkbook();//创建excel文件
        } else {
            workbook = new HSSFWorkbook();//创建excel文件
        }
        init();
    }

    public void init() {
        sheet = workbook.createSheet("sheet1");//创建工作表
        row = sheet.createRow(crow);
    }

    /**
     * @param ccol 切换的列数
     * @return 当前类
     * @Date 2017/12/6 9:18
     * @Author dxcr
     * @Description 切换列
     */
    public Eexcel toCol(int ccol) {
        this.ccol = ccol;
        return this;
    }

    /**
     * @param crow 切换的行数
     * @return 当前类
     * @Date 2017/12/6 9:05
     * @Author dxcr
     * @Description 切换行
     */
    public Eexcel toRow(int crow) {
        this.crow = crow;
        /*判断切换行是否创建*/
        row = sheet.getRow(crow);
        if (row == null) {
            row = sheet.createRow(crow);
        }
        this.row = row;
        /*切换当前列为最后列*/
        int ccol = row.getLastCellNum();
        /*小于0 是刚创建行切换到第一列*/
        if (ccol < 0) {
            ccol = 0;
        }
        this.ccol = ccol;
        return this;
    }

    /**
     * @param value 内容
     * @return
     * @Date 2017/12/6 14:36
     * @Author dxcr
     * @Description 设置当前单元格内容
     */
    public Eexcel setCell(String value) {
        cell.setCellValue(value);
        return this;
    }

    /**
     * @param value 内容
     * @param col   列数
     * @return
     * @Date 2017/12/6 14:36
     * @Author dxcr
     * @Description 设置当前行某个单元格内容
     */
    public Eexcel setCell(String value, int col) {
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        cell.setCellValue(value);
        return this;
    }

    /**
     * @param value 内容
     * @param col   列数
     * @param row   行数
     * @return
     * @Date 2017/12/6 14:36
     * @Author dxcr
     * @Description 设置某个单元格内容
     */
    public Eexcel setCell(String value, int col, int row) {
        Row row1 = sheet.getRow(row);
        if (row1 == null) {
            row1 = sheet.createRow(row);
        }
        Cell cell = row1.getCell(col);
        if (cell == null) {
            cell = row1.createCell(col);
        }
        cell.setCellValue(value);
        return this;
    }

    public Eexcel setCellStyle(CellStyle style) {
        int num = isMergedRegion(sheet, ccol, crow);
        if (num != -1) {
            MergedRegionStyle(num, style);
        } else {
            cell.setCellStyle(style);
        }
        return this;
    }

    public Eexcel setCellStyle(CellStyle style, int col) {
        int num = isMergedRegion(sheet, col, crow);
        if (num != -1) {
            MergedRegionStyle(num, style);
        } else {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellStyle(style);
        }
        return this;
    }

    public Eexcel setCellStyle(CellStyle style, int col, int row) {
        int num = isMergedRegion(sheet, col, row);
        if (num != -1) {
            MergedRegionStyle(num, style);
        } else {
            Row row1 = sheet.getRow(row);
            if (row1 == null) {
                row1 = sheet.createRow(row);
            }
            Cell cell = row1.getCell(col);
            if (cell == null) {
                cell = row1.createCell(col);
            }
            cell.setCellStyle(style);
        }
        return this;
    }

    /**
     * @param style 样式
     * @param cols  开始行
     * @param cole  结束行
     * @param rows  开始列
     * @param rowe  结束列
     * @return
     * @Date 2017/12/6 15:52
     * @Author dxcr
     * @Description 范围设置单元格样式
     */
    public void setRegionStyle(CellStyle style, int cols, int cole, int rows, int rowe) {
        for (int i = rows; i <= rowe; i++) {
            for (int j = cols; j <= cole; j++) {
                Row row1 = sheet.getRow(i);
                if (row1 == null) {
                    row1 = sheet.createRow(i);
                }
                Cell cell = row1.getCell(j);
                if (cell == null) {
                    cell = row1.createCell(j);
                }
                cell.setCellStyle(style);
            }
        }
    }

    /**
     * @param value 单元格值
     * @param scol  合并列数
     * @param srow  合并行数
     * @return
     * @Date 2017/12/6 9:07
     * @Author dxcr
     * @Description 追加列
     */
    public Eexcel append(String value, int scol, int srow) {
        if (scol < 1) {
            scol = 1;
        }
        if (srow < 1) {
            srow = 1;
        }
        int i = isMergedRegion(sheet, ccol, crow);
        if (cell != null) {
            if (i != -1) {
                MergedRegionChange(sheet, scol, srow);
            } else {
                MergedRegionChange(sheet, scol, srow);
            }
        } else {
            /*是合并单元格的副格*/
            if (i != -1) {
                MergedRegionChange(sheet, scol, srow);
            } else {
                MergedRegionChange(sheet, scol, srow);
            }
        }
        while (true) {
            cell = row.getCell(ccol);
            if (cell != null) {
                ccol++;
            } else {
                break;
            }
        }
        cell = row.createCell(ccol);
        cell.setCellValue(value);
        if (scol > 1 || srow > 1) {
            sheet.addMergedRegion(new CellRangeAddress(crow, crow + srow - 1, ccol, ccol + scol - 1));//合并单元格
        }
        ccol += scol;
        return this;
    }

    public Eexcel append(String value, int scol, int srow, CellStyle style) {
        if (scol < 1) {
            scol = 1;
        }
        if (srow < 1) {
            srow = 1;
        }
        int b = isMergedRegion(sheet, ccol, crow);
        /*未创建并且是合并单元格*/
        if (cell != null) {
            if (b != -1) {
                MergedRegionChange(sheet, scol, srow);
            } else {
                MergedRegionChange(sheet, scol, srow);
            }
        } else {
            /*是合并单元格的副格*/
            if (b != -1) {
                MergedRegionChange(sheet, scol, srow);
            } else {
                MergedRegionChange(sheet, scol, srow);
            }
        }
        while (true) {
            cell = row.getCell(ccol);
            if (cell != null) {
                ccol++;
            } else {
                break;
            }
        }

        for (int i = 0; i < srow; i++) {
            for (int j = 0; j < scol; j++) {
                if (i == 0 && j == 0) {
                    cell = row.createCell(ccol);
                    cell.setCellValue(value);
                    cell.setCellStyle(style);
                } else {
                    Row r = sheet.getRow(crow + i);
                    if (r == null) {
                        r = sheet.createRow(crow + i);
                    }
                    Cell c = r.getCell(ccol + j);
                    if (c == null) {
                        c = r.createCell(ccol + j);
                    }
                    c.setCellStyle(style);
                }
            }
        }
        if (scol > 1 || srow > 1) {
            sheet.addMergedRegion(new CellRangeAddress(crow, crow + srow - 1, ccol, ccol + scol - 1));//合并单元格
        }
        ccol += scol;
        return this;
    }

    public Eexcel append(String value, CellStyle style) {
        int i = isMergedRegion(sheet, ccol, crow);
        if (i != -1) {
            int j = 0;
            while (j != -1) {
                ccol++;
                j = isMergedRegion(sheet, ccol, crow);
            }
        }
        while (true) {
            cell = row.getCell(ccol);
            if (cell != null) {
                ccol++;
            } else {
                break;
            }
        }
        cell = row.createCell(ccol);
        cell.setCellValue(value);
        cell.setCellStyle(style);
        ccol += 1;
        return this;
    }

    public Eexcel append(String value) {
        int i = isMergedRegion(sheet, ccol, crow);
        /*未创建并且是合并单元格*/
        if (i != -1) {
            int j = 0;
            while (j != -1) {
                ccol++;
                j = isMergedRegion(sheet, ccol, crow);
            }
        }
        while (true) {
            cell = row.getCell(ccol);
            if (cell != null) {
                ccol++;
            } else {
                break;
            }
        }
        cell = row.createCell(ccol);
        cell.setCellValue(value);
        ccol += 1;
        return this;
    }

    /**
     * @return java.lang.String
     * @Param
     * @Date 2020/10/21
     * @Author dxcr
     * @Description
     */
    public String read() {
        Cell mcell = row.getCell(ccol);
        String value = getCellValue(mcell);
        step();
        return value;
    }

    /**
    * @Param file
    * @return java.util.List
    * @Date 2020/10/21
    * @Author dxcr
    * @Description  读取execl转为list
    */
    public static List<List<String>> readToList(MultipartFile file) {
        List<List<String>> dataList = new ArrayList<>();

        String fileName = file.getOriginalFilename();
        if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
            return dataList;
        }
        Workbook workbook = null;
        try {
            InputStream is = file.getInputStream();
            if (fileName.endsWith("xlsx")) {
//                FileInputStream is = new FileInputStream(new File(path));
                workbook = new XSSFWorkbook(is);
            }
            if (fileName.endsWith("xls")) {
//                FileInputStream is = new FileInputStream(new File(path));
                workbook = new HSSFWorkbook(is);
            }
            if (workbook != null) {

                //默认读取第一个sheet
                Sheet sheet = workbook.getSheetAt(0);

                boolean firstRow = true;
                for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    //首行  提取注解
                    if (firstRow) {
                        List list = new ArrayList();
                        for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                            Cell cell = row.getCell(j);
                            String cellValue = getCellValue(cell);
                            list.add(cellValue);
                        }
                        dataList.add(list);
                        firstRow = false;
                    } else {
                        //忽略空白行
                        if (row == null) {
                            continue;
                        }
                        List list = new ArrayList();
                        //判断是否为空白行
                        boolean allBlank = true;
                        for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                            Cell cell = row.getCell(j);
                            String cellValue = getCellValue(cell);
                            list.add(cellValue);
                        }
                        dataList.add(list);
                    }
                }
            }
        } catch (Exception e) {
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                }
            }
        }
        return dataList;
    }

    public Eexcel step() {
        ccol += 1;
        cell = row.getCell(ccol);
        return this;
    }

    public Eexcel next() {
        crow++;
        ccol = 0;
        row = sheet.getRow(crow);
        if (row == null) {
            row = sheet.createRow(crow);
        }
        return this;
    }

    private static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return "";
        }
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                cellValue = DateFormatUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
            } else {
                NumberFormat nf = NumberFormat.getInstance();
                cellValue = String.valueOf(nf.format(cell.getNumericCellValue())).replace(",", "");
            }
        } else if (cell.getCellTypeEnum() == CellType.STRING) {
            cellValue = String.valueOf(cell.getStringCellValue());
        } else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellTypeEnum() == CellType.ERROR) {
            cellValue = "error";
        } else {
            cellValue = "";
        }
        return cellValue;
    }

    /**
     * @param sheet
     * @param row
     * @param column
     * @return
     * @Date 2017/12/6 9:22
     * @Author dxcr
     * @Description 是否合并单元格 是:当前列+1并重新运行
     */
    private int isMergedRegion(Sheet sheet, int column, int row) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return i;
                }
            }
        }
        return -1;
    }

    private void MergedRegionStyle(int num, CellStyle style) {
        CellRangeAddress ca = sheet.getMergedRegion(num);
        int firstColumn = ca.getFirstColumn();
        int lastColumn = ca.getLastColumn();
        int firstRow = ca.getFirstRow();
        int lastRow = ca.getLastRow();
        for (int i = firstRow; i <= lastRow; i++) {
            for (int j = firstColumn; j <= lastColumn; j++) {
                Row r = sheet.getRow(i);
                if (r == null) {
                    r = sheet.createRow(i);
                }
                Cell c = r.getCell(j);
                if (c == null) {
                    c = r.createCell(j);
                }
                c.setCellStyle(style);
            }
        }
    }

    private void MergedRegionChange(Sheet sheet, int scol, int srow) {
        for (int i1 = 0; i1 < scol; i1++) {
            for (int j1 = 0; j1 < srow; j1++) {
                int j = isMergedRegion(sheet, ccol + j1, crow + i1);
                if (j != -1) {
                    ccol++;
                    MergedRegionChange(sheet, ccol, crow);
                }
            }
        }
    }

    private boolean isCellExisting(Row row, int col) {
        Cell cell = row.getCell(col);
        if (cell == null) {
            return false;
        } else {
            return true;
        }
    }

    private boolean isRowExisting(Sheet sheet, int row) {
        Row row1 = sheet.getRow(row);
        if (row1 == null) {
            return false;
        } else {
            return true;
        }
    }

}
