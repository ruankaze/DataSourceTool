package cn.songlin.service;

import cn.songlin.dto.TableColumnDto;
import cn.songlin.entity.TableColumn;
import cn.songlin.mapper.TableColumnMapper;
import cn.songlin.utils.ExcelUtil;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Service
public class ExcelService {

    @Autowired
    private TableColumnMapper columnMapper;

    private Logger logger = Logger.getLogger(this.getClass());

    /**
     * @Function: TODO
     * @description: 创建多表结构
     * @author: NingZe
     * @date: 2020/5/21 0021 11:13
     * @params: [map, tableColumnDto]
     * @version: 02.06
     * @return: void
     */
    public void createTableColumnExcel(Map<String, List<TableColumn>> map, TableColumnDto tableColumnDto) throws Exception {
        // 原版
        //createTableColumnExcels1(map, tableColumnDto);
        // 美化版
        createTableColumnExcels2(map, tableColumnDto);
    }

    /**
     * @Function: TODO
     * @description: To Excel 原版
     *
     * @author: liusonglin
     * @date: 2020/5/21 0021 11:16
     * @params: [map, tableColumnDto]
     * @version: 02.06
     * @return: void
     */
    public void createTableColumnExcels1(Map<String, List<TableColumn>> map, TableColumnDto tableColumnDto)
            throws Exception {
        if (StringUtils.isEmpty(tableColumnDto.getExcelFileName())) {
            throw new Exception("未输入导出文件名");
        } else {
            File desktopDir = FileSystemView.getFileSystemView().getHomeDirectory();
            tableColumnDto.setExcelFileName(desktopDir + "/" + tableColumnDto.getExcelFileName() + ".xlsx");
        }

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();

        // 默认 列宽（实际数值为 24） 。
        int defaultwidth = 24 - 4;

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(tableColumnDto.getSheetName());
        sheet.setDefaultColumnWidth(defaultwidth);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        // 声明列对象
        HSSFCell cell = null;

        HSSFRow row = sheet.createRow(0);
        row.setHeight((short) (50 * 20));// 设置行高
        cell = row.createCell(2);
        cell.setCellValue(tableColumnDto.getDataSourceName() + "数据库所有表的表结构");
        cell.setCellStyle(style);

        // POI导出EXCEL设置跨行跨列（在所有数据行和列创建完成后再执行）
        CellRangeAddress range = new CellRangeAddress(0, 0, 2, 9);
        sheet.addMergedRegion(range);

        // 数据处理
        // logger.info(sheet.getLastRowNum());
        // row = sheet.createRow(sheet.getLastRowNum()+3);//向下偏移3行
        String[][] content = null;
        Set<String> keySet = map.keySet();
        for (String key : keySet) {

            String tabcomment = columnMapper.getTabComment(tableColumnDto.getDataSourceName(), key);
            row = sheet.createRow(sheet.getLastRowNum() + 2);// 向下偏移2行
            row.createCell(2).setCellValue("表名：" + tabcomment + "（" + key + "）");
            row = sheet.createRow(sheet.getLastRowNum() + 1);// 向下偏移1行
            // 创建标题
            for (int i = 0; i < tableColumnDto.getExcelTitle().length; i++) {
                cell = row.createCell(i + 2);// 单元格从第二列开始
                cell.setCellValue(tableColumnDto.getExcelTitle()[i]);
            }
            List<TableColumn> tableColumns = map.get(key);
            content = new String[tableColumns.size()][tableColumnDto.getExcelTitle().length];

            // 创建内容
            setTabVals(content, tableColumns);

            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i < content.length; i++) {
                row = sheet.createRow(i + 1 + lastRowNum);
                for (int j = 0; j < content[i].length; j++) {
                    // 将内容按顺序赋给对应的列对象
                    row.createCell(j + 2).setCellValue(content[i][j]);
                }
            }
        }

        // 导入指定地址的Excel
        OutputStream os = new FileOutputStream(tableColumnDto.getExcelFileName());
        wb.write(os);
        os.close();
    }

    /**
     * @Function: TODO
     * @description: To Excel 美化版
     * @author: NingZe
     * @date: 2020/5/21 0021 11:15
     * @params: [map, tableColumnDto]
     * @version: 02.06
     * @return: void
     */
    public void createTableColumnExcels2(Map<String, List<TableColumn>> map, TableColumnDto tableColumnDto)
            throws Exception {
        if (StringUtils.isEmpty(tableColumnDto.getExcelFileName())) {
            throw new Exception("未输入导出文件名");
        } else {
            File desktopDir = FileSystemView.getFileSystemView().getHomeDirectory();
            tableColumnDto.setExcelFileName(desktopDir + "/" + tableColumnDto.getExcelFileName() + ".xlsx");
        }

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();

        // 表头字体
        HSSFFont headfont = wb.createFont();
        headfont.setFontName("微软雅黑");
        headfont.setFontHeightInPoints((short) 15);
        // 标题字体
        HSSFFont titlefont = wb.createFont();
        titlefont.setFontName("微软雅黑");
        titlefont.setColor(HSSFColor.WHITE.index);
        titlefont.setFontHeightInPoints((short) 11);
        titlefont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 通用字体
        HSSFFont font = wb.createFont();
        font.setFontName("微软雅黑");
        // 通用样式（水平居中、垂直居中、字体）
        HSSFCellStyle style = createHSSFCellStyle(wb, font, (short) 0);
        // 表头样式（水平居中、垂直居中、背景色、字体）
        HSSFCellStyle headstyle = createHSSFCellStyle(wb, headfont, HSSFColor.LIGHT_GREEN.index);
        // 标题样式（水平居中、垂直居中、背景色、字体）
        HSSFCellStyle titlestyle = createHSSFCellStyle(wb, titlefont, HSSFColor.LIGHT_BLUE.index);
        // 列1样式（水平居中、垂直居中、背景色、字体）
        HSSFCellStyle colstyle1 = createHSSFCellStyle(wb, font, HSSFColor.PALE_BLUE.index);
        // 列2样式（水平居中、垂直居中、背景色、字体）
        HSSFCellStyle colstyle2 = createHSSFCellStyle(wb, font, HSSFColor.LIGHT_YELLOW.index);
        // 拿到palette颜色板 - 替换颜色
        HSSFPalette palette = wb.getCustomPalette();
        // HSSFColor.LIGHT_GREEN.index 替换为 RGB(198,239,206)
        palette.setColorAtIndex(HSSFColor.LIGHT_GREEN.index, (byte) 198, (byte) 239, (byte) 206);
        // HSSFColor.LIGHT_BLUE.index 替换为 RGB(79,129,189)
        palette.setColorAtIndex(HSSFColor.LIGHT_BLUE.index, (byte) 79, (byte) 129, (byte) 189);
        // HSSFColor.PALE_BLUE.index 替换为 RGB(184,204,228)
        palette.setColorAtIndex(HSSFColor.PALE_BLUE.index, (byte) 184, (byte) 204, (byte) 228);
        // HSSFColor.LIGHT_YELLOW.index 替换为 RGB(220,230,241)
        palette.setColorAtIndex(HSSFColor.LIGHT_YELLOW.index, (byte) 220, (byte) 230, (byte) 241);

        // 统计个数
        int total = 1;
        // 数据处理
        String[][] content = null;
        Set<String> keySet = map.keySet();
        for (String key : keySet) {

            // 表注释
            String tabcomment = columnMapper.getTabComment(tableColumnDto.getDataSourceName(), key);

            // 默认 列宽、行高（实际数值为 33、20） 。
            int defaultwidth = 33 - 4;
            short defaultheight = 20 * 20;

            // 第二步，在workbook中添加一个sheet, 对应Excel文件中的sheet
            HSSFSheet sheet = wb.createSheet(total + " " + tabcomment);
            sheet.setDefaultColumnWidth(defaultwidth);
            sheet.setDefaultRowHeight(defaultheight);

            // 声明列对象
            HSSFCell cell = null;

            // 第四步，创建单元格，设置表头值，设置表头样式
            HSSFRow row = sheet.createRow(0);
            row.setHeight((short) (50 * 20));
            cell = row.createCell(0);
            cell.setCellValue(key + " " + tabcomment);
            cell.setCellStyle(headstyle);
            // 表头合并（0-0行，0-7列合并）
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));

            // 向下偏移1行
            row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.setHeight(defaultheight);
            // 创建标题
            for (int i = 0; i < tableColumnDto.getExcelTitle().length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(tableColumnDto.getExcelTitle()[i]);
                cell.setCellStyle(titlestyle);
            }
            List<TableColumn> tableColumns = map.get(key);
            content = new String[tableColumns.size()][tableColumnDto.getExcelTitle().length];

            // 创建内容
            setTabVals(content, tableColumns);

            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i < content.length; i++) {
                row = sheet.createRow(i + 1 + lastRowNum);
                row.setHeight(defaultheight);
                for (int j = 0; j < content[i].length; j++) {
                    // 将内容按顺序赋给对应的列对象
                    cell = row.createCell(j);
                    cell.setCellValue(content[i][j]);
                    if (i % 2 == 0) {
                        cell.setCellStyle(colstyle1);
                    } else {
                        cell.setCellStyle(colstyle2);
                    }
                }
            }


            // 导入指定地址的Excel
            OutputStream os = new FileOutputStream(tableColumnDto.getExcelFileName());
            wb.write(os);
            os.close();

            total++;

        }

    }

    /**
     * 创建常规单表结构
     *
     * @param tableColumns
     * @param tableColumnDto
     * @author liusonglin
     * @date 2018年7月25日
     */

    public void createSingleTableColumnExcel(List<TableColumn> tableColumns, TableColumnDto tableColumnDto)
            throws Exception {
        if (StringUtils.isEmpty(tableColumnDto.getExcelFileName())) {
            throw new Exception("未输入导出文件名");
        } else {
            File desktopDir = FileSystemView.getFileSystemView().getHomeDirectory();
            tableColumnDto.setExcelFileName(desktopDir + "/" + tableColumnDto.getExcelFileName() + ".xlsx");
        }

        String[][] content = new String[tableColumns.size()][tableColumnDto.getExcelTitle().length];

        // 创建内容
        setTabVals(content, tableColumns);

        // 创建HSSFWorkbook
        HSSFWorkbook wb = ExcelUtil.getHSSFWorkbook(tableColumnDto.getSheetName(), tableColumnDto.getExcelTitle(),
                content, null);

        // 响应到客户端
        OutputStream os = new FileOutputStream(tableColumnDto.getExcelFileName());
        wb.write(os);
        os.close();

    }

    /**
     * 创建样式
     *
     * @param wb
     * @param font
     * @param colorshort
     * @return
     */
    private HSSFCellStyle createHSSFCellStyle(HSSFWorkbook wb, HSSFFont font, short colorshort) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        if (colorshort >= 0) {
            style.setFillForegroundColor(colorshort);
        }
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFont(font);
        return style;
    }

    /**
     * 创建内容
     *
     * @param content
     * @param tableColumns
     */
    private void setTabVals(String[][] content, List<TableColumn> tableColumns) {
        for (int i = 0; i < tableColumns.size(); i++) {
            TableColumn tableColumn = tableColumns.get(i);
            content[i][0] = tableColumn.getColumnName();
            content[i][1] = tableColumn.getColumnType();
            content[i][2] = tableColumn.getDataType();
            content[i][3] = tableColumn.getCharacterMaximumLength();
            content[i][4] = tableColumn.getColumnKey();
            content[i][5] = tableColumn.getIsNullable();
            content[i][6] = tableColumn.getColumnDefault();
            content[i][7] = tableColumn.getColumnComment();
        }
    }

}
