package com.wp.velocity;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFRegionUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.awt.*;
import java.io.IOException;
import java.util.*;
import java.util.List;

/**
 * 将html格式的列表以excel导出
 * 类名称：ParseHtmlToXls
 * 类描述：
 * 创建人：wangping
 * 修改人：wangping
 * 修改时间： 2015年3月26日 下午8:04:39
 * 修改备注：
 */
public class ParseHtmlToXls {

    /**
     * 每一个sheet的最大行数
     */
    public static final int SHEET_MAX_ROWS = 20000;

    /**
     * 每一列的最小宽度
     */
    private static final int COL_MIN_WIDTH = 5 * 256;

    /**
     * 每一列的最大宽度
     */
    private static final int COL_MAX_WIDTH = 35 * 256;

    /**
     * 保存已经添加的颜色和序号的映射表
     */
    private static Map<Color, Short> colorMap = new HashMap<Color, Short>();

    /**
     * 保存已经添加的颜色序号映射
     */
    private static Map<Short, Short> colorExistsMap = new HashMap<Short, Short>();

    /**
     * 存放已经保存对应颜色的样式映射表
     */
    private static Map<Color, HSSFCellStyle> colorStyleMap = new HashMap<Color, HSSFCellStyle>();

    /**
     * 保存所有可以进行修改颜色序号的数组
     */
    public static final short[] CUSTOM_INDEX_ARR =
            new short[]{HSSFColor.PLUM.index, HSSFColor.BROWN.index, HSSFColor.OLIVE_GREEN.index, HSSFColor.DARK_GREEN.index,
                    HSSFColor.SEA_GREEN.index, HSSFColor.DARK_TEAL.index, HSSFColor.GREY_40_PERCENT.index, HSSFColor.BLUE_GREY.index,
                    HSSFColor.ORANGE.index, HSSFColor.LIGHT_ORANGE.index, HSSFColor.GOLD.index, HSSFColor.LIME.index, HSSFColor.AQUA.index,
                    HSSFColor.LIGHT_BLUE.index, HSSFColor.TAN.index, HSSFColor.LAVENDER.index, HSSFColor.ROSE.index, HSSFColor.PALE_BLUE.index,
                    HSSFColor.LIGHT_YELLOW.index, HSSFColor.LIGHT_GREEN.index, HSSFColor.LIGHT_TURQUOISE.index, HSSFColor.SKY_BLUE.index,
                    HSSFColor.BLUE.index, HSSFColor.TEAL.index, HSSFColor.DARK_RED.index};

    /**
     * 取得标题的样式
     *
     * @param wb
     * @return
     * @throws
     */
    private static HSSFCellStyle getTitleStyle(HSSFWorkbook wb) {
        // 设置字体
        HSSFFont font = wb.createFont();
        font.setFontName("宋体");
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 加粗

        HSSFCellStyle titleStyle = wb.createCellStyle();
        titleStyle.setBorderLeft((short) 1);
        titleStyle.setBorderRight((short) 1);
        titleStyle.setBorderBottom((short) 1);
//		titleStyle.setBorderTop(HSSFCellStyle.BORDER_DOUBLE);
        titleStyle.setFillForegroundColor((short) 46);
        titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        titleStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);// 左右居中
        titleStyle.setFont(font);
        return titleStyle;
    }

    private static HSSFCellStyle getHeadStyle(HSSFWorkbook wb) {
        // 设置字体
        HSSFFont font = wb.createFont();
        font.setFontName("宋体");
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 加粗

        HSSFCellStyle headStyle = wb.createCellStyle();
        headStyle.setBorderLeft((short) 1);
        headStyle.setBorderRight((short) 1);
        headStyle.setBorderBottom((short) 1);
//		headStyle.setBorderTop(HSSFCellStyle.BORDER_DOUBLE);
        headStyle.setFillForegroundColor((short) 159);
        headStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        headStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);// 左右居中
        headStyle.setFont(font);
        return headStyle;
    }

    /**
     * 取得列表的样式
     *
     * @param wb
     * @return
     * @throws
     */
    private static HSSFCellStyle getContentStyle(HSSFWorkbook wb) {
        HSSFFont contentFont = wb.createFont();
        contentFont.setFontName("宋体");

        HSSFCellStyle contentSonStyle = wb.createCellStyle();
        contentSonStyle.setBorderBottom((short) 1);
        contentSonStyle.setBorderLeft((short) 1);
        contentSonStyle.setBorderRight((short) 1);
        contentSonStyle.setBorderBottom((short) 1);
        contentSonStyle
                .setFillForegroundColor((short) HSSFColor.LEMON_CHIFFON.index);
        contentSonStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        contentSonStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);// 左居中
        contentSonStyle.setFont(contentFont);
        contentSonStyle.setWrapText(true);

        return contentSonStyle;
    }


    /**
     * 通用单标题列表解析（将html转为工作表对象）
     *
     * @param htmlStr   html文本
     * @param sheetName sheet名称
     * @return 返回工作表对象
     * @throws IOException
     * @throws
     */
    public static HSSFWorkbook parseHtmlToXlsForCommon(String htmlStr, String sheetName)
            throws IOException {

        Document doc = Jsoup.parse(htmlStr);
        // 创建 xls
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = null;

        // 所有行
        Elements eleTrs = doc.select("tr");
        Iterator<?> itTrs = eleTrs.iterator();
        Iterator<?> itTds = null;
        Element eleTd = null;
        Element eleTr = null;
        Element headEle = null;
        HSSFRow row = null;
        HSSFCell cell = null;

        HSSFCellStyle titleStyle = getTitleStyle(wb);
        HSSFCellStyle contentSonStyle = getContentStyle(wb);

        int sheetMaxLength = 20000;
        int sheetNum = 0;
        int iTr = 0;
        int rowNum = 0;
        int jTd = 0;
        Map<Integer, RowNextColumn> nextColMap = new HashMap<Integer, RowNextColumn>();
        List<CellRangeAddress> mergeList = new ArrayList<CellRangeAddress>();
        if (itTrs.hasNext()) {
            headEle = (Element) itTrs.next();
        }
        while (itTrs.hasNext()) {
            eleTr = (Element) itTrs.next();
            if (sheetNum * sheetMaxLength <= iTr) {
                sheetNum++;
                for (CellRangeAddress merge : mergeList) {
                    if (null != sheet) sheet.addMergedRegion(merge);
                }
                if (StringUtils.isEmpty(sheetName)) {
                    sheet = wb.createSheet("Sheet" + sheetNum);
                } else {
                    sheet = wb.createSheet(sheetName);
                }
                mergeList = new ArrayList<CellRangeAddress>();
                insertHeadRow(headEle, sheet, titleStyle, mergeList, nextColMap);
                rowNum = 1;
            }
            row = sheet.createRow(rowNum++);
            itTds = eleTr.children().iterator();
            while (itTds.hasNext()) {
                eleTd = (Element) itTds.next();
                updateMergeList(eleTd, mergeList, rowNum - 1, jTd, nextColMap);
                if (null != nextColMap.get(rowNum - 1)) {
                    jTd = nextColMap.get(rowNum - 1).pollNextColIndex();
                    cell = row.createCell(jTd);
                    jTd = nextColMap.get(rowNum - 1).getFirstNextColIndex();
                } else {
                    if (null != row.getCell(jTd)) {
                        jTd++;
                    }
                    cell = row.createCell(jTd++);
                }
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                cell.setCellStyle(getNewCellStyle(eleTd, contentSonStyle, wb));
                cell.setCellValue(StringUtils.trim(eleTd.text()));
            }
            nextColMap.put(iTr, null);
            iTr++;
            jTd = 0;
        }

        for (CellRangeAddress merge : mergeList) {
            if (null != sheet) {
                sheet.addMergedRegion(merge);
                updateStyle(sheet, merge);
            }
        }

        colorMap.clear();
        colorExistsMap.clear();
        colorStyleMap.clear();

        return wb;
    }

    /**
     * 对多标题列表的解析：所有的标题都需要加上my-title，所有的表头都需要加上my-head
     *
     * @param htmlStr   html字符串
     * @param sheetName sheet名称
     * @return
     * @throws
     */
    public static HSSFWorkbook parseHtmlToXlsForMultiTitle(String htmlStr, String sheetName) {
        Document doc = Jsoup.parse(htmlStr);
        // 创建 xls
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = null;

        // 所有行
        Elements eleTrs = doc.select("tr");
        Iterator<?> itTrs = eleTrs.iterator();
        Iterator<?> itTds = null;
        Element eleTd = null;
        Element eleTr = null;
        HSSFRow row = null;
        HSSFCell cell = null;

        HSSFCellStyle titleStyle = getTitleStyle(wb);
        HSSFCellStyle headStyle = getHeadStyle(wb);
        HSSFCellStyle contentSonStyle = getContentStyle(wb);

        int sheetNum = 0;
        int iTr = 0;
        int rowNum = 0;
        int jTd = 0;
        Map<Integer, RowNextColumn> nextColMap = new HashMap<Integer, RowNextColumn>();
        Map<Integer, Integer> colMaxWidthMap = new HashMap<Integer, Integer>();//保存每一列的最大宽度
        List<CellRangeAddress> mergeList = new ArrayList<CellRangeAddress>();
        while (itTrs.hasNext()) {
            eleTr = (Element) itTrs.next();
            if (sheetNum * SHEET_MAX_ROWS <= iTr) {
                sheetNum++;
                for (CellRangeAddress merge : mergeList) {
                    if (null != sheet) {
                        sheet.addMergedRegion(merge);
                        updateStyle(sheet, merge);
                    }
                }

                updateAllColWidthWithMap(colMaxWidthMap, sheet);

                sheet = wb.createSheet(StringUtils.isEmpty(sheetName) ? "Sheet" : sheetName);
                mergeList = new ArrayList<CellRangeAddress>();
                colMaxWidthMap = new HashMap<Integer, Integer>();
                rowNum = 0;
            }
            row = sheet.createRow(rowNum++);
            itTds = eleTr.children().iterator();
            while (itTds.hasNext()) {
                eleTd = (Element) itTds.next();
                updateMergeList(eleTd, mergeList, rowNum - 1, jTd, nextColMap);
                if (null != nextColMap.get(rowNum - 1)) {
                    jTd = nextColMap.get(rowNum - 1).pollNextColIndex();
                    if (StringUtils.isNotEmpty(eleTd.attr("rowspan"))
                            || StringUtils.isNotEmpty(eleTd.attr("colspan"))) {
                        cell = row.createCell(jTd);
                        updateColMaxWidthMap(colMaxWidthMap, eleTd, jTd);
                    } else {
                        cell = row.createCell(jTd);
                        updateColMaxWidthMap(colMaxWidthMap, eleTd, rowNum - 1);
                    }
                    jTd = nextColMap.get(rowNum - 1).getFirstNextColIndex();
                } else {
                    updateColMaxWidthMap(colMaxWidthMap, eleTd, jTd);
                    if (null != row.getCell(jTd)) {
                        jTd++;
                    }
                    cell = row.createCell(jTd++);
                }
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                if (eleTd.hasAttr("my-title")) {
                    cell.setCellStyle(getNewCellStyle(eleTd, titleStyle, wb));
                } else if (eleTd.hasAttr("my-head")) {
                    cell.setCellStyle(getNewCellStyle(eleTd, headStyle, wb));
                } else {
                    cell.setCellStyle(getNewCellStyle(eleTd, contentSonStyle, wb));
                }
                cell.setCellValue(StringUtils.trim(eleTd.text()));
            }
            nextColMap.put(iTr, null);
            iTr++;
            jTd = 0;
        }

        for (CellRangeAddress merge : mergeList) {
            if (null != sheet) {
                sheet.addMergedRegion(merge);
                updateStyle(sheet, merge);
            }
        }

        updateAllColWidthWithMap(colMaxWidthMap, sheet);

        colorMap.clear();
        colorExistsMap.clear();
        colorStyleMap.clear();

        return wb;
    }

    /**
     * 更新每一列的宽度
     * @param colMaxWidthMap
     * @param sheet
     */
    private static void updateAllColWidthWithMap(Map<Integer, Integer> colMaxWidthMap, HSSFSheet sheet) {
        int colWidth = 0;
        for (Map.Entry<Integer, Integer> entry : colMaxWidthMap.entrySet()) {
            colWidth = COL_MIN_WIDTH < entry.getValue() * 256 ? entry.getValue() * 256 : COL_MIN_WIDTH;
            colWidth = COL_MAX_WIDTH >= colWidth ? colWidth : COL_MAX_WIDTH;
            colWidth = (int) (colWidth * 1.14388);
            sheet.setColumnWidth(entry.getKey(), colWidth);
        }
    }

    /**
     * 更新每列最大的长度map
     *
     * @param colMaxWidthMap
     * @param eleTd          列对象
     * @param col            列序号
     * @throws
     */
    private static void updateColMaxWidthMap(Map<Integer, Integer> colMaxWidthMap, Element eleTd, int col) {
        if (null == colMaxWidthMap.get(col)) {
            colMaxWidthMap.put(col, StringUtils.trim(eleTd.text()).getBytes().length);
        } else if (colMaxWidthMap.get(col) < StringUtils.trim(eleTd.text()).getBytes().length) {
            colMaxWidthMap.put(col, StringUtils.trim(eleTd.text()).getBytes().length);
        }
    }

    /**
     * 插入标题行
     *
     * @param headEle    标题行对象
     * @param sheet      当前sheet对象
     * @param titleStyle 标题行样式对象
     * @param mergeList  合并列表
     * @return 返回sheet对象
     * @throws
     */
    private static HSSFSheet insertHeadRow(Element headEle, HSSFSheet sheet,
                                           HSSFCellStyle titleStyle, List<CellRangeAddress> mergeList, Map<Integer, RowNextColumn> nextColMap) {
        HSSFRow row = null;
        HSSFCell cell = null;
        Iterator<?> itTds = null;
        Element eleTd = null;
        int iTr = 0;
        row = sheet.createRow(iTr);
        itTds = headEle.children().iterator();
        int jTd = 0;
        while (itTds.hasNext()) {
            eleTd = (Element) itTds.next();
            updateMergeList(eleTd, mergeList, iTr, jTd, nextColMap);

            if (null != nextColMap.get(iTr)) {
                if (StringUtils.isNotEmpty(eleTd.attr("rowspan")) || StringUtils.isNotEmpty(eleTd.attr("colspan"))) {
                    cell = row.createCell(jTd);
                } else {
                    cell = row.createCell(nextColMap.get(iTr).getFirstNextColIndex());
                }
                jTd = nextColMap.get(iTr).pollNextColIndex();
            } else {
                cell = row.createCell(jTd++);
            }
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellStyle(getNewCellStyle(eleTd, titleStyle, sheet.getWorkbook()));
            cell.setCellValue(StringUtils.trim(eleTd.text()));

        }
        return sheet;
    }

    /**
     * 更新合并列表，将所有td中的合并操作添加到合并列表中
     * @param eleTd 当前td对象
     * @param mergeList 合并对象
     * @param iTr 当前tr序号
     * @param jTd 当前td序号
     * @param nextColMap 每一行对应的下一个位置对象
     */
    private static void updateMergeList(Element eleTd, List<CellRangeAddress> mergeList, int iTr, int jTd,
                                        Map<Integer, RowNextColumn> nextColMap) {
        String rowspan = "";
        int rowspan_int = 0;
        String colspan = "";
        int colspan_int = 0;
        CellRangeAddress mergeTemp = null;
        rowspan = eleTd.attr("rowspan");
        colspan = eleTd.attr("colspan");

        try {
            if (StringUtils.isEmpty(rowspan) && StringUtils.isEmpty(colspan)) {//没有rowspan和colspan
                if (null == nextColMap.get(iTr)) {//若当前行没有对象
                    nextColMap.put(iTr, new RowNextColumn(iTr));
                } else {
                    jTd = nextColMap.get(iTr).getFirstNextColIndex();
                }

                if (null != nextColMap.get(iTr).getLastNextColIndex() && nextColMap.get(iTr).getLastNextColIndex() > jTd + 1) {
                    return;
                }
                nextColMap.get(iTr).offBatchArgs(jTd, jTd + 1);

            } else if (StringUtils.isNotEmpty(rowspan) && StringUtils.isEmpty(colspan)) {
                rowspan_int = new Integer(rowspan.trim());
                mergeTemp = new CellRangeAddress(iTr, iTr + rowspan_int - 1, jTd, jTd);
                mergeList.add(mergeTemp);

                if (null == nextColMap.get(iTr)) {
                    nextColMap.put(iTr, new RowNextColumn(iTr));
                }
                nextColMap.get(iTr).offBatchArgs(jTd, jTd+1);
                for (int i = iTr + 1; i < iTr + rowspan_int; i++) {
                    if (null == nextColMap.get(i)) {
                        nextColMap.put(i, new RowNextColumn(i));
                    }
                    if (null != nextColMap.get(i).getLastNextColIndex() && nextColMap.get(i).getLastNextColIndex() == jTd) {
                        nextColMap.get(i).popNextColIndex();
                    }else if (null != nextColMap.get(i).getLastNextColIndex()) {
                        nextColMap.get(i).offBatch(nextColMap.get(i).popNextColIndex(), jTd-1);
                    }

                    nextColMap.get(i).offBatchArgs(jTd + 1);
                }
            } else if (StringUtils.isNotEmpty(colspan) && StringUtils.isEmpty(rowspan)) {
                colspan_int = new Integer(colspan.trim());
                mergeTemp = new CellRangeAddress(iTr, iTr, jTd, jTd + colspan_int - 1);
                mergeList.add(mergeTemp);
                if (null == nextColMap.get(iTr)) {
                    nextColMap.put(iTr, new RowNextColumn(iTr));
                }
                nextColMap.get(iTr).offBatchArgs(jTd, jTd + colspan_int);
            } else if (StringUtils.isNotEmpty(rowspan) && StringUtils.isNotEmpty(colspan)) {
                rowspan_int = new Integer(rowspan.trim());
                colspan_int = new Integer(colspan.trim());
                mergeTemp = new CellRangeAddress(iTr, iTr + rowspan_int - 1, jTd, jTd + colspan_int - 1);
                mergeList.add(mergeTemp);
                if (null == nextColMap.get(iTr)) {
                    nextColMap.put(iTr, new RowNextColumn(iTr));
                }
                nextColMap.get(iTr).offBatchArgs(jTd, jTd+rowspan_int);
                for (int i = iTr + 1; i < iTr + rowspan_int; i++) {
                    if (null == nextColMap.get(i)) {
                        nextColMap.put(i, new RowNextColumn(iTr));
                    }

                    if (null != nextColMap.get(i).getLastNextColIndex() && nextColMap.get(i).getLastNextColIndex() == jTd) {
                        nextColMap.get(i).popNextColIndex();
                    }else if (null != nextColMap.get(i).getLastNextColIndex()) {
                        nextColMap.get(i).offBatch(nextColMap.get(i).popNextColIndex(), jTd-1);
                    }

                    nextColMap.get(i).offBatchArgs(jTd + rowspan_int);
                }
            }
        } catch (NumberFormatException e) {
            e.printStackTrace();
        }
    }

    /**
     * 更新当前合并单元格边框长度
     *
     * @param sheet 当前sheet对象
     * @param merge 当前合并单元格对象
     * @throws
     */
    private static void updateStyle(HSSFSheet sheet, CellRangeAddress merge) {
        HSSFRegionUtil.setBorderBottom(1, merge, sheet, sheet.getWorkbook());
        HSSFRegionUtil.setBorderTop(1, merge, sheet, sheet.getWorkbook());
        HSSFRegionUtil.setBorderLeft(1, merge, sheet, sheet.getWorkbook());
        HSSFRegionUtil.setBorderRight(1, merge, sheet, sheet.getWorkbook());
    }

    /**
     * 根据当前单元格对象修改样式对象
     *
     * @param eleTd     单元格对象
     * @param cellStyle 样式对象
     * @param wb        工作表对象
     * @return 返回修改后的样式对象
     * @throws
     */
    private static HSSFCellStyle getNewCellStyle(Element eleTd, HSSFCellStyle cellStyle, HSSFWorkbook wb) {
        HSSFPalette customPalette = wb.getCustomPalette();
        HSSFCellStyle newStyle = null;
        if (eleTd.hasAttr("my-color") && StringUtils.isNotEmpty(eleTd.attr("my-color"))) {
            Color c = ColorUtil.String2Color(eleTd.attr("my-color"));
            short clrIndex = 0;
            if (null != c) {
                if (null != colorMap.get(c)) {
                    clrIndex = colorMap.get(c);
                } else {
                    clrIndex = CUSTOM_INDEX_ARR[colorExistsMap.size()];
                    colorMap.put(c, clrIndex);
                    colorExistsMap.put(clrIndex, clrIndex);
                }
                customPalette.setColorAtIndex(clrIndex, (byte) c.getRed(), (byte) c.getGreen(), (byte) c.getBlue());
                if (null == colorStyleMap.get(c)) {
                    newStyle = wb.createCellStyle();
                    newStyle.cloneStyleFrom(cellStyle);
                    newStyle.setFillForegroundColor(clrIndex);
                    colorStyleMap.put(c, newStyle);
                } else newStyle = colorStyleMap.get(c);
                cellStyle = newStyle;
            }
        }
        return cellStyle;
    }

}

/**
 * 存放每一行中下一个cell的位置对象
 * @author wangping
 * @version 1.0
 * @since 2016/3/15
 */
class RowNextColumn {
    /**
     * 行号
     */
    private int rowIndex = 0;
    /**
     * 一个既可以作为队列（先进先出），也可以作为堆栈（后进先出）使用的链表
     */
    private LinkedList<Integer> nextColIndexList = new LinkedList<Integer>();

    public RowNextColumn(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    /**
     * 在队列最前面插入元素（多个）
     * @param indexs
     */
    public void offBatchArgs(Integer... indexs) {
        for (Integer index : indexs) {
            if (!this.nextColIndexList.contains(index)) {
                this.nextColIndexList.offer(index);
            }
        }
    }

    /**
     * 在堆栈最后插入元素（多个）
     * @param indexs
     */
    public void pushBatchArgs(Integer... indexs) {
        for (Integer index : indexs) {
            if (!this.nextColIndexList.contains(index)) {
                this.nextColIndexList.push(index);
            }
        }
    }

    /**
     * 在队列中批量插入元素（顺序）
     * @param startIndex
     * @param endIndex
     */
    public void offBatch(int startIndex, int endIndex) {
        for (int j = startIndex; j <= endIndex; j++) {
            if (!this.nextColIndexList.contains(j)) {
                this.nextColIndexList.offer(j);
            }
        }
    }

    /**
     * 在堆栈中批量插入元素（顺序）
     * @param startIndex
     * @param endIndex
     */
    public void pushBatch(int startIndex, int endIndex) {
        for (int j = endIndex; j >= startIndex; j--) {
            if (!this.nextColIndexList.contains(j)) {
                this.nextColIndexList.push(j);
            }
        }
    }

    /**
     * 判断在队列中是否存在当前元素
     * @param nextColIndex
     * @return
     */
    public boolean containKey(Integer nextColIndex) {
        return nextColIndexList.contains(nextColIndex);
    }

    /**
     * 弹出最后的元素，并返回对应的值
     * @return
     */
    public Integer popNextColIndex() {
        return nextColIndexList.pollLast();
    }

    /**
     * 弹出最前的元素，并返回对应的值
     * @return
     */
    public Integer pollNextColIndex() {
        return nextColIndexList.pollFirst();
    }

    /**
     * 返回最前面的元素值
     * @return
     */
    public Integer getFirstNextColIndex() {
        Integer nextColIndex = nextColIndexList.peekFirst();
        return nextColIndex;
    }

    /**
     * 返回最后的元素值
     * @return
     */
    public Integer getLastNextColIndex() {
        Integer nextColIndex = nextColIndexList.peekLast();
        return nextColIndex;
    }

}
