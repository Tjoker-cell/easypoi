/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package cn.afterturn.easypoi.word.parse.excel;

import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.entity.params.ExcelForEachParams;
import cn.afterturn.easypoi.util.PoiElUtil;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import cn.afterturn.easypoi.util.PoiWordStyleUtil;
import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import com.google.common.base.Strings;
import com.google.common.collect.Maps;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static cn.afterturn.easypoi.util.PoiElUtil.*;

/**
 * 处理和生成Map 类型的数据变成表格
 *
 * @author JueYue
 * 2014年8月9日 下午10:28:46
 */
public final class ExcelMapParse {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelMapParse.class);

    /**
     * 添加图片
     *
     * @param obj
     * @param currentRun
     * @throws Exception
     * @author JueYue
     * 2013-11-20
     */
    public static void addAnImage(ImageEntity obj, XWPFRun currentRun) {
        try {
            Object[] isAndType = PoiPublicUtil.getIsAndType(obj);
            String   picId;
            picId = currentRun.getDocument().addPictureData((byte[]) isAndType[0],
                    (Integer) isAndType[1]);
            if (obj.getLocationType() == ImageEntity.EMBED) {
                ((MyXWPFDocument) currentRun.getDocument()).createPicture(currentRun,
                        picId, currentRun.getDocument()
                                .getNextPicNameNumber((Integer) isAndType[1]),
                        obj.getWidth(), obj.getHeight());
            } else if (obj.getLocationType() == ImageEntity.ABOVE) {
                ((MyXWPFDocument) currentRun.getDocument()).createPicture(currentRun,
                        picId, currentRun.getDocument()
                                .getNextPicNameNumber((Integer) isAndType[1]),
                        obj.getWidth(), obj.getHeight(), false);
            }  else if (obj.getLocationType() == ImageEntity.BEHIND) {
                ((MyXWPFDocument) currentRun.getDocument()).createPicture(currentRun,
                        picId, currentRun.getDocument()
                                .getNextPicNameNumber((Integer) isAndType[1]),
                        obj.getWidth(), obj.getHeight(), true);
            }


        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }

    }


    private static void addAnImage(ImageEntity obj, XWPFTableCell cell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        XWPFParagraph newPara = paragraphs.get(0);
        XWPFRun imageCellRun = newPara.createRun();
        addAnImage(obj,imageCellRun);
    }
    /**
     * 解析参数行,获取参数列表
     *
     * @param currentRow
     * @return
     * @author JueYue
     * 2013-11-18
     */
    private static String[] parseCurrentRowGetParams(XWPFTableRow currentRow) {
        List<XWPFTableCell> cells  = currentRow.getTableCells();
        String[]            params = new String[cells.size()];
        String              text;
        for (int i = 0; i < cells.size(); i++) {
            text = cells.get(i).getText();
            params[i] = text == null ? ""
                    : text.trim().replace(START_STR, EMPTY).replace(END_STR, EMPTY);
        }
        return params;
    }

    /**
     * 解析参数行,获取参数列表
     *
     * @param currentRow
     * @return
     * @author JueYue
     * 2013-11-18
     */
    private static List<ExcelForEachParams> parseCurrentRowGetParamsEntity(XWPFTableRow currentRow) {
        List<XWPFTableCell>      cells  = currentRow.getTableCells();
        List<ExcelForEachParams> params = new ArrayList<>();
        String                   text;
        for (int i = 0; i < cells.size(); i++) {
            ExcelForEachParams param = new ExcelForEachParams();
            text = cells.get(i).getText();
            param.setName(text == null ? ""
                    : text.trim().replace(START_STR, EMPTY).replace(END_STR, EMPTY));
            if (cells.get(i).getCTTc().getTcPr().getGridSpan() != null) {
                if (cells.get(i).getCTTc().getTcPr().getGridSpan().getVal() != null) {
                    param.setColspan(cells.get(i).getCTTc().getTcPr().getGridSpan().getVal().intValue());
                }
            }
            param.setHeight((short) currentRow.getHeight());
            params.add(param);
        }
        return params;
    }

    /**
     * 解析下一行,并且生成更多的行
     *
     * @param table
     * @param index
     * @param list
     */

    public static void parseNextRowAndAddRow(XWPFTable table, int index, List<Object> list, int col) throws Exception {
        XWPFTableRow currentRow = table.getRow(index);
        String[] params = parseCurrentRowGetParams(currentRow);
        String listname = params[col];
        boolean isCreate = !listname.contains("!fe:");
        listname = listname.replace("!fe:", "").replace("$fe:", "").replace("fe:", "").replace("{{", "");
        String[] keys = listname.replaceAll("\\s{1,}", " ").trim().split(" ");
        params[col] = keys[1];
        List<XWPFTableCell> tempCellList = new ArrayList();
        tempCellList.addAll(table.getRow(index).getTableCells());
//        int cellIndex = false;
        Map<String, Object> tempMap = Maps.newHashMap();
        LOGGER.debug("start for each data list :{}", list.size());
        Iterator var11 = list.iterator();

        while (var11.hasNext()) {
            Object obj = var11.next();
            currentRow = isCreate ? table.insertNewTableRow(index++) : table.getRow(index++);
            tempMap.put("t", obj);

            //如果有合并单元格情况，会导致params越界,这里需要补齐
            String[] paramsNew = (String[]) ArrayUtils.clone(params);
            if (params.length < currentRow.getTableCells().size()) {
                for (int i = 0; i < currentRow.getTableCells().size() - params.length; i++) {
                    paramsNew = (String[]) ArrayUtils.add(paramsNew, 0, "placeholderLc_" + i);
                }
            }

            String val;
            int cellIndex;
            for (cellIndex = 0; cellIndex < currentRow.getTableCells().size(); ++cellIndex) {
                val = PoiElUtil.eval(paramsNew[cellIndex], tempMap).toString();
                //源代码的bug 此方法无法删除单元格中的内容
                //currentRow.getTableCells().get(cellIndex).setText("");
                //使用此方法清空单元格内容
                if (!Strings.isNullOrEmpty(val)) {
                    currentRow.getTableCells().get(cellIndex).getParagraphs().forEach(p -> p.getRuns().forEach(r -> r.setText("", 0)));
                }
                PoiWordStyleUtil.copyCellAndSetValue(cellIndex >= tempCellList.size() ? tempCellList.get(tempCellList.size() - 1) : tempCellList.get(cellIndex)
                        , currentRow.getTableCells().get(cellIndex), val);
            }

            while (cellIndex < paramsNew.length) {
                val = PoiElUtil.eval(paramsNew[cellIndex], tempMap).toString();
                PoiWordStyleUtil.copyCellAndSetValue((XWPFTableCell) tempCellList.get(cellIndex), currentRow.createCell(), val);
                ++cellIndex;
            }
        }

        table.removeRow(index);
    }


    private static void clearParagraphText(List<XWPFParagraph> paragraphs) {
        paragraphs.forEach(pp -> {
            if (pp.getRuns() != null) {
                for (int i = pp.getRuns().size() - 1; i >= 0; i--) {
                    pp.removeRun(i);
                }
            }
        });
    }

}
