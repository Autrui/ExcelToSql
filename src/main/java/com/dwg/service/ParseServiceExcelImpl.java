package com.dwg.service;

import cn.hutool.json.JSONObject;
import com.dwg.entity.FireFacilityConfig;
import com.dwg.entity.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Autrui
 * @date 2023/5/23
 * @apiNote
 */
public class ParseServiceExcelImpl {
    public static void parseExcel(String filepath) throws IOException {

        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        List<FireFacilityConfig> list = null;
        String cellData = null;
        //读取Excel文件
        wb = readExcel(filepath);
        //如果文件不为空
        if (wb != null) {
            //用来存放表中数据
            list = new ArrayList<FireFacilityConfig>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            //循环行
            for (int i = 1; i < rownum; i++) {
                FireFacilityConfig camera = new FireFacilityConfig();
                row = sheet.getRow(i);
                if (row != null) {
                    //循环列
                    List<JSONObject> jsList = new ArrayList<JSONObject>();
                    for (int j = 0; j < colnum; j++) {
                        cellData = (String) getCellFormatValue(row.getCell(j));
                        JSONObject json = new JSONObject();

                        switch (j) {
                            case 0:// name
                                camera.setName(cellData);
                                break;
                            case 1:// code
                                camera.setCode(cellData);
                                break;
                            case 2:// 外观
                                if (StringUtils.isBlank(cellData)) {
                                    continue;
                                }
                                json = new JSONObject();
                                json.putOnce("type", 3);
                                getJson(cellData, json);
                                jsList.add(json);
                                break;
                            case 3:// 功能
                                if (StringUtils.isBlank(cellData)) {
                                    continue;
                                }
                                json = new JSONObject();
                                json.putOnce("type", 1);
                                getJson(cellData, json);
                                jsList.add(json);
                                break;
                            case 4:// 保养
                                if (StringUtils.isBlank(cellData)) {
                                    continue;
                                }
                                json = new JSONObject();
                                json.putOnce("type", 2);
                                getJson(cellData, json);
                                jsList.add(json);
                                break;
                            case 5:// 联动
                                if (StringUtils.isBlank(cellData)) {
                                    continue;
                                }
                                json = new JSONObject();
                                json.putOnce("type", 4);
                                getJson(cellData, json);
                                jsList.add(json);
                                break;
                            default:
                                break;
                        }
                    }
                    camera.setMaintenance_config(jsList);
                    //放入集合
                    list.add(camera);
                } else {
                    break;
                }
            }
        }
        //定义一个文件，用来存数据；
        System.out.println("number of camera: " + list.size());
        String fileName = filepath.substring(0, filepath.length() - 4) + "_update_camera" + ".sql";
        PrintWriter ps = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fileName), "UTF-8")));
        if (list.size() > 0) {
            //遍历解析出来的list
            for (FireFacilityConfig camera : list) {

                String name = camera.getName() == null ? "null" : "'" + camera.getName() + "'";
                String code = camera.getCode() == null ? "null" : "'" + camera.getCode() + "'";
                String type = camera.getType() == null ? "null" : "'" + camera.getType() + "'";
                String delFlag = camera.getDel_flag() == null ? "null" : "'" + camera.getDel_flag() + "'";
                String maintenance_config = camera.getMaintenance_config() == null ? "null" : "'" + camera.getMaintenance_config() + "'";

                String strSQL = String.format("UPDATE t_fire_facility_config set maintenance_config= %s WHERE code=%s and type = 1 and del_flag=0;", maintenance_config, code);
                ps.println(strSQL);
            }
        }
        ps.close();
    }

    private static void getJson(String cellData, JSONObject json) {
        if (cellData.equals("月")) {
            json.putOnce("preiod", 1);
        } else if (cellData.equals("季度")) {
            json.putOnce("preiod", 2);
        } else if (cellData.equals("半年")) {
            json.putOnce("preiod", 3);
        } else if (cellData.equals("年")) {
            json.putOnce("preiod", 4);
        }
    }

    //读取excel
    @SuppressWarnings("resource")
    public static Workbook readExcel(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        //文件后缀名
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            //如果文件后缀名为xls
            if (".xls".equals(extString)) {
                return wb = new HSSFWorkbook(is);
            }//如果文件后缀名为xlsx
            else if (".xlsx".equals(extString)) {
                return wb = new XSSFWorkbook(is);
            } else {
                return wb = null;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    @SuppressWarnings("deprecation")
    public static Object getCellFormatValue(Cell cell) {
        Object cellValue = null;
        if (cell != null) {
            //判断cell类型
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA: {
                    //判断cell是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    } else {
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }
}
