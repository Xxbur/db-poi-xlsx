package com.dbPoiXlsx;

import java.sql.*;
import java.text.ParseException;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;

public class PoiToXsl {
    public static void main(String[] args) throws Exception {
        GenerateTime generateTs = new GenerateTime();
        jdbcConfig db = new jdbcConfig();
        Connection conn = db.connectPostgre();
        Scanner scanner = new Scanner(System.in);
        System.out.println("开始时间 yyyy-MM-dd:");
        String start = scanner.next();
        Timestamp startTime= null;
        try {
            startTime = generateTs.generateTs(start);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        System.out.println("结束时间(不包含) yyyy-MM-dd:");
        String end = scanner.next();
        Timestamp endTime= generateTs.generateTs(end);
        scanner.close();
        String sql = "select brand_name,delivery_no,create_time,customer_name,address,receiver_phone from r_delivery_info_201801 where channel is null and create_time>=? and  create_time<? limit ? offset ?";
        String countSql = "select count(*) from r_delivery_info_201801 where channel is null and create_time>=? and  create_time<?";
        String sheetName = "测试Excel格式";
        String sheetTitle = "测试Excel格式";
        List<String> columnNames = new LinkedList<>();
        columnNames.add("brand_name");
        columnNames.add("delivery_no");
        columnNames.add("create_time");
        columnNames.add("customer_name");
        columnNames.add("address");
        columnNames.add("receiver_phone");

        int k = 1;

        PreparedStatement countStatement = conn.prepareStatement(countSql);
        countStatement.setTimestamp(1, startTime);
        countStatement.setTimestamp(2, endTime);
        ResultSet countRs = countStatement.executeQuery();
        countRs.next();
        long count = (long) countRs.getObject(1);
        int loop = (int) ((count - 1) / 1000000 + 1);
        System.out.println(loop);


        PreparedStatement statement = conn.prepareStatement(sql);
        ResultSet rs;
        for (int j = 7; j < loop; j++) {
            statement.setTimestamp(1, startTime);
            statement.setTimestamp(2, endTime);
            statement.setInt(3, 1000000);
            statement.setInt(4, 1000000*j);
            rs = statement.executeQuery();
            ResultSetMetaData metaData = rs.getMetaData();
            int columnCount = metaData.getColumnCount();
            List<List<Object>> list = new LinkedList<>();
            while(rs.next()) {
                List<Object> rowData = new LinkedList<>();
                for (int i = 1; i <= columnCount; i++) {
                    rowData.add(rs.getObject(i));
                }
                list.add(rowData);
            }
            System.out.println(list.size());

            poiIntoXls poiIntoXls = new poiIntoXls();
            poiIntoXls.writeExcelTitle("E:\\temp", "1.9-1.10-"+k, sheetName, columnNames, sheetTitle);

            try {
                poiIntoXls.writeExcelData("E:\\temp", "1.9-1.10-"+k, sheetName, list);
            } catch (Exception e) {
                e.printStackTrace();
            }

            poiIntoXls.dispose();
            k++;
            list.clear();
        }
        conn.close();
    }
}
