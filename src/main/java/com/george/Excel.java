package com.george;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import static javafx.scene.input.KeyCode.T;

/**
 * Created by Yuan Wei Cheng on 2016/5/19.
 * This is for reading excel content to a POJO .
 * The excel should satisfy below two points
 * 1, the file should be 2003 version and end with xls
 * 2, the first row of the first sheet should be the column name of database and it should the same as the database
 */
public class Excel<T> {

    private Logger logger  = LoggerFactory.getLogger(Excel.class);

    public List<T> readExcel(String filePath, Class<T>  t) {
        List<T> entityList = new ArrayList<T>();
        logger.info("读文件内容 -------开始---------");
        try {
            FileInputStream fis = new FileInputStream(filePath);
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            for(int i=1;i<sheet.getLastRowNum();i++){
                T entity = t.newInstance();
                for(int j = 0;j<sheet.getRow(0).getLastCellNum();j++){
                    try {
                        if(sheet.getRow(i).getCell(j)!=null && !"".equals(sheet.getRow(i).getCell(j).toString().trim())) {
                            //get the column name of excel
                            String columnName = sheet.getRow(0).getCell(j).toString().trim();
                            //get the entity field name according to column name
                            Field field = entity.getClass().getDeclaredField(columnName);
                            //make the column of entity accessible
                            field.setAccessible(true);
                            //get the column value of column name
                            Object columnValue = sheet.getRow(i).getCell(j).toString().trim();
                            //set the column value to the entity
                            field.set(entity,columnValue);
                        }else{
                            //set the value as "" if the value is null or "" in excel
                            //get the column name of excel
                            String columnName = sheet.getRow(0).getCell(j).toString().trim();
                            //get the entity field name according to column name
                            Field field = entity.getClass().getDeclaredField(columnName);
                            //make the column of entity accessible
                            field.setAccessible(true);
                            //set the column value to the entity
                            field.set(entity,"");
                        }
                    } catch (NoSuchFieldException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
                entityList.add(entity);
            }
            fis.close();
            logger.info("总共行数 :"+entityList.size());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        logger.info("读文件内容 -------结束---------");
        return entityList;
    }


}
