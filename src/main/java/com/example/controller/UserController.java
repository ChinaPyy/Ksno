package com.example.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.springframework.stereotype.Controller;

import java.io.FileOutputStream;

/**
 * 创建人:楚克旺
 * 创建时间:2022/1/5 20:02
 **/
@Controller
public class UserController {

    @Test
    //生成并输入内容
    public void test01() throws Exception {
        //创建一个新的工作簿
        Workbook workbook = new HSSFWorkbook();
        //新建一个工作表
        Sheet sheet = workbook.createSheet("导出excel");
        //创建行
        Row row = sheet.createRow(0);
        //创建单元格（列和行交错）
        Cell cell = row.createCell(0);
        //给第一个单元格设置值(数据内容1)
        cell.setCellValue("张山");
        //创建第二个单元格
        Cell cell1 = row.createCell(1);
        //给第二个单元格设置值(数据内容1)
        cell1.setCellValue("男");
        //创建第二行的两列数据
        Row row1 = sheet.createRow(1);
        Cell cell2 = row1.createCell(0);
        //(数据内容1)
        cell2.setCellValue("李四");
        Cell cell3 = row1.createCell(1);
        //(数据内容1)
        cell3.setCellValue("女");
        //创建输出流 值: 要创建文件的路径
        FileOutputStream fs = new FileOutputStream("E:/temp/" + "03版排名.xls");
        //将相应的excel工作簿存盘
        workbook.write(fs);
        fs.close();
        System.out.println("文件生成成功");
    }





}
