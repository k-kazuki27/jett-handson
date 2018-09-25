/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.jett.handson;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import net.sf.jett.transform.ExcelTransformer;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author kawabata
 */
public class JettSample {

    public static void main(String args[]) {

        JettSample jett = new JettSample();
        File actual = new File("src/test/resources/filestest/testExportExcel.xlsx");
        try {
            FileUtils.writeByteArrayToFile(actual, jett.toByte(jett.exportExcel()));
        } catch (IOException ex) {
            Logger.getLogger(JettSample.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public Workbook exportExcel() {
        String path = "/template/sample_template.xlsx";

        List<String> templateSheetNames = Lists.newArrayList("test");
        List<String> newSheetNames = Lists.newArrayList("テスト");

        Map<String, Object> beanMap = Maps.newHashMap();
        beanMap.put("name", "Hello World!!");
        List<Map<String, Object>> sheetBeans = Lists.newArrayList();
        sheetBeans.add(beanMap);
        return this.createWorkbook(templateSheetNames, newSheetNames, sheetBeans, path);
    }

    public byte[] toByte(Workbook workbook) {
        if (workbook == null) {
            return null;
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            workbook.write(out);
        } catch (IOException ex) {
            Logger.getLogger(JettSample.class.getName()).log(Level.SEVERE, null, ex);
        }
        return out.toByteArray();
    }

    private Workbook createWorkbook(List<String> templateSheetNames, List<String> newSheetNames,
            List<Map<String, Object>> sheetBeans, String templateUrl) {
        Workbook workbook = null;
        try (InputStream is = JettSample.class.getResourceAsStream(templateUrl)) {

            ExcelTransformer transformer = new ExcelTransformer();
            transformer.setForceRecalculationOnOpening(true);
            workbook = transformer.transform(is, templateSheetNames, newSheetNames, sheetBeans);
        } catch (InvalidFormatException | IOException ex) {
            Logger.getLogger(JettSample.class.getName()).log(Level.SEVERE, null, ex);
        }
        return workbook;
    }

}
