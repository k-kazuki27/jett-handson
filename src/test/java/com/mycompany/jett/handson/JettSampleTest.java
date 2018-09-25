/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.jett.handson;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import static org.hamcrest.CoreMatchers.is;
import org.junit.After;
import org.junit.AfterClass;
import static org.junit.Assert.assertThat;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author kawabata
 */
public class JettSampleTest {

    JettSample target;

    public JettSampleTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
        target = new JettSample();
    }

    @After
    public void tearDown() {
    }

    @Test
    public void testExportExcel() {
        Workbook workbook = target.exportExcel();
        Sheet sheet = workbook.getSheet("テスト");
        Row row = sheet.getRow(0);
        assertThat(row.getCell(0).getStringCellValue(), is("Hello World!!"));
    }

}
