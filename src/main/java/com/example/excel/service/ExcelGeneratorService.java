package com.example.excel.service;

import com.example.excel.util.ExcelUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

public class ExcelGeneratorService {

    public static void generate(String fileName, Map<String, List<String>> data) throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();

        Sheet dataSheet = workbook.createSheet("Data");
        Sheet listSheet = workbook.createSheet("Lists");

        /* ---------------- Header Row ---------------- */
        Row header = dataSheet.createRow(0);

        int colIndex = 0;
        for (String headerName : data.keySet()) {
            header.createCell(colIndex++).setCellValue(headerName);
        }

        /* ---------------- Create Lists Sheet ---------------- */
        int listCol = 0;

        for (Map.Entry<String, List<String>> entry : data.entrySet()) {

            List<String> items = entry.getValue();

            // Skip empty lists (no dropdown needed)
            if (items == null || items.isEmpty()) {
                listCol++;
                continue;
            }

            String safeName = ExcelUtil.safeName(entry.getKey());

            listSheet.createRow(0).createCell(listCol).setCellValue(safeName);

            for (int i = 0; i < items.size(); i++) {
                Row row = listSheet.getRow(i + 1);
                if (row == null) row = listSheet.createRow(i + 1);
                row.createCell(listCol).setCellValue(items.get(i));
            }

            Name name = workbook.createName();
            name.setNameName(safeName);
            name.setRefersToFormula(
                    "Lists!$" + ExcelUtil.col(listCol) + "$2:$" +
                    ExcelUtil.col(listCol) + "$" + (items.size() + 1)
            );

            /* ---------------- Apply Dropdown ---------------- */
            DataValidationHelper helper = dataSheet.getDataValidationHelper();
            DataValidationConstraint constraint =
                    helper.createFormulaListConstraint(safeName);

            DataValidation validation = helper.createValidation(
                    constraint,
                    new CellRangeAddressList(1, 500, listCol, listCol)
            );
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);

            dataSheet.addValidationData(validation);

            listCol++;
        }

        workbook.setSheetHidden(workbook.getSheetIndex(listSheet), true);

        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        }

        workbook.close();
    }
}
