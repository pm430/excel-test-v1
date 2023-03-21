package org.example;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {
    public static void main(String[] args) throws Exception {
        // 새 엑셀 워크북 생성
        XSSFWorkbook workbook = new XSSFWorkbook();

        // 첫 번째 시트 생성
        Sheet sheet = workbook.createSheet("Sheet1");

        // 첫 번째 열 첫 다섯 행에 1 채우기
        for (int i = 0; i < 5; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(1);
        }

        // 두 번째 열 첫 행부터 열 줄에 알파벳 채우기
        for (int i = 0; i < 26; i++) {
            Row row = sheet.getRow(i % 26);
            if (row == null) {
                row = sheet.createRow(i % 26);
            }
            Cell cell = row.createCell(1);
            cell.setCellValue((char) ('a' + i));
        }

        for (int i = 0; i < 10; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            Cell cell = row.createCell(2);
            cell.setCellValue(33);
        }

        // 파일에 저장
        FileOutputStream fileOut = new FileOutputStream("example.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // 메모리에서 제거
        workbook.close();
    }
}
