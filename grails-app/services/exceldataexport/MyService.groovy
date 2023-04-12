package exceldataexport

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.springframework.jdbc.core.JdbcTemplate


class MyService {

    JdbcTemplate jdbcTemplate

    void generateExcelFile(String sql, String fileName) {
//        List<Map<String, Object>> rows = jdbcTemplate.queryForList(sql)
        List<Map<String, Object>> rows = jdbcTemplate.query(sql)


        Workbook workbook = new XSSFWorkbook()
        Sheet sheet = workbook.createSheet(tableName)
        int rowNum = 0
        Row headerRow = sheet.createRow(rowNum++)
        int colNum = 0
        for (String header : rows[0].keySet()) {
            Cell headerCell = headerRow.createCell(colNum++)
            headerCell.setCellValue(header)
        }
        for (Map<String, Object> row : rows) {
            Row dataRow = sheet.createRow(rowNum++)
            colNum = 0
            for (Object value : row.values()) {
                Cell dataCell = dataRow.createCell(colNum++)
                if (value != null) {
                    if (value instanceof Number) {
                        dataCell.setCellValue((double) value)
                    } else {
                        dataCell.setCellValue(value.toString())
                    }
                }
            }
        }
        FileOutputStream outputStream = new FileOutputStream(fileName)
        workbook.write(outputStream)
        workbook.close()
        outputStream.close()
    }
}
