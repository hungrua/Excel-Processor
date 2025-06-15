package org.example;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.DataRow.*;

import java.io.*;

import java.util.*;
import java.util.stream.Collectors;

public class FileProcesserService {
    private static final Logger log = LogManager.getLogger();
    private static final String[] DEFAULT_COLUMNS_INDEX = {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ",
            "AK", "AL", "AM", "AN"
    };

    // Cần copy 1 file mới trong đó lấy nguyên sheet master data từ nguồn
    public void processing(String srcPath, String tartPath) throws IOException {
        // Mở file src để lấy dữ liệu yêu cầu của các sheet cần fill dữ liệu
        List<DataSheet> dataSheets = readSheetsAfterMasterData(srcPath);
        fillData(srcPath, tartPath, dataSheets);
        for (DataSheet dataSheet : dataSheets) {
            System.out.println(dataSheet.getRows().size());
        }
        // Mở file src để đọc dữ liệu từ sheet Master data

    }

    /**
     * @param filePath -- đường dẫn đến file nguồn
     * @return danh sách các sheet
     */
    public List<DataSheet> readSheetsAfterMasterData(String filePath) throws IOException {
        List<DataSheet> sheetList = new ArrayList<>();
        try (BufferedInputStream fis = new BufferedInputStream(new FileInputStream(filePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
//            Iterator<Sheet> sheetIterator = workbook.iterator();
//            boolean foundMasterData = false;
//
//            while (sheetIterator.hasNext()) {
//                Sheet sheet = sheetIterator.next();
//                String sheetName = sheet.getSheetName();
//
//                if (!foundMasterData) {
//                    if ("Master data".equalsIgnoreCase(sheetName.trim())) {
//                        foundMasterData = true;
//                    }
//                    continue;
//                }
//
//                // ✅ Lấy trực tiếp dòng 3 và 4
//                Row row3 = sheet.getRow(3);
//                Row row4 = sheet.getRow(4);
//                if (row3 == null || row4 == null) {
//                    System.out.println("⚠️ Không tìm thấy dòng 3 hoặc 4 trong sheet: " + sheetName);
//                    continue;
//                }
//                int maxCol = Math.max(row3.getLastCellNum(), row4.getLastCellNum());
//                List<DataColumn> dataColumnList = new ArrayList<>();
//                for (int colIndex = 1; colIndex < maxCol; colIndex++) { // bỏ cột đầu
//                    Cell cell3 = row3.getCell(colIndex);
//                    Cell cell4 = row4.getCell(colIndex);
//
//                    String header = getCellValue(cell3);
//                    String position = getCellValue(cell4);
//                    String hexColor = getCellFillColorHex(cell3);
//                    ExcelCellColorType colorType = ExcelCellColorType.fromHex(hexColor);
//
//                    dataColumnList.add(new DataColumn(position, header, colorType));
//                }
//                sheetList.add(new DataSheet(sheetName, dataColumnList));
//            }
            List<Sheet> sheetsToProcess = new ArrayList<>();
            boolean foundMasterData = false;

            // Lấy ra các sheet sau sheet có tên là Master data
            for (Sheet sheet : workbook) {
                String sheetName = sheet.getSheetName();
                if (!foundMasterData) {
                    if ("Master data".equalsIgnoreCase(sheetName.trim())) {
                        foundMasterData = true;
                    }
                    continue;
                }
                sheetsToProcess.add(sheet); // lưu lại
            }
            // Duyệt song song các sheet để điền thông tin các sheet
            sheetList = sheetsToProcess.parallelStream()
                    .map(sheet -> {
                        String sheetName = sheet.getSheetName();

                        Row row3 = sheet.getRow(3); // Dòng 3 chứa thông tin tên trường
                        Row row4 = sheet.getRow(4); // Dòng 4 là cột chứa dữ liệu trong master data
                        if (row3 == null || row4 == null) {
                            log.error("Không tìm thấy dòng 3 hoặc 4 trong sheet: " + sheetName);
                            return null;
                        }

                        int maxCol = Math.max(row3.getLastCellNum(), row4.getLastCellNum());
                        List<DataColumn> dataColumnList = new ArrayList<>();
                        // Lặp theo cột để lấy các thông tin về tên trường và cột lấy dữ liệu
                        for (int colIndex = 1; colIndex < maxCol; colIndex++) {
                            Cell cell3 = row3.getCell(colIndex);
                            Cell cell4 = row4.getCell(colIndex);
                            String header = getCellValue(cell3); // Tên trường
                            String position = getCellValue(cell4); // Vị trí cột lấy dữ liệu
                            String hexColor = getCellFillColorHex(cell3); // Giá trị màu biểu thị cho điều kiện
                            ExcelCellColorType colorType = ExcelCellColorType.fromHex(hexColor);
                            dataColumnList.add(new DataColumn(position, header, colorType));
                        }
                        return new DataSheet(sheetName, dataColumnList);
                    })
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());
            workbook.close();
        }
        return sheetList;
    }

    /**
     * @param srcPath    --- đường dẫn file nguồn
     * @param tartPath   --- đường dẫn file đích
     * @param dataSheets --- danh sách các sheet cần fill thông tin
     */

    public void fillData(String srcPath, String tartPath, List<DataSheet> dataSheets) throws IOException {

        // Mở file và đọc file
        try (FileInputStream fis = new FileInputStream(srcPath);
             Workbook workbook = StreamingReader.builder()
                     .rowCacheSize(100)
                     .bufferSize(4096)
                     .open(fis)) {

            // Đọc dữ liệu từ sheet Master data
            Sheet sourceSheet = workbook.getSheet("Master data");
            // Duyệt từng dòng trong master data
            for (Row row : sourceSheet) {
                if (row.getRowNum() < 3) continue; // chưa chứa dữ liệu

                if (isRowEmptyFast(row)) continue; // Bỏ qua dòng không có dữ liệu
                // 👉 Bỏ qua dòng trống
                log.debug("Bắt đầu xét row thứ :" + row.getRowNum());

                // Lấy dữ liệu cho từng sheet
                for (DataSheet dataSheet : dataSheets) {
                    log.debug("Bắt đầu xét sheet :" + dataSheet.getSheet_name());

                    // Khởi tạo 1 set để check trường hợp có điều kiện cần check trùng
                    Set<String> existValue = dataSheet.getColumns().stream()
                            .anyMatch(col -> col.getCondition() == ExcelCellColorType.NOT_DUPLICATED)
                            ? new HashSet<>()
                            : null;
                    DataRow dataRow = new DataRow();
                    boolean isValidRow = true; // biến đánh dấu xem row có hợp lệ cho sheet data không
                    // Duyệt dữ liệu trên từng Cell của row
                    for (Cell cell : row) {
                        int columnIndex = cell.getColumnIndex(); // Vị trí cell theo index
                        log.debug("Bắt đầu xét cell thứ :" + DEFAULT_COLUMNS_INDEX[columnIndex]);
                        // Nếu index mà nằm trong ds index mà sheet này có thì thực hiện lấy dữ liệu
                        if (dataSheet.getIndexColumns().contains(DEFAULT_COLUMNS_INDEX[columnIndex])) {
                            DataColumn dataColumn = dataSheet.getColumnIndexMap().get(DEFAULT_COLUMNS_INDEX[columnIndex]); // Lấy dữ liệu cột
                            String cellValue = getCellValue(cell); // Gía trị tại cell
                            log.debug("Xét cell {} của {} có giá trị {}  :", DEFAULT_COLUMNS_INDEX[columnIndex], dataSheet.getSheet_name(), cellValue);
                            // Nếu điều kiện của cell là REQUIRED nhưng giá trị tại cell là rỗng thì bỏ qua dữ liệu dòng này
                            if (cellValue.isEmpty() && ExcelCellColorType.REQUIRED.equals(dataColumn.getCondition())) {
                                isValidRow = false;
                                break;
                            }
                            // Nếu điều kiện của cell là NOT_DUPLICATED nhưng giá trị tại cell lại bị lặp thì bỏ qua dữ liệu dòng này
                            if (ExcelCellColorType.NOT_DUPLICATED.equals(dataColumn.getCondition()) &&
                                    !cellValue.isEmpty()
                                    && existValue.contains(cellValue)) {
                                isValidRow = false;
                                break;
                            }
                            if (existValue != null) existValue.add(cellValue);
                            dataRow.add(new DataCell(dataColumn, cellValue));
                        }
                    }
                    if (isValidRow) dataSheet.add(dataRow);
                }
            }
        }
        log.info("Đã đọc xong");
        try (SXSSFWorkbook workbook = new SXSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(tartPath)) {
            //Tạo style cho cell border
            CellStyle borderedStyle = workbook.createCellStyle();
            borderedStyle.setBorderTop(BorderStyle.THIN);
            borderedStyle.setBorderBottom(BorderStyle.THIN);
            borderedStyle.setBorderLeft(BorderStyle.THIN);
            borderedStyle.setBorderRight(BorderStyle.THIN);

            //Tạo style header kế thừa border + thêm màu + in đậm
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(borderedStyle); // 👈 Kế thừa border
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = workbook.createFont();
            font.setBold(true);
            headerStyle.setFont(font);
            //Ghi sheet
            for (DataSheet dataSheet : dataSheets) {
                Sheet sheet = workbook.createSheet(dataSheet.getSheet_name());
                int rowIndex = 0;

                // Header row
                Row headerRow = sheet.createRow(rowIndex++);
                for (int i = 0; i < dataSheet.getColumns().size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(dataSheet.getColumns().get(i).getColumn_name());
                    cell.setCellStyle(headerStyle);
                }

                // Dữ liệu
                for (DataRow rowData : dataSheet.getRows()) {
                    Row row = sheet.createRow(rowIndex++);
                    List<DataCell> values = rowData.getValues();
                    for (int i = 0; i < values.size(); i++) {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(values.get(i).getContent());
                        cell.setCellStyle(borderedStyle); //Áp dụng border cho data
                    }
                }
            }
            workbook.write(fos); // Ghi dữ liệu ra file
            workbook.dispose();  // Giải phóng các sheet tạm trong đĩa
        }
    }


    /**
     * @param cell -- ô excel
     * @return giá trị dạng chuỗi trong ô excel
     */
    private static String getCellValue(Cell cell) {
        String value = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING -> value = cell.getStringCellValue();
                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = cell.getDateCellValue().toString();
                    } else {
                        value = String.valueOf(cell.getNumericCellValue());
                    }
                }
                case BOOLEAN -> value = String.valueOf(cell.getBooleanCellValue());
                case FORMULA -> {
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING -> value = cell.getStringCellValue();
                        case NUMERIC -> value = String.valueOf(cell.getNumericCellValue());
                        case BOOLEAN -> value = String.valueOf(cell.getBooleanCellValue());
                    }
                }
            }
        }
        return value;
    }

    /**
     * @param cell -- ô excel
     * @return giá trị dạng chuỗi hexa của màu fill ô cần xét
     */
    private String getCellFillColorHex(Cell cell) {
        if (cell == null) return "";

        CellStyle style = cell.getCellStyle();
        if (!(style instanceof XSSFCellStyle)) return "";

        XSSFCellStyle xssfStyle = (XSSFCellStyle) style;
        XSSFColor color = xssfStyle.getFillForegroundColorColor();
        if (color == null) return "";

        byte[] rgb = color.getRGB();
        if (rgb == null) return "";
        return String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
    }


    /**
     * @param row -- dòng excel
     * @return dòng có trống hoàn toàn không
     */
    public boolean isRowEmptyFast(Row row) {
        if (row == null || row.getPhysicalNumberOfCells() == 0) {
            return true;
        }

        for (int cn = row.getFirstCellNum(); cn < row.getLastCellNum(); cn++) {
            Cell cell = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false; // Có ô không trắng
            }
        }
        return true;
    }
}
