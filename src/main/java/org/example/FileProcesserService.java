package org.example;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.Data.*;

import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

public class FileProcesserService {
    private static final Logger log = LogManager.getLogger(FileProcesserService.class);
    private static final DecimalFormat decimalFormat = new DecimalFormat("#.###############");
    private static final String[] DEFAULT_COLUMNS_INDEX = {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ",
            "AK", "AL", "AM", "AN"
    };

    public void processing(String srcPath, String tartPath, String configPath) throws IOException {
        // Mở file src để lấy dữ liệu yêu cầu của các sheet cần fill dữ liệu
        List<DataSheet> dataSheets = readSheetsAfterMasterData(configPath);
        fillData(srcPath, tartPath, dataSheets);
    }

    /**
     * @param fileConfigPath -- đường dẫn đến file cấu hình các sheet nguồn
     * @return danh sách các sheet
     */
    public List<DataSheet> readSheetsAfterMasterData(String fileConfigPath) throws IOException {
        List<DataSheet> sheetList;
        try (InputStream fis = new BufferedInputStream(new FileInputStream(fileConfigPath), 32768);
             Workbook workbook = StreamingReader.builder()
                     .rowCacheSize(100)
                     .bufferSize(8192)
                     .open(fis)) {
            List<Sheet> sheetsToProcess = new ArrayList<>();

            for (Sheet sheet : workbook) {
                sheetsToProcess.add(sheet);
            }

            // Xử lý song song các sheet
            sheetList = sheetsToProcess.parallelStream()
                    .map(sheet -> {
                        String sheetName = sheet.getSheetName();
                        Row row3 = null;
                        Row row4 = null;

                        // Duyệt tuần tự để lấy dòng 3 và 4
                        int rowIndex = 0;
                        for (Row row : sheet) {
                            if (row.getRowNum() == 3) {
                                row3 = row;
                            } else if (row.getRowNum() == 4) {
                                row4 = row;
                            }
                            if (row3 != null && row4 != null) {
                                break; // Thoát khi đã lấy đủ dòng 3 và 4
                            }
                            rowIndex++;
                            if (rowIndex > 4) {
                                break; // Tránh duyệt quá nhiều dòng
                            }
                        }

                        if (row3 == null || row4 == null) {
                            log.error("Không tìm thấy dòng 3 hoặc 4 trong sheet: " + sheetName);
                            return null;
                        }

                        //Lấy ra số cột lớn nhất giữa dòng 3 và 4
                        int maxCol = Math.max(row3.getLastCellNum(), row4.getLastCellNum());
                        List<DataColumn> dataColumnList = new ArrayList<>();

                        for (int colIndex = 1; colIndex < maxCol; colIndex++) {
                            Cell cell3 = row3.getCell(colIndex);
                            Cell cell4 = row4.getCell(colIndex);
                            // Lấy ra thông tin hader và position từ dòng 3 và 4
                            String header = getCellValue(cell3).trim();
                            String position = getCellValue(cell4).trim();
                            log.debug("Trong sheet {}, cột {} có header: '{}' và position: '{}'", sheetName, colIndex, header, position);
                            // Lấy ra mã màu của các trường để xác định điều kiện ràng buộc
                            String hexColor = getCellFillColorHex(cell3);
                            ExcelCellColorType colorType = ExcelCellColorType.fromHex(hexColor);
                            dataColumnList.add(new DataColumn(position, colIndex, header, colorType));
                        }
                        // Khởi tạo thông tin sheet với danh sách các trường
                        return new DataSheet(sheetName, dataColumnList);
                    })
                    .filter(Objects::nonNull)
                    .collect(Collectors.toList());
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
        try (InputStream fis = new BufferedInputStream(new FileInputStream(srcPath), 32768);
             Workbook workbook = StreamingReader.builder()
                     .rowCacheSize(100)
                     .bufferSize(8192)
                     .open(fis)) {
            // Đọc dữ liệu từ sheet Master data
            Sheet sourceSheet = workbook.getSheet("Master data");
            // Duyệt từng dòng trong master data
            // Khởi tạo Map để lưu trữ HashSet cho từng DataSheet
            Map<String, HashMap<String, HashSet<String>>> sheetExistValues = new HashMap<>();
            for (DataSheet dataSheet : dataSheets) {
                String sheetName = dataSheet.getSheet_name();
                // Khởi tạo HashMap cho từng sheet
                HashMap<String, HashSet<String>> existValue = new HashMap<>();
                // Tạo HashSet cho từng cột có điều kiện NOT_DUPLICATED
                dataSheet.getColumns().stream()
                        .filter(col -> col.getCondition() == ExcelCellColorType.NOT_DUPLICATED)
                        .forEach(col -> {
                            String position = col.getColumn_position();
                            existValue.put(position, new HashSet<>());
                            log.debug("Khởi tạo HashSet cho cột: {} trong sheet {}", position, sheetName);
                        });
                sheetExistValues.put(sheetName, existValue);
                log.debug("HashMap existValue cho sheet {} sau khi khởi tạo: {}", sheetName, existValue.keySet());
            }

            // Duyệt qua từng dòng trong sourceSheet
            for (Row row : sourceSheet) {
                if (row.getRowNum() < 3) {
                    log.debug("Bỏ qua dòng {} vì chưa chứa dữ liệu", row.getRowNum() + 1);
                    continue;
                }

                if (isRowEmptyFast(row)) {
                    log.info("Dòng {} đã hết dữ liệu", row.getRowNum() + 1);
                    break; // Dòng không có dữ liệu, thoát vòng lặp
                }

                log.debug("Bắt đầu xét row thứ: {}", row.getRowNum() + 1);

                // Lấy dữ liệu cho từng sheet
                for (DataSheet dataSheet : dataSheets) {
                    String sheetName = dataSheet.getSheet_name();
                    log.debug("Bắt đầu xét sheet: {}", sheetName);

                    // Lấy HashMap cho sheet hiện tại
                    HashMap<String, HashSet<String>> existValue = sheetExistValues.get(sheetName);

                    DataRow dataRow = new DataRow();
                    boolean isContainAccepedValue = false; // Cờ đánh dấu dòng có chứa giá trị hợp lệ cho sheet
                    boolean isValidRow = true; // Cờ đánh dấu dòng có hợp lệ cho sheet không

                    // Duyệt qua từng ô trong dòng
                    for (Cell cell : row) {
                        if (cell == null) {
                            log.debug("Bỏ qua ô null tại dòng {}", row.getRowNum() + 1);
                            continue;
                        }
                        int columnIndex = cell.getColumnIndex(); // Lấy chỉ số ô
                        log.debug("Xử lý ô tại dòng {}, cột index: {}", row.getRowNum() + 1, columnIndex);
                        // Kiểm tra nếu chỉ số cột nằm trong danh sách chỉ số cột của sheet
                        String columnKey = DEFAULT_COLUMNS_INDEX[columnIndex];
                        if (dataSheet.getIndexColumns().contains(columnKey)) {
                            DataColumn dataColumn = dataSheet.getColumnIndexMap().get(columnKey);
                            if (dataColumn == null) {
                                log.error("Không tìm thấy cột tại chỉ số {} trong sheet {}", columnKey, sheetName);
                                isValidRow = false;
                                break;
                            }
                            String cellValue = getCellValue(cell); // Lấy giá trị ô
                            log.debug("Giá trị ô tại cột {}: '{}'", dataColumn.getColumn_name(), cellValue);
                            // Xử lý trường hợp cellValue là null

                            // Kiểm tra nếu ô REQUIRED nhưng giá trị rỗng thực hiện điền "null"
                            if (ExcelCellColorType.REQUIRED.equals(dataColumn.getCondition())) {
                                log.debug("Điền dữ liệu null tại dòng {} vào sheet {} do cột {} có giá trị rỗng ",
                                        row.getRowNum() + 1, sheetName, dataColumn.getColumn_name());
                            }
                            if (!cellValue.equals("null")) isContainAccepedValue = true;
                            // Kiểm tra trùng lặp cho cột NOT_DUPLICATED
                            if (ExcelCellColorType.NOT_DUPLICATED.equals(dataColumn.getCondition())) {
                                String columnPosition = dataColumn.getColumn_position();
                                HashSet<String> valueSet = existValue.get(columnPosition);

                                // Kiểm tra xem HashSet có tồn tại không
                                if (valueSet == null) {
                                    log.error("HashSet cho cột {} không tồn tại trong existValue của sheet {}, khởi tạo lại",
                                            columnPosition, sheetName);
                                    valueSet = new HashSet<>();
                                    existValue.put(columnPosition, valueSet);
                                }

                                // Ghi log debug trước khi kiểm tra trùng
                                log.debug("Check trùng cho giá trị: '{}' tại cột {} trong sheet {}", cellValue, columnPosition, sheetName);
                                log.debug("Cột {} trong sheet {} hiện có {} giá trị: {}", columnPosition, sheetName, valueSet.size(), valueSet);

                                // Kiểm tra giá trị trùng
                                if (!"null".equals(cellValue) && valueSet.contains(cellValue)) {
                                    log.error("Dữ liệu tại dòng {} không thể điền vào sheet {} do cột {} có giá trị trùng: '{}'",
                                            row.getRowNum() + 1, sheetName, dataColumn.getColumn_name(), cellValue);
                                    isValidRow = false;
                                    break;
                                }

                                // Thêm giá trị vào HashSet và kiểm tra kết quả
                                boolean added = valueSet.add(cellValue);
                                log.debug("Thêm giá trị '{}' vào HashSet của cột {} trong sheet {}: {}",
                                        cellValue, columnPosition, sheetName, added ? "Thành công" : "Thất bại (đã tồn tại)");
                                log.debug("HashSet sau khi thêm: {}", valueSet);
                            }

                            // Kiểm tra định dạng ngày cho cột DATE_FORMAT
                            if (ExcelCellColorType.DATE_FORMAT.equals(dataColumn.getCondition())) {
                                if (!cellValue.isEmpty() && !"null".equals(cellValue)) {
                                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                                    sdf.setLenient(false); // Không cho phép parse ngày không hợp lệ
                                    try {
                                        sdf.parse(cellValue);
                                        log.debug("Định dạng ngày '{}' hợp lệ cho cột {} trong sheet {}",
                                                cellValue, dataColumn.getColumn_name(), sheetName);
                                    } catch (ParseException e) {
                                        log.error("Dữ liệu tại dòng {} không thể điền vào sheet {} do cột {} không đúng định dạng yyyy-MM-dd: '{}'",
                                                row.getRowNum() + 1, sheetName, dataColumn.getColumn_name(), cellValue);
                                        isValidRow = false;
                                        break;
                                    }
                                } else {
                                    log.error("Dữ liệu tại dòng {} không thể điền vào sheet {} do cột {} có giá trị rỗng (yêu cầu định dạng ngày)",
                                            row.getRowNum() + 1, sheetName, dataColumn.getColumn_name());
                                    isValidRow = false;
                                    break;
                                }
                            }

                            // Thêm ô vào DataRow
                            dataRow.add(new DataCell(dataColumn, cellValue));
                            log.debug("Đã thêm DataCell cho cột {} với giá trị '{}' trong sheet {}",
                                    dataColumn.getColumn_name(), cellValue, sheetName);
                        }
                    }

                    // Nếu dòng hợp lệ, thêm vào sheet
                    if (isValidRow && isContainAccepedValue) {
                        dataSheet.add(dataRow);
                        log.debug("Đã thêm DataRow vào sheet {} tại dòng {}", sheetName, row.getRowNum() + 1);
                    } else {
                        log.debug("Dòng {} không hợp lệ, không thêm vào sheet {}", row.getRowNum() + 1, sheetName);
                    }
                }
            }
        } catch (FileNotFoundException e) {
            log.error("Lỗi không tìm thấy file nguồn");
            throw e;
        } catch (IOException exception) {
            log.error("Lỗi đọc file nguồn");
            throw exception;
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
            headerStyle.cloneStyleFrom(borderedStyle); // Kế thừa border
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = workbook.createFont();
            font.setBold(true);
            headerStyle.setFont(font);
            //Ghi sheet
            for (DataSheet dataSheet : dataSheets) {
                Sheet sheet = workbook.createSheet(dataSheet.getSheet_name());
                int rowIndex = 0;
                List<DataColumn> sheetHeader = dataSheet.getColumns();

                // Header ghi header row
                Row headerRow = sheet.createRow(rowIndex++);
                for (int i = 0; i < dataSheet.getColumns().size(); i++) {
                    Cell cell = headerRow.createCell(i);
                    log.debug("Đang tạo header cho cột: {} của sheet {} {}", dataSheet.getColumns().get(i).getColumn_name(), dataSheet.getSheet_name(), dataSheet.getColumns().get(i).getColumn_index());
                    cell.setCellValue(dataSheet.getColumns().get(i).getColumn_name());
                    cell.setCellStyle(headerStyle);
                }

                // Ghi dữ liệu vào sheet
                for (DataRow rowData : dataSheet.getRows()) {
                    Row row = sheet.createRow(rowIndex++);
                    // Lấy ra danh sách các ô trong DataRow
                    List<DataCell> values = rowData.getValues();

                    // Tạo Map để truy cập nhanh theo column_index
                    Map<Integer, DataCell> cellMap = values.stream()
                            .collect(Collectors.toMap(
                                    (DataCell dc) -> dc.getColumn().getColumn_index(),
                                    Function.identity()
                            ));
                    // Duyệt đúng thứ tự hiển thị từ cấu hình
                    for (int i = 0; i < dataSheet.getColumns().size(); i++) {
                        DataColumn column = dataSheet.getColumns().get(i);
                        Cell cell = row.createCell(i);
                        // Thực hiện ánh xạ theo column_index từ datacell tới data column
                        DataCell dataCell = cellMap.get(column.getColumn_index());
                        // Nếu có ánh xạ thành công, ghi giá trị vào ô
                        if (dataCell != null && dataCell.getContent() != null) {
                            cell.setCellValue(dataCell.getContent());
                        } else { // Trường hợp không có dữ liệu (không có ánh xạ tới cột nào trong master data)
                            cell.setCellValue(""); // hoặc "null" nếu bạn muốn
                        }
                        cell.setCellStyle(borderedStyle);
                    }
                }

            }
            workbook.write(fos); // Ghi dữ liệu ra file
            workbook.dispose();  // Giải phóng các sheet tạm trong đĩa
        } catch (FileNotFoundException e) {
            log.error("Lỗi không tìm thấy file đích");
            throw e;
        } catch (IOException exception) {
            log.error("Lỗi đọc file đích");
            throw exception;
        }
    }

    /**
     * @param cell -- ô excel
     * @return giá trị dạng chuỗi trong ô excel
     */
    private static String getCellValue(Cell cell) {
        String value = "null";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING -> value = cell.getStringCellValue();
                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = cell.getDateCellValue().toString();
                    } else {
                        double numericValue = cell.getNumericCellValue();
                        value = (numericValue == Math.floor(numericValue))
                                ? String.valueOf((long) numericValue)
                                : String.valueOf(numericValue);
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

            // ✅ Kiểm tra sau cùng
            if (value == null || value.isBlank() || value.trim().equals("''")) {
                return "null";
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
