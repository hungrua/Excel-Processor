package org.example.DataRow;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.*;
import java.util.stream.Collectors;

public class DataSheet {
    private String sheet_name; // Tên sheet
    private List<DataColumn> columns; // Danh sách các column của sheet
    private List<DataRow> rows; // Danh sách các row dữ liệu
    private Set<String> indexColumns; // Danh sách index theo chuỗi trong excel
    private Map<String, DataColumn> columnIndexMap; // Map để lấy tham chiếu dữ liệu lấy Data Columns bằng index theo chữ cái
    private static final Logger log = LogManager.getLogger();

    public Map<String, DataColumn> getColumnIndexMap() {
        if (columnIndexMap == null) {
            columnIndexMap = this.columns.stream()
                    .collect(Collectors.toMap(DataColumn::getColumn_position, dc -> dc));
        }
        return columnIndexMap;
    }

    public DataSheet(String sheet_name, List<DataColumn> columns) {
        this.sheet_name = sheet_name;
        this.columns = columns;
        this.rows = new ArrayList<>();
        this.indexColumns = getIndexColumnsSet(this.columns);
    }

    public String getSheet_name() {
        return sheet_name;
    }

    public void setSheet_name(String sheet_name) {
        this.sheet_name = sheet_name;
    }

    public List<DataColumn> getColumns() {
        return columns;
    }

    public void setColumns(List<DataColumn> columns) {
        this.columns = columns;
    }

    public List<DataRow> getRows() {
        return rows;
    }

    public Set<String> getIndexColumns() {
        return indexColumns;
    }

    private Set<String> getIndexColumnsSet(List<DataColumn> dataColumnList) {
        return new HashSet<>(dataColumnList.stream().map(DataColumn::getColumn_position).toList());
    }

    public void add(DataRow val) {
        rows.add(val);
    }

    @Override
    public String toString() {
        StringBuilder tmp = new StringBuilder(sheet_name + " ");
        columns.forEach(column -> tmp.append(column.toString() + " "));
        return tmp.toString();
    }
}
