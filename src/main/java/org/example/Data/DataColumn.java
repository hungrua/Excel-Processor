package org.example.Data;

public class DataColumn {
    private String column_position;
    private int column_index;
    private String column_name;
    private ExcelCellColorType condition;

    public DataColumn(String column_position, Integer index, String column_name, ExcelCellColorType condition) {
        this.column_position = column_position;
        this.column_index = index;
        this.column_name = column_name;
        this.condition = condition;
    }

    public String getColumn_name() {
        return column_name;
    }

    public String getColumn_position() {
        return column_position;
    }

    public int getColumn_index() {
        return column_index;
    }

    public ExcelCellColorType getCondition() {
        return condition;
    }

    public String toString() {
        return column_position + " " + column_name + " " + condition;
    }
}
