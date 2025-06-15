package org.example.DataRow;

public class DataColumn {
    private String column_position;
    private String column_name;
    private ExcelCellColorType condition;

    public DataColumn(String column_position, String column_name, ExcelCellColorType condition) {
        this.column_position = column_position;
        this.column_name = column_name;
        this.condition = condition;
    }

    public String getColumn_name() {
        return column_name;
    }

    public String getColumn_position() {
        return column_position;
    }

    public ExcelCellColorType getCondition() {
        return condition;
    }

    public String toString() {
        return column_position + " " + column_name + " " + condition;
    }
}
