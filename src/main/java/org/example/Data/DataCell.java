package org.example.Data;

public class DataCell {
    private DataColumn column;
    private String content;

    public DataCell(DataColumn column, String content) {
        this.column = column;
        this.content = content;
    }

    public DataColumn getColumn() {
        return column;
    }

    public String getContent() {
        return content;
    }
}
