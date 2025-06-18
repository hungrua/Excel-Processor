package org.example.Data;

import java.util.ArrayList;
import java.util.List;

public class DataRow {
    private List<DataCell> values;

    public DataRow() {
        this.values = new ArrayList<>();
    }

    public void add(DataCell val) {
        values.add(val);
    }

    public List<DataCell> getValues() {
        return values;
    }
}
