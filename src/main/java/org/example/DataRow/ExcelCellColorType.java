package org.example.DataRow;

public enum ExcelCellColorType {
    NOT_DUPLICATED("#FFA6A6"),
    REQUIRED("#FFFFA6"),
    SELECTION("#B4C7DC"),
    YES_OR_NO("#E0C2CD"),
    DATE_FORMAT("#D4EA6B");


    private final String hexColor;

    ExcelCellColorType(String hexColor) {
        this.hexColor = hexColor;
    }

    public String getHexColor() {
        return hexColor;
    }

    /**
     * @param hex --- Color code in hex
     * @return type of condition
     */
    public static ExcelCellColorType fromHex(String hex) {
        for (ExcelCellColorType type : values()) {
            if (type.hexColor.equalsIgnoreCase(hex)) {
                return type;
            }
        }
        return null;
    }
}
