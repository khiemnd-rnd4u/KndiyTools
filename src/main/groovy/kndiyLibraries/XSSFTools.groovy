package kndiyLibraries

import org.apache.poi.ss.usermodel.*

class XSSFTools {
    static final String BOLD_FONT = "Bold"
    static final String ITALIC_FONT = "Italic"
    static final String NORMAL_FONT = "Normal"
    static final String BOLD_ITALIC_FONT = "Bold Italic"
    static Row getOrSetRow(Sheet sheet, int rowNum) {
        Row row = sheet.getRow(rowNum)
        if (!row) {
            row = sheet.createRow(rowNum)
        }

        return row
    }

    static Cell getOrSetCell(Row row, int cellNum, CellStyle cellStyle = null, CellType cellType = CellType.STRING) {
        Cell cell = row.getCell(cellNum)
        if (!cell) {
            cell = row.createCell(cellNum, cellType)
            if (cellStyle) {
                cell.setCellStyle(cellStyle)
            }
        }

        return cell
    }

    static Font getFont(Workbook workbook,
                        String fontStyle = "Normal",
                        short colorIdx = IndexedColors.BLACK.getIndex(),
                        String fontName = "Calibri",
                        int size = 11) {
        Font font = workbook.createFont()
        font.setFontHeightInPoints((short) size)
        font.setFontName(fontName)
        font.setColor(colorIdx)
        switch (fontStyle) {
            case BOLD_FONT :
                font.setBold(true)
                break
            case NORMAL_FONT :
                font.setBold(false)
                font.setItalic(false)
                break
            case ITALIC_FONT :
                font.setItalic(true)
                break
            case BOLD_ITALIC_FONT :
                font.setBold(true)
                font.setItalic(true)
                break
        }

        return font
    }

    static CellStyle getCellStyleWithFormat(Workbook workbook,
                                            CreationHelper creationHelper,
                                            String formatPattern) {
        CellStyle cellStyle = workbook.createCellStyle()
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(formatPattern))

        return cellStyle
    }
}
