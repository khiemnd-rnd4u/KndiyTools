package services.kndiyLibraries

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil

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
                        Short colorIdx = null,
                        String fontName = null,
                        int size = 11) {
        if (!colorIdx) {
            colorIdx = IndexedColors.BLACK.getIndex()
        }
        if (!fontName) {
            fontName = "Calibri Light"
        }

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
                                            String formatPattern,
                                            Font font = null) {
        CellStyle cellStyle = workbook.createCellStyle()
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(formatPattern))

        if (font) {
            cellStyle.setFont(font)
        }

        return cellStyle
    }

    static void addBorderToStyles(List<CellStyle> cellStyles) {
        cellStyles.each { CellStyle cellStyle ->
            addBorderToStyle(cellStyle)
        }
    }

    static CellStyle addBorderToStyle(CellStyle cellStyle) {
        cellStyle.setBorderTop(BorderStyle.THIN)
        cellStyle.setBorderBottom(BorderStyle.THIN)
        cellStyle.setBorderLeft(BorderStyle.THIN)
        cellStyle.setBorderRight(BorderStyle.THIN)

        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex())
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex())
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex())
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex())

        return cellStyle
    }

    static void setCenterAlignmentToStyles(CellStyle cellStyle) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER)
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER)
    }

    static void setRegionsBorder(Sheet sheet,
                                List<CellRangeAddress> regions) {
        regions.each { CellRangeAddress region ->
            setRegionBorder(sheet, region)
        }
    }

    static void setRegionBorder(Sheet sheet,
                                CellRangeAddress region) {
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet)
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet)
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet)
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet)
    }

    static Cell styleMergeCells(Sheet sheet, int fromRow, int toRow, int fromCell, int toCell, CellStyle cellStyle, value, Float heightInPoint = null) {
        (fromRow..toRow).each { int rowNum ->
            Row row = getOrSetRow(sheet, rowNum)
           setRowHeight(row, heightInPoint)

            (fromCell..toCell).each { int cellNum ->
                Cell cell = getOrSetCell(row, cellNum)
                cell.setCellStyle(cellStyle)
            }
        }

        sheet.addMergedRegion(new CellRangeAddress(fromRow, toRow, fromCell, toCell))
        return setCellValue(sheet, fromRow, fromCell, cellStyle, value)
    }

    static Cell setCellValue(Sheet sheet, int rowNum, int cellNum,
                             CellStyle cellStyle,
                             value,
                             Float heightInPoint = null) {
        Row row = getOrSetRow(sheet, rowNum)
        Cell cell = getOrSetCell(row, cellNum, cellStyle)
        setRowHeight(row, heightInPoint)
        cell.setCellValue(value)

        return cell
    }

    static Row setRowHeight(Row row, Float heightInPoint = null) {
        if (heightInPoint == 0) {
            row.setZeroHeight(true)
        }
        else if (heightInPoint != null) {
            row.setHeightInPoints(heightInPoint)
        }

        return row
    }
}
