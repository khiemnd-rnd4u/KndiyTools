package services.sensorService

import entities.sensor.Inspection
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import services.kndiyLibraries.XSSFTools

abstract class KnDiyWorkbook implements KnDiyWorkbookInterface {
    protected XSSFWorkbook workbook
    protected CreationHelper creationHelper
    protected Font BOLD_FONT
    protected Font NORMAL_FONT
    protected Font UNDER_LINE_FONT
    protected Font ITALIC_FONT
    protected Font BOLD_ITALIC_FONT
    protected Font TITLE_FONT
    protected Font TITLE_ITALIC_FONT
    protected Font HEADER_FONT
    protected Font HEADER_ITALIC_FONT

    protected Font RED_BOLD_FONT
    protected Font RED_NORMAL_FONT
    protected Font RED_ITALIC_FONT
    protected Font RED_BOLD_ITALIC_FONT
    protected Font RED_TITLE_FONT
    protected Font RED_TITLE_ITALIC_FONT
    protected Font RED_HEADER_FONT
    protected Font RED_HEADER_ITALIC_FONT

    protected CellStyle WRAP_NUMBER_BOLD_ITALIC_STYLE
    protected CellStyle WRAP_NUMBER_ITALIC_STYLE
    protected CellStyle WRAP_NUMBER_BOLD_STYLE
    protected CellStyle WRAP_NUMBER_STYLE
    protected CellStyle WRAP_STYLE
    protected CellStyle WRAP_BOLD_STYLE
    protected CellStyle WRAP_ITALIC_STYLE
    protected CellStyle WRAP_BOLD_ITALIC_STYLE
    protected CellStyle TITLE_STYLE
    protected CellStyle TITLE_ITALIC_STYLE
    protected CellStyle HEADER_STYLE
    protected CellStyle HEADER_ITALIC_STYLE
    protected CellStyle HEADER_UNDERLINE_STYLE

    protected CellStyle RED_FILL_STYLE
    protected CellStyle RED_FILL_BOLD_STYLE
    protected CellStyle RED_FILL_ITALIC_STYLE
    protected CellStyle RED_FILL_BOLD_ITALIC_STYLE

    protected CellStyle RED_FILL_NUMBER_STYLE
    protected CellStyle RED_FILL_NUMBER_BOLD_STYLE
    protected CellStyle RED_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle RED_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected CellStyle ROSE_FILL_STYLE
    protected CellStyle ROSE_FILL_BOLD_STYLE
    protected CellStyle ROSE_FILL_ITALIC_STYLE
    protected CellStyle ROSE_FILL_BOLD_ITALIC_STYLE

    protected CellStyle ROSE_FILL_NUMBER_STYLE
    protected CellStyle ROSE_FILL_NUMBER_BOLD_STYLE
    protected CellStyle ROSE_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle ROSE_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected CellStyle GREEN_FILL_STYLE
    protected CellStyle GREEN_FILL_BOLD_STYLE
    protected CellStyle GREEN_FILL_ITALIC_STYLE
    protected CellStyle GREEN_FILL_BOLD_ITALIC_STYLE

    protected CellStyle GREEN_FILL_NUMBER_STYLE
    protected CellStyle GREEN_FILL_NUMBER_BOLD_STYLE
    protected CellStyle GREEN_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle GREEN_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected CellStyle GREY_FILL_STYLE
    protected CellStyle GREY_FILL_BOLD_STYLE
    protected CellStyle GREY_FILL_ITALIC_STYLE
    protected CellStyle GREY_FILL_BOLD_ITALIC_STYLE

    protected CellStyle GREY_FILL_NUMBER_STYLE
    protected CellStyle GREY_FILL_NUMBER_BOLD_STYLE
    protected CellStyle GREY_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle GREY_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected CellStyle ORANGE_FILL_STYLE
    protected CellStyle ORANGE_FILL_BOLD_STYLE
    protected CellStyle ORANGE_FILL_ITALIC_STYLE
    protected CellStyle ORANGE_FILL_BOLD_ITALIC_STYLE

    protected CellStyle ORANGE_FILL_NUMBER_STYLE
    protected CellStyle ORANGE_FILL_NUMBER_BOLD_STYLE
    protected CellStyle ORANGE_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle ORANGE_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected CellStyle YELLOW_FILL_STYLE
    protected CellStyle YELLOW_FILL_BOLD_STYLE
    protected CellStyle YELLOW_FILL_ITALIC_STYLE
    protected CellStyle YELLOW_FILL_BOLD_ITALIC_STYLE

    protected CellStyle BLUE_FILL_STYLE
    protected CellStyle BLUE_FILL_BOLD_STYLE
    protected CellStyle BLUE_FILL_ITALIC_STYLE
    protected CellStyle BLUE_FILL_BOLD_ITALIC_STYLE

    protected CellStyle YELLOW_FILL_NUMBER_STYLE
    protected CellStyle YELLOW_FILL_NUMBER_BOLD_STYLE
    protected CellStyle YELLOW_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle YELLOW_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected CellStyle BLUE_FILL_NUMBER_STYLE
    protected CellStyle BLUE_FILL_NUMBER_BOLD_STYLE
    protected CellStyle BLUE_FILL_NUMBER_ITALIC_STYLE
    protected CellStyle BLUE_FILL_NUMBER_BOLD_ITALIC_STYLE

    protected Integer NORMAL_FONT_SIZE
    protected Integer TITLE_FONT_SIZE
    protected Integer HEADER_FONT_SIZE

    protected Integer IMPACT_ROW_HEIGHT = 30
    protected Integer HEADER_ROW_HEIGHT = 35

    static Integer LS_MAX_CHARACTERS = 120
    static Integer PT_MAX_CHARACTERS = 85

    KnDiyWorkbook(String fontName = "Calibri Light",
                  int fontSizeIncrease = 0,
                  String numberFormat = "#0.0") {
        this.workbook = new XSSFWorkbook()
        this.creationHelper = workbook.getCreationHelper()

        NORMAL_FONT_SIZE = 12 + fontSizeIncrease
        HEADER_FONT_SIZE = 12 + fontSizeIncrease
        TITLE_FONT_SIZE = 14 + fontSizeIncrease

        NORMAL_FONT = XSSFTools.getFont(workbook, XSSFTools.NORMAL_FONT, IndexedColors.BLACK.getIndex(), fontName, NORMAL_FONT_SIZE)
        UNDER_LINE_FONT = XSSFTools.getFont(workbook, XSSFTools.NORMAL_FONT, IndexedColors.BLACK.getIndex(), fontName, NORMAL_FONT_SIZE)
        UNDER_LINE_FONT.setUnderline(XSSFFont.U_SINGLE)
        ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.ITALIC_FONT, IndexedColors.BLACK.getIndex(), fontName, NORMAL_FONT_SIZE)
        BOLD_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT, IndexedColors.BLACK.getIndex(), fontName, NORMAL_FONT_SIZE)
        BOLD_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_ITALIC_FONT, IndexedColors.BLACK.getIndex(), fontName, NORMAL_FONT_SIZE)
        TITLE_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT, IndexedColors.BLACK.getIndex(), fontName, TITLE_FONT_SIZE)
        TITLE_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_ITALIC_FONT, IndexedColors.BLACK.getIndex(), fontName, TITLE_FONT_SIZE)
        HEADER_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT, IndexedColors.BLACK.getIndex(), fontName, HEADER_FONT_SIZE)
        HEADER_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_ITALIC_FONT, IndexedColors.BLACK.getIndex(), fontName, HEADER_FONT_SIZE)

        RED_NORMAL_FONT = XSSFTools.getFont(workbook, XSSFTools.NORMAL_FONT, IndexedColors.RED.getIndex(), fontName, NORMAL_FONT_SIZE)
        RED_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.ITALIC_FONT, IndexedColors.RED.getIndex(), fontName, NORMAL_FONT_SIZE)
        RED_BOLD_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT, IndexedColors.RED.getIndex(), fontName, NORMAL_FONT_SIZE)
        RED_BOLD_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_ITALIC_FONT, IndexedColors.RED.getIndex(), fontName, NORMAL_FONT_SIZE)
        RED_TITLE_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT, IndexedColors.RED.getIndex(), fontName, TITLE_FONT_SIZE)
        RED_TITLE_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_ITALIC_FONT, IndexedColors.RED.getIndex(), fontName, TITLE_FONT_SIZE)
        RED_HEADER_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT, IndexedColors.RED.getIndex(), fontName, HEADER_FONT_SIZE)
        RED_HEADER_ITALIC_FONT = XSSFTools.getFont(workbook, XSSFTools.BOLD_ITALIC_FONT, IndexedColors.RED.getIndex(), fontName, HEADER_FONT_SIZE)

        WRAP_STYLE = createStyleWithFont(NORMAL_FONT)
        WRAP_BOLD_STYLE = createStyleWithFont(BOLD_FONT)
        WRAP_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT)
        WRAP_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT)

        WRAP_NUMBER_STYLE = createStyleWithFont(NORMAL_FONT, numberFormat)
        WRAP_NUMBER_BOLD_STYLE = createStyleWithFont(BOLD_FONT, numberFormat)
        WRAP_NUMBER_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT, numberFormat)
        WRAP_NUMBER_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT, numberFormat)

        TITLE_STYLE = createStyleWithFont(TITLE_FONT)
        TITLE_ITALIC_STYLE = createStyleWithFont(TITLE_ITALIC_FONT)

        HEADER_STYLE = createStyleWithFont(HEADER_FONT)
        HEADER_ITALIC_STYLE = createStyleWithFont(HEADER_ITALIC_FONT)
        HEADER_UNDERLINE_STYLE = createStyleWithFont(UNDER_LINE_FONT)

        RED_FILL_STYLE = createStyleWithFont(NORMAL_FONT)
        RED_FILL_BOLD_STYLE = createStyleWithFont(BOLD_FONT)
        RED_FILL_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT)
        RED_FILL_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT)

        ROSE_FILL_STYLE = createStyleWithFont(NORMAL_FONT)
        ROSE_FILL_BOLD_STYLE = createStyleWithFont(BOLD_FONT)
        ROSE_FILL_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT)
        ROSE_FILL_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT)

        GREEN_FILL_STYLE = createStyleWithFont(NORMAL_FONT)
        GREEN_FILL_BOLD_STYLE = createStyleWithFont(BOLD_FONT)
        GREEN_FILL_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT)
        GREEN_FILL_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT)

        GREY_FILL_STYLE = createStyleWithFont(NORMAL_FONT)
        GREY_FILL_BOLD_STYLE = createStyleWithFont(BOLD_FONT)
        GREY_FILL_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT)
        GREY_FILL_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT)

        ORANGE_FILL_STYLE = createStyleWithFont(NORMAL_FONT)
        ORANGE_FILL_BOLD_STYLE = createStyleWithFont(BOLD_FONT)
        ORANGE_FILL_ITALIC_STYLE = createStyleWithFont(ITALIC_FONT)
        ORANGE_FILL_BOLD_ITALIC_STYLE = createStyleWithFont(BOLD_ITALIC_FONT)

        YELLOW_FILL_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        YELLOW_FILL_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        YELLOW_FILL_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        YELLOW_FILL_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        BLUE_FILL_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        BLUE_FILL_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        BLUE_FILL_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        BLUE_FILL_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)


        RED_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        RED_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        RED_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        RED_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        ROSE_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        ROSE_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        ROSE_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        ROSE_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        GREEN_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        GREEN_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        GREEN_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        GREEN_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        GREY_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        GREY_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        GREY_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        GREY_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        ORANGE_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        ORANGE_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        ORANGE_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        ORANGE_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        YELLOW_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        YELLOW_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        YELLOW_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        YELLOW_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        BLUE_FILL_NUMBER_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, NORMAL_FONT)
        BLUE_FILL_NUMBER_BOLD_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_FONT)
        BLUE_FILL_NUMBER_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, ITALIC_FONT)
        BLUE_FILL_NUMBER_BOLD_ITALIC_STYLE = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat, BOLD_ITALIC_FONT)

        [
                WRAP_STYLE, WRAP_BOLD_STYLE, WRAP_ITALIC_STYLE, WRAP_BOLD_ITALIC_STYLE,
                WRAP_NUMBER_STYLE, WRAP_NUMBER_BOLD_STYLE, WRAP_NUMBER_ITALIC_STYLE, WRAP_NUMBER_BOLD_ITALIC_STYLE,
                TITLE_STYLE, TITLE_ITALIC_STYLE,
                HEADER_STYLE, HEADER_ITALIC_STYLE, HEADER_UNDERLINE_STYLE,
                RED_FILL_STYLE, RED_FILL_BOLD_STYLE, RED_FILL_ITALIC_STYLE, RED_FILL_BOLD_ITALIC_STYLE,
                GREEN_FILL_STYLE, GREEN_FILL_BOLD_STYLE, GREEN_FILL_ITALIC_STYLE, GREEN_FILL_BOLD_ITALIC_STYLE,
                GREY_FILL_STYLE, GREY_FILL_BOLD_STYLE, GREY_FILL_ITALIC_STYLE, GREY_FILL_BOLD_ITALIC_STYLE,
                ORANGE_FILL_STYLE, ORANGE_FILL_BOLD_STYLE, ORANGE_FILL_ITALIC_STYLE, ORANGE_FILL_BOLD_ITALIC_STYLE,
                RED_FILL_NUMBER_STYLE, RED_FILL_NUMBER_BOLD_STYLE, RED_FILL_NUMBER_ITALIC_STYLE, RED_FILL_NUMBER_BOLD_ITALIC_STYLE,
                GREEN_FILL_NUMBER_STYLE, GREEN_FILL_NUMBER_BOLD_STYLE, GREEN_FILL_NUMBER_ITALIC_STYLE, GREEN_FILL_NUMBER_BOLD_ITALIC_STYLE,
                GREY_FILL_NUMBER_STYLE, GREY_FILL_NUMBER_BOLD_STYLE, GREY_FILL_NUMBER_ITALIC_STYLE, GREY_FILL_NUMBER_BOLD_ITALIC_STYLE,
                ORANGE_FILL_NUMBER_STYLE, ORANGE_FILL_NUMBER_BOLD_STYLE, ORANGE_FILL_NUMBER_ITALIC_STYLE, ORANGE_FILL_NUMBER_BOLD_ITALIC_STYLE,
                BLUE_FILL_STYLE, BLUE_FILL_BOLD_STYLE, BLUE_FILL_ITALIC_STYLE, BLUE_FILL_BOLD_ITALIC_STYLE,
                BLUE_FILL_NUMBER_STYLE, BLUE_FILL_NUMBER_BOLD_STYLE, BLUE_FILL_NUMBER_ITALIC_STYLE, BLUE_FILL_NUMBER_BOLD_ITALIC_STYLE,
                YELLOW_FILL_STYLE, YELLOW_FILL_BOLD_STYLE, YELLOW_FILL_ITALIC_STYLE, YELLOW_FILL_BOLD_ITALIC_STYLE,
                YELLOW_FILL_NUMBER_STYLE, YELLOW_FILL_NUMBER_BOLD_STYLE, YELLOW_FILL_NUMBER_ITALIC_STYLE, YELLOW_FILL_NUMBER_BOLD_ITALIC_STYLE,
                ROSE_FILL_STYLE, ROSE_FILL_BOLD_STYLE, ROSE_FILL_ITALIC_STYLE, ROSE_FILL_BOLD_ITALIC_STYLE,
                ROSE_FILL_NUMBER_STYLE, ROSE_FILL_NUMBER_BOLD_STYLE, ROSE_FILL_NUMBER_ITALIC_STYLE, ROSE_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ].each { CellStyle cellStyle ->
            XSSFTools.addBorderToStyle(cellStyle)
            XSSFTools.setCenterAlignmentToStyles(cellStyle)
            cellStyle.setWrapText(true)
        }

        applyForeGroundColor([
                RED_FILL_STYLE, RED_FILL_BOLD_STYLE, RED_FILL_ITALIC_STYLE, RED_FILL_BOLD_ITALIC_STYLE,
                RED_FILL_NUMBER_STYLE, RED_FILL_NUMBER_BOLD_STYLE, RED_FILL_NUMBER_ITALIC_STYLE, RED_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ], [ 244, 67, 54 ])

        applyForeGroundColor([
                GREEN_FILL_STYLE, GREEN_FILL_BOLD_STYLE, GREEN_FILL_ITALIC_STYLE, GREEN_FILL_BOLD_ITALIC_STYLE,
                GREEN_FILL_NUMBER_STYLE, GREEN_FILL_NUMBER_BOLD_STYLE, GREEN_FILL_NUMBER_ITALIC_STYLE, GREEN_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ], [ 33, 150, 243 ])

        applyForeGroundColor([
                GREY_FILL_STYLE, GREY_FILL_BOLD_STYLE, GREY_FILL_ITALIC_STYLE, GREY_FILL_BOLD_ITALIC_STYLE,
                GREY_FILL_NUMBER_STYLE, GREY_FILL_NUMBER_BOLD_STYLE, GREY_FILL_NUMBER_ITALIC_STYLE, GREY_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ], [ 238, 238, 238 ])

        applyForeGroundColor([
                ORANGE_FILL_STYLE, ORANGE_FILL_BOLD_STYLE, ORANGE_FILL_ITALIC_STYLE, ORANGE_FILL_BOLD_ITALIC_STYLE,
                ORANGE_FILL_NUMBER_STYLE, ORANGE_FILL_NUMBER_BOLD_STYLE, ORANGE_FILL_NUMBER_ITALIC_STYLE, ORANGE_FILL_NUMBER_BOLD_ITALIC_STYLE
        ], [ 255, 217, 102 ])

        applyForeGroundColor([
                YELLOW_FILL_STYLE, YELLOW_FILL_BOLD_STYLE, YELLOW_FILL_ITALIC_STYLE, YELLOW_FILL_BOLD_ITALIC_STYLE,
                YELLOW_FILL_NUMBER_STYLE, YELLOW_FILL_NUMBER_BOLD_STYLE, YELLOW_FILL_NUMBER_ITALIC_STYLE, YELLOW_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ], [ 255, 242, 204 ])

        applyForeGroundColor([
                BLUE_FILL_STYLE, BLUE_FILL_BOLD_STYLE, BLUE_FILL_ITALIC_STYLE, BLUE_FILL_BOLD_ITALIC_STYLE,
                BLUE_FILL_NUMBER_STYLE, BLUE_FILL_NUMBER_BOLD_STYLE, BLUE_FILL_NUMBER_ITALIC_STYLE, BLUE_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ], [ 221, 235, 247 ])

        applyForeGroundColor([
                ROSE_FILL_STYLE, ROSE_FILL_BOLD_STYLE, ROSE_FILL_ITALIC_STYLE, ROSE_FILL_BOLD_ITALIC_STYLE,
                ROSE_FILL_NUMBER_STYLE, ROSE_FILL_NUMBER_BOLD_STYLE, ROSE_FILL_NUMBER_ITALIC_STYLE, ROSE_FILL_NUMBER_BOLD_ITALIC_STYLE,
        ], [ 252, 228, 214 ])
    }

    private void applyForeGroundColor(List<CellStyle> cellStyles, List rgbValues) {
        XSSFColor color = new XSSFColor(new byte[]{(byte)rgbValues[0], (byte)rgbValues[1], (byte)rgbValues[2]})

        cellStyles.each { CellStyle cellStyle ->
            cellStyle.setFillForegroundColor(color)
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        }
    }

    private CellStyle createStyleWithFont(Font font, String numberFormat = null) {
        CellStyle cellStyle = numberFormat
                ? XSSFTools.getCellStyleWithFormat(workbook, creationHelper, numberFormat)
                : workbook.createCellStyle()
        cellStyle.setFont(font)

        return cellStyle
    }

    private void saveWorkBookAsFile(String pathToSaveWithTyped) {
        try (FileOutputStream outputStream = new FileOutputStream(new File(pathToSaveWithTyped))) {
            workbook.write(outputStream)
        }
        catch (Exception ex) {
            println(ex.getMessage())
            ex.printStackTrace()
        }
    }

    void createAndSaveWorkbook(String pathToSave) {
        println("")
        println("++++++++++++++++++++")
        println("++++++++++++++++++++")
        println("")
        println("--Start Writing Workbook at ${pathToSave}")
        writeIntoWorkBook()
        println("--Finished Writing Workbook")
        saveWorkBookAsFile(pathToSave)
        println("--> SAVED Workbook to ${pathToSave}")
    }

    static void setColumnWidth(Sheet sheet,
                               int columnIdx,
                               int columnWidthInChars) {
        Integer columnWidth = (columnWidthInChars * 256).toBigInteger().toInteger()
        sheet.setColumnWidth(columnIdx, columnWidth)
    }

    static CellStyle getCellStyleForMaxMinValue(Inspection inspection,
                                                BigDecimal value,
                                                CellStyle maxStyle,
                                                CellStyle minStyle,
                                                CellStyle defaultStyle,
                                                boolean isTemp) {
        BigDecimal allTimeMaxTemp = inspection.getAllTimeMaxTemperature()
        BigDecimal allTimeMinTemp = inspection.getAllTimeMinTemperature()
        BigDecimal allTimeMaxHum = inspection.getAllTimeMaxHumidity()
        BigDecimal allTimeMinHum = inspection.getAllTimeMinHumidity()

        boolean isMax = isTemp ? value == allTimeMaxTemp : value == allTimeMaxHum
        boolean isMin = isTemp ? value == allTimeMinTemp : value == allTimeMinHum

        return (isMax
                ? maxStyle
                : isMin
                ? minStyle
                : defaultStyle)
    }

    static String getAltitudeDesc(String altitude) {
        if (altitude == "H") {
            return "Cao / High"
        }
        if (altitude == "M") {
            return "Giữa / Middle"
        }
        if (altitude == "L") {
            return "Thấp / Low"
        }

        return null
    }
}
