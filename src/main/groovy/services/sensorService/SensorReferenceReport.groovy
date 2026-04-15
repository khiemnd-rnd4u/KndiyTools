package services.sensorService

import entities.sensor.Inspection
import entities.sensor.LogTagCalibration
import entities.sensor.LogTagSensor
import org.apache.poi.ss.usermodel.PageMargin
import org.apache.poi.ss.usermodel.PaperSize
import org.apache.poi.ss.usermodel.PrintSetup
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFSheet
import services.kndiyLibraries.XSSFTools
import services.kndiyLibraries.DateTimeResolver

class SensorReferenceReport extends KnDiyWorkbook {
    private static String ASSET_SHEET_NAME = "DataLoggerDetails"
    private static String TEMPERATURE_REF_SHEET_NAME = "ReferenceTemperature"
    private static String HUMIDITY_REF_SHEET_NAME = "ReferenceHumidity"
    private static String REPORT_DATE_TIME_PATTERN = "yyyy-MM-dd HH:mm"
    private XSSFSheet assetSheet
    private XSSFSheet refTemperatureSheet
    private XSSFSheet refHumiditySheet

    private Integer TITLE_ROW_NUM
    private Integer HEADER_ROW_NUM
    private Integer HEADER_EN_ROW_NUM
    private Integer DATA_START_ROW_NUM

    private Integer NO_CELL_NUM
    private Integer ASSET_ID_CELL_NUM
    private Integer ASSET_SERIAL_CELL_NUM
    private Integer ASSET_LOCATION_CELL_NUM
    private Integer ASSET_ALTITUDE_CELL_NUM
    private Integer ASSET_INSPECTION_DATE_CELL_NUM
    private Integer ASSET_CALIBRATION_ID_CELL_NUM
    private Integer ASSET_CALIBRATION_DATE_CELL_NUM
    private Integer ASSET_CALIBRATION_DUE_DATE_CELL_NUM
    private Integer ASSET_END_CELL_NUM

    private Integer REF_LOCATION_CELL_NUM
    private Integer REF_AVG_CELL_NUM
    private Integer REF_MAX_CELL_NUM
    private Integer REF_MAX_TIME_CELL_NUM
    private Integer REF_MIN_CELL_NUM
    private Integer REF_MIN_TIME_CELL_NUM
    private Integer REF_CONFORM_CELL_NUM
    private Integer REF_END_CELL_NUM

    private Inspection inspection
    private boolean hasAltitude

    SensorReferenceReport(Inspection inspection,
                          boolean hasAltitude) {
        super()

        this.hasAltitude = hasAltitude
        this.inspection = inspection

        int rowNum = 0
        TITLE_ROW_NUM = rowNum
        HEADER_ROW_NUM = ++rowNum
        HEADER_EN_ROW_NUM = ++rowNum
        DATA_START_ROW_NUM = ++rowNum

        int cellNum = 0
        NO_CELL_NUM = cellNum
        ASSET_ID_CELL_NUM = ++cellNum
        ASSET_SERIAL_CELL_NUM = ++cellNum
        ASSET_LOCATION_CELL_NUM = ++cellNum
        if (hasAltitude) {
            ASSET_ALTITUDE_CELL_NUM = ++cellNum
        }
        ASSET_CALIBRATION_ID_CELL_NUM = ++cellNum
        ASSET_CALIBRATION_DATE_CELL_NUM = ++cellNum
        ASSET_INSPECTION_DATE_CELL_NUM = ++cellNum
        ASSET_CALIBRATION_DUE_DATE_CELL_NUM = ++cellNum
        ASSET_END_CELL_NUM = cellNum

        cellNum = 0
        REF_LOCATION_CELL_NUM = ++cellNum
        REF_AVG_CELL_NUM = ++cellNum
        REF_MAX_CELL_NUM = ++cellNum
        REF_MAX_TIME_CELL_NUM = ++cellNum
        REF_MIN_CELL_NUM = ++cellNum
        REF_MIN_TIME_CELL_NUM = ++cellNum
        REF_CONFORM_CELL_NUM = ++cellNum
        REF_END_CELL_NUM = cellNum
    }

    @Override
    void writeIntoWorkBook() {
        createSheets()
        writeAssets()
        writeReferenceTemperature()
        writeReferenceHumidity()
    }

    private void createSheets() {
        assetSheet = createSheets(ASSET_SHEET_NAME)
        println("Created Asset Sheet")
        refTemperatureSheet = createSheets(TEMPERATURE_REF_SHEET_NAME)
        println("Created refTemperatureSheet Sheet")
        refHumiditySheet = createSheets(HUMIDITY_REF_SHEET_NAME)
        println("Created refHumiditySheet Sheet")
    }

    private XSSFSheet createSheets(String sheetName) {
        XSSFSheet sheet = workbook.createSheet(sheetName)

        sheet.setMargin(PageMargin.TOP, 0.5)
        sheet.setMargin(PageMargin.BOTTOM, 0.5)
        sheet.setMargin(PageMargin.LEFT, 0.5)
        sheet.setMargin(PageMargin.RIGHT, 0.5)

        PrintSetup printSetup = sheet.getPrintSetup();

        printSetup.setFitWidth((short) 0)
        printSetup.setFitHeight((short) (0))
        printSetup.setPaperSize(PaperSize.A4_PAPER)
        printSetup.setLandscape(false)

        sheet.setHorizontallyCenter(true)

        if (sheetName == ASSET_SHEET_NAME) {
            setColumnWidth(sheet, NO_CELL_NUM, 5)
            setColumnWidth(sheet, ASSET_ID_CELL_NUM, 5)
            setColumnWidth(sheet, ASSET_SERIAL_CELL_NUM, 13)
            setColumnWidth(sheet, ASSET_LOCATION_CELL_NUM, 14)

            setColumnWidth(sheet, ASSET_CALIBRATION_ID_CELL_NUM, 12)
            setColumnWidth(sheet, ASSET_CALIBRATION_DATE_CELL_NUM, 12)
            setColumnWidth(sheet, ASSET_INSPECTION_DATE_CELL_NUM, 19)
            setColumnWidth(sheet, ASSET_CALIBRATION_DUE_DATE_CELL_NUM, 12)
        }
        else if (sheetName in [ TEMPERATURE_REF_SHEET_NAME, HUMIDITY_REF_SHEET_NAME ]) {
            setColumnWidth(sheet, NO_CELL_NUM, 5)
            setColumnWidth(sheet, REF_LOCATION_CELL_NUM, 13)
            setColumnWidth(sheet, REF_MAX_CELL_NUM, 10)
            setColumnWidth(sheet, REF_MAX_TIME_CELL_NUM, 12)
            setColumnWidth(sheet, REF_AVG_CELL_NUM, 10)
            setColumnWidth(sheet, REF_MIN_CELL_NUM, 10)
            setColumnWidth(sheet, REF_MIN_TIME_CELL_NUM, 12)
            setColumnWidth(sheet, REF_CONFORM_CELL_NUM, 20)
        }

        return sheet
    }

    private void writeAssets() {
        writeAssetTitle()
        writeAssetHeader()
        writeAssetData()
    }

    private void writeAssetTitle() {
        XSSFTools.styleMergeCells(
                assetSheet, TITLE_ROW_NUM, TITLE_ROW_NUM, NO_CELL_NUM, ASSET_END_CELL_NUM,
                GREY_FILL_BOLD_STYLE,
                "Chi tiết cảm biến nhiệt độ, độ ẩm dùng cho thẩm định/\n Details of data-loggers used for qualification",
                HEADER_ROW_HEIGHT
        )
    }

    private void writeAssetHeader() {
        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, NO_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "STT",
                HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, NO_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "No.",
                HEADER_ROW_HEIGHT
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_ID_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Log"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_ID_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "ID"
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_SERIAL_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Số Serial"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_SERIAL_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Serial No."
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_LOCATION_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Vị trí"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_LOCATION_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Loc."
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_CALIBRATION_ID_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "ID Hiệu chuẩn"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_CALIBRATION_ID_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Calib. ID"
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_CALIBRATION_DATE_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Ngày hiệu chuẩn"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_CALIBRATION_DATE_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Calib. Date"
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_INSPECTION_DATE_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Ngày thực hiện"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_INSPECTION_DATE_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Qualification Date"
        )

        XSSFTools.setCellValue(
                assetSheet, HEADER_ROW_NUM, ASSET_CALIBRATION_DUE_DATE_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Ngày hết hạn"
        )
        XSSFTools.setCellValue(
                assetSheet, HEADER_EN_ROW_NUM, ASSET_CALIBRATION_DUE_DATE_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Calib. Due Date"
        )
    }

    private void writeAssetData() {
        int idx = 0
        inspection.getLogTagSensorByLocation()
                .each { String location, LogTagSensor logTag ->
                    writeAssetData(logTag, ++idx)
                }
        writeAssetData(inspection.getEnvironmentSensor(), ++idx)
    }

    private void writeAssetData(LogTagSensor logTag, int idx) {
        int currentRow = DATA_START_ROW_NUM + idx - 1
        XSSFTools.setCellValue(
                assetSheet, currentRow, NO_CELL_NUM, WRAP_STYLE, idx
        )
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_ID_CELL_NUM, WRAP_STYLE, logTag.getId()
        )
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_SERIAL_CELL_NUM, WRAP_STYLE, logTag.getSerial()
        )

        String location = logTag.getLocation()
        String altitude = logTag.getAltitudeLevel()
        String locationData = Inspection.getLocationKey(location, altitude)
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_LOCATION_CELL_NUM, WRAP_STYLE, locationData
        )
        if (hasAltitude) {
            String altitudeDesc = SensorDistributionReport.getAltitudeDesc(altitude)
            XSSFTools.setCellValue(
                    assetSheet, currentRow, ASSET_ALTITUDE_CELL_NUM, WRAP_STYLE, altitudeDesc
            )
        }

        LogTagCalibration calibration = logTag.getCalibration()
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_CALIBRATION_ID_CELL_NUM, WRAP_STYLE, calibration.getCalibrationId()
        )
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_CALIBRATION_DATE_CELL_NUM, WRAP_STYLE,
                DateTimeResolver.getDateString(calibration.getCalibrationDate())
        )
        String fromTime = DateTimeResolver.getDateTimeString(inspection.getFromTime(), REPORT_DATE_TIME_PATTERN)
        String toTime = DateTimeResolver.getDateTimeString(inspection.getToTime(), REPORT_DATE_TIME_PATTERN)
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_INSPECTION_DATE_CELL_NUM, WRAP_STYLE,
                "${fromTime} - ${toTime}"
        )
        XSSFTools.setCellValue(
                assetSheet, currentRow, ASSET_CALIBRATION_DUE_DATE_CELL_NUM, WRAP_STYLE,
                DateTimeResolver.getDateString(calibration.getCalibrationDueDate())
        )
    }

    private void writeReferenceTemperature() {
        writeReference(refTemperatureSheet, true)
    }

    private void writeReferenceHumidity() {
        writeReference(refHumiditySheet, false)
    }

    private void writeReference(Sheet sheet, boolean isTemp) {
        writeReferenceTitle(sheet, isTemp)
        writeReferenceHeader(sheet, isTemp)
        writeReferenceData(sheet, isTemp)
    }

    private void writeReferenceTitle(Sheet sheet, boolean isTemp) {
        String title = isTemp
                ? "Sơ lược giá trị nhiệt độ / Summary of Temperature readings"
                : "Sơ lược giá trị độ ẩm / Summary of Humidity readings"

        XSSFTools.styleMergeCells(
                sheet, TITLE_ROW_NUM, TITLE_ROW_NUM, NO_CELL_NUM, REF_END_CELL_NUM,
                GREY_FILL_BOLD_STYLE,
                title,
                HEADER_ROW_HEIGHT
        )
    }

    private void writeReferenceHeader(Sheet sheet, boolean isTemp) {
        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, NO_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "STT",
                HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, NO_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "No.",
                HEADER_ROW_HEIGHT
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_LOCATION_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Vị trí"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_LOCATION_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Loc."
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_AVG_CELL_NUM, GREY_FILL_BOLD_STYLE,
                isTemp ? "Nhiệt độ trung bình" : "Độ ẩm trung bình"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_AVG_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                isTemp ? "T. Avg (°C)" : "RH Avg (%)"
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_MAX_CELL_NUM, GREY_FILL_BOLD_STYLE,
                isTemp ? "Nhiệt độ cao nhất" : "Độ ẩm cao nhất"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_MAX_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                isTemp ? "T. Max (°C)" : "RH Max (%)"
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_MAX_TIME_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Ngày giờ"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_MAX_TIME_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                isTemp ? "T. Max Time" : "RH Max Time"
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_MIN_CELL_NUM, GREY_FILL_BOLD_STYLE,
                isTemp ? "Nhiệt độ thấp nhất" : "Độ ẩm thấp nhất"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_MIN_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                isTemp ? "T. Min (°C)" : "RH Min (%)"
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_MIN_TIME_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "Ngày giờ"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_MIN_TIME_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                isTemp ? "T. Min Time" : "RH Min Time"
        )

        XSSFTools.setCellValue(
                sheet, HEADER_ROW_NUM, REF_CONFORM_CELL_NUM, GREY_FILL_BOLD_STYLE,
                "ĐẠT/\nKHÔNG ĐẠT"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_EN_ROW_NUM, REF_CONFORM_CELL_NUM, GREY_FILL_BOLD_ITALIC_STYLE,
                "Conformed/\nNone Conformed"
        )
    }

    private void writeReferenceData(Sheet sheet, boolean isTemp) {
        int idx = 0
        inspection.getLogTagSensorByLocation()
                .each {String location, LogTagSensor logTag ->
                    writeReferenceData(sheet, logTag, ++idx, isTemp)
                }

        writeReferenceData(sheet, inspection.getEnvironmentSensor(), ++idx, isTemp)
    }

    private void writeReferenceData(Sheet sheet, LogTagSensor logTag, int idx, boolean isTemp) {
        int rowNum = DATA_START_ROW_NUM + idx - 1
        BigDecimal max = isTemp ? logTag.getMaxTemperature() : logTag.getMaxHumidity()
        BigDecimal min = isTemp ? logTag.getMinTemperature() : logTag.getMinHumidity()

        XSSFTools.setCellValue(
                sheet, rowNum, NO_CELL_NUM, WRAP_STYLE, idx
        )

        String location = logTag.getLocation()
        if (location.contains("ENV")) {
            location += "/\nMôi trường"
        }
        XSSFTools.setCellValue(
                sheet, rowNum, REF_LOCATION_CELL_NUM, WRAP_STYLE, location
        )

        XSSFTools.setCellValue(
                sheet, rowNum, REF_AVG_CELL_NUM, WRAP_NUMBER_STYLE, isTemp ? logTag.getAvgTemperature() : logTag.getAvgHumidity()
        )

        XSSFTools.setCellValue(
                sheet, rowNum, REF_MAX_CELL_NUM,
                getCellStyleForMaxMinValue(
                        inspection, max, RED_FILL_NUMBER_STYLE, GREEN_FILL_STYLE, WRAP_NUMBER_STYLE, isTemp
                ),
                max
        )

        XSSFTools.setCellValue(
                sheet, rowNum, REF_MAX_TIME_CELL_NUM,
                WRAP_STYLE,
                DateTimeResolver.getDateTimeString(
                        isTemp ? logTag.getMaxTemperatureTime() : logTag.getMaxHumidityTime(),
                        REPORT_DATE_TIME_PATTERN
                )
        )

        XSSFTools.setCellValue(
                sheet, rowNum, REF_MIN_CELL_NUM,
                getCellStyleForMaxMinValue(
                        inspection, min, RED_FILL_NUMBER_STYLE, GREEN_FILL_STYLE, WRAP_NUMBER_STYLE, isTemp
                ),
                min
        )

        XSSFTools.setCellValue(
                sheet, rowNum, REF_MIN_TIME_CELL_NUM, WRAP_STYLE,
                DateTimeResolver.getDateTimeString(
                        isTemp ? logTag.getMinTemperatureTime() : logTag.getMinHumidityTime(),
                        REPORT_DATE_TIME_PATTERN
                )
        )

        XSSFTools.setCellValue(
                sheet, rowNum, REF_CONFORM_CELL_NUM, WRAP_STYLE,
                Inspection.getConformedState(
                        inspection, logTag, isTemp
                )
        )
    }
}
