package services.sensorService

import entities.sensor.Inspection
import entities.sensor.LogTagReading
import entities.sensor.LogTagSensor
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.PageMargin
import org.apache.poi.ss.usermodel.PrintSetup
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xddf.usermodel.chart.AxisCrosses
import org.apache.poi.xddf.usermodel.chart.AxisPosition
import org.apache.poi.xddf.usermodel.chart.ChartTypes
import org.apache.poi.xddf.usermodel.chart.LegendPosition
import org.apache.poi.xddf.usermodel.chart.MarkerStyle
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis
import org.apache.poi.xddf.usermodel.chart.XDDFChartData
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis
import org.apache.poi.xssf.usermodel.XSSFAnchor
import org.apache.poi.xssf.usermodel.XSSFChart
import org.apache.poi.xssf.usermodel.XSSFDrawing
import org.apache.poi.xssf.usermodel.XSSFSheet
import services.kndiyLibraries.DateTimeResolver
import services.kndiyLibraries.XSSFTools

import java.math.RoundingMode
import java.time.ZonedDateTime

class SensorDistributionReport extends KnDiyWorkbook {
    private static String TEMPERATURE_CHART_SHEET_NAME = "TemperatureChart"
    private static String TEMPERATURE_SHEET_NAME = "Temperature"
    private static String HUMIDITY_CHART_SHEET_NAME = "HumidityChart"
    private static String HUMIDITY_SHEET_NAME = "Humidity"

    private XSSFSheet temperatureChartSheet
    private XSSFSheet temperatureSheet
    private XSSFSheet humidityChartSheet
    private XSSFSheet humiditySheet

    private int TITLE_ROW = 0
    private int TITLE_ROW_EN = 1

    private int CHART_START_ROW_NUMBER = 3
    private int CHART_END_ROW_NUMBER = 50
    private int CHART_START_CELL_NUMBER = 0
    private int CHART_END_CELL_NUMBER = 20

    private int HEADER_LOG_TAG_SERIAL_ROW = 2
    private Integer HEADER_ALTITUDE_DESC_ROW
    private int HEADER_LOCATION_AND_ALTITUDE_ROW

    private int MAX_ROW_NUMBER
    private int AVG_ROW_NUMBER
    private int MIN_ROW_NUMBER
    private int MKT_ROW_NUMBER
    private int DATA_START_TEMP_ROW_NUMBER
    private int DATA_START_RH_ROW_NUMBER
    private int DATA_END_TEMP_ROW_NUMBER
    private int DATA_END_RH_ROW_NUMBER

    private int INDEX_CELL_NUMBER = 0
    private int TIME_CELL_NUMBER = 1

    private int DATA_START_CELL_NUMBER = 2
    private int DATA_END_CELL_NUMBER
    private Integer MIN_TOLERANCE_CELL_NUMBER
    private Integer MAX_TOLERANCE_CELL_NUMBER

    private BigDecimal ALL_TIME_MAX_TEMPERATURE
    private BigDecimal ALL_TIME_MIN_TEMPERATURE
    private BigDecimal ALL_TIME_MAX_HUMIDITY
    private BigDecimal ALL_TIME_MIN_HUMIDITY
    private BigDecimal ALL_TIME_OFF_SET_TEMP = 0.5
    private BigDecimal ALL_TIME_OFF_SET_HUM = 2.5

    private Inspection inspection

    private Map<String, Integer> cellNumberByLogTagSerial
    private boolean hasAltitude

    private static String REPORT_TIME_PATTERN = "yyyy-MM-dd HH:mm"

    SensorDistributionReport(Inspection inspection,
                             boolean hasAltitude = true) {
        super()
        this.inspection = inspection
        this.hasAltitude = hasAltitude
        cellNumberByLogTagSerial = [ : ]

        Integer sensorCount = inspection.getSensorCount()
        DATA_END_CELL_NUMBER = DATA_START_CELL_NUMBER + sensorCount - 1
        MIN_TOLERANCE_CELL_NUMBER = DATA_END_CELL_NUMBER + 1
        MAX_TOLERANCE_CELL_NUMBER = DATA_END_CELL_NUMBER + 2

        int currentRowNumber = HEADER_LOG_TAG_SERIAL_ROW
        HEADER_ALTITUDE_DESC_ROW = hasAltitude ? ++currentRowNumber : null
        HEADER_LOCATION_AND_ALTITUDE_ROW = ++currentRowNumber

        MAX_ROW_NUMBER = ++currentRowNumber
        AVG_ROW_NUMBER = ++currentRowNumber
        MIN_ROW_NUMBER = ++currentRowNumber
        MKT_ROW_NUMBER = ++currentRowNumber
        DATA_START_TEMP_ROW_NUMBER = ++currentRowNumber
        DATA_START_RH_ROW_NUMBER = --currentRowNumber
        DATA_END_TEMP_ROW_NUMBER = DATA_START_TEMP_ROW_NUMBER + inspection.getMaxDataCount() - 1
        DATA_END_RH_ROW_NUMBER = DATA_START_RH_ROW_NUMBER + inspection.getMaxDataCount() - 1

        ALL_TIME_MAX_TEMPERATURE = inspection.getAllTimeMaxTemperature()
        ALL_TIME_MIN_TEMPERATURE = inspection.getAllTimeMinTemperature()
        ALL_TIME_MAX_HUMIDITY = inspection.getAllTimeMaxHumidity()
        ALL_TIME_MIN_HUMIDITY = inspection.getAllTimeMinHumidity()
    }

    @Override
    void writeIntoWorkBook() {
        createSheets()
        println("Created Sheets")

        writeHeaders()
        println("Wrote Headers")

        writeIndicators()
        println("Wrote Indicators")

        writeData()
        println("Wrote Data")

        createCharts()
        println("Created Charts")
    }

    private XSSFSheet createSheets() {
        temperatureChartSheet = createSheets(TEMPERATURE_CHART_SHEET_NAME, true)
        temperatureSheet = createSheets(TEMPERATURE_SHEET_NAME)

        humidityChartSheet = createSheets(HUMIDITY_CHART_SHEET_NAME, true)
        humiditySheet = createSheets(HUMIDITY_SHEET_NAME)
    }

    private XSSFSheet createSheets(String sheetName,
                                   boolean isChartSheet = false,
                                   boolean isSimpleSheet = false) {
        XSSFSheet sheet = workbook.createSheet(sheetName)

        sheet.setMargin(PageMargin.TOP, 0.3)
        sheet.setMargin(PageMargin.BOTTOM, 0.3)
        sheet.setMargin(PageMargin.LEFT, 0.3)
        sheet.setMargin(PageMargin.RIGHT, 0.3)

        PrintSetup printSetup = sheet.getPrintSetup();

        printSetup.setFitWidth((short) 1)
        printSetup.setFitHeight((short) (isChartSheet ? 1 : 0))

        printSetup.setLandscape(true)

        sheet.setFitToPage(true)
        sheet.setHorizontallyCenter(true)
        sheet.setVerticallyCenter(isChartSheet)

        if (!isChartSheet) {
            sheet.setColumnWidth(INDEX_CELL_NUMBER, 8 * 256)
            sheet.setColumnWidth(TIME_CELL_NUMBER, 24 * 256)
            (DATA_START_CELL_NUMBER..MAX_TOLERANCE_CELL_NUMBER).each { int cellNum ->
                sheet.setColumnWidth(cellNum, 15 * 256)
            }
        }

        if (sheetName in [ TEMPERATURE_SHEET_NAME, HUMIDITY_SHEET_NAME ]) {
            sheet.setRepeatingRows(new CellRangeAddress(TITLE_ROW, HEADER_LOCATION_AND_ALTITUDE_ROW, -1, -1))
        }

        return sheet
    }

    private void writeHeaders() {
        writeTitle(temperatureSheet, true, false)
        writeTitle(humiditySheet, false, false)

        writeTitle(temperatureChartSheet, true, true)
        writeTitle(humidityChartSheet, false, true)

        writeStaticHeaders(temperatureSheet, true)
        writeStaticHeaders(humiditySheet, false)

        writeMaxToleranceData(temperatureSheet)
        writeMinToleranceData(temperatureSheet)
        writeMaxToleranceData(humiditySheet, false)
        writeMinToleranceData(humiditySheet, false)

        writeDataLogMeta(temperatureSheet)
        writeDataLogMeta(humiditySheet)
    }

    private void writeTitle(Sheet sheet, boolean isTemp, boolean isChart) {
        String inspectionName = inspection.getName()
        ZonedDateTime fromTime = inspection.getFromTime()
        String fromTimeStr = DateTimeResolver.getDateTimeString(fromTime, REPORT_TIME_PATTERN)
        ZonedDateTime toTime = inspection.getToTime()
        String toTimeStr = DateTimeResolver.getDateTimeString(toTime, REPORT_TIME_PATTERN)

        String data = "${inspectionName} từ ${fromTimeStr} đến ${toTimeStr}"
        String enData = "${inspectionName} from ${fromTimeStr} to ${toTimeStr}"
        String title = isTemp
                ? "Báo cáo phân bố nhiệt độ (°C) ${data}"
                : "Báo cáo phân bố độ ẩm (%RH) ${data}"
        String enTitle = isTemp
                ? "Temperature (°C) distribution of ${enData}"
                : "Humidity (%RH) distribution of ${enData}"

        if (isChart) {
            XSSFTools.styleMergeCells(sheet, TITLE_ROW, TITLE_ROW, CHART_START_CELL_NUMBER, CHART_END_CELL_NUMBER, TITLE_STYLE, title, IMPACT_ROW_HEIGHT)
            XSSFTools.styleMergeCells(sheet, TITLE_ROW_EN, TITLE_ROW_EN, CHART_START_CELL_NUMBER, CHART_END_CELL_NUMBER, TITLE_ITALIC_STYLE, enTitle, IMPACT_ROW_HEIGHT)
        }
        else {
            XSSFTools.styleMergeCells(sheet, TITLE_ROW, TITLE_ROW, INDEX_CELL_NUMBER, MAX_TOLERANCE_CELL_NUMBER, TITLE_STYLE, title, IMPACT_ROW_HEIGHT)
            XSSFTools.styleMergeCells(sheet, TITLE_ROW_EN, TITLE_ROW_EN, INDEX_CELL_NUMBER, MAX_TOLERANCE_CELL_NUMBER, TITLE_ITALIC_STYLE, enTitle, IMPACT_ROW_HEIGHT)
        }
    }

    private void writeStaticHeaders(Sheet sheet, boolean isTemp) {
        XSSFTools.styleMergeCells(
                sheet, HEADER_LOCATION_AND_ALTITUDE_ROW + 1,
                isTemp ? DATA_START_TEMP_ROW_NUMBER - 1 : DATA_START_RH_ROW_NUMBER - 1,
                MIN_TOLERANCE_CELL_NUMBER, MAX_TOLERANCE_CELL_NUMBER,
                WRAP_STYLE,
                ""
        )

        XSSFTools.styleMergeCells(
                sheet, HEADER_LOG_TAG_SERIAL_ROW, HEADER_LOG_TAG_SERIAL_ROW, INDEX_CELL_NUMBER, TIME_CELL_NUMBER,
                WRAP_BOLD_ITALIC_STYLE,
                "Mã số quản lý /\n Datalogger Serial",
                HEADER_ROW_HEIGHT
        )

        XSSFTools.setCellValue(sheet, HEADER_LOCATION_AND_ALTITUDE_ROW, INDEX_CELL_NUMBER, GREY_FILL_BOLD_ITALIC_STYLE, "STT /\n No.", HEADER_ROW_HEIGHT)
        XSSFTools.setCellValue(sheet, HEADER_LOCATION_AND_ALTITUDE_ROW, TIME_CELL_NUMBER, GREY_FILL_BOLD_ITALIC_STYLE, "Ngày và giờ /\n Date & Time", HEADER_ROW_HEIGHT)

        if (hasAltitude) {
            XSSFTools.styleMergeCells(
                    sheet, HEADER_ALTITUDE_DESC_ROW, HEADER_ALTITUDE_DESC_ROW, INDEX_CELL_NUMBER, TIME_CELL_NUMBER,
                    WRAP_BOLD_ITALIC_STYLE,
                    "Mức chiều cao /\n Height level",
                    HEADER_ROW_HEIGHT
            )
        }

        XSSFTools.styleMergeCells(
                sheet, HEADER_LOG_TAG_SERIAL_ROW, HEADER_LOCATION_AND_ALTITUDE_ROW - 1, MIN_TOLERANCE_CELL_NUMBER, MAX_TOLERANCE_CELL_NUMBER,
                WRAP_BOLD_ITALIC_STYLE,
                "Giới hạn chấp nhận /\nTolerance",
        )
        XSSFTools.setCellValue(
                sheet, HEADER_LOCATION_AND_ALTITUDE_ROW, MIN_TOLERANCE_CELL_NUMBER,
                WRAP_BOLD_ITALIC_STYLE,
                "Cận dưới /\n Min"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_LOCATION_AND_ALTITUDE_ROW, MAX_TOLERANCE_CELL_NUMBER,
                WRAP_BOLD_ITALIC_STYLE,
                "Cận trên /\n Max"
        )

        XSSFTools.styleMergeCells(
                sheet, MIN_ROW_NUMBER, MIN_ROW_NUMBER, INDEX_CELL_NUMBER, TIME_CELL_NUMBER, WRAP_BOLD_ITALIC_STYLE,
                "Giá trị cực tiểu /\nMinimum",
                HEADER_ROW_HEIGHT
        )
        XSSFTools.styleMergeCells(
                sheet, AVG_ROW_NUMBER, AVG_ROW_NUMBER, INDEX_CELL_NUMBER, TIME_CELL_NUMBER, WRAP_BOLD_ITALIC_STYLE,
                "Giá trị trung bình /\nAverage",
                HEADER_ROW_HEIGHT
        )
        XSSFTools.styleMergeCells(
                sheet, MAX_ROW_NUMBER, MAX_ROW_NUMBER, INDEX_CELL_NUMBER, TIME_CELL_NUMBER, WRAP_BOLD_ITALIC_STYLE,
                "Giá trị cực đại /\nMaximum",
                HEADER_ROW_HEIGHT
        )

        if (isTemp) {
            XSSFTools.styleMergeCells(
                    sheet, MKT_ROW_NUMBER, MKT_ROW_NUMBER, INDEX_CELL_NUMBER, TIME_CELL_NUMBER, WRAP_BOLD_ITALIC_STYLE,
                    "Nhiệt động học trung bình/\nMean Kinetic Temperature",
                    HEADER_ROW_HEIGHT
            )
        }
    }

    private void writeMinToleranceData(Sheet sheet, boolean isTemp = true) {
        BigDecimal value = isTemp ? inspection.getMinTemperatureTolerance() : inspection.getMinHumidityTolerance()

        ((isTemp ? DATA_START_TEMP_ROW_NUMBER : DATA_START_RH_ROW_NUMBER)..DATA_END_TEMP_ROW_NUMBER).each { Integer rowNumber ->
            XSSFTools.setCellValue(sheet, rowNumber, MIN_TOLERANCE_CELL_NUMBER, WRAP_NUMBER_STYLE, value)
        }
    }

    private void writeMaxToleranceData(Sheet sheet, boolean isTemp = true) {
        BigDecimal value = isTemp ? inspection.getMaxTemperatureTolerance() : inspection.getMaxHumidityTolerance()

        ((isTemp ? DATA_START_TEMP_ROW_NUMBER : DATA_START_RH_ROW_NUMBER)..DATA_END_TEMP_ROW_NUMBER).each { Integer rowNumber ->
            XSSFTools.setCellValue(sheet, rowNumber, MAX_TOLERANCE_CELL_NUMBER, WRAP_NUMBER_STYLE, value)
        }
    }

    private void writeDataLogMeta(Sheet sheet) {
        int currentCellNumber = DATA_START_CELL_NUMBER
        inspection.getLogTagSensorByLocation().each { String location, LogTagSensor logTag ->
            writeDataLogMeta(sheet, currentCellNumber, logTag)
            currentCellNumber++
        }
        writeDataLogMeta(sheet, currentCellNumber, inspection.getEnvironmentSensor())
    }

    private void writeDataLogMeta(Sheet sheet, int currentCellNumber, LogTagSensor logTag) {
        String serial = logTag.getSerial()
        cellNumberByLogTagSerial[ serial ] = currentCellNumber

        XSSFTools.setCellValue(sheet, HEADER_LOG_TAG_SERIAL_ROW, currentCellNumber, WRAP_BOLD_STYLE, serial)

        String location = logTag.getLocation()
        String altitude = logTag.getAltitudeLevel()
        String locationData = altitude ? "${location}_${altitude}" : location
        if (currentCellNumber == DATA_END_CELL_NUMBER) {
            XSSFTools.setCellValue(sheet, HEADER_LOCATION_AND_ALTITUDE_ROW, currentCellNumber, ORANGE_FILL_BOLD_STYLE, locationData)
        }
        else {
            XSSFTools.setCellValue(sheet, HEADER_LOCATION_AND_ALTITUDE_ROW, currentCellNumber, GREY_FILL_BOLD_STYLE, locationData)
        }

        if (hasAltitude) {
            XSSFTools.setCellValue(sheet, HEADER_ALTITUDE_DESC_ROW, currentCellNumber, WRAP_BOLD_STYLE, getAltitudeDesc(altitude))
        }
    }

    private void writeIndicators() {
        inspection.getLogTagSensorByLocation().each { String location, LogTagSensor logTag ->
            writeIndicators(temperatureSheet, logTag, true)
            writeIndicators(humiditySheet, logTag, false)
        }
        writeIndicators(temperatureSheet, inspection.getEnvironmentSensor(), true)
        writeIndicators(humiditySheet, inspection.getEnvironmentSensor(), false)
    }

    private void writeIndicators(Sheet sheet, LogTagSensor logTag, boolean isTemp) {
        Integer cellNum = getCellNumber(logTag)

        BigDecimal min = isTemp ? logTag.getMinTemperature() : logTag.getMinHumidity()
        BigDecimal avg = isTemp ? logTag.getAvgTemperature() : logTag.getAvgHumidity()
        BigDecimal max = isTemp ? logTag.getMaxTemperature() : logTag.getMaxHumidity()

        boolean isEnvCellNum = cellNum == DATA_END_CELL_NUMBER
        CellStyle maxStyle = isEnvCellNum
                ? ORANGE_FILL_NUMBER_BOLD_STYLE
                : getCellStyleForMaxMinValue(
                        inspection,
                        max,
                        RED_FILL_NUMBER_BOLD_STYLE,
                        GREEN_FILL_NUMBER_BOLD_STYLE,
                        ROSE_FILL_NUMBER_BOLD_STYLE,
                        isTemp
                )
        CellStyle minStyle = isEnvCellNum
                ? ORANGE_FILL_NUMBER_BOLD_STYLE
                : getCellStyleForMaxMinValue(
                        inspection,
                        min,
                        RED_FILL_NUMBER_BOLD_STYLE,
                        GREEN_FILL_NUMBER_BOLD_STYLE,
                        BLUE_FILL_NUMBER_BOLD_STYLE,
                        isTemp
                )
        CellStyle avgStyle = isEnvCellNum
                ? ORANGE_FILL_NUMBER_BOLD_STYLE
                : YELLOW_FILL_NUMBER_BOLD_STYLE

        XSSFTools.setCellValue(sheet, MIN_ROW_NUMBER, cellNum, minStyle, min)
        XSSFTools.setCellValue(sheet, AVG_ROW_NUMBER, cellNum, avgStyle, avg)
        XSSFTools.setCellValue(sheet, MAX_ROW_NUMBER, cellNum, maxStyle, max)
        if (isTemp) {
            CellStyle mktStyle = isEnvCellNum ? ORANGE_FILL_NUMBER_BOLD_STYLE : GREY_FILL_NUMBER_BOLD_STYLE
            XSSFTools.setCellValue(sheet, MKT_ROW_NUMBER, cellNum, mktStyle, logTag.getMeanKineticTemperature())
        }
    }

    private void writeData() {
        Integer envCellNum = getCellNumber(inspection.getEnvironmentSensor())

        inspection.getLogTagSensorByLocation().each { String location, LogTagSensor logTag ->
            Integer tempRowNum = DATA_START_TEMP_ROW_NUMBER
            Integer rhRowNum = DATA_START_RH_ROW_NUMBER
            Integer cellNum = getCellNumber(logTag)
            Integer idx = 1
            logTag.getLogTagReadingByDateTime().each { ZonedDateTime dateTime, LogTagReading reading ->
                LogTagReading envReading = inspection.getEnvironmentSensor().getReading(dateTime)
                writeData(temperatureSheet, reading, tempRowNum, cellNum, true, idx, dateTime)
                writeData(temperatureSheet, envReading, tempRowNum, envCellNum,  true, idx, dateTime)
                writeData(humiditySheet, reading, rhRowNum, cellNum,  false, idx, dateTime)
                writeData(humiditySheet, envReading, rhRowNum, envCellNum,  false, idx, dateTime)

                tempRowNum ++
                rhRowNum ++
                idx ++
            }
        }
    }

    private Integer getCellNumber(LogTagSensor sensor) {
        String serial = sensor.getSerial()
        return cellNumberByLogTagSerial?.getAt(serial)
    }

    private void writeData(Sheet sheet, LogTagReading reading,
                           Integer rowNum, Integer cellNum,
                           boolean isTemp, Integer idx,
                           ZonedDateTime dateTime) {
        BigDecimal value = isTemp ? reading.getTemperature() : reading.getHumidity()
        CellStyle valueStyle = getCellStyleForMaxMinValue(
                inspection, value,
                RED_FILL_NUMBER_STYLE,
                GREEN_FILL_NUMBER_STYLE,
                WRAP_NUMBER_STYLE,
                isTemp
        )

        XSSFTools.setCellValue(sheet, rowNum, cellNum, valueStyle, value)
        XSSFTools.setCellValue(sheet, rowNum, INDEX_CELL_NUMBER, WRAP_STYLE, idx)
        XSSFTools.setCellValue(sheet, rowNum, TIME_CELL_NUMBER, WRAP_STYLE, DateTimeResolver.getDateTimeString(dateTime, "yyyy-MM-dd HH:mm"))
    }

    private void createCharts() {
        createLineChart(temperatureChartSheet)
        createLineChart(humidityChartSheet, false)
    }

    private void createLineChart(XSSFSheet sheet, boolean isTemp = true) {
        Sheet sourceSheet = isTemp ? temperatureSheet : humiditySheet

        XSSFDrawing drawing = sheet.createDrawingPatriarch()
        XSSFAnchor anchor = drawing.createAnchor(
                0, 0,
                0, 0,
                CHART_START_CELL_NUMBER, CHART_START_ROW_NUMBER,
                CHART_END_CELL_NUMBER + 1, CHART_END_ROW_NUMBER
        )

        XSSFChart chart = drawing.createChart(anchor)

        XDDFChartLegend legend = chart.getOrAddLegend()
        legend.setPosition(LegendPosition.RIGHT)

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM)
        bottomAxis.setTitle("Time")
        bottomAxis.setCrosses(AxisCrosses.MIN)

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT)
        leftAxis.setTitle(isTemp ? "Temp °C" : "Hum %RH")
        leftAxis.setMinimum(findMinimumRoundedValueForLeftAxis(isTemp))
        leftAxis.setMaximum(findMaximumRoundedValueForLeftAxis(isTemp))
        leftAxis.setMinorUnit(0.1)
        leftAxis.setMajorUnit(isTemp ? 0.5 : 2.5)
        leftAxis.setCrosses(AxisCrosses.MIN)

        XDDFDataSource timeSource = XDDFDataSourcesFactory.fromStringCellRange(sourceSheet, new CellRangeAddress(
                (isTemp ? DATA_START_TEMP_ROW_NUMBER : DATA_START_RH_ROW_NUMBER), DATA_END_TEMP_ROW_NUMBER, TIME_CELL_NUMBER, TIME_CELL_NUMBER
        ))

        Row headerRow = sourceSheet.getRow(HEADER_LOCATION_AND_ALTITUDE_ROW)
        XDDFChartData data = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis)
        (DATA_START_CELL_NUMBER..DATA_END_CELL_NUMBER).each { int cellNumber ->
            XDDFNumericalDataSource valueSource = XDDFDataSourcesFactory.fromNumericCellRange(sourceSheet, new CellRangeAddress(
                    (isTemp ? DATA_START_TEMP_ROW_NUMBER : DATA_START_RH_ROW_NUMBER), DATA_END_TEMP_ROW_NUMBER, cellNumber, cellNumber
            ))

            XDDFChartData.Series series = data.addSeries(timeSource, valueSource)
            if (DATA_END_CELL_NUMBER == cellNumber) {
                series.setMarkerStyle(MarkerStyle.X)
            }
            else {
                series.setMarkerStyle(MarkerStyle.NONE)
            }

            Cell headerCell = headerRow.getCell(cellNumber)
            String seriesName = headerCell.getStringCellValue()
            if (seriesName.contains("ENV")) {
                seriesName += " /\nMôi trường"
            }
            series.setTitle(seriesName)
        }

        [MIN_TOLERANCE_CELL_NUMBER, MAX_TOLERANCE_CELL_NUMBER ].each { int cellNumber ->
            XDDFNumericalDataSource valueSource = XDDFDataSourcesFactory.fromNumericCellRange(sourceSheet, new CellRangeAddress(
                    (isTemp ? DATA_START_TEMP_ROW_NUMBER : DATA_START_RH_ROW_NUMBER), DATA_END_TEMP_ROW_NUMBER, cellNumber, cellNumber
            ))

            XDDFChartData.Series series = data.addSeries(timeSource, valueSource)
            series.setMarkerStyle(MarkerStyle.NONE)

            Cell headerCell = headerRow.getCell(cellNumber)
            String seriesName = headerCell.getStringCellValue()
            series.setTitle(seriesName)
        }

        data.setVaryColors(true)

        chart.plot(data)
    }

    private BigDecimal findMinimumRoundedValueForLeftAxis(boolean isTemperature) {
        BigDecimal value = isTemperature ? ALL_TIME_MIN_TEMPERATURE : ALL_TIME_MIN_HUMIDITY
        BigDecimal tolerance = isTemperature ? inspection.getMinTemperatureTolerance() : inspection.getMinHumidityTolerance()
        BigDecimal env = isTemperature ? inspection.getEnvironmentSensor().getMinTemperature() : inspection.getEnvironmentSensor().getMinHumidity()

        BigDecimal min = [ value, tolerance, env ].min()

        return roundToPointZeroFive(min - (isTemperature ? ALL_TIME_OFF_SET_TEMP : ALL_TIME_OFF_SET_HUM))
    }

    private BigDecimal findMaximumRoundedValueForLeftAxis(boolean isTemperature) {
        BigDecimal value = isTemperature ? ALL_TIME_MAX_TEMPERATURE : ALL_TIME_MAX_HUMIDITY
        BigDecimal tolerance = isTemperature ? inspection.getMaxTemperatureTolerance() : inspection.getMaxHumidityTolerance()
        BigDecimal env = isTemperature ? inspection.getEnvironmentSensor().getMaxTemperature() : inspection.getEnvironmentSensor().getMaxHumidity()

        BigDecimal max = [ value, tolerance, env ].max()

        return roundToPointZeroFive(max + (isTemperature ? ALL_TIME_OFF_SET_TEMP : ALL_TIME_OFF_SET_HUM))
    }

    private BigDecimal roundToPointZeroFive(BigDecimal value) {
        if (!value) {
            return value
        }
        if (value % 0.5 == 0.0) {
            return value
        }

        value = value * 2
        value = value.setScale(0, RoundingMode.HALF_UP)

        return (value / 2).setScale(2, RoundingMode.HALF_UP)
    }
}
