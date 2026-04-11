package sensorService

import kndiyLibraries.XSSFTools
import cache.Cache
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.joda.time.DateTime
import org.joda.time.format.DateTimeFormat
import org.joda.time.format.DateTimeFormatter

import java.math.RoundingMode

class SensorFolderParser {
    private String folderPath
    private String savePathWithName
    private TreeMap dataBySerialByDateTime
    private XSSFWorkbook workbook
    private String TEMPERATURE_SHEET_NAME = "Temperature"
    private String HUMIDITY_SHEET_NAME = "Humidity"
    private XSSFSheet temperatureSheet
    private XSSFSheet humiditySheet
    private int CHART_START_ROW_NUMBER = 2
    private int CHART_END_ROW_NUMBER = 30
    private int AVG_ROW_NUMBER = 32
    private int MIN_ROW_NUMBER = 33
    private int MAX_ROW_NUMBER = 34
    private int HEADER_ROW_NUMBER = 35
    private int DATA_START_ROW_NUMBER = 36
    private int DATA_END_ROW_NUMBER = 36

    private int CHART_START_CELL_NUMBER = 0
    private int CHART_END_CELL_NUMBER = 14

    private int INDEX_CELL_NUMBER = 0
    private int DATE_CELL_NUMBER = 1
    private int TIME_CELL_NUMBER = 2

    private int DATA_START_CELL_NUMBER = 3
    private int DATA_END_CELL_NUMBER = 3
    private Integer DATA_MIN_CELL_NUMBER = null
    private Integer DATA_MAX_CELL_NUMBER = null

    private BigDecimal ALL_TIME_MIN_TEMP = 9999999
    private BigDecimal ALL_TIME_MAX_TEMP = 0
    private BigDecimal ALL_TIME_MIN_HUM = 9999999
    private BigDecimal ALL_TIME_MAX_HUM = 0
    private BigDecimal ALL_TIME_OFF_SET_TEMP = 0.5
    private BigDecimal ALL_TIME_OFF_SET_HUM = 1.5

    private SensorDataParser sensorDataParser
    private DateTimeFormatter dateTimeFormatter
    private CreationHelper creationHelper
    private CellStyle dateStyle
    private CellStyle timeStyle
    private CellStyle numberStyle
    private CellStyle boldTitleStyle
    private CellStyle WRAP_TEXT
    private CellStyle WRAP_TEXT_BOLD

    private BigDecimal MAX_TEMPERATURE
    private BigDecimal MIN_TEMPERATURE
    private BigDecimal MAX_HUMIDITY
    private BigDecimal MIN_HUMIDITY

    SensorFolderParser(String folderPath, String savePathWithName,
                       int columnSizeWithoutEvent,
                       int yearIdx, int monthIdx, int dayIdx,
                       int tempIdx, int humidIdx, int minuteInterval,
                       BigDecimal minTemp = 15.0,
                       BigDecimal maxTemp = 25.0,
                       BigDecimal minHum = 0.0,
                       BigDecimal maxHum = 75.0) {
        this.folderPath = folderPath
        this.savePathWithName = savePathWithName
        Cache.initiateCache()

        workbook = new XSSFWorkbook()
        creationHelper = workbook.getCreationHelper()
        dateStyle = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, "yyyy-MM-dd")
        timeStyle = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, "HH:mm:ss")
        numberStyle = XSSFTools.getCellStyleWithFormat(workbook, creationHelper, "#.00")

        boldTitleStyle = workbook.createCellStyle()
        boldTitleStyle.setFont(XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT))

        WRAP_TEXT = workbook.createCellStyle()
        WRAP_TEXT.setWrapText(true)

        WRAP_TEXT_BOLD = workbook.createCellStyle()
        WRAP_TEXT_BOLD.setWrapText(true)
        WRAP_TEXT_BOLD.setFont(XSSFTools.getFont(workbook, XSSFTools.BOLD_FONT))
        
        MAX_TEMPERATURE = maxTemp.setScale(2, RoundingMode.HALF_UP)
        MIN_TEMPERATURE = minTemp.setScale(2, RoundingMode.HALF_UP)
        MAX_HUMIDITY = maxHum.setScale(2, RoundingMode.HALF_UP)
        MIN_HUMIDITY = minHum.setScale(2, RoundingMode.HALF_UP)

        dateTimeFormatter = DateTimeFormat.forPattern("yyyy-MM-dd HH:mm:ss")
        sensorDataParser = new SensorDataParser(columnSizeWithoutEvent, yearIdx, monthIdx, dayIdx, tempIdx, humidIdx, minuteInterval)
        dataBySerialByDateTime = new TreeMap((DateTime d1, DateTime d2) -> {

            return d1 <=> d2
        })
    }

    void parseFilesInFolder() {
        File folder = new File(folderPath)
        List files = folder.listFiles()
        for (String filePath : files) {
            if (!filePath.endsWith(".pdf")) {
                continue
            }

            println("Started Parsing ${filePath}")
            sensorDataParser.parseFileAndUpdateDataBySerialByDateTime(
                    dataBySerialByDateTime,
                    filePath
            )
            println("Done.")
        }

        println("Started Writing to ${savePathWithName} Workbook")
        writeIntoWorkBook()
        createLineChart()

        saveWorkBookAsFile()
        println("Finished!")
    }

    private void writeIntoWorkBook() {
        temperatureSheet = workbook.createSheet(TEMPERATURE_SHEET_NAME)
        humiditySheet = workbook.createSheet(HUMIDITY_SHEET_NAME)

        int currentDataRow = DATA_START_ROW_NUMBER
        for (DateTime dateTime : dataBySerialByDateTime?.keySet()) {
            TreeMap dataBySerial = dataBySerialByDateTime?.getAt(dateTime)
            writeMinMaxAvgHeader(temperatureSheet)
            writeMinMaxAvgHeader(humiditySheet)

            for (String serial : dataBySerial?.keySet()) {
                Map data = dataBySerial?.getAt(serial)
                int cellNumber = getCellNumber(serial)
                DATA_END_CELL_NUMBER = [ DATA_END_CELL_NUMBER, cellNumber ].max()

                writeHeader(serial, dateTime, currentDataRow, cellNumber, data)
                writeData(data, currentDataRow, temperatureSheet, SensorDataParser.TEMPERATURE, cellNumber)
                writeData(data, currentDataRow, humiditySheet, SensorDataParser.HUMIDITY, cellNumber)
            }
            currentDataRow ++
        }
        DATA_END_ROW_NUMBER = currentDataRow - 1
        DATA_MIN_CELL_NUMBER = DATA_END_CELL_NUMBER + 1
        DATA_MAX_CELL_NUMBER = DATA_END_CELL_NUMBER + 2

        resolveAllTimeMaxMinData()

        writeMinMaxAvgHeader(temperatureSheet, true)
        writeMinMaxAvgHeader(humiditySheet, true)

        writeMinData(temperatureSheet)
        writeMinData(humiditySheet, false)
        writeMaxData(temperatureSheet)
        writeMaxData(humiditySheet, false)
    }

    private void resolveAllTimeMaxMinData() {
        ALL_TIME_MIN_TEMP = [ ALL_TIME_MIN_TEMP, MIN_TEMPERATURE ].min()
        ALL_TIME_MAX_TEMP = [ ALL_TIME_MAX_TEMP, MAX_TEMPERATURE ].max()
        ALL_TIME_MIN_HUM = [ ALL_TIME_MIN_HUM, MIN_HUMIDITY ].min()
        ALL_TIME_MAX_HUM = [ ALL_TIME_MAX_HUM, MAX_HUMIDITY ].max()

        ALL_TIME_MIN_TEMP = roundToPointZeroFive(ALL_TIME_MIN_TEMP - ALL_TIME_OFF_SET_TEMP)
        ALL_TIME_MAX_TEMP = roundToPointZeroFive(ALL_TIME_MAX_TEMP + ALL_TIME_OFF_SET_TEMP)
        ALL_TIME_MIN_HUM = roundToPointZeroFive(ALL_TIME_MIN_HUM - ALL_TIME_OFF_SET_HUM)
        ALL_TIME_MAX_HUM = roundToPointZeroFive(ALL_TIME_MAX_HUM + ALL_TIME_OFF_SET_HUM)

        println("All Time Data ${[ ALL_TIME_MIN_TEMP, ALL_TIME_MAX_TEMP, ALL_TIME_MIN_HUM, ALL_TIME_MAX_HUM ]}")
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
        println("After x2 ${value}")

        return (value / 2).setScale(2, RoundingMode.HALF_UP)
    }

    private void writeMinData(Sheet sheet, boolean isTemp = true) {
        BigDecimal value = isTemp ? MIN_TEMPERATURE : MIN_HUMIDITY

        (DATA_START_ROW_NUMBER..DATA_END_ROW_NUMBER).each { Integer rowNumber ->
            Row row = XSSFTools.getOrSetRow(sheet, rowNumber)
            Cell cell = XSSFTools.getOrSetCell(row, DATA_MIN_CELL_NUMBER)

            cell.setCellValue(value)
        }
    }

    private void writeMaxData(Sheet sheet, boolean isTemp = true) {
        BigDecimal value = isTemp ? MAX_TEMPERATURE : MAX_HUMIDITY

        (DATA_START_ROW_NUMBER..DATA_END_ROW_NUMBER).each { Integer rowNumber ->
            Row row = XSSFTools.getOrSetRow(sheet, rowNumber)
            Cell cell = XSSFTools.getOrSetCell(row, DATA_MAX_CELL_NUMBER)

            cell.setCellValue(value)
        }
    }

    private void writeMinMaxAvgHeader(Sheet sheet, boolean isForData = false) {
        Row minRow = XSSFTools.getOrSetRow(sheet, isForData ? HEADER_ROW_NUMBER : MIN_ROW_NUMBER)
        Row maxRow = XSSFTools.getOrSetRow(sheet, isForData ? HEADER_ROW_NUMBER : MAX_ROW_NUMBER)
        Row avgRow = XSSFTools.getOrSetRow(sheet, AVG_ROW_NUMBER)

        Cell minCell = XSSFTools.getOrSetCell(minRow, isForData ? DATA_MIN_CELL_NUMBER : TIME_CELL_NUMBER, boldTitleStyle)
        Cell maxCell = XSSFTools.getOrSetCell(maxRow, isForData ? DATA_MAX_CELL_NUMBER : TIME_CELL_NUMBER, boldTitleStyle)
        Cell avgCell = XSSFTools.getOrSetCell(avgRow, TIME_CELL_NUMBER, boldTitleStyle)

        minCell.setCellValue("Min")
        maxCell.setCellValue("Max")
        avgCell.setCellValue("Average")
    }

    private void createLineChart() {
        createLineChart(temperatureSheet)
        createLineChart(humiditySheet, false)
    }

    private void createLineChart(XSSFSheet sheet, boolean isTemperature = true) {
        XSSFDrawing drawing = sheet.createDrawingPatriarch()
        XSSFAnchor anchor = drawing.createAnchor(
                0, 0,
                0, 0,
                CHART_START_CELL_NUMBER, CHART_START_ROW_NUMBER,
                CHART_END_CELL_NUMBER, CHART_END_ROW_NUMBER
        )

        XSSFChart chart = drawing.createChart(anchor)
        String title = isTemperature ? "Temperature" : "Humidity"
        chart.setTitleText("${title} Chart")
        chart.setTitleOverlay(false)

        XDDFChartLegend legend = chart.getOrAddLegend()
        legend.setPosition(LegendPosition.BOTTOM)

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM)
        bottomAxis.setTitle("Time")

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT)
        leftAxis.setTitle(isTemperature ? "Temp °C" : "Hum %RH")
        leftAxis.setMaximum(isTemperature ? ALL_TIME_MAX_TEMP : ALL_TIME_MAX_HUM)
        leftAxis.setMinimum(isTemperature ? ALL_TIME_MIN_TEMP : ALL_TIME_MIN_HUM)
        leftAxis.setMinorUnit(0.1)
        leftAxis.setMajorUnit(isTemperature ? 0.5 : 2.5)

        XDDFDataSource timeSource = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(
                DATA_START_ROW_NUMBER, DATA_END_ROW_NUMBER, TIME_CELL_NUMBER, TIME_CELL_NUMBER
        ))

        Row headerRow = sheet.getRow(HEADER_ROW_NUMBER)
        XDDFChartData data = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis)
        (DATA_START_CELL_NUMBER..(DATA_END_CELL_NUMBER - 1)).each { int cellNumber ->
            XDDFNumericalDataSource valueSource = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(
                    DATA_START_ROW_NUMBER, DATA_END_ROW_NUMBER, cellNumber, cellNumber
            ))

            XDDFChartData.Series series = data.addSeries(timeSource, valueSource)
            series.setMarkerStyle(MarkerStyle.NONE)

            Cell headerCell = headerRow.getCell(cellNumber)
            String seriesName = headerCell.getStringCellValue()
            series.setTitle(seriesName)
        }

        [ DATA_MIN_CELL_NUMBER, DATA_MAX_CELL_NUMBER ].each { int cellNumber ->
            XDDFNumericalDataSource valueSource = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(
                    DATA_START_ROW_NUMBER, DATA_END_ROW_NUMBER, cellNumber, cellNumber
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

    private int getCellNumber(String serial) {
        return sensorDataParser.getSerials()?.indexOf(serial) + DATA_START_CELL_NUMBER
    }

    private void writeHeader(String serial, DateTime dateTime, int currentDataRow, int cellNumber, Map data) {
        String dateTimeString = dateTime.toString(dateTimeFormatter)
        List<String> dateTimeCompos = dateTimeString?.split(" ")

        writeIndex(temperatureSheet, currentDataRow)
        writeDate(temperatureSheet, currentDataRow, dateTimeCompos[0])
        writeTime(temperatureSheet, currentDataRow, dateTimeCompos[1])
        writeTableHeader(temperatureSheet, cellNumber, serial)
        writeMinOrMax(temperatureSheet, MAX_ROW_NUMBER, cellNumber, data, true)
        writeMinOrMax(temperatureSheet, MIN_ROW_NUMBER, cellNumber, data, true, false)
        writeAvg(temperatureSheet, cellNumber, data, true)

        writeIndex(humiditySheet, currentDataRow)
        writeDate(humiditySheet, currentDataRow, dateTimeCompos[0])
        writeTime(humiditySheet, currentDataRow, dateTimeCompos[1])
        writeTableHeader(humiditySheet, cellNumber, serial)
        writeMinOrMax(humiditySheet, MAX_ROW_NUMBER, cellNumber, data, false)
        writeMinOrMax(humiditySheet, MIN_ROW_NUMBER, cellNumber, data, false, false)
        writeAvg(temperatureSheet, cellNumber, data, false)
    }

    private void writeAvg(Sheet sheet, int cellNumber, Map data, boolean isTemp) {
        Row row = XSSFTools.getOrSetRow(sheet, AVG_ROW_NUMBER)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, boldTitleStyle)

        BigDecimal value = data[ isTemp ? SensorDataParser.TEMPERATURE : SensorDataParser.HUMIDITY ] as BigDecimal
        BigDecimal currentAvg
        try {
            currentAvg = cell.getNumericCellValue()
        }
        catch (ignored) {
            currentAvg = value
        }
        BigDecimal newAvg = (value + currentAvg) / 2.0

        cell.setCellValue(newAvg)
    }

    private void writeMinOrMax(Sheet sheet, int rowNumber, int cellNumber, Map data, boolean isTemp, boolean isMax = true) {
        Row row = XSSFTools.getOrSetRow(sheet, rowNumber)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, boldTitleStyle)

        BigDecimal value = data[ isTemp ? SensorDataParser.TEMPERATURE : SensorDataParser.HUMIDITY ] as BigDecimal
        BigDecimal cellValue = null
        try {
            cellValue = cell.getNumericCellValue() ?: value
        }
        catch (ex) {
        }
        if (!cellValue) {
            cellValue = value
        }

        BigDecimal newValue = isMax
                ? [ value, cellValue ].max()
                : [ value, cellValue ].min()

        if (isTemp && cellNumber != DATA_END_CELL_NUMBER) {
            if (isMax) {
                ALL_TIME_MAX_TEMP = [ ALL_TIME_MAX_TEMP, newValue ].max()
            }
            else {
                ALL_TIME_MIN_TEMP = [ ALL_TIME_MIN_TEMP, newValue ].min()
            }
        }
        else if (cellNumber != DATA_END_CELL_NUMBER) {
            if (isMax) {
                ALL_TIME_MAX_HUM = [ ALL_TIME_MAX_HUM, newValue ].max()
            }
            else {
                ALL_TIME_MIN_HUM = [ ALL_TIME_MIN_HUM, newValue ].min()
            }
        }

        cell.setCellValue(newValue)
    }

    private void writeTableHeader(Sheet sheet, int cellNumber, String serial) {
        Row row = XSSFTools.getOrSetRow(sheet, HEADER_ROW_NUMBER)
        row.setHeightInPoints(62)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, boldTitleStyle)

        cell.setCellStyle(WRAP_TEXT_BOLD)
        cell.setCellValue(serial)
    }

    private void writeIndex(Sheet sheet, int rowNum) {
        Row row = XSSFTools.getOrSetRow(sheet, rowNum)
        Cell cell = XSSFTools.getOrSetCell(row, INDEX_CELL_NUMBER)
        cell.setCellValue(rowNum - HEADER_ROW_NUMBER)
    }

    private void writeDate(Sheet sheet, int rowNum, String date) {
        Row row = XSSFTools.getOrSetRow(sheet, rowNum)
        Cell cell = XSSFTools.getOrSetCell(row, DATE_CELL_NUMBER, dateStyle)
        cell.setCellValue(date)
    }

    private void writeTime(Sheet sheet, int rowNum, String time) {
        Row row = XSSFTools.getOrSetRow(sheet, rowNum)
        Cell cell = XSSFTools.getOrSetCell(row, TIME_CELL_NUMBER, timeStyle)
        cell.setCellValue(time)
    }

    private void writeData(Map dataByKey, int currentDataRow, Sheet sheet, String key, int cellNumber) {
        Row row = XSSFTools.getOrSetRow(sheet, currentDataRow)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, numberStyle, CellType.NUMERIC)
        BigDecimal dataValue = dataByKey[key] as BigDecimal
        if (dataValue) {
            cell.setCellValue(dataValue)
        }
    }

    private void saveWorkBookAsFile() {
        try (FileOutputStream outputStream = new FileOutputStream(new File("${savePathWithName}"))) {
            workbook.write(outputStream)
        }
        catch (Exception ex) {
            println(ex.getMessage())
            ex.printStackTrace()
        }
    }

    TreeMap resetDataBySerialByDateTime() {
        dataBySerialByDateTime = new TreeMap()
    }

    TreeMap getDataBySerialByDateTime() {
        return dataBySerialByDateTime
    }
}
