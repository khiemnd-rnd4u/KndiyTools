package sensorService

import kndiyLibraries.XSSFTools
import cache.Cache
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xddf.usermodel.chart.AxisPosition
import org.apache.poi.xddf.usermodel.chart.ChartTypes
import org.apache.poi.xddf.usermodel.chart.LegendPosition
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
    private int HEADER_ROW_NUMBER = 34
    private int MIN_ROW_NUMBER = 32
    private int MAX_ROW_NUMBER = 33
    private int DATA_START_ROW_NUMBER = 35
    private int DATA_END_ROW_NUMBER = 35

    private int CHART_START_CELL_NUMBER = 0
    private int CHART_END_CELL_NUMBER = 12

    private int INDEX_CELL_NUMBER = 0
    private int DATE_CELL_NUMBER = 1
    private int TIME_CELL_NUMBER = 2

    private int DATA_START_CELL_NUMBER = 3
    private int DATA_END_CELL_NUMBER = 3

    private SensorDataParser sensorDataParser
    private DateTimeFormatter dateTimeFormatter
    private CreationHelper creationHelper
    private CellStyle dateStyle
    private CellStyle timeStyle
    private CellStyle numberStyle
    private CellStyle boldTitleStyle

    SensorFolderParser(String folderPath, String savePathWithName,
                       int columnSizeWithoutEvent,
                       int yearIdx, int monthIdx, int dayIdx,
                       int tempIdx, int humidIdx, int minuteInterval) {
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
            writeMinOrMaxHeader(temperatureSheet)
            writeMinOrMaxHeader(humiditySheet)

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
    }

    private void writeMinOrMaxHeader(Sheet sheet) {
        Row minRow = XSSFTools.getOrSetRow(sheet, MIN_ROW_NUMBER)
        Row maxRow = XSSFTools.getOrSetRow(sheet, MAX_ROW_NUMBER)

        Cell minCell = XSSFTools.getOrSetCell(minRow, TIME_CELL_NUMBER)
        Cell maxCell = XSSFTools.getOrSetCell(maxRow, TIME_CELL_NUMBER)

        minCell.setCellValue("Min")
        maxCell.setCellValue("Max")
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
        chart.setTitleText(isTemperature ? "Temperature Chart" : "Humidity Chart")
        chart.setTitleOverlay(false)

        XDDFChartLegend legend = chart.getOrAddLegend()
        legend.setPosition(LegendPosition.BOTTOM)

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM)
        bottomAxis.setTitle("Time")

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT)
        leftAxis.setTitle(isTemperature ? "Temp °C" : "Hum %RH")

        XDDFDataSource timeSource = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(
                DATA_START_ROW_NUMBER, DATA_END_ROW_NUMBER, TIME_CELL_NUMBER, TIME_CELL_NUMBER
        ))

        Row headerRow = sheet.getRow(HEADER_ROW_NUMBER)
        XDDFChartData data = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis)
        (DATA_START_CELL_NUMBER..(DATA_END_CELL_NUMBER - 1)).each { int columnNumber ->
            XDDFNumericalDataSource valueSource = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(
                    DATA_START_ROW_NUMBER, DATA_END_ROW_NUMBER, columnNumber, columnNumber
            ))

            XDDFChartData.Series series = data.addSeries(timeSource, valueSource)

            Cell headerCell = headerRow.getCell(columnNumber)
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
        writeTitle(temperatureSheet, cellNumber, serial)
        writeMinOrMax(temperatureSheet, MAX_ROW_NUMBER, cellNumber, data, SensorDataParser.TEMPERATURE)
        writeMinOrMax(temperatureSheet, MIN_ROW_NUMBER, cellNumber, data, SensorDataParser.TEMPERATURE, false)

        writeIndex(humiditySheet, currentDataRow)
        writeDate(humiditySheet, currentDataRow, dateTimeCompos[0])
        writeTime(humiditySheet, currentDataRow, dateTimeCompos[1])
        writeTitle(humiditySheet, cellNumber, serial)
        writeMinOrMax(humiditySheet, MAX_ROW_NUMBER, cellNumber, data, SensorDataParser.HUMIDITY)
        writeMinOrMax(humiditySheet, MIN_ROW_NUMBER, cellNumber, data, SensorDataParser.HUMIDITY, false)
    }

    private void writeMinOrMax(Sheet sheet, int rowNumber, int cellNumber, Map data, String key, boolean isMax = true) {
        Row row = XSSFTools.getOrSetRow(sheet, rowNumber)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, boldTitleStyle)

        BigDecimal value = data[ key ] as BigDecimal
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

        cell.setCellValue(newValue)
    }

    private void writeTitle(Sheet sheet, int cellNumber, String serial) {
        Row row = XSSFTools.getOrSetRow(sheet, HEADER_ROW_NUMBER)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, boldTitleStyle)
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
