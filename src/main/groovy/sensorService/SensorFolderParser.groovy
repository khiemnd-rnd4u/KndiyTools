package sensorService

import kndiyLibraries.XSSFTools
import cache.Cache
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.joda.time.DateTime
import org.joda.time.format.DateTimeFormat
import org.joda.time.format.DateTimeFormatter

class SensorFolderParser {
    private String folderPath
    private String savePathWithName
    private TreeMap dataBySerialByDateTime
    private Workbook workbook
    private String TEMPERATURE_SHEET_NAME = "Temperature"
    private String HUMIDITY_SHEET_NAME = "Humidity"
    private Sheet temperatureSheet
    private Sheet humiditySheet
    private int TITLE_ROW_NUMBER = 0
    private int INDEX_CELL_NUMBER = 0
    private int DATE_CELL_NUMBER = 1
    private int TIME_CELL_NUMBER = 2
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
        dataBySerialByDateTime = new TreeMap()
    }

    void parseFilesInFolder() {
        File folder = new File(folderPath)
        List files = folder.listFiles()
        for (String filePath : files) {
            println("Started Parsing ${filePath}")
            sensorDataParser.parseFileAndUpdateDataBySerialByDateTime(
                    dataBySerialByDateTime,
                    filePath
            )
            println("Done.")
        }

        println("Started Writing to ${savePathWithName} Workbook")
        writeIntoWorkBook()
        saveWorkBookAsFile()
        println("Finished!")
    }

    private void writeIntoWorkBook() {
        temperatureSheet = workbook.createSheet(TEMPERATURE_SHEET_NAME)
        humiditySheet = workbook.createSheet(HUMIDITY_SHEET_NAME)

        int currentDataRow = TITLE_ROW_NUMBER + 1
        for (DateTime dateTime : dataBySerialByDateTime?.keySet()) {
            TreeMap dataBySerial = dataBySerialByDateTime?.getAt(dateTime)
            for (String serial : dataBySerial?.keySet()) {
                Map data = dataBySerial?.getAt(serial)
                int cellNumber = getCellNumber(serial)
                writeHeader(serial, dateTime, currentDataRow, cellNumber)
                writeData(data, currentDataRow, temperatureSheet, SensorDataParser.TEMPERATURE, cellNumber)
                writeData(data, currentDataRow, humiditySheet, SensorDataParser.HUMIDITY, cellNumber)
            }
            currentDataRow ++
        }
    }

    private int getCellNumber(String serial) {
        return sensorDataParser.getSerials()?.indexOf(serial) + 3
    }

    private void writeHeader(String serial, DateTime dateTime, int currentDataRow, int cellNumber) {
        String dateTimeString = dateTime.toString(dateTimeFormatter)
        List<String> dateTimeCompos = dateTimeString?.split(" ")

        writeIndex(temperatureSheet, currentDataRow)
        writeDate(temperatureSheet, currentDataRow, dateTimeCompos[0])
        writeTime(temperatureSheet, currentDataRow, dateTimeCompos[1])
        writeTitle(temperatureSheet, cellNumber, serial)

        writeIndex(humiditySheet, currentDataRow)
        writeDate(humiditySheet, currentDataRow, dateTimeCompos[0])
        writeTime(humiditySheet, currentDataRow, dateTimeCompos[1])
        writeTitle(humiditySheet, cellNumber, serial)
    }

    private void writeTitle(Sheet sheet, int cellNumber, String serial) {
        Row row = XSSFTools.getOrSetRow(sheet, TITLE_ROW_NUMBER)
        Cell cell = XSSFTools.getOrSetCell(row, cellNumber, boldTitleStyle)
        cell.setCellValue(serial)
    }

    private void writeIndex(Sheet sheet, int rowNum) {
        Row row = XSSFTools.getOrSetRow(sheet, rowNum)
        Cell cell = XSSFTools.getOrSetCell(row, INDEX_CELL_NUMBER)
        cell.setCellValue(rowNum - TITLE_ROW_NUMBER)
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
