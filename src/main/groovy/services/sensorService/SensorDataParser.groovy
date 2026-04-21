package services.sensorService

import services.kndiyLibraries.DataStructure
import services.kndiyLibraries.DateTimeResolver

import cache.Cache
import com.itextpdf.kernel.pdf.PdfDocument
import com.itextpdf.kernel.pdf.PdfPage
import com.itextpdf.kernel.pdf.PdfReader
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor
import org.codehaus.groovy.runtime.typehandling.GroovyCastException

import java.time.ZonedDateTime
import java.util.zip.DataFormatException

class SensorDataParser {
    static final String TEMPERATURE = "temperature"
    static final String HUMIDITY = "humidity"
    static final String EVENT = "event"
    private int columnSizeWithoutEvent
    private int yearIdx
    private int monthIdx
    private int dayIdx
    private int temperatureIdx
    private int humidityIdx
    private int minuteInterval
    private Set serialSet
    List serials
    private String serialAndDesc
    boolean currentLineLeftSideError = false
    boolean currentLineRightSideError = false
    int leftSizeXEventIdx

    SensorDataParser(int columnSizeWithoutEvent, int yearIdx, int monthIdx, int dayIdx,
                     int temperatureIdx, int humidityIdx,
                     int minuteInterval) {
        this.serialSet = [ ]
        this.columnSizeWithoutEvent = columnSizeWithoutEvent
        this.yearIdx = yearIdx
        this.monthIdx = monthIdx
        this.dayIdx = dayIdx
        this.temperatureIdx = temperatureIdx
        this.humidityIdx = humidityIdx
        this.minuteInterval = minuteInterval
    }

    TreeMap parseFileAndUpdateDataByTypeByDateTime(TreeMap dataByTypeByDateTime,
                                                   File file) {
        try (PdfReader reader = new PdfReader(file.getAbsolutePath())) {
            try (PdfDocument document = new PdfDocument(reader)) {
                parseDocumentAndUpdateMap(
                        dataByTypeByDateTime,
                        document
                )
            }
        }

        return dataByTypeByDateTime
    }

    private TreeMap parseDocumentAndUpdateMap(TreeMap dataByTypeByDateTime,
                                              PdfDocument document) {
        int numOfPage = document.getNumberOfPages()
        for (int i = 2; i <= numOfPage; i++) {
            PdfPage page = document.getPage(i)
            String text = PdfTextExtractor.getTextFromPage(page)

            parseTextAndUpdatedataByTypeByDateTime(dataByTypeByDateTime, text)
        }

        return dataByTypeByDateTime
    }

    private TreeMap parseTextAndUpdatedataByTypeByDateTime(TreeMap dataByTypeByDateTime,
                                                             String text) {
        List lines = text.split("\n")
        for (String line : lines) {
            line = fixSpacing(line)
            println(line)

            currentLineLeftSideError = false
            currentLineRightSideError = false

            List words = line?.split(" ")
            if (!isDataLine(words)) {
                continue
            }
            println(words)
            try {
                parseWordsAndUpdateDataByTypeByDateTime(dataByTypeByDateTime, words)
            }
            catch (DataFormatException ignored) {
            }
        }

        return dataByTypeByDateTime
    }

    String fixSpacing(String inputString) {
        inputString = inputString.replace("         ", " ")
        inputString = inputString.replace("        ", " ")
        inputString = inputString.replace("       ", " ")
        inputString = inputString.replace("      ", " ")
        inputString = inputString.replace("     ", " ")
        inputString = inputString.replace("    ", " ")
        inputString = inputString.replace("   ", " ")
        inputString = inputString.replace("  ", " ")

        return inputString
    }

    boolean isDataLine(List words) {
        if (words.size() !in [5, 6, 7, 10, 11, 12, 13, 14]) {
            return false
        }
        try {
            BigDecimal test = words[0] as BigDecimal
            return true
        }
        catch (Exception ignored) {
            return false
        }
    }

    private TreeMap parseWordsAndUpdateDataByTypeByDateTime(TreeMap dataByDateTime, List words) {
        List leftSideWords = getLeftSideData(words)
        List rightSideWords = getRightSideData(words)

        println(leftSideWords)
        println(rightSideWords)

        addData(dataByDateTime, leftSideWords)
        addData(dataByDateTime, rightSideWords)

        return dataByDateTime
    }

    private List getLeftSideData(List words) {
        int xPosition = words.indexOf("X")
        if (xPosition < words.size() - 1 && xPosition > 0) {
            currentLineLeftSideError = true
            leftSizeXEventIdx = xPosition

            return words.subList(0, xPosition + 1)
        }

        int columnSize = columnSizeWithoutEvent <= words.size() ? columnSizeWithoutEvent : words.size()
        if (leftSideHasEvent(words)) {
            return words.subList(0, columnSize + 1)
        }
        else {
            return words.subList(0, columnSize)
        }
    }

    private List getRightSideData(List words) {
        if (columnSizeWithoutEvent >= words.size() - 1) {
            return [ ]
        }
        else if (currentLineLeftSideError) {
            return words.subList(leftSizeXEventIdx + 1, words.size())
        }
        else if (leftSideHasEvent(words)) {
            return words.subList(columnSizeWithoutEvent + 1, words.size())
        }
        else {
            return words.subList(columnSizeWithoutEvent, words.size())
        }
    }

    private boolean leftSideHasEvent(words) {
        try {
            BigDecimal test = words[5] as BigDecimal
            return false
        }
        catch (Exception ignored) {
            return true
        }
    }

    private TreeMap addData(TreeMap dataByTypeByDateTime, List words) {
        if (!words) {
            return dataByTypeByDateTime
        }

        BigDecimal humidity
        BigDecimal temp
        String event
        try {
            humidity = words[humidityIdx] as BigDecimal
            temp = words[temperatureIdx] as BigDecimal
            event = words?.getAt(5)
        }
        catch (Exception ignored) {
            try {
                humidity = words[humidityIdx + 1] as BigDecimal
                temp = words[temperatureIdx + 1] as BigDecimal
                event = words?.getAt(6)
            }
            catch (NumberFormatException exception) {
                Cache.errorMessages.add("${serialAndDesc}: no Temp and Humid data;")
            }
        }
        ZonedDateTime dateTime = getDateTime(words)
        Map dataByType = dataByTypeByDateTime.get(dateTime, [ : ])

        if (temp != null) {
            dataByType[TEMPERATURE] = temp
        }
        if (humidity != null) {
            dataByType[HUMIDITY] = humidity
        }
        if (event) {
            dataByType[EVENT] = event
        }

        return dataByTypeByDateTime
    }

    private ZonedDateTime getDateTime(List words) {
        List dateCompos = [ ]
        int i = 1
        while (dateCompos.size() <= 1) {
            dateCompos = words[i].split("-")
            if (dateCompos.size() <= 1) {
                dateCompos = words[i].split("/")
            }
            i ++
        }
        println(dateCompos)
        List timeCompos = words[i].split(":")
        Boolean isAM = checkAMOrPM(words, ++i)

        int year = (dateCompos[yearIdx] as BigInteger).toInteger()
        int tempYearIdx = yearIdx
        int tempMonthIdx = monthIdx
        int tempDayIdx = dayIdx
        if (year < 2000) {
            tempYearIdx = 0
            tempMonthIdx = 1
            tempDayIdx = 2
        }

        year = (dateCompos[tempYearIdx] as BigInteger).toInteger()
        int month
        try {
            month = (dateCompos[tempMonthIdx] as BigInteger).toInteger()
        }
        catch (Exception ignored) {
            month = (DateTimeResolver.MONTH_BY_ABBREVIATION[dateCompos[tempMonthIdx].toLowerCase()] as BigInteger).toInteger()
        }
        int day = (dateCompos[tempDayIdx] as BigInteger).toInteger()

        int hour = (timeCompos[0] as BigInteger).toInteger()
        int min = (timeCompos[1] as BigInteger).toInteger()
        int sec = 0

        if (isAM != null && !isAM && hour != 12) {
            hour += 12
        }
        if (isAM != null && isAM && hour == 12) {
            hour -= 12
        }

        int modMin = min % minuteInterval
        if (modMin <= minuteInterval / 2) {
            min = min - modMin
        }
        else {
            min = min + (minuteInterval - modMin)
        }
        if (min >= 60) {
            min = 0
            hour ++
        }
        if (hour >= 24) {
            hour = 0
            day ++
        }
        if (day > getDaysInMonth(month, year)) {
            day = 1
            month ++
        }
        if (month > 12) {
            month = 1
            year ++
        }
        String monthString = getNumberStringWithLengthOf2(month)
        String dayString = getNumberStringWithLengthOf2(day)
        String hourString = getNumberStringWithLengthOf2(hour)
        String minString = getNumberStringWithLengthOf2(min)
        String secString = getNumberStringWithLengthOf2(sec)

        String dateTime = "${year}-${monthString}-${dayString}T${hourString}:${minString}:${secString}"

        return DateTimeResolver.getZonedDateTime(dateTime)
    }

    String getNumberStringWithLengthOf2(input) {
        String placeHolder = "0000${input.toString()}"
        int len = placeHolder.length()

        return placeHolder.substring(len - 2, len)
    }

    private int getDaysInMonth(int month, int year) {
        try {
            if (year % 4 == 0) {
                return DateTimeResolver.DAYS_BY_MONTH_LEAP[month]
            }

            return DateTimeResolver.DAYS_BY_MONTH[month]
        }
        catch (GroovyCastException exception) {
            println("${serialAndDesc}: year-month-day ordering is wrong;")
            Cache.errorMessages.add("${serialAndDesc}: year-month-day ordering is wrong;")
            return 0
        }
    }

    private Boolean checkAMOrPM(List words, int i) {
        String amOrPm = words[i]
        if (amOrPm == "AM") {
            return true
        }
        else if (amOrPm == "PM") {
            return false
        }

        return null
    }
}