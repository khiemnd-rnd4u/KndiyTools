package sensorService

import kndiyLibraries.DataStructure
import kndiyLibraries.TimeAndDate
import cache.Cache
import com.itextpdf.kernel.pdf.PdfDocument
import com.itextpdf.kernel.pdf.PdfPage
import com.itextpdf.kernel.pdf.PdfReader
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor
import org.codehaus.groovy.runtime.typehandling.GroovyCastException
import org.joda.time.DateTime

import java.util.zip.DataFormatException

class SensorDataParser {
    static final String TEMPERATURE = "temperature"
    static final String HUMIDITY = "humidity"
    static final String SERIAL_SEPARATOR = "-"
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

    TreeMap parseFileAndUpdateDataBySerialByDateTime(TreeMap dataBySerialByDateTime,
                                                     String pathToPdf) {
        PdfReader reader = new PdfReader(pathToPdf)
        PdfDocument document = new PdfDocument(reader)

        parseDocumentAndUpdateMap(
                dataBySerialByDateTime,
                document
        )
        document.close()
        serials = (serialSet as List)?.sort()

        return dataBySerialByDateTime
    }

    private TreeMap parseDocumentAndUpdateMap(TreeMap dataBySerialByDateTime,
                                              PdfDocument document) {
        String serial = findSerial(document)

        int numOfPage = document.getNumberOfPages()
        for (int i = 2; i <= numOfPage; i++) {
            PdfPage page = document.getPage(i)
            String text = PdfTextExtractor.getTextFromPage(page)

            parseTextAndUpdatedataBySerialByDateTime(dataBySerialByDateTime, text, serial)
        }

        return dataBySerialByDateTime
    }

    private String findSerial(PdfDocument document) {
        PdfPage firstPage = document.getPage(1)
        String text = PdfTextExtractor.getTextFromPage(firstPage)
        List lines = text.split("\n")
        String description = ""
        String serial = ""
        println("+++++++++++++++++++++++++++")
        for (String line : lines) {
            line = line.replace("         ", " ")
            line = line.replace("        ", " ")
            line = line.replace("       ", " ")
            line = line.replace("      ", " ")
            line = line.replace("     ", " ")
            line = line.replace("    ", " ")
            line = line.replace("   ", " ")
            line = line.replace("  ", " ")
            println(line)
            if (line.contains("Serial #:")) {
                List serialLineComponents = line.split(" ")
                try {
                    serialLineComponents[3].toString().toBigInteger()
                    serial = serialLineComponents[3]
                }
                catch (Exception e) {
                    serial = serialLineComponents[2]
                }
            }

            if (line.contains("Description:")) {
                description = line - "Description: "
                List descParts = description.split(" ")
                if (descParts.size() == 1) {
                    descParts = [ ]
                    for (int i = 0; i < description.length(); i ++) {
                        String cur = description.substring(i, i + 1)
                        try {
                            Integer num = cur.toBigInteger().toInteger()
                            if (!descParts[1]) {
                                descParts[1] = num
                            }
                            else {
                                descParts[1] = "${descParts[1]}${num}"
                            }
                        }
                        catch (ignored) {
                            if (!descParts[0]) {
                                descParts[0] = cur
                            }
                            else {
                                descParts[0] = "${descParts[0]}${cur}"
                            }
                        }
                    }
                }
                int numPart = (descParts[1] as String).toBigInteger().toInteger()
                String finalNum = String.format("%03d", numPart)
                description = "${descParts[0]}-${finalNum}"
            }

            if (description != "" && serial != "") {
                String finalSerial = "${description}${SERIAL_SEPARATOR}${serial}"
                println("---------------------------------------")
                println(finalSerial)
                println("---------------------------------------")
                serialAndDesc = finalSerial
                serialSet << finalSerial
                println("+++++++++++++++++++++++++++")

                return finalSerial
            }
        }
        println("SERIAL NOT FOUND!! SERIAL NOT FOUND!! SERIAL NOT FOUND!! SERIAL NOT FOUND!! SERIAL NOT FOUND!!")


        return ""
    }

    private TreeMap parseTextAndUpdatedataBySerialByDateTime(TreeMap dataBySerialByDateTime,
                                                                String text,
                                                                String serial) {
        List lines = text.split("\n")
        for (String line : lines) {
            line = line.replace("         ", " ")
            line = line.replace("        ", " ")
            line = line.replace("       ", " ")
            line = line.replace("      ", " ")
            line = line.replace("     ", " ")
            line = line.replace("    ", " ")
            line = line.replace("   ", " ")
            line = line.replace("  ", " ")
            println(line)

            currentLineLeftSideError = false
            currentLineRightSideError = false

            List words = line?.split(" ")
            if (!isDataLine(words)) {
                continue
            }
            println(words)
            try {
                parseWordsAndUpdateDataBySerialByDateTime(dataBySerialByDateTime, words, serial)
            }
            catch (DataFormatException ignored) {
            }
        }

        return dataBySerialByDateTime
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

    private TreeMap parseWordsAndUpdateDataBySerialByDateTime(TreeMap dataByDateTime, List words, String serial) {
        List leftSideWords = getLeftSideData(words)
        List rightSideWords = getRightSideData(words)

        println(leftSideWords)
        println(rightSideWords)

        addData(dataByDateTime, leftSideWords, serial)
        addData(dataByDateTime, rightSideWords, serial)

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

    private TreeMap addData(TreeMap dataBySerialByDateTime, List words, String serial) {
        if (!words) {
            return dataBySerialByDateTime
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
        DateTime dateTime = getDateTime(words)
        TreeMap dataBySerial = DataStructure.getOrCreateObject(dataBySerialByDateTime, dateTime, DataStructure.DATA_STRUCTURE_TREE_MAP)
        Map data = DataStructure.getOrCreateObject(dataBySerial, serial)

        if (temp != null) {
            data[TEMPERATURE] = temp
        }
        if (humidity != null) {
            data[HUMIDITY] = humidity
        }
        if (event) {
            data[EVENT] = event
        }

        return dataBySerialByDateTime
    }

    private DateTime getDateTime(List words) {
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
        int month
        try {
            month = (dateCompos[monthIdx] as BigInteger).toInteger()
        }
        catch (Exception ignored) {
            month = (TimeAndDate.MONTH_BY_ABBREVIATION[dateCompos[monthIdx].toLowerCase()] as BigInteger).toInteger()
        }
        int day = (dateCompos[dayIdx] as BigInteger).toInteger()

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

        return new DateTime(year, month, day, hour, min, sec)
    }

    private int getDaysInMonth(int month, int year) {
        try {
            if (year % 4 == 0) {
                return TimeAndDate.DAYS_BY_MONTH_LEAP[month]
            }

            return TimeAndDate.DAYS_BY_MONTH[month]
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