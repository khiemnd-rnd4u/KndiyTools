package services.sensorService

import java.time.ZonedDateTime

class SensorFolderParser {
    private String folderPath
    private TreeMap<String, TreeMap<ZonedDateTime, Map>> dataByTypeByDateTimeByFileName

    private SensorDataParser sensorDataParser

    SensorFolderParser(String folderPath,
                       int columnSizeWithoutEvent,
                       int yearIdx, int monthIdx, int dayIdx,
                       int tempIdx, int humidIdx, int minuteInterval) {
        this.folderPath = folderPath

        sensorDataParser = new SensorDataParser(columnSizeWithoutEvent, yearIdx, monthIdx, dayIdx, tempIdx, humidIdx, minuteInterval)
        dataByTypeByDateTimeByFileName = new TreeMap<>()

        parseFilesInFolderAndGetRawData()
    }

    TreeMap<String, TreeMap<ZonedDateTime, Map>> parseFilesInFolderAndGetRawData() {
        File folder = new File(folderPath)
        folder.listFiles().each { File file ->
            String fileName = file.getName()
            if (!fileName.endsWith(".pdf")) {
                return
            }
            fileName = fileName.replace(".pdf", "")

            TreeMap<ZonedDateTime, Map> dataByTypeByDateTime = new TreeMap<>((t1, t2) -> {
                return t1 <=> t2
            })
            dataByTypeByDateTimeByFileName[ fileName ] = dataByTypeByDateTime

            println("Started Parsing ${file.getAbsolutePath()}")
            sensorDataParser.parseFileAndUpdateDataByTypeByDateTime(
                    dataByTypeByDateTime,
                    file
            )
        }
        println("Done Parsing ALL Files")

        return dataByTypeByDateTimeByFileName
    }
}
