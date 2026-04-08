package pricefx

import kndiyLibraries.DataStructure
import org.slf4j.Logger
import org.slf4j.LoggerFactory

import java.nio.file.*
import java.time.ZoneId
import java.time.ZonedDateTime
import java.time.format.DateTimeFormatter
import java.time.format.ResolverStyle

class Metcash {
    Map getIntegrationFileByStateByDsName(String folderPath) {
        Map integrationFileByStateByDsName = [ : ]
        File folder = new File(folderPath)
        println("Parsing Folder: ${folder.getName()}")
        for (File file : folder.listFiles()) {
            if (!file.isFile()) {
                continue
            }
            println("Parsing File: ${file.getName()}")

            String fileName = file.getName()
            String nameFirstPart = fileName.substring(0, fileName.indexOf("."))
            String state = nameFirstPart.substring(1, 4)
            String dsName = nameFirstPart.substring(4)

            Map integrationFileByState = DataStructure.getOrCreateObject(integrationFileByStateByDsName, dsName)
            integrationFileByState[ state ] = file
        }

        return integrationFileByStateByDsName
    }

    void copyFilesToFolder(Map integrationFileByStateByDsName, String movePath) {
        integrationFileByStateByDsName.each { String dsName, Map fileByState ->
            fileByState.each { String state, File file ->
                Path sourceFile = Paths.get(file.getPath())
                String pathToMove = movePath.endsWith("\\/")
                        ? "${movePath}${dsName}"
                        : "${movePath}/${dsName}"
                Path moveToPath = Paths.get(pathToMove)
                Files.createDirectories(moveToPath)

                String fileName = sourceFile.getFileName()
                fileName = fileName.substring(0, fileName.indexOf("__"))
                Path moveToFile = moveToPath.resolve(fileName)

                try {
                    Files.copy(sourceFile, moveToFile)
                    println("Copy completed for ${moveToFile.getFileName()}")
                }
                catch (IOException ex) {
                    println("Copy failed: ${ex.getMessage()}")
                }
            }
        }
    }

    String convertUtcDateTimeToBrisbaneDateTime(String inputDateTime) {
        ZoneId UTC_ZONE_ID = ZoneId.of("UTC")
        String DEFAULT_DATE_FORMAT = "yyyy-MM-dd'T'HH:mm:ss"
        DateTimeFormatter DF_DATE_FORMATTER =
                DateTimeFormatter.ofPattern(DEFAULT_DATE_FORMAT)
                        .withResolverStyle(ResolverStyle.SMART)
                        .withZone(UTC_ZONE_ID)

        ZoneId BRISBANE_ZONE_ID = ZoneId.of("Australia/Brisbane")
        String OUTPUT_DATE_FORMAT = "yyyy-MM-dd"

        DateTimeFormatter DF_OUTPUT_DATE_LOG_FORMATTER =
                DateTimeFormatter.ofPattern(DEFAULT_DATE_FORMAT)
                        .withResolverStyle(ResolverStyle.SMART)

        DateTimeFormatter DF_OUTPUT_DATE_FORMATTER =
                DateTimeFormatter.ofPattern(OUTPUT_DATE_FORMAT)
                        .withResolverStyle(ResolverStyle.SMART)

        ZonedDateTime utcTime = ZonedDateTime
                .parse(String.valueOf(inputDateTime), DF_DATE_FORMATTER)
        println("Input DateTime Convert for IntegrationLogs ${inputDateTime}")

        ZonedDateTime brisbaneTime = utcTime.withZoneSameInstant(BRISBANE_ZONE_ID)

        final String brisbaneTimeString = brisbaneTime.format(DF_OUTPUT_DATE_LOG_FORMATTER)
        println("Output DateTime Convert for IntegrationLogs ${brisbaneTimeString}")

        final String brisbaneDateString = brisbaneTime.format(DF_OUTPUT_DATE_FORMATTER)
        println("Output Date ${brisbaneDateString}")

        return brisbaneDateString
    }
}
