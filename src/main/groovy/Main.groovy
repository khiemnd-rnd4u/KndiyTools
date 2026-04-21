import entities.Matrix
import entities.sensor.Inspection
import services.kndiyLibraries.DateTimeResolver
import services.novelParsingService.NovelProcessor
import services.sensorService.SensorDataResolver
import services.sensorService.SensorDistributionReport
import services.sensorService.SensorFolderParser
import services.sensorService.SensorReferenceReport

import java.time.ZonedDateTime

static void main(String[] args) {
    runSensorParser()
}

static void runMatrixTesting() {
    Matrix matrix = new Matrix(2, 3, [[11111, 2, 7],[5, 6, 8]])
    matrix.printMatrix("Matrix A")

    Matrix transposedMatrix = matrix.transposeMatrix()
    transposedMatrix.printMatrix("Matrix B")

    Matrix product = matrix * transposedMatrix
    product.printMatrix("Product Matrix")
}

static void runNovelParser() {
    runNovelParser(
            "/run/media/kndiy/Mana/Documents/Kindle Books/WN CN The Dark King/source/",
            "/run/media/kndiy/Mana/Documents/Kindle Books/WN CN The Dark King/WN CN The Dark King.docx",
            "/run/media/kndiy/Mana/Documents/Kindle Books/Template.docx",
            "If you find any errors", "THE DARK KING",
            "Chapter ",
            "The Dark King", "古羲 - Gu Xi"
    )
}

static void runNovelParser(String sourcePath, String outputPath, String templatePath,
                           String skippableStartSign, String skippableEndSign,
                           String chapterStartsWithText,
                           String bookName, String authorName) {
    new NovelProcessor(
            sourcePath, outputPath, templatePath,
            skippableStartSign, skippableEndSign,
            chapterStartsWithText,
            bookName, authorName
    )
}

static void runSensorParser() {
    String sourceFolder = "D:\\Documents\\Personal\\00 Family Small\\Projects\\260407 Kiem Dinh Kho Viet An\\Final Arrangement\\Official Test\\PhongNgoai"
    String savedFileName = "Kho_Viet_An"
    String savedPath = "D:\\"

    String inspectionId = "VietAn_20260407"
    String inspectionAddress = "HCM"
    String inspectionDate = "2026-04-04"

    BigDecimal minTemperatureTolerance = 15
    BigDecimal maxTemperatureTolerance = 30
    BigDecimal minHumidityTolerance = 0
    BigDecimal maxHumidityTolerance = 75
    Integer locationIdx = 0
    Integer serialIdx = 2
    Integer altitudeIdx = null

    SensorFolderParser sensorFolderParser = new SensorFolderParser(
            sourceFolder,
            5,
            2, 0, 1,
            4, 3,
            15
    )
    ////////////////////// DON'T TOUCH BELOW
    ////////////////////// DON'T TOUCH BELOW
    ////////////////////// DON'T TOUCH BELOW
    String dateTime = DateTimeResolver.getDateTimeString(ZonedDateTime.now(), "yyyyMMdd_HHmm")

    TreeMap rawData = sensorFolderParser.parseFilesInFolderAndGetRawData()

    SensorDataResolver sensorDataResolver = new SensorDataResolver(
            rawData,
            inspectionId,
            savedFileName,
            inspectionAddress,
            inspectionDate,
            minTemperatureTolerance,
            maxTemperatureTolerance,
            minHumidityTolerance,
            maxHumidityTolerance,
            locationIdx,
            serialIdx,
            altitudeIdx
    )

    Inspection inspection = sensorDataResolver.getInspection()
    new SensorDistributionReport(inspection, false)
            .createAndSaveWorkbook("${savedPath}${dateTime}_${savedFileName}.xlsx")
    new SensorReferenceReport(inspection, false)
            .createAndSaveWorkbook("${savedPath}${dateTime}_ReferenceData_${savedFileName}.xlsx")
}