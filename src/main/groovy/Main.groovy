import entities.Matrix
import entities.sensor.Inspection
import services.kndiyLibraries.DateTimeResolver
import services.novelParsingService.NovelProcessor
import services.sensorService.SensorDataResolver
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

    String dateTime = DateTimeResolver.getDateTimeString(ZonedDateTime.now(), "yyyyMMdd_HHmm")

    SensorFolderParser sensorFolderParser = new SensorFolderParser(
            sourceFolder,
            5,
            2, 0, 1,
            4, 3,
            15
    )
    TreeMap rawData = sensorFolderParser.parseFilesInFolderAndGetRawData()

    SensorDataResolver sensorDataResolver = new SensorDataResolver(
            rawData,
            "VietAn_20260407",
            savedFileName,
            "HCM",
            "2026-03-29"
    )
    Inspection inspection = sensorDataResolver.getInspection()
//    new SensorDistributionReport(inspection, false)
//            .createAndSaveWorkbook("${savedPath}${dateTime}_${savedFileName}.xlsx")
    new SensorReferenceReport(inspection, false)
            .createAndSaveWorkbook("${savedPath}${dateTime}_ReferenceData_${savedFileName}.xlsx")
}