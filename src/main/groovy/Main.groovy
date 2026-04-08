import chineseTextService.ChineseNumberConverter
import entities.Matrix
import novelParsingService.NovelProcessor
import readTxtNovelParsingService.NovelConverter
import sensorService.SensorFolderParser
import cache.Cache

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

static void main(String[] args) {
    SensorFolderParser folderParser = new SensorFolderParser(
            "D:\\Documents\\Personal\\00 Family Small\\Projects\\260329 Kiem Dinh Xe Van Chuyen\\FinalArrangement\\Xe1\\Lan1",
            "D:\\Documents\\Personal\\00 Family Small\\Projects\\260329 Kiem Dinh Xe Van Chuyen\\FinalArrangement\\Xe1\\Lan1.xlsx",
            5,
            0, 1, 2,
            4, 3,
            10
    )

    folderParser.parseFilesInFolder()

//    String movePath = "D:\\AllFiles\\All"
//    File folder = new File("D:\\AllFiles")
//    folder.listFiles().each { File file ->
//        Path sourceFile = Paths.get(file.getPath())
//        String pathToMove = movePath
//        Path moveToPath = Paths.get(pathToMove)
//        Files.createDirectories(moveToPath)
//
//        String fileName = sourceFile.getFileName()
//        fileName = fileName?.split("__")?.first()
//        Path moveToFile = moveToPath.resolve(fileName)
//        Files.copy(sourceFile, moveToFile)
//        println("copy for ${moveToFile}")
//    }

//    new NovelConverter(
//            "D:\\Documents\\Kindle Books\\RAW Chinese Books\\远瞳 《黎明之剑》 Sword of Daybreaker.txt",
//            "D:\\Documents\\Kindle Books\\RAW Chinese Books\\远瞳 《黎明之剑》 Sword of Daybreaker.docx",
//            "D:\\Documents\\Kindle Books\\RAW Chinese Books\\Template.docx",
//            "远瞳", "黎明之剑"
//    )
//

//    Metcash metcash = new Metcash()
//    metcash.convertUtcDateTimeToBrisbaneDateTime("2026-02-04T21:00:00")
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

static void runSensorParser() {
    runSensorParser(
            "D:\\Documents\\Personal\\00 Family Small\\Projects\\250710 Tham Dinh Kho\\kho thường 7 ngày\\kho thường 7 ngày PDF",
            "D:\\Documents\\Personal\\00 Family Small\\Projects\\250710 Tham Dinh Kho\\KhoThuong7Ngay.xlsx",
            6,
            2, 0, 1,
            4, 3, 15
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

static void runSensorParser(String folderPath, String savePathWithName,
                       int columnSizeWithoutEvent,
                       int yearIdx, int monthIdx, int dayIdx,
                       int temperatureIdx, int humidityIdx,
                       int minuteInterval) {
    SensorFolderParser sensorFolderParser = new SensorFolderParser(
            folderPath, savePathWithName,
            columnSizeWithoutEvent,
            yearIdx, monthIdx, dayIdx,
            temperatureIdx, humidityIdx,
            minuteInterval
    )
    sensorFolderParser.parseFilesInFolder()

    Cache.printErrorMessages()
}