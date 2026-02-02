import entities.Matrix
import novelParsingService.NovelProcessor
import pricefx.Metcash
import sensorService.SensorFolderParser
import cache.Cache

static void main(String[] args) {
    Metcash metcash = new Metcash()
    Map fileByStateByDsName = metcash.getIntegrationFileByStateByDsName("/home/kndiy/metcashIntegration/files/Source/")
    metcash.copyFilesToFolder(
            fileByStateByDsName,
            "/home/kndiy/metcashIntegration/files/Target/"
    )
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