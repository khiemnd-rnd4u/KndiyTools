import entities.Matrix
import entities.sensor.Inspection
import services.inventories.InventoryReader
import services.inventories.InventoryReport
import services.novelParsingService.NovelProcessor
import services.sensorService.SensorDataResolver
import services.sensorService.SensorDistributionReport
import services.sensorService.SensorFolderParser
import services.sensorService.SensorReferenceReport

static void main(String[] args) {
    runInventoryParser()
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
    String temp = "3"
    String sourceFolder = "D:\\Documents\\Personal\\00 Family Small\\Projects\\260407 Kiem Dinh Kho Viet An\\Final Arrangement\\Xe\\Lan${temp}"
    String savedFileName = "XeVanChuyen_51K768.20_Lan${temp}"
    String savedPath = "D:\\Documents\\Personal\\00 Family Small\\Projects\\260407 Kiem Dinh Kho Viet An\\Final Arrangement"

    String inspectionId = savedFileName
    String inspectionName = "Xe vận chuyển 51K-768.20 Lần ${temp}"
    String inspectionAddress = "32/47 Phan Sào Nam, phường Bảy Hiền, TP.HCM"
    String inspectionDate = "2026-03-29"

    BigDecimal minTemperatureTolerance = 15
    BigDecimal maxTemperatureTolerance = 30
    BigDecimal minHumidityTolerance = 0
    BigDecimal maxHumidityTolerance = 75
    Integer locationIdx = 0
    Integer serialIdx = 2
    Integer altitudeIdx = null

    boolean hasAltitude = false

    SensorFolderParser sensorFolderParser = new SensorFolderParser(
            sourceFolder,
            5,
            2, 0, 1,
            4, 3,
            10
    )
    ////////////////////// DON'T TOUCH BELOW
    ////////////////////// DON'T TOUCH BELOW
    ////////////////////// DON'T TOUCH BELOW
    TreeMap rawData = sensorFolderParser.parseFilesInFolderAndGetRawData()

    SensorDataResolver sensorDataResolver = new SensorDataResolver(
            rawData,
            inspectionId,
            inspectionName,
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

    if (!savedPath.endsWith("\\")) {
        savedPath += "\\"
    }

    Inspection inspection = sensorDataResolver.getInspection()
    new SensorDistributionReport(inspection, hasAltitude)
            .createAndSaveWorkbook("${savedPath}${inspectionDate}_${savedFileName}.xlsx")
    new SensorReferenceReport(inspection, hasAltitude)
            .createAndSaveWorkbook("${savedPath}${inspectionDate}_${savedFileName}_ReferenceData.xlsx")
}

static void runInventoryParser() {
    InventoryReader inventoryReader = new InventoryReader(
            "D:\\Documents\\Personal\\00 Family Small\\Projects\\260503 Nhap Xuat Ton Ario\\NXT\\2025",
            "D:\\Documents\\Personal\\00 Family Small\\Projects\\260503 Nhap Xuat Ton Ario\\ProductMaster.xlsx"
    )

    InventoryReport inventoryReport = new InventoryReport(
            inventoryReader,
            ""
    )
    inventoryReport.createAndSaveWorkbook("D:\\Documents\\Personal\\00 Family Small\\Projects\\260503 Nhap Xuat Ton Ario\\Test.xlsx")
}

