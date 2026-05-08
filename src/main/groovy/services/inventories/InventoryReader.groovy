package services.inventories

import entities.inventories.CustomerMaster
import entities.inventories.MaterialNorm
import entities.inventories.ProductMaster
import entities.inventories.StockContent
import entities.inventories.StockEntry
import groovyjarjarantlr4.runtime.tree.Tree
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import services.kndiyLibraries.KndiyFormatter
import services.kndiyLibraries.XSSFTools

import java.text.SimpleDateFormat
import java.time.ZonedDateTime

class InventoryReader {
    final static String STOCK_IN_STARTS_WITH = "MUA_VAO"
    final static String STOCK_OUT_STARTS_WITH = "BAN_RA"

    final static String MAIN_SHEET_NAME = "Smart_KTSC_OK"

    private static final int PM_SKU_CELL_NUM = 1
    private static final int PM_FAMILY_CELL_NUM = 2
    private static final int PM_LABEL_CELL_NUM = 3
    private static final int PM_UOM_CELL_NUM = 4
    private static final int PM_PROCESSING_TARGET_CELL_NUM = 5

    private static final int MN_DATE_CELL_NUM = 0
    private static final int MN_ITEM_LABEL_CELL_NUM = 1
    private static final int MN_MATERIAL_LABEL_CELL_NUM = 2
    private static final int MN_MATERIAL_UOM_CELL_NUM = 3
    private static final int MN_NORM_AMOUNT_CELL_NUM = 4

    private int VOUCHER_ID_CELL_NUM
    private int DATE_CELL_NUM
    private int ITEM_LABEL_CELL_NUM
    private int UOM_CELL_NUM
    private int QUANTITY_CELL_NUM
    private int UNIT_PRICE_CELL_NUM
    private int VALUE_CELL_NUM
    private int TAX_RATE_CELL_NUM
    private int TAX_CELL_NUM
    private int VALUE_WITH_TAX_CELL_NUM
    private int CUSTOMER_ID_CELL_NUM
    private int CUSTOMER_NAME_CELL_NUM
    private int CUSTOMER_ADDRESS_CELL_NUM

    SimpleDateFormat DATE_CELL_FORMATTER

    List<XSSFWorkbook> stockInBooks
    List<XSSFWorkbook> stockOutBooks
    XSSFWorkbook productMasterBook

    private Map productMasterBySku
    private Map productMasterByItemLabel
    private Map customerMasterByName

    private TreeMap<ZonedDateTime, Map<String, List<StockContent>>> outStockContentsByItemLabelByDate
    private TreeMap<ZonedDateTime, TreeMap<String, List<StockContent>>> inStockContentsByItemLabelByDate

    private TreeMap<ZonedDateTime, Map<String, MaterialNorm>> materialNormByTargetSkuByDate

    InventoryReader(String folderPath,
                    String productMasterPath) {
        File folder = new File(folderPath)
        stockInBooks = [ ]
        stockOutBooks = [ ]

        folder.listFiles().each { File file ->
            String fileName = file.getName()
            String filePath = file.getAbsolutePath()
            if (fileName.startsWith(STOCK_IN_STARTS_WITH)) {
                stockInBooks.add(XSSFTools.getWorkbook(filePath))
            }
            else if (fileName.startsWith(STOCK_OUT_STARTS_WITH)) {
                stockOutBooks.add(XSSFTools.getWorkbook(filePath))
            }
        }
        productMasterBook = XSSFTools.getWorkbook(productMasterPath)

        productMasterByItemLabel = [: ]
        productMasterBySku = [: ]
        customerMasterByName = [ : ]

        outStockContentsByItemLabelByDate = new TreeMap<>((ZonedDateTime d1, ZonedDateTime d2) -> { return d1 <=> d2 })
        inStockContentsByItemLabelByDate = new TreeMap<>((ZonedDateTime d1, ZonedDateTime d2) -> { return d1 <=> d2 })
        materialNormByTargetSkuByDate = new TreeMap<>((ZonedDateTime d1, ZonedDateTime d2) -> { return d1 <=> d2 })

        DATE_CELL_FORMATTER = new SimpleDateFormat("yyyy-MM-dd")

        parseData()
    }

    private void parseData() {
        buildProductMasterData()
        println("Built Product Master By Label")

        buildStockData(stockInBooks, StockEntry.TYPE_STOCK_IN)
        println("Built Stock in Data")

        buildStockData(stockOutBooks, StockEntry.TYPE_STOCK_OUT)
        println("Built Stock out Data")

        buildMaterialNormData()
        println("Built Material Norm Data")
    }

    private void buildStockData(List<XSSFWorkbook> workbooks,
                                String type) {
        workbooks.each { XSSFWorkbook workbook ->
            buildStockData(workbook, type)
        }
    }
    
    private void buildStockData(XSSFWorkbook workbook, 
                                String type) {
        XSSFSheet sheet = workbook.getSheet(MAIN_SHEET_NAME)
        buildCellLocation(sheet)

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue
            }

            String customerName = KndiyFormatter.getStringWithoutRedundantSpaces(row.getCell(CUSTOMER_NAME_CELL_NUM).getStringCellValue())
            String customerId = KndiyFormatter.getStringWithoutRedundantSpaces(row.getCell(CUSTOMER_ID_CELL_NUM).getStringCellValue())
            String customerAddress = KndiyFormatter.getStringWithoutRedundantSpaces(row.getCell(CUSTOMER_ADDRESS_CELL_NUM).getStringCellValue())

            CustomerMaster customerMaster = customerMasterByName.get(customerName, new CustomerMaster(customerId, customerName, customerAddress))
            StockEntry stockEntry = addStockEntry(customerMaster, row, type)
            addStockContent(stockEntry, row, type)
        }
    }

    private StockEntry addStockEntry(CustomerMaster customerMaster, Row row, String type) {
        String voucherId = KndiyFormatter.getStringWithoutRedundantSpaces(row.getCell(VOUCHER_ID_CELL_NUM).getStringCellValue())
        ZonedDateTime date = XSSFTools.readDateTimeFromCell(row, DATE_CELL_NUM, DATE_CELL_FORMATTER)

        return customerMaster.getOrSetStockEntry(voucherId, type, date)
    }

    private StockContent addStockContent(StockEntry stockEntry, Row row, String type) {
        String itemLabel = KndiyFormatter.getStringWithoutRedundantSpaces(row.getCell(ITEM_LABEL_CELL_NUM).getStringCellValue())
        String uom = KndiyFormatter.getStringWithoutRedundantSpaces(row.getCell(UOM_CELL_NUM).getStringCellValue())
        BigDecimal quantity = XSSFTools.getNumericCellValue(row.getCell(QUANTITY_CELL_NUM))
        BigDecimal unitPrice = XSSFTools.getNumericCellValue(row.getCell(UNIT_PRICE_CELL_NUM))
        BigDecimal value = XSSFTools.getNumericCellValue(row.getCell(VALUE_CELL_NUM))
        BigDecimal tax = XSSFTools.getNumericCellValue(row.getCell(TAX_CELL_NUM))
        BigDecimal taxRate = XSSFTools.getNumericCellValue(row.getCell(TAX_RATE_CELL_NUM))
        BigDecimal valueWithTax = XSSFTools.getNumericCellValue(row.getCell(VALUE_WITH_TAX_CELL_NUM))

        ProductMaster productMaster = productMasterByItemLabel[ itemLabel] as ProductMaster
        if (!productMaster) {
            String tempSku = itemLabel
            productMaster = new ProductMaster(tempSku, itemLabel, "None", uom)
            productMasterByItemLabel[ itemLabel] = productMaster
            productMasterBySku[ tempSku ] = productMaster
        }
        StockContent stockContent = stockEntry.addStockContent(productMaster, uom, quantity, unitPrice, value, tax, valueWithTax, taxRate)

        if (type == StockEntry.TYPE_STOCK_IN) {
            TreeMap<String, List<StockContent>> inStockContentsByItemLabel = inStockContentsByItemLabelByDate.get(
                    stockEntry.getDate(),
                    new TreeMap<>((String label1, String label2) ->{
                        ProductMaster productMaster1 = productMasterByItemLabel[ label1 ]
                        ProductMaster productMaster2 = productMasterByItemLabel[ label2 ]

                        String family1 = productMaster1.getItemFamily()
                        String family2 = productMaster2.getItemFamily()
                        if (family1 == ProductMaster.FAMILY_PROCESSED_GOODS) {
                            return 1
                        }
                        if (family2 == ProductMaster.FAMILY_PROCESSED_GOODS) {
                            return -1
                        }

                        return label1 <=> label2
                    })
            )
            List<StockContent> stockInContents = inStockContentsByItemLabel[ itemLabel ]
            if (!stockInContents) {
                stockInContents = [ ]
                inStockContentsByItemLabel[ itemLabel ] = stockInContents
            }
            stockInContents.add(stockContent)
        }
        else if (type == StockEntry.TYPE_STOCK_OUT) {
            Map<String, List<StockContent>>  outStockContentsByItemLabel = outStockContentsByItemLabelByDate.get(stockEntry.getDate(), [ : ])
            List<StockContent> stockOutContents = outStockContentsByItemLabel.get(itemLabel, [ ])
            stockOutContents.add(stockContent)
        }

        return stockContent
    }

    private void buildCellLocation(XSSFSheet sheet) {
        Row header = sheet.getRow(0)
        int idx = -1
        for (Cell cell : header) {
            String cellValue = cell.getStringCellValue()
            setCellLocation(cellValue, ++idx)
        }
    }

    private void setCellLocation(String cellValue, int idx) {
        switch (cellValue) {
            case "SOCT" :
                VOUCHER_ID_CELL_NUM = idx
                break
            case "NGAY_HD" :
                DATE_CELL_NUM = idx
                break
            case "TENDM" :
                ITEM_LABEL_CELL_NUM = idx
                break
            case "DONVI" :
                UOM_CELL_NUM = idx
                break
            case "LUONG" :
                QUANTITY_CELL_NUM = idx
                break
            case "DGVND" :
                UNIT_PRICE_CELL_NUM = idx
                break
            case "TTVND" :
                VALUE_CELL_NUM = idx
                break
            case "THUEVND" :
                TAX_CELL_NUM = idx
                break
            case "TS_GTGT" :
                TAX_RATE_CELL_NUM = idx
                break
            case "TTVND_TT" :
                VALUE_WITH_TAX_CELL_NUM = idx
                break
            case "KHACHHANG" :
                CUSTOMER_NAME_CELL_NUM = idx
                break
            case "MS_DN" :
                CUSTOMER_ID_CELL_NUM = idx
                break
            case "DIACHI" :
                CUSTOMER_ADDRESS_CELL_NUM = idx
        }
    }

    private void buildProductMasterData() {
        Sheet sheet = productMasterBook.getSheet("ProductMaster")
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue
            }
            String sku = XSSFTools.readStringFromCell(row, PM_SKU_CELL_NUM)
            String label = XSSFTools.readStringFromCell(row, PM_LABEL_CELL_NUM)
            String family = XSSFTools.readStringFromCell(row, PM_FAMILY_CELL_NUM)
            String uom = XSSFTools.readStringFromCell(row, PM_UOM_CELL_NUM)
            String processingTarget = XSSFTools.readStringFromCell(row, PM_PROCESSING_TARGET_CELL_NUM)

            ProductMaster productMaster = productMasterBySku.get(sku, new ProductMaster(sku, label, family, uom, processingTarget))
            productMasterByItemLabel[ label ] = productMaster
        }
    }

    private void buildMaterialNormData() {
        XSSFSheet sheet = productMasterBook.getSheet("MaterialNorm")
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue
            }
            ZonedDateTime date = XSSFTools.readDateTimeFromCell(row, MN_DATE_CELL_NUM, DATE_CELL_FORMATTER)
            String targetLabel = XSSFTools.readStringFromCell(row, MN_ITEM_LABEL_CELL_NUM)
            String materialLabel = XSSFTools.readStringFromCell(row, MN_MATERIAL_LABEL_CELL_NUM)
            BigDecimal normAmount = XSSFTools.readNumberFromCell(row, MN_NORM_AMOUNT_CELL_NUM)

            ProductMaster targetProductMaster = productMasterByItemLabel[ targetLabel ] as ProductMaster
            ProductMaster materialProductMaster = productMasterByItemLabel[ materialLabel ] as ProductMaster

            Map materialNormByTargetSku = materialNormByTargetSkuByDate.get(date, [ : ])
            MaterialNorm materialNorm = materialNormByTargetSku.get(targetProductMaster.getSku(), new MaterialNorm(targetProductMaster))
            if (!materialProductMaster) {
                println("Failed to find Product Master of label: ${materialLabel}")
            }
            if (!targetProductMaster) {
                println("Failed to find Product Master of label: ${targetLabel}")
            }
            materialNorm.addMaterialComponent(materialProductMaster, normAmount)
        }
    }

    Map getProductMasterByItemLabel() {
        return productMasterByItemLabel
    }

    Map getProductMasterBySku() {
        return productMasterBySku
    }

    Map getCustomerMasterByName() {
        return customerMasterByName
    }

    TreeMap<ZonedDateTime, Map<String, List<StockContent>>> getOutStockContentsByItemLabelByDate() {
        return outStockContentsByItemLabelByDate
    }

    TreeMap<ZonedDateTime, Map<String, List<StockContent>>> getInStockContentsByItemLabelByDate() {
        return inStockContentsByItemLabelByDate
    }

    TreeMap<ZonedDateTime, Map<String, MaterialNorm>> getMaterialNormByTargetSkuByDate() {
        return materialNormByTargetSkuByDate
    }
}
