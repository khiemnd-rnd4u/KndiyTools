package services.inventories

import entities.inventories.ProductMaster
import entities.inventories.StockContent
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.PageMargin
import org.apache.poi.ss.usermodel.PrintSetup
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFSheet
import services.kndiyLibraries.DateTimeResolver
import services.kndiyLibraries.KnDiyWorkbook
import services.kndiyLibraries.XSSFTools

import java.time.ZonedDateTime

class InventoryReport extends KnDiyWorkbook{
    private String TRADE_GOODS_IN_OUT_SHEET_NAME = "Nhập xuất HH"
    private String TRADE_GOODS_INVENTORY_SHEET_NAME = "Tồn HH"
    private String MATERIAL_IN_OUT_SHEET_NAME = "Nhập xuất NVL"
    private String MATERIAL_INVENTORY_SHEET_NAME = "Tồn NVL"

    private String lastInventoryReportPath

    private XSSFSheet TRADE_GOODS_IN_OUT_SHEET
    private XSSFSheet TRADE_GOODS_INVENTORY_SHEET
    private XSSFSheet MATERIAL_IN_OUT_SHEET
    private XSSFSheet MATERIAL_INVENTORY_SHEET

    private int TITLE_ROW_NUM
    private int HEADER_FIRST_ROW_NUM
    private int HEADER_SECOND_ROW_NUM
    private int HEADER_ID_ROW_NUM
    private int SUMMARY_ROW_NUM
    private int DATA_START_ROW_NUM
    private int DATA_END_ROW_NUM

    private int VOUCHER_ID_CELL_NUM
    private int DATE_CELL_NUM
    private int CUSTOMER_CELL_NUM
    private int CUSTOMER_ID_CELL_NUM
    private int CONTENT_CELL_NUM
    private int SKU_CELL_NUM
    private int ITEM_LABEL_CELL_NUM
    private int UOM_CELL_NUM
    private int IN_QUANTITY_CELL_NUM
    private int IN_UNIT_PRICE_CELL_NUM
    private int IN_VALUE_CELL_NUM
    private int OUT_QUANTITY_CELL_NUM
    private int OUT_UNIT_PRICE_CELL_NUM
    private int OUT_VALUE_CELL_NUM
    private int SOLD_QUANTITY_CELL_NUM
    private int SOLD_UNIT_PRICE_CELL_NUM
    private int SOLD_VALUE_CELL_NUM
    private int ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM
    private int ACCUMULATE_INVENTORY_VALUE_CELL_NUM
    private int DATA_END_CELL_NUM

    private InventoryReader inventoryReader
    private TreeMap<ZonedDateTime, Map<String, List<StockContent>>> outStockContentsByItemLabelByDate
    private TreeMap<ZonedDateTime, Map<String, List<StockContent>>> inStockContentsByItemLabelByDate

    Map currentDataRowBySheetName

    Map accumulateInventoryDataBySku
    static final String AI_QUANTITY = "InventoryQuantity"
    static final String AI_SUM_VALUE = "InventorySumValue"

    TreeMap<String, Map> inventoryDataBySku
    static final String ID_OPEN_QUANTITY = "OpenQuantity"
    static final String ID_OPEN_VALUE = "OpenValue"
    static final String ID_IN_QUANTITY = "InQuantity"
    static final String ID_IN_VALUE = "InValue"
    static final String ID_OUT_QUANTITY = "OutQuantity"
    static final String ID_OUT_VALUE = "OutValue"

    private int IV_SKU_CELL_NUM
    private int IV_ITEM_LABEL_CELL_NUM
    private int IV_UOM_CELL_NUM
    private int IV_OPEN_QUANTITY_CELL_NUM
    private int IV_OPEN_VALUE_CELL_NUM
    private int IV_IN_QUANTITY_CELL_NUM
    private int IV_IN_VALUE_CELL_NUM
    private int IV_OUT_QUANTITY_CELL_NUM
    private int IV_OUT_VALUE_CELL_NUM
    private int IV_CLOSE_QUANTITY_CELL_NUM
    private int IV_CLOSE_VALUE_CELL_NUM
    private int IV_CLOSE_UNIT_PRICE_CELL_NUM

    InventoryReport(InventoryReader inventoryReader,
                    String lastInventoryReportPath) {
        super(null, null, "#,##0")

        this.lastInventoryReportPath = lastInventoryReportPath
        this.inventoryReader = inventoryReader
        outStockContentsByItemLabelByDate = inventoryReader.getOutStockContentsByItemLabelByDate()
        inStockContentsByItemLabelByDate = inventoryReader.getInStockContentsByItemLabelByDate()

        int rowNum = -1
        TITLE_ROW_NUM = ++rowNum
        SUMMARY_ROW_NUM = ++rowNum
        HEADER_FIRST_ROW_NUM = ++rowNum
        HEADER_SECOND_ROW_NUM = ++rowNum
        HEADER_ID_ROW_NUM = ++rowNum
        DATA_START_ROW_NUM = ++rowNum

        /////////////////////////// IN OUT SHEETS
        int cellNum = -1
        VOUCHER_ID_CELL_NUM = ++cellNum
        DATE_CELL_NUM = ++cellNum
        CUSTOMER_CELL_NUM = ++cellNum
        CUSTOMER_ID_CELL_NUM = ++cellNum
        CONTENT_CELL_NUM = ++cellNum
        SKU_CELL_NUM = ++cellNum
        ITEM_LABEL_CELL_NUM = ++cellNum
        UOM_CELL_NUM = ++cellNum
        IN_QUANTITY_CELL_NUM = ++cellNum
        IN_UNIT_PRICE_CELL_NUM = ++cellNum
        IN_VALUE_CELL_NUM = ++cellNum
        OUT_QUANTITY_CELL_NUM = ++cellNum
        OUT_UNIT_PRICE_CELL_NUM = ++cellNum
        OUT_VALUE_CELL_NUM = ++cellNum
        SOLD_QUANTITY_CELL_NUM = ++cellNum
        SOLD_UNIT_PRICE_CELL_NUM = ++cellNum
        SOLD_VALUE_CELL_NUM = ++cellNum
        ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM = ++cellNum
        ACCUMULATE_INVENTORY_VALUE_CELL_NUM = ++cellNum
        DATA_END_CELL_NUM = cellNum

        //////////////////////////// INVENTORY SHEETS
        cellNum = -1
        IV_SKU_CELL_NUM = ++cellNum
        IV_ITEM_LABEL_CELL_NUM = ++cellNum
        IV_UOM_CELL_NUM = ++cellNum
        IV_OPEN_QUANTITY_CELL_NUM = ++cellNum
        IV_OPEN_VALUE_CELL_NUM = ++cellNum
        IV_IN_QUANTITY_CELL_NUM = ++cellNum
        IV_IN_VALUE_CELL_NUM = ++cellNum
        IV_OUT_QUANTITY_CELL_NUM = ++cellNum
        IV_OUT_VALUE_CELL_NUM = ++cellNum
        IV_CLOSE_QUANTITY_CELL_NUM = ++cellNum
        IV_CLOSE_VALUE_CELL_NUM = ++cellNum
        IV_CLOSE_UNIT_PRICE_CELL_NUM = ++cellNum

        inventoryDataBySku = new TreeMap<>()
        accumulateInventoryDataBySku = [ : ]
        currentDataRowBySheetName = [
                (TRADE_GOODS_IN_OUT_SHEET_NAME) : DATA_START_ROW_NUM,
                (MATERIAL_IN_OUT_SHEET_NAME) : DATA_START_ROW_NUM,
                (TRADE_GOODS_INVENTORY_SHEET_NAME) : DATA_START_ROW_NUM,
                (MATERIAL_INVENTORY_SHEET_NAME) : DATA_START_ROW_NUM
        ]
    }

    @Override
    void writeIntoWorkBook() {
        initiateInventoryDataBySku()
        createSheets()

        writeInOutSheetHeaders(MATERIAL_IN_OUT_SHEET)
        writeInOutSheetHeaders(TRADE_GOODS_IN_OUT_SHEET)
        writeInOutData()

        writeInventorySheetHeader(MATERIAL_INVENTORY_SHEET)
        writeInventorySheetHeader(TRADE_GOODS_INVENTORY_SHEET)
        writeInventoryData()

        addSumRows()
        addFilters()
    }

    private void initiateInventoryDataBySku() {

    }

    private void createSheets() {
        TRADE_GOODS_IN_OUT_SHEET = createSheets(TRADE_GOODS_IN_OUT_SHEET_NAME, true)
        TRADE_GOODS_INVENTORY_SHEET = createSheets(TRADE_GOODS_INVENTORY_SHEET_NAME, true)
        MATERIAL_IN_OUT_SHEET = createSheets(MATERIAL_IN_OUT_SHEET_NAME, true)
        MATERIAL_INVENTORY_SHEET = createSheets(MATERIAL_INVENTORY_SHEET_NAME, true)
    }

    private XSSFSheet createSheets(String sheetName,
                                   boolean isLandscape) {
        XSSFSheet sheet = workbook.createSheet(sheetName)

        sheet.setMargin(PageMargin.TOP, 0.3)
        sheet.setMargin(PageMargin.BOTTOM, 0.3)
        sheet.setMargin(PageMargin.LEFT, 0.3)
        sheet.setMargin(PageMargin.RIGHT, 0.3)

        PrintSetup printSetup = sheet.getPrintSetup();

        printSetup.setFitWidth((short) 1)
        printSetup.setFitHeight((short) 0)

        printSetup.setLandscape(isLandscape)

        sheet.setFitToPage(true)
        sheet.setHorizontallyCenter(true)

        if (sheetName in [ TRADE_GOODS_IN_OUT_SHEET_NAME, MATERIAL_IN_OUT_SHEET_NAME ]) {
            sheet.setColumnWidth(VOUCHER_ID_CELL_NUM, 5 * 256)
            sheet.setColumnWidth(DATE_CELL_NUM, 12 * 256)
            sheet.setColumnWidth(CUSTOMER_CELL_NUM, 30 * 256)
            sheet.setColumnWidth(CUSTOMER_ID_CELL_NUM, 14 * 256)
            sheet.setColumnWidth(CONTENT_CELL_NUM, 13 * 256)
            sheet.setColumnWidth(SKU_CELL_NUM, 10 * 256)
            sheet.setColumnWidth(ITEM_LABEL_CELL_NUM, 20 * 256)
            sheet.setColumnWidth(UOM_CELL_NUM, 8 * 256)

            int numberCellSize = 16
            sheet.setColumnWidth(IN_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IN_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IN_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(OUT_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(OUT_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(OUT_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(SOLD_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(SOLD_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(SOLD_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(ACCUMULATE_INVENTORY_VALUE_CELL_NUM, numberCellSize * 256)
        }

        if (sheetName in [ TRADE_GOODS_INVENTORY_SHEET_NAME, MATERIAL_INVENTORY_SHEET_NAME ]) {
            sheet.setColumnWidth(IV_SKU_CELL_NUM, 50 * 256)
            sheet.setColumnWidth(IV_ITEM_LABEL_CELL_NUM, 50 * 256)
            sheet.setColumnWidth(IV_UOM_CELL_NUM, 8 * 256)

            int numberCellSize = 16
            sheet.setColumnWidth(IV_OPEN_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_OPEN_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_IN_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_IN_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_OUT_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_OUT_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_CLOSE_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_CLOSE_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(IV_CLOSE_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
        }

        return sheet
    }

    private void writeInOutSheetHeaders(Sheet sheet) {
        XSSFTools.styleMergeCells(
                sheet,
                0, TITLE_ROW_NUM, 0, ACCUMULATE_INVENTORY_VALUE_CELL_NUM,
                TITLE_STYLE,
                sheet == TRADE_GOODS_IN_OUT_SHEET ? "BẢNG KÊ NHẬP XUẤT HÀNG HÓA" : "BẢNG KÊ NHẬP XUẤT NGUYÊN VẬT LIỆU",
                HEADER_ROW_HEIGHT
        )

        //////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, VOUCHER_ID_CELL_NUM, DATE_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Chứng từ",
                HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, VOUCHER_ID_CELL_NUM, WRAP_BOLD_STYLE, "Số hóa đơn", HEADER_ROW_HEIGHT)
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, DATE_CELL_NUM, WRAP_BOLD_STYLE, "Ngày tháng")

        //////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, CUSTOMER_CELL_NUM, CUSTOMER_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Tên doanh nghiệp"
        )
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, CUSTOMER_ID_CELL_NUM, CUSTOMER_ID_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Mã số thuế"
        )

        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, CONTENT_CELL_NUM, CONTENT_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Nội dung"
        )
        ////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, SKU_CELL_NUM, SKU_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Mã vật tư"
        )
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, ITEM_LABEL_CELL_NUM, ITEM_LABEL_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Tên vật tư"
        )
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, UOM_CELL_NUM, UOM_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Đơn vị"
        )

        /////////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, IN_QUANTITY_CELL_NUM, IN_VALUE_CELL_NUM,
                BLUE_FILL_BOLD_STYLE,
                "Nhập trong kỳ"
        )
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, IN_QUANTITY_CELL_NUM, BLUE_FILL_BOLD_STYLE, "Số lượng nhập")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, IN_UNIT_PRICE_CELL_NUM, BLUE_FILL_BOLD_STYLE, "Đơn giá nhập")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, IN_VALUE_CELL_NUM, BLUE_FILL_BOLD_STYLE, "Thành tiền")

        /////////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, OUT_QUANTITY_CELL_NUM, OUT_VALUE_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE,
                "Xuất trong kỳ - GHI NHẬN GIÁ VỐN"
        )
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, OUT_QUANTITY_CELL_NUM, ORANGE_FILL_BOLD_STYLE, "Số lượng xuất")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, OUT_UNIT_PRICE_CELL_NUM, ORANGE_FILL_BOLD_STYLE, "Đơn giá xuất")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, OUT_VALUE_CELL_NUM, ORANGE_FILL_BOLD_STYLE, "Thành tiền giá vốn")

        /////////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, SOLD_QUANTITY_CELL_NUM, SOLD_VALUE_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE,
                "Xuất trong kỳ - GHI NHẬN GIÁ BÁN"
        )
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, SOLD_QUANTITY_CELL_NUM, YELLOW_FILL_BOLD_STYLE, "Số lượng bán")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, SOLD_UNIT_PRICE_CELL_NUM, YELLOW_FILL_BOLD_STYLE, "Đơn giá bán")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, SOLD_VALUE_CELL_NUM, YELLOW_FILL_BOLD_STYLE, "Thành tiền giá bán")

        /////////////////////////
        XSSFTools.styleMergeCells(
                sheet,
                HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, ACCUMULATE_INVENTORY_VALUE_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Kiểm tồn cuối kỳ"
        )
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, WRAP_BOLD_STYLE, "Số lượng")
        XSSFTools.setCellValue(sheet, HEADER_SECOND_ROW_NUM, ACCUMULATE_INVENTORY_VALUE_CELL_NUM, WRAP_BOLD_STYLE, "Thành tiền")
    }

    private void writeInOutData() {
        while (inStockContentsByItemLabelByDate || outStockContentsByItemLabelByDate) {
            ZonedDateTime inDate = inStockContentsByItemLabelByDate ? inStockContentsByItemLabelByDate?.firstKey() : null
            ZonedDateTime outDate = outStockContentsByItemLabelByDate ? outStockContentsByItemLabelByDate?.firstKey() : null

            if (inDate <= outDate && inStockContentsByItemLabelByDate) {
                Map<String, List<StockContent>> inStockContentsByItemLabel = inStockContentsByItemLabelByDate.pollFirstEntry().getValue()
                writeInOutData(inStockContentsByItemLabel, true)
            }
            else if (outStockContentsByItemLabelByDate) {
                Map<String, List<StockContent>> outStockContentsByItemLabel = outStockContentsByItemLabelByDate.pollFirstEntry().getValue()
                writeInOutData(outStockContentsByItemLabel, false)
            }
            else {
                Map<String, List<StockContent>> inStockContentsByItemLabel = inStockContentsByItemLabelByDate.pollFirstEntry().getValue()
                writeInOutData(inStockContentsByItemLabel, true)
            }
        }
    }

    private void writeInOutData(Map<String, List<StockContent>> stockContentsByItemLabel, boolean isInStock) {
        stockContentsByItemLabel.each { String itemLabel, List<StockContent> stockContents ->
            stockContents.each { StockContent stockContent ->
                boolean isProcessedGoods = checkIfIsProcessedGoods(stockContent)
                if (isProcessedGoods) {
                    // TODO: extend later
                    return
                }
                boolean isTradeGoods = checkIfIsTradeGoods(stockContent)
                String sheetName = isTradeGoods ? TRADE_GOODS_IN_OUT_SHEET_NAME : MATERIAL_IN_OUT_SHEET_NAME
                XSSFSheet sheet = isTradeGoods ? TRADE_GOODS_IN_OUT_SHEET : MATERIAL_IN_OUT_SHEET

                int rowNum = currentDataRowBySheetName?.getAt(sheetName)
                writeGeneralData(sheet, stockContent, rowNum)

                if (isInStock) {
                    writeInData(sheet, stockContent, rowNum)
                }
                else {
                    writeOutData(sheet, stockContent, rowNum)
                }

                currentDataRowBySheetName[ sheetName ] = ++rowNum
            }
        }
    }

    boolean checkIfIsProcessedGoods(StockContent stockContent) {
        return checkIfIsProcessedGoods(stockContent.getProductMaster())
    }

    boolean checkIfIsProcessedGoods(ProductMaster productMaster) {
        return productMaster.getItemFamily() == ProductMaster.FAMILY_PROCESSED_GOODS
    }

    boolean checkIfIsTradeGoods(StockContent stockContent) {
        return checkIfIsTradeGoods(stockContent.getProductMaster())
    }

    boolean checkIfIsTradeGoods(ProductMaster productMaster) {
        return productMaster.getItemFamily() == ProductMaster.FAMILY_TRADE_GOODS
    }

    private void writeGeneralData(XSSFSheet sheet, StockContent stockContent, int rowNum) {
        XSSFTools.setCellValue(sheet, rowNum, VOUCHER_ID_CELL_NUM, WRAP_STYLE, stockContent.getStockEntry().getVoucherId(), HEADER_ROW_HEIGHT)
        XSSFTools.setCellValue(sheet, rowNum, DATE_CELL_NUM, WRAP_STYLE, DateTimeResolver.getDateString(stockContent.getStockEntry().getDate()))
        XSSFTools.setCellValue(sheet, rowNum, CUSTOMER_CELL_NUM, WRAP_STYLE, stockContent.getStockEntry().getCustomerMaster().getCustomer())
        XSSFTools.setCellValue(sheet, rowNum, CUSTOMER_ID_CELL_NUM, WRAP_STYLE, stockContent.getStockEntry().getCustomerMaster().getCustomerId())
        XSSFTools.setCellValue(sheet, rowNum, SKU_CELL_NUM, WRAP_STYLE, stockContent.getProductMaster().getSku())
        XSSFTools.setCellValue(sheet, rowNum, ITEM_LABEL_CELL_NUM, WRAP_STYLE, stockContent.getProductMaster().getLabel())
        XSSFTools.setCellValue(sheet, rowNum, UOM_CELL_NUM, WRAP_STYLE, stockContent.getProductMaster().getUom())
    }

    private void writeInData(XSSFSheet sheet, StockContent stockContent, int rowNum) {
        XSSFTools.setCellValue(sheet, rowNum, CONTENT_CELL_NUM, WRAP_STYLE, "Nhập hàng hóa TM")

        BigDecimal quantity = stockContent.getQuantity()
        BigDecimal unitPrice = stockContent.getUnitPrice()
        BigDecimal value = stockContent.getValue()
        String sku = stockContent.getProductMaster().getSku()

        XSSFTools.setCellValue(sheet, rowNum, IN_QUANTITY_CELL_NUM, BLUE_FILL_NUMBER_STYLE, quantity)
        XSSFTools.setCellValue(sheet, rowNum, IN_UNIT_PRICE_CELL_NUM, BLUE_FILL_NUMBER_STYLE, unitPrice)
        XSSFTools.setCellValue(sheet, rowNum, IN_VALUE_CELL_NUM, BLUE_FILL_NUMBER_STYLE, value)

        Map accumulateInventoryData = accumulateInventoryDataBySku.get(sku, [
                (AI_QUANTITY) : 0.0,
                (AI_SUM_VALUE): 0.0
        ])
        accumulateInventoryData[ AI_QUANTITY ] += quantity
        accumulateInventoryData[ AI_SUM_VALUE ] += value

        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, WRAP_NUMBER_STYLE, accumulateInventoryData[ AI_QUANTITY ])
        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_VALUE_CELL_NUM, WRAP_NUMBER_STYLE, accumulateInventoryData[ AI_SUM_VALUE ])

        XSSFTools.setCellValue(sheet, rowNum, OUT_QUANTITY_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, OUT_UNIT_PRICE_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, OUT_VALUE_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, "")

        XSSFTools.setCellValue(sheet, rowNum, SOLD_QUANTITY_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, SOLD_UNIT_PRICE_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, SOLD_VALUE_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, "")

        Map inventoryData = inventoryDataBySku.get(sku, [
                (ID_IN_QUANTITY) : 0.0,
                (ID_IN_VALUE): 0.0,
                (ID_OUT_QUANTITY) : 0.0,
                (ID_OUT_VALUE): 0.0
        ])
        inventoryData[ ID_IN_QUANTITY ] += quantity
        inventoryData[ ID_IN_VALUE ] += value
    }

    private void writeOutData(XSSFSheet sheet, StockContent stockContent, int rowNum) {
        XSSFTools.setCellValue(sheet, rowNum, CONTENT_CELL_NUM, WRAP_STYLE, "Xuất bán hàng TM")

        BigDecimal quantity = stockContent.getQuantity()
        BigDecimal unitPrice = stockContent.getUnitPrice()
        BigDecimal value = stockContent.getValue()
        String sku = stockContent.getProductMaster().getSku()

        XSSFTools.setCellValue(sheet, rowNum, SOLD_QUANTITY_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, quantity)
        XSSFTools.setCellValue(sheet, rowNum, SOLD_UNIT_PRICE_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, unitPrice)
        XSSFTools.setCellValue(sheet, rowNum, SOLD_VALUE_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, value)

        Map accumulateInventoryData = accumulateInventoryDataBySku[ sku ] as Map
        BigDecimal accumulativeQuantity = (accumulateInventoryData?.getAt(AI_QUANTITY) as BigDecimal) ?: 0.0
        BigDecimal accumulativeValue = (accumulateInventoryData?.getAt(AI_SUM_VALUE) as BigDecimal) ?: 0.0
        BigDecimal outUnitPrice = accumulativeValue && accumulativeQuantity ? accumulativeValue / accumulativeQuantity : 0.0
        BigDecimal outValue = outUnitPrice * quantity

        XSSFTools.setCellValue(sheet, rowNum, OUT_QUANTITY_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, quantity)
        XSSFTools.setCellValue(sheet, rowNum, OUT_UNIT_PRICE_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, outUnitPrice)
        XSSFTools.setCellValue(sheet, rowNum, OUT_VALUE_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, outValue)

        accumulativeQuantity -= quantity
        accumulativeValue -= outValue

        if (accumulateInventoryData) {
            accumulateInventoryData << [
                    (AI_QUANTITY) : accumulativeQuantity,
                    (AI_SUM_VALUE): accumulativeValue
            ]
        }

        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, WRAP_NUMBER_STYLE, accumulativeQuantity)
        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_VALUE_CELL_NUM, WRAP_NUMBER_STYLE, accumulativeValue)

        XSSFTools.setCellValue(sheet, rowNum, IN_QUANTITY_CELL_NUM, BLUE_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, IN_UNIT_PRICE_CELL_NUM, BLUE_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, IN_VALUE_CELL_NUM, BLUE_FILL_NUMBER_STYLE, "")

        Map inventoryData = inventoryDataBySku.get(sku, [
                (ID_IN_QUANTITY) : 0.0,
                (ID_IN_VALUE): 0.0,
                (ID_OUT_QUANTITY) : 0.0,
                (ID_OUT_VALUE): 0.0
        ])
        inventoryData[ ID_OUT_QUANTITY ] += quantity
        inventoryData[ ID_OUT_VALUE ] += outValue
    }

    /////////////////////////////
    private void writeInventorySheetHeader(XSSFSheet sheet) {
        XSSFTools.styleMergeCells(
                sheet, TITLE_ROW_NUM, TITLE_ROW_NUM, IV_SKU_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM,
                TITLE_STYLE, sheet == TRADE_GOODS_INVENTORY_SHEET ? "BẢNG TỔNG HỢP NHẬP XUẤT TỒN HÀNG HÓA" : "BẢNG TỔNG HỢP NHẬP XUẤT TỒN NGUYÊN VẬT LIỆU",
                HEADER_ROW_HEIGHT
        )

        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, IV_SKU_CELL_NUM, IV_SKU_CELL_NUM,
                WRAP_BOLD_STYLE, "Mã Vật Tư"
        )
        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, IV_ITEM_LABEL_CELL_NUM, IV_ITEM_LABEL_CELL_NUM,
                WRAP_BOLD_STYLE, "Tên Vật Tư"
        )
        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, IV_UOM_CELL_NUM, IV_UOM_CELL_NUM,
                WRAP_BOLD_STYLE, "Đơn Vị"
        )

        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, IV_OPEN_QUANTITY_CELL_NUM, IV_OPEN_VALUE_CELL_NUM,
                GREY_FILL_BOLD_STYLE, "Tồn Đầu Kỳ", HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_OPEN_QUANTITY_CELL_NUM,
                GREY_FILL_BOLD_STYLE, "Số lượng"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_OPEN_VALUE_CELL_NUM,
                GREY_FILL_BOLD_STYLE, "Thành tiền"
        )

        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, IV_IN_QUANTITY_CELL_NUM, IV_IN_VALUE_CELL_NUM,
                BLUE_FILL_BOLD_STYLE, "Nhập Trong Kỳ"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_IN_QUANTITY_CELL_NUM,
                BLUE_FILL_BOLD_STYLE, "Số lượng"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_IN_VALUE_CELL_NUM,
                BLUE_FILL_BOLD_STYLE, "Thành tiền"
        )

        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, IV_OUT_QUANTITY_CELL_NUM, IV_OUT_VALUE_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE, "Xuất Trong Kỳ"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_OUT_QUANTITY_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE, "Số lượng"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_OUT_VALUE_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE, "Thành tiền"
        )

        XSSFTools.styleMergeCells(
                sheet, HEADER_FIRST_ROW_NUM, HEADER_FIRST_ROW_NUM, IV_CLOSE_QUANTITY_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE, "Tồn Cuối Kỳ"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_CLOSE_QUANTITY_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE, "Số lượng"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_CLOSE_VALUE_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE, "Thành tiền"
        )
        XSSFTools.setCellValue(
                sheet, HEADER_SECOND_ROW_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE, "Đơn giá"
        )
    }

    private void writeInventoryData() {
        inventoryDataBySku.each { String sku, Map inventoryData ->
            ProductMaster productMaster = inventoryReader.getProductMasterBySku()?.getAt(sku)
            if (!productMaster) {
                println(" - Could not Write Inventory Data for Sku: ${sku} :: NULL ProductMaster")
                return
            }
            boolean isTradeGoods = checkIfIsTradeGoods(productMaster)
            String sheetName = isTradeGoods ? TRADE_GOODS_INVENTORY_SHEET_NAME : MATERIAL_INVENTORY_SHEET_NAME
            XSSFSheet sheet = isTradeGoods ? TRADE_GOODS_INVENTORY_SHEET : MATERIAL_INVENTORY_SHEET

            Integer rowNum = currentDataRowBySheetName[ sheetName ] as Integer
            writeInventoryData(sheet, inventoryData, productMaster, rowNum)

            currentDataRowBySheetName[ sheetName ] = ++rowNum
        }
    }

    private void writeInventoryData(XSSFSheet sheet,
                                    Map inventoryData,
                                    ProductMaster productMaster,
                                    Integer rowNum) {
        BigDecimal openQuantity = (inventoryData[ ID_OPEN_QUANTITY ] ?: 0.0) as BigDecimal
        BigDecimal openValue = (inventoryData[ ID_OPEN_VALUE ] ?: 0.0) as BigDecimal
        BigDecimal inQuantity = (inventoryData[ ID_IN_QUANTITY ] ?: 0.0) as BigDecimal
        BigDecimal inValue = (inventoryData[ ID_IN_VALUE ] ?: 0.0) as BigDecimal
        BigDecimal outQuantity = (inventoryData[ ID_OUT_QUANTITY ] ?: 0.0) as BigDecimal
        BigDecimal outValue = (inventoryData[ ID_OUT_VALUE ] ?: 0.0) as BigDecimal
        BigDecimal closeQuantity = openQuantity + inQuantity - outQuantity
        BigDecimal closeValue = openValue + inValue - outValue
        BigDecimal closeUnitPrice = closeValue && closeQuantity ? closeValue / closeQuantity : 0.0

        XSSFTools.setCellValue(
                sheet, rowNum, IV_SKU_CELL_NUM, WRAP_NUMBER_STYLE, productMaster.getSku(), HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_ITEM_LABEL_CELL_NUM, WRAP_NUMBER_STYLE, productMaster.getLabel()
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_UOM_CELL_NUM, WRAP_NUMBER_STYLE, productMaster.getUom()
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_OPEN_QUANTITY_CELL_NUM, GREY_FILL_NUMBER_STYLE, openQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_OPEN_VALUE_CELL_NUM, GREY_FILL_NUMBER_STYLE, openValue
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_IN_QUANTITY_CELL_NUM, BLUE_FILL_NUMBER_STYLE, inQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_IN_VALUE_CELL_NUM, BLUE_FILL_NUMBER_STYLE, inValue
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_OUT_QUANTITY_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, outQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_OUT_VALUE_CELL_NUM, ORANGE_FILL_NUMBER_STYLE, outValue
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_CLOSE_QUANTITY_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, closeQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_CLOSE_VALUE_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, closeValue
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_CLOSE_UNIT_PRICE_CELL_NUM, YELLOW_FILL_NUMBER_STYLE, closeUnitPrice
        )
    }

    ////////////////////////////
    private void addSumRows() {
        addSumRow(
                TRADE_GOODS_IN_OUT_SHEET,
                (IN_QUANTITY_CELL_NUM..SOLD_VALUE_CELL_NUM),
                [
                        IN_UNIT_PRICE_CELL_NUM, OUT_UNIT_PRICE_CELL_NUM, SOLD_UNIT_PRICE_CELL_NUM,
                        IN_QUANTITY_CELL_NUM, OUT_QUANTITY_CELL_NUM, SOLD_QUANTITY_CELL_NUM
                ] as Set
        )
        addSumRow(
                MATERIAL_IN_OUT_SHEET,
                (IN_QUANTITY_CELL_NUM..SOLD_VALUE_CELL_NUM),
                [
                        IN_UNIT_PRICE_CELL_NUM, OUT_UNIT_PRICE_CELL_NUM, SOLD_UNIT_PRICE_CELL_NUM,
                        IN_QUANTITY_CELL_NUM, OUT_QUANTITY_CELL_NUM, SOLD_QUANTITY_CELL_NUM
                ] as Set
        )

        addSumRow(
                TRADE_GOODS_INVENTORY_SHEET,
                (IV_OPEN_QUANTITY_CELL_NUM..IV_CLOSE_UNIT_PRICE_CELL_NUM),
                [
                        IV_OPEN_QUANTITY_CELL_NUM, IV_IN_QUANTITY_CELL_NUM, IV_OUT_QUANTITY_CELL_NUM,
                        IV_CLOSE_QUANTITY_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM
                ] as Set
        )
        addSumRow(
                MATERIAL_INVENTORY_SHEET,
                (IV_OPEN_QUANTITY_CELL_NUM..IV_CLOSE_UNIT_PRICE_CELL_NUM),
                [
                        IV_OPEN_QUANTITY_CELL_NUM, IV_IN_QUANTITY_CELL_NUM, IV_OUT_QUANTITY_CELL_NUM,
                        IV_CLOSE_QUANTITY_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM
                ] as Set
        )
    }

    private void addSumRow(XSSFSheet sheet, List fromCellNumToCellNum, Set excludedCellNums) {
        Integer firstExcludedInARowCellNum = null

        fromCellNumToCellNum.each { Integer cellNum ->
            Row dataRow = XSSFTools.getOrSetRow(sheet, DATA_START_ROW_NUM)
            Cell dataCell = XSSFTools.getOrSetCell(dataRow, cellNum)
            CellStyle cellStyle = workbook.createCellStyle()
            cellStyle.cloneStyleFrom(dataCell.getCellStyle())
            cellStyle.setFont(BOLD_FONT)

            Cell sumCell = XSSFTools.setCellValue(
                    sheet, SUMMARY_ROW_NUM, cellNum,
                    cellStyle, "", HEADER_ROW_HEIGHT
            )
            if (cellNum in excludedCellNums) {
                if (!firstExcludedInARowCellNum) {
                    firstExcludedInARowCellNum = cellNum
                }
                return
            }

            if (firstExcludedInARowCellNum && cellNum - 1 == firstExcludedInARowCellNum) {
                XSSFTools.setCellValue(sheet, SUMMARY_ROW_NUM, cellNum - 1, cellStyle, "SUM")
                firstExcludedInARowCellNum = null
            }
            else if (firstExcludedInARowCellNum && cellNum - 1 > firstExcludedInARowCellNum) {
                XSSFTools.styleMergeCells(sheet, SUMMARY_ROW_NUM, SUMMARY_ROW_NUM, firstExcludedInARowCellNum, cellNum - 1, cellStyle, "SUM")
                firstExcludedInARowCellNum = null
            }

            String sheetName = sheet.getSheetName()
            String cellLetter = CellReference.convertNumToColString(cellNum)
            Integer endRowNum = currentDataRowBySheetName[ sheetName ] as Integer
            String range = "${cellLetter}${DATA_START_ROW_NUM}:${cellLetter}${endRowNum}"
            sumCell.setCellFormula("SUM(${range})")
        }
    }

    private void addFilters() {
        TRADE_GOODS_IN_OUT_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_FIRST_ROW_NUM, currentDataRowBySheetName[ TRADE_GOODS_IN_OUT_SHEET_NAME ] as Integer,
                        VOUCHER_ID_CELL_NUM, ACCUMULATE_INVENTORY_VALUE_CELL_NUM
                )
        )

        MATERIAL_IN_OUT_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_FIRST_ROW_NUM, currentDataRowBySheetName[ MATERIAL_IN_OUT_SHEET_NAME ] as Integer,
                        VOUCHER_ID_CELL_NUM, ACCUMULATE_INVENTORY_VALUE_CELL_NUM
                )
        )

        TRADE_GOODS_INVENTORY_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_FIRST_ROW_NUM, currentDataRowBySheetName[ TRADE_GOODS_INVENTORY_SHEET_NAME ] as Integer,
                        IV_SKU_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM
                )
        )

        MATERIAL_INVENTORY_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_FIRST_ROW_NUM, currentDataRowBySheetName[ MATERIAL_INVENTORY_SHEET_NAME ] as Integer,
                        IV_SKU_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM
                )
        )
    }
}
