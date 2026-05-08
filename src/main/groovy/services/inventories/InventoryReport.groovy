package services.inventories

import entities.inventories.MaterialNorm
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import services.kndiyLibraries.DateTimeResolver
import services.kndiyLibraries.KnDiyWorkbook
import services.kndiyLibraries.XSSFTools

import java.math.RoundingMode
import java.time.ZonedDateTime

class InventoryReport extends KnDiyWorkbook{
    private String TRADE_GOODS_IN_OUT_SHEET_NAME = "Nhập xuất HH"
    private String TRADE_GOODS_INVENTORY_SHEET_NAME = "Tồn HH"
    private String MATERIAL_IN_OUT_SHEET_NAME = "Nhập xuất NVL"
    private String MATERIAL_INVENTORY_SHEET_NAME = "Tồn NVL"

    private String NORM_USAGE_SHEET_NAME = "Định mưc Sử dụng"

    private String lastInventoryReportPath

    private XSSFWorkbook lastInventoryReport

    private XSSFSheet TRADE_GOODS_IN_OUT_SHEET
    private XSSFSheet TRADE_GOODS_INVENTORY_SHEET
    private XSSFSheet MATERIAL_IN_OUT_SHEET
    private XSSFSheet MATERIAL_INVENTORY_SHEET

    private XSSFSheet NORM_USAGE_SHEET

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

    private ZonedDateTime fromDate = DateTimeResolver.getZonedDateTime("9999-12-31")
    private ZonedDateTime toDate = DateTimeResolver.getZonedDateTime("1000-01-01")

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

    TreeMap<ZonedDateTime, Map> materialUsageDataByMaterialSkuBySkuByDate
    static final String MD_VOUCHER_ID = "VoucherId"
    static final String MD_VOUCHER_DATE = "VoucherDate"
    static final String MD_NORM_AMOUNT = "NormAmount"
    static final String MD_INVENTORY_QUANTITY = "InventoryQuantity"
    static final String MD_INVENTORY_VALUE = "InventoryValue"
    static final String MD_INVENTORY_UNIT_PRICE = "InventoryUnitPrice"
    static final String MD_OUT_QUANTITY = "Quantity"
    static final String MD_OUT_VALUE = "Value"
    static final String MD_OUT_UNIT_PRICE = "UnitPrice"

    private int NU_IDX_CELL_NUM
    private int NU_DATE_CELL_NUM
    private int NU_VOUCHER_ID_CELL_NUM
    private int NU_ITEM_LABEL_CELL_NUM
    private int NU_SKU_CELL_NUM
    private int NU_MATERIAL_LABEL_CELL_NUM
    private int NU_MATERIAL_SKU_CELL_NUM
    private int NU_MATERIAL_UOM_CELL_NUM
    private int NU_NORM_AMOUNT_CELL_NUM
    private int NU_INVENTORY_QUANTITY_CELL_NUM
    private int NU_INVENTORY_VALUE_CELL_NUM
    private int NU_INVENTORY_UNIT_PRICE_CELL_NUM
    private int NU_OUT_QUANTITY_CELL_NUM
    private int NU_OUT_VALUE_CELL_NUM
    private int NU_OUT_UNIT_PRICE_CELL_NUM
    private int NU_CLOSE_QUANTITY_CELL_NUM
    private int NU_CLOSE_UNIT_PRICE_CELL_NUM
    private int NU_CLOSE_VALUE_CELL_NUM

    InventoryReport(InventoryReader inventoryReader,
                    String lastInventoryReportPath) {
        super(null, null, "#,##0")

        if (lastInventoryReportPath) {
            lastInventoryReport = XSSFTools.getWorkbook(lastInventoryReportPath)
        }

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

        //////////////////////// NORM USAGE SHEET
        cellNum = -1
        NU_IDX_CELL_NUM = ++cellNum
        NU_VOUCHER_ID_CELL_NUM = ++cellNum
        NU_DATE_CELL_NUM = ++cellNum
        NU_SKU_CELL_NUM = ++cellNum
        NU_ITEM_LABEL_CELL_NUM = ++cellNum
        NU_MATERIAL_LABEL_CELL_NUM = ++cellNum
        NU_MATERIAL_SKU_CELL_NUM = ++cellNum
        NU_MATERIAL_UOM_CELL_NUM = ++cellNum
        NU_NORM_AMOUNT_CELL_NUM = ++cellNum
        NU_INVENTORY_QUANTITY_CELL_NUM = ++cellNum
        NU_INVENTORY_UNIT_PRICE_CELL_NUM = ++cellNum
        NU_INVENTORY_VALUE_CELL_NUM = ++cellNum
        NU_OUT_QUANTITY_CELL_NUM = ++cellNum
        NU_OUT_UNIT_PRICE_CELL_NUM = ++cellNum
        NU_OUT_VALUE_CELL_NUM = ++cellNum
        NU_CLOSE_QUANTITY_CELL_NUM = ++cellNum
        NU_CLOSE_UNIT_PRICE_CELL_NUM = ++cellNum
        NU_CLOSE_VALUE_CELL_NUM = ++cellNum

        //////////////////////////
        inventoryDataBySku = new TreeMap<>()
        accumulateInventoryDataBySku = [ : ]
        currentDataRowBySheetName = [
                (TRADE_GOODS_IN_OUT_SHEET_NAME) : DATA_START_ROW_NUM,
                (MATERIAL_IN_OUT_SHEET_NAME) : DATA_START_ROW_NUM,
                (TRADE_GOODS_INVENTORY_SHEET_NAME) : DATA_START_ROW_NUM,
                (MATERIAL_INVENTORY_SHEET_NAME) : DATA_START_ROW_NUM,
                (NORM_USAGE_SHEET_NAME) : DATA_START_ROW_NUM
        ]

        materialUsageDataByMaterialSkuBySkuByDate = new TreeMap<>((ZonedDateTime d1, ZonedDateTime d2) -> { return d1 <=> d2 })
    }

    @Override
    void writeIntoWorkBook() {
        initiateInventoryData()
        createSheets()

        writeInOutSheetHeaders(MATERIAL_IN_OUT_SHEET)
        writeInOutSheetHeaders(TRADE_GOODS_IN_OUT_SHEET)
        writeInOutData()

        writeInventorySheetHeaders(MATERIAL_INVENTORY_SHEET)
        writeInventorySheetHeaders(TRADE_GOODS_INVENTORY_SHEET)
        writeInventoryData()

        writeNormUsageHeaders()
        writeNormUsageData()

        addSumRows()
        addHeaderIdRows()
        addFilters()
    }

    private void initiateInventoryData() {
        if (!lastInventoryReport) {
            return
        }
        initiateInventoryData(lastInventoryReport.getSheet(TRADE_GOODS_INVENTORY_SHEET_NAME))
        initiateInventoryData(lastInventoryReport.getSheet(MATERIAL_INVENTORY_SHEET_NAME))
    }

    private void initiateInventoryData(XSSFSheet sheet) {
        for (Row row : sheet) {
            int rowNum = row.getRowNum()
            if (rowNum < DATA_START_ROW_NUM) {
                continue
            }

            String sku = XSSFTools.readStringFromCell(sheet, rowNum, IV_SKU_CELL_NUM)
            BigDecimal quantity = XSSFTools.readNumberFromCell(sheet, rowNum, IV_CLOSE_QUANTITY_CELL_NUM)
            BigDecimal value = XSSFTools.readNumberFromCell(sheet, rowNum, IV_CLOSE_VALUE_CELL_NUM)

            inventoryDataBySku[ sku ] = [
                    (ID_OPEN_QUANTITY) : quantity,
                    (ID_OPEN_VALUE) : value,
                    (ID_IN_QUANTITY) : 0.0,
                    (ID_IN_VALUE) : 0.0,
                    (ID_OUT_QUANTITY) : 0.0,
                    (ID_OUT_VALUE) : 0.0,
            ]
            accumulateInventoryDataBySku[ sku ] = [
                    (AI_QUANTITY) : quantity,
                    (AI_SUM_VALUE) : value
            ]
        }
    }

    private void createSheets() {
        TRADE_GOODS_IN_OUT_SHEET = createSheets(TRADE_GOODS_IN_OUT_SHEET_NAME, true)
        TRADE_GOODS_INVENTORY_SHEET = createSheets(TRADE_GOODS_INVENTORY_SHEET_NAME, true)
        MATERIAL_IN_OUT_SHEET = createSheets(MATERIAL_IN_OUT_SHEET_NAME, true)
        MATERIAL_INVENTORY_SHEET = createSheets(MATERIAL_INVENTORY_SHEET_NAME, true)
        NORM_USAGE_SHEET = createSheets(NORM_USAGE_SHEET_NAME, true)
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

        if (sheetName == NORM_USAGE_SHEET_NAME) {
            sheet.setColumnWidth(NU_IDX_CELL_NUM, 5 * 256)
            sheet.setColumnWidth(NU_VOUCHER_ID_CELL_NUM, 5 * 256)
            sheet.setColumnWidth(NU_DATE_CELL_NUM, 12 * 256)
            sheet.setColumnWidth(NU_SKU_CELL_NUM, 10 * 256)
            sheet.setColumnWidth(NU_ITEM_LABEL_CELL_NUM, 20 * 256)
            sheet.setColumnWidth(NU_MATERIAL_SKU_CELL_NUM, 10 * 256)
            sheet.setColumnWidth(NU_MATERIAL_LABEL_CELL_NUM, 20 * 256)
            sheet.setColumnWidth(NU_MATERIAL_UOM_CELL_NUM, 8 * 256)

            int numberCellSize = 16
            sheet.setColumnWidth(NU_NORM_AMOUNT_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_INVENTORY_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_INVENTORY_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_INVENTORY_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_OUT_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_OUT_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_OUT_VALUE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_CLOSE_QUANTITY_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_CLOSE_UNIT_PRICE_CELL_NUM, numberCellSize * 256)
            sheet.setColumnWidth(NU_CLOSE_VALUE_CELL_NUM, numberCellSize * 256)
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

            fromDate = DateTimeResolver.getMinDate([ fromDate, inDate, outDate ])
            toDate = DateTimeResolver.getMaxDate([ toDate, inDate, outDate ])

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
                    writeProcessedGoodsData(stockContent)
                    return
                }
                boolean isTradeGoods = checkIfIsTradeGoods(stockContent)
                String sheetName = isTradeGoods ? TRADE_GOODS_IN_OUT_SHEET_NAME : MATERIAL_IN_OUT_SHEET_NAME
                XSSFSheet sheet = isTradeGoods ? TRADE_GOODS_IN_OUT_SHEET : MATERIAL_IN_OUT_SHEET

                int rowNum = currentDataRowBySheetName?.getAt(sheetName)
                writeGeneralData(sheet, stockContent, rowNum)

                if (isInStock) {
                    writeInData(sheet, stockContent, rowNum, isTradeGoods)
                }
                else {
                    writeOutData(sheet, stockContent, rowNum, isTradeGoods)
                }

                currentDataRowBySheetName[ sheetName ] = ++rowNum
            }
        }
    }

    private void writeProcessedGoodsData(StockContent stockContent) {
        String targetSku = stockContent.getProductMaster().getProcessingTarget()
        ZonedDateTime date = stockContent.getStockEntry().getDate()
        MaterialNorm materialNorm = inventoryReader.getMaterialNormByTargetSkuByDate()?.getAt(date)?.getAt(targetSku)
        if (!materialNorm) {
            println("Missing Material Norm for Sku: ${targetSku} on ${DateTimeResolver.getDateTimeString(date)}")
            return
        }
        ProductMaster targetProductMaster = materialNorm.getTargetProductMaster()
        Map materialUsageDataByMaterialSkuBySku = materialUsageDataByMaterialSkuBySkuByDate.get(date, [ : ])
        Map materialUsageDataByMaterialSku = materialUsageDataByMaterialSkuBySku.get(targetSku, [ : ])

        materialNorm.getNormAmountByMaterialSku().each { String materialSku, BigDecimal normAmount ->
            ProductMaster materialProductMaster = materialNorm.getMaterialProductMasterByMaterialSku()[ materialSku ]
            writeOutMaterialDataForProcessing(
                    stockContent,
                    targetProductMaster, materialProductMaster,
                    normAmount,
                    materialUsageDataByMaterialSku
            )
        }

        writeInProcessedGoodsData(stockContent, targetProductMaster, materialUsageDataByMaterialSku)
    }

    private void writeOutMaterialDataForProcessing(StockContent stockContent,
                                                   ProductMaster targetProductMaster,
                                                   ProductMaster materialProductMaster,
                                                   BigDecimal normAmount,
                                                   Map materialUsageDataByMaterialSku) {
        String materialSku = materialProductMaster.getSku()
        String materialUom = materialProductMaster.getUom()
        BigDecimal quantity = stockContent.getQuantity()
        BigDecimal processFee = stockContent.getValue()
        BigDecimal processUnitPrice = stockContent.getUnitPrice()

        Map accumulateMaterialData = accumulateInventoryDataBySku[ materialSku ] as Map
        BigDecimal accumulativeMaterialQuantity = (accumulateMaterialData?.getAt(AI_QUANTITY) as BigDecimal) ?: 0.0
        BigDecimal accumulativeMaterialValue = (accumulateMaterialData?.getAt(AI_SUM_VALUE) as BigDecimal) ?: 0.0
        BigDecimal accumulativeMaterialUnitPrice = accumulativeMaterialValue && accumulativeMaterialQuantity
                ? accumulativeMaterialValue / accumulativeMaterialQuantity
                : 0.0
        BigDecimal outMaterialQuantity = normAmount * quantity
        if (checkIfIsWholeUom(materialUom)) {
            outMaterialQuantity = outMaterialQuantity.setScale(0, RoundingMode.HALF_UP)
        }
        BigDecimal outMaterialValue = accumulativeMaterialUnitPrice * outMaterialQuantity
        BigDecimal normUnitPrice = accumulativeMaterialUnitPrice * normAmount

        //////////////////////// UPDATE HOLDERS
        materialUsageDataByMaterialSku[ materialSku ] = [
                (MD_VOUCHER_ID) : stockContent.getStockEntry().getVoucherId(),
                (MD_VOUCHER_DATE) : stockContent.getStockEntry().getDate(),
                (MD_INVENTORY_QUANTITY) : accumulativeMaterialQuantity,
                (MD_INVENTORY_VALUE) : accumulativeMaterialValue,
                (MD_INVENTORY_UNIT_PRICE) : accumulativeMaterialUnitPrice,
                (MD_NORM_AMOUNT) : normAmount,
                (MD_OUT_QUANTITY) : outMaterialQuantity,
                (MD_OUT_VALUE) : outMaterialValue,
                (MD_OUT_UNIT_PRICE): normUnitPrice
        ]

        accumulativeMaterialQuantity -= outMaterialQuantity
        accumulativeMaterialValue -= outMaterialValue
        if (accumulateMaterialData) {
            accumulateMaterialData << [
                    (AI_QUANTITY) : accumulativeMaterialQuantity,
                    (AI_SUM_VALUE) : accumulativeMaterialValue
            ]
        }

        Map inventoryData = inventoryDataBySku.get(materialSku, [
                (ID_IN_QUANTITY) : 0.0,
                (ID_IN_VALUE): 0.0,
                (ID_OUT_QUANTITY) : 0.0,
                (ID_OUT_VALUE): 0.0
        ])
        inventoryData[ ID_OUT_QUANTITY ] += outMaterialQuantity
        inventoryData[ ID_OUT_VALUE ] += outMaterialValue

        //////////////////////// WRITE TO SHEET
        Integer rowNum = currentDataRowBySheetName[ MATERIAL_IN_OUT_SHEET_NAME ] as Integer
        writeGeneralData(
                MATERIAL_IN_OUT_SHEET, rowNum,
                stockContent.getStockEntry().getVoucherId(),
                DateTimeResolver.getDateString(stockContent.getStockEntry().getDate()),
                stockContent.getStockEntry().getCustomerMaster().getCustomer(),
                stockContent.getStockEntry().getCustomerMaster().getCustomerId(),
                materialProductMaster.getSku(),
                materialProductMaster.getLabel(),
                materialProductMaster.getUom()
        )

        writeOutData(
                MATERIAL_IN_OUT_SHEET, rowNum, false, targetProductMaster.getLabel(), materialProductMaster.getUom(),
                outMaterialQuantity, accumulativeMaterialUnitPrice, outMaterialValue,
                outMaterialQuantity, accumulativeMaterialUnitPrice, outMaterialValue,
                accumulativeMaterialQuantity, accumulativeMaterialValue
        )

        currentDataRowBySheetName[ MATERIAL_IN_OUT_SHEET_NAME ] = ++rowNum
    }

    private boolean checkIfIsWholeUom(String materialUom) {
        return materialUom in [ "Bịch", "Gói","Cái" ]
    }

    private void writeInProcessedGoodsData(StockContent stockContent,
                                           ProductMaster targetProductMaster,
                                           Map materialUsageDataByMaterialSku) {
        BigDecimal quantity = stockContent.getQuantity()
        BigDecimal processFee = stockContent.getValue()
        BigDecimal processUnitPrice = stockContent.getUnitPrice()
        String sku = targetProductMaster.getSku()
        String uom = targetProductMaster.getUom()

        BigDecimal materialSumValue = 0.0
        materialUsageDataByMaterialSku.each { String materialSku, Map materialUsageData ->
            BigDecimal materialValue = (materialUsageData[ MD_OUT_VALUE ] as BigDecimal) ?: 0.0
            materialSumValue += materialValue
        }
        BigDecimal value = processFee + materialSumValue
        BigDecimal unitPrice = value && quantity ? value / quantity : 0.0

        Map accumulateInventoryData = accumulateInventoryDataBySku.get(sku, [
                (AI_QUANTITY) : 0.0,
                (AI_SUM_VALUE): 0.0
        ])
        accumulateInventoryData[ AI_QUANTITY ] += quantity
        accumulateInventoryData[ AI_SUM_VALUE ] += value

        Map inventoryData = inventoryDataBySku.get(sku, [
                (ID_IN_QUANTITY) : 0.0,
                (ID_IN_VALUE): 0.0,
                (ID_OUT_QUANTITY) : 0.0,
                (ID_OUT_VALUE): 0.0
        ])
        inventoryData[ ID_IN_QUANTITY ] += quantity
        inventoryData[ ID_IN_VALUE ] += value

        //////////////////////// WRITE TO SHEET
        Integer rowNum = currentDataRowBySheetName[ TRADE_GOODS_IN_OUT_SHEET_NAME ] as Integer
        writeGeneralData(
                TRADE_GOODS_IN_OUT_SHEET, rowNum,
                stockContent.getStockEntry().getVoucherId(),
                DateTimeResolver.getDateString(stockContent.getStockEntry().getDate()),
                stockContent.getStockEntry().getCustomerMaster().getCustomer(),
                stockContent.getStockEntry().getCustomerMaster().getCustomerId(),
                sku,
                targetProductMaster.getLabel(),
                targetProductMaster.getUom()
        )
        writeInData(
                TRADE_GOODS_IN_OUT_SHEET, rowNum, false, true, uom,
                quantity, unitPrice, value,
                accumulateInventoryData[ AI_QUANTITY ] as BigDecimal, accumulateInventoryData[ AI_SUM_VALUE ] as BigDecimal
        )

        currentDataRowBySheetName[ TRADE_GOODS_IN_OUT_SHEET_NAME ] = ++rowNum
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
        writeGeneralData(
                sheet, rowNum,
                stockContent.getStockEntry().getVoucherId(),
                DateTimeResolver.getDateString(stockContent.getStockEntry().getDate()),
                stockContent.getStockEntry().getCustomerMaster().getCustomer(),
                stockContent.getStockEntry().getCustomerMaster().getCustomerId(),
                stockContent.getProductMaster().getSku(),
                stockContent.getProductMaster().getLabel(),
                stockContent.getProductMaster().getUom()
        )
    }

    private void writeGeneralData(Sheet sheet, int rowNum,
                                  String voucherId, String date,
                                  String customerName, String customerId,
                                  String sku, String label, String uom) {
        XSSFTools.setCellValue(sheet, rowNum, VOUCHER_ID_CELL_NUM, WRAP_STYLE, voucherId, HEADER_ROW_HEIGHT)
        XSSFTools.setCellValue(sheet, rowNum, DATE_CELL_NUM, WRAP_STYLE, date)
        XSSFTools.setCellValue(sheet, rowNum, CUSTOMER_CELL_NUM, WRAP_STYLE, customerName)
        XSSFTools.setCellValue(sheet, rowNum, CUSTOMER_ID_CELL_NUM, WRAP_STYLE, customerId)
        XSSFTools.setCellValue(sheet, rowNum, SKU_CELL_NUM, WRAP_STYLE, sku)
        XSSFTools.setCellValue(sheet, rowNum, ITEM_LABEL_CELL_NUM, WRAP_STYLE, label)
        XSSFTools.setCellValue(sheet, rowNum, UOM_CELL_NUM, WRAP_STYLE, uom)
    }

    private void writeInData(XSSFSheet sheet, StockContent stockContent, int rowNum, boolean isTradeGoods) {
        BigDecimal quantity = stockContent.getQuantity()
        BigDecimal unitPrice = stockContent.getUnitPrice()
        BigDecimal value = stockContent.getValue()
        String sku = stockContent.getProductMaster().getSku()
        String uom = stockContent.getProductMaster().getUom()

        Map accumulateInventoryData = accumulateInventoryDataBySku.get(sku, [
                (AI_QUANTITY) : 0.0,
                (AI_SUM_VALUE): 0.0
        ])
        accumulateInventoryData[ AI_QUANTITY ] += quantity
        accumulateInventoryData[ AI_SUM_VALUE ] += value

        Map inventoryData = inventoryDataBySku.get(sku, [
                (ID_IN_QUANTITY) : 0.0,
                (ID_IN_VALUE): 0.0,
                (ID_OUT_QUANTITY) : 0.0,
                (ID_OUT_VALUE): 0.0
        ])
        inventoryData[ ID_IN_QUANTITY ] += quantity
        inventoryData[ ID_IN_VALUE ] += value

        writeInData(
                sheet, rowNum, isTradeGoods, false, uom,
                quantity, unitPrice, value,
                accumulateInventoryData[ AI_QUANTITY ] as BigDecimal,
                accumulateInventoryData[ AI_SUM_VALUE ] as BigDecimal
        )
    }

    private void writeInData(XSSFSheet sheet, int rowNum, boolean isTradeGoods, boolean isProcessed, String uom,
                             BigDecimal quantity, BigDecimal unitPrice, BigDecimal value,
                             BigDecimal accumulativeQuantity, BigDecimal accumulativeValue) {
        String content = isProcessed
                ? "Nhập hàng hoá gia công TM"
                : isTradeGoods
                        ? "Nhập hàng hóa TM"
                        : "Mua NVL"

        boolean isWholeUom = checkIfIsWholeUom(uom)

        XSSFTools.setCellValue(sheet, rowNum, CONTENT_CELL_NUM, WRAP_STYLE, content)

        XSSFTools.setCellValue(sheet, rowNum, IN_QUANTITY_CELL_NUM, isWholeUom ? BLUE_FILL_STYLE : BLUE_FILL_NUMBER_STYLE, quantity)
        XSSFTools.setCellValue(sheet, rowNum, IN_UNIT_PRICE_CELL_NUM, BLUE_FILL_STYLE, unitPrice)
        XSSFTools.setCellValue(sheet, rowNum, IN_VALUE_CELL_NUM, BLUE_FILL_STYLE, value)

        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, isWholeUom ? WRAP_STYLE : WRAP_NUMBER_STYLE, accumulativeQuantity)
        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_VALUE_CELL_NUM, WRAP_STYLE, accumulativeValue)

        XSSFTools.setCellValue(sheet, rowNum, OUT_QUANTITY_CELL_NUM, isWholeUom ? ORANGE_FILL_STYLE : ORANGE_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, OUT_UNIT_PRICE_CELL_NUM, ORANGE_FILL_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, OUT_VALUE_CELL_NUM, ORANGE_FILL_STYLE, "")

        XSSFTools.setCellValue(sheet, rowNum, SOLD_QUANTITY_CELL_NUM, isWholeUom ? YELLOW_FILL_STYLE : YELLOW_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, SOLD_UNIT_PRICE_CELL_NUM, YELLOW_FILL_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, SOLD_VALUE_CELL_NUM, YELLOW_FILL_STYLE, "")
    }

    private void writeOutData(XSSFSheet sheet, StockContent stockContent, int rowNum, boolean isTradeGoods) {
        BigDecimal quantity = stockContent.getQuantity()
        BigDecimal unitPrice = stockContent.getUnitPrice()
        BigDecimal value = stockContent.getValue()
        String sku = stockContent.getProductMaster().getSku()
        String uom = stockContent.getProductMaster().getUom()

        Map accumulateInventoryData = accumulateInventoryDataBySku[ sku ] as Map
        BigDecimal accumulativeQuantity = (accumulateInventoryData?.getAt(AI_QUANTITY) as BigDecimal) ?: 0.0
        BigDecimal accumulativeValue = (accumulateInventoryData?.getAt(AI_SUM_VALUE) as BigDecimal) ?: 0.0
        BigDecimal outUnitPrice = accumulativeValue && accumulativeQuantity ? accumulativeValue / accumulativeQuantity : 0.0
        BigDecimal outValue = outUnitPrice * quantity

        accumulativeQuantity -= quantity
        accumulativeValue -= outValue

        if (accumulateInventoryData) {
            accumulateInventoryData << [
                    (AI_QUANTITY) : accumulativeQuantity,
                    (AI_SUM_VALUE): accumulativeValue
            ]
        }

        Map inventoryData = inventoryDataBySku.get(sku, [
                (ID_IN_QUANTITY) : 0.0,
                (ID_IN_VALUE): 0.0,
                (ID_OUT_QUANTITY) : 0.0,
                (ID_OUT_VALUE): 0.0
        ])
        inventoryData[ ID_OUT_QUANTITY ] += quantity
        inventoryData[ ID_OUT_VALUE ] += outValue

        writeOutData(
                sheet, rowNum, isTradeGoods, null, uom,
                quantity, unitPrice, value,
                quantity, outUnitPrice, outValue,
                accumulativeQuantity, accumulativeValue
        )
    }

    private void writeOutData(Sheet sheet, int rowNum, boolean isTradeGoods, String processedTargetLabel, String uom,
                              BigDecimal soldQuantity, BigDecimal soldUnitPrice, BigDecimal soldValue,
                              BigDecimal outQuantity, BigDecimal outUnitPrice, BigDecimal outValue,
                              BigDecimal accumulativeQuantity, BigDecimal accumulativeValue) {
        String content = processedTargetLabel
                ? "Xuất NVL gia công (${processedTargetLabel})"
                :  isTradeGoods
                        ? "Xuất bán hàng TM"
                        : "Xuất NVL"

        boolean isWholeUom = checkIfIsWholeUom(uom)

        XSSFTools.setCellValue(sheet, rowNum, CONTENT_CELL_NUM, WRAP_STYLE, content)

        XSSFTools.setCellValue(sheet, rowNum, SOLD_QUANTITY_CELL_NUM, isWholeUom ? YELLOW_FILL_STYLE : YELLOW_FILL_NUMBER_STYLE, soldQuantity)
        XSSFTools.setCellValue(sheet, rowNum, SOLD_UNIT_PRICE_CELL_NUM, YELLOW_FILL_STYLE, soldUnitPrice)
        XSSFTools.setCellValue(sheet, rowNum, SOLD_VALUE_CELL_NUM, YELLOW_FILL_STYLE, soldValue)

        XSSFTools.setCellValue(sheet, rowNum, OUT_QUANTITY_CELL_NUM, isWholeUom ? ORANGE_FILL_STYLE : ORANGE_FILL_NUMBER_STYLE, outQuantity)
        XSSFTools.setCellValue(sheet, rowNum, OUT_UNIT_PRICE_CELL_NUM, ORANGE_FILL_STYLE, outUnitPrice)
        XSSFTools.setCellValue(sheet, rowNum, OUT_VALUE_CELL_NUM, ORANGE_FILL_STYLE, outValue)

        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_QUANTITY_CELL_NUM, isWholeUom ? WRAP_STYLE : WRAP_NUMBER_STYLE, accumulativeQuantity)
        XSSFTools.setCellValue(sheet, rowNum, ACCUMULATE_INVENTORY_VALUE_CELL_NUM, WRAP_STYLE, accumulativeValue)

        XSSFTools.setCellValue(sheet, rowNum, IN_QUANTITY_CELL_NUM, isWholeUom ? BLUE_FILL_STYLE : BLUE_FILL_NUMBER_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, IN_UNIT_PRICE_CELL_NUM, BLUE_FILL_STYLE, "")
        XSSFTools.setCellValue(sheet, rowNum, IN_VALUE_CELL_NUM, BLUE_FILL_STYLE, "")
    }

    /////////////////////////////
    private void writeInventorySheetHeaders(XSSFSheet sheet) {
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

        String uom = productMaster.getUom()
        boolean isWholeUom = checkIfIsWholeUom(uom)

        XSSFTools.setCellValue(
                sheet, rowNum, IV_SKU_CELL_NUM, WRAP_STYLE, productMaster.getSku(), HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_ITEM_LABEL_CELL_NUM, WRAP_STYLE, productMaster.getLabel()
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_UOM_CELL_NUM, WRAP_STYLE, productMaster.getUom()
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_OPEN_QUANTITY_CELL_NUM, isWholeUom ? GREY_FILL_STYLE : GREY_FILL_NUMBER_STYLE, openQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_OPEN_VALUE_CELL_NUM, GREY_FILL_STYLE, openValue
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_IN_QUANTITY_CELL_NUM, isWholeUom ? BLUE_FILL_STYLE : BLUE_FILL_NUMBER_STYLE, inQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_IN_VALUE_CELL_NUM, BLUE_FILL_STYLE, inValue
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_OUT_QUANTITY_CELL_NUM, isWholeUom ? ORANGE_FILL_STYLE : ORANGE_FILL_NUMBER_STYLE, outQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_OUT_VALUE_CELL_NUM, ORANGE_FILL_STYLE, outValue
        )

        XSSFTools.setCellValue(
                sheet, rowNum, IV_CLOSE_QUANTITY_CELL_NUM, isWholeUom ? YELLOW_FILL_STYLE : YELLOW_FILL_NUMBER_STYLE, closeQuantity
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_CLOSE_VALUE_CELL_NUM, YELLOW_FILL_STYLE, closeValue
        )
        XSSFTools.setCellValue(
                sheet, rowNum, IV_CLOSE_UNIT_PRICE_CELL_NUM, YELLOW_FILL_STYLE, closeUnitPrice
        )
    }

    /////////////////////////// NORM USAGE
    /////////////////////////// NORM USAGE
    /////////////////////////// NORM USAGE
    private void writeNormUsageHeaders() {
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, TITLE_ROW_NUM, SUMMARY_ROW_NUM, NU_IDX_CELL_NUM, NU_CLOSE_VALUE_CELL_NUM,
                WRAP_BOLD_STYLE,
                "BÁO CÁO ĐỊNH MỨC SỬ DỤNG",
                HEADER_ROW_HEIGHT / 2
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_IDX_CELL_NUM, NU_IDX_CELL_NUM,
                WRAP_BOLD_STYLE,
                "No.",
                HEADER_ROW_HEIGHT
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_VOUCHER_ID_CELL_NUM, NU_VOUCHER_ID_CELL_NUM,
                WRAP_BOLD_STYLE,
                "HĐ Số"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_DATE_CELL_NUM, NU_DATE_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Date"
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_SKU_CELL_NUM, NU_SKU_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Mã Thành Phẩm"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_ITEM_LABEL_CELL_NUM, NU_ITEM_LABEL_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Tên Thành Phẩm"
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_MATERIAL_SKU_CELL_NUM, NU_MATERIAL_SKU_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Mã NVL"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_MATERIAL_LABEL_CELL_NUM, NU_MATERIAL_LABEL_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Tên NVL"
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_MATERIAL_UOM_CELL_NUM, NU_MATERIAL_UOM_CELL_NUM,
                WRAP_BOLD_STYLE,
                "Đơn vị NVL"
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_NORM_AMOUNT_CELL_NUM, NU_NORM_AMOUNT_CELL_NUM,
                ROSE_FILL_BOLD_STYLE,
                "Định Mức"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_INVENTORY_QUANTITY_CELL_NUM, NU_INVENTORY_QUANTITY_CELL_NUM,
                GREY_FILL_BOLD_STYLE,
                "Tồn NVL"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_INVENTORY_UNIT_PRICE_CELL_NUM, NU_INVENTORY_UNIT_PRICE_CELL_NUM,
                GREY_FILL_BOLD_STYLE,
                "Đơn Giá Vốn NVL Tồn"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_INVENTORY_VALUE_CELL_NUM, NU_INVENTORY_VALUE_CELL_NUM,
                GREY_FILL_BOLD_STYLE,
                "TT NVL Tồn"
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_OUT_QUANTITY_CELL_NUM, NU_OUT_QUANTITY_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE,
                "Xuất NVL"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_OUT_UNIT_PRICE_CELL_NUM, NU_OUT_UNIT_PRICE_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE,
                "Đơn Giá Định Mức NVL / 1 Thành Phẩm"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_OUT_VALUE_CELL_NUM, NU_OUT_VALUE_CELL_NUM,
                ORANGE_FILL_BOLD_STYLE,
                "Đơn giá Xuất NVL"
        )

        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_CLOSE_QUANTITY_CELL_NUM, NU_CLOSE_QUANTITY_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE,
                "NVL Còn Lại"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_CLOSE_UNIT_PRICE_CELL_NUM, NU_CLOSE_UNIT_PRICE_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE,
                "Đơn Giá NVL Còn Lại"
        )
        XSSFTools.styleMergeCells(
                NORM_USAGE_SHEET, HEADER_FIRST_ROW_NUM, HEADER_SECOND_ROW_NUM, NU_CLOSE_VALUE_CELL_NUM, NU_CLOSE_VALUE_CELL_NUM,
                YELLOW_FILL_BOLD_STYLE,
                "TT NVL Còn Lại"
        )
    }

    private void writeNormUsageData() {
        materialUsageDataByMaterialSkuBySkuByDate.each { ZonedDateTime date, Map materialUsageDataByMaterialSkuBySku ->
            materialUsageDataByMaterialSkuBySku.each { String sku, Map materialUsageDataByMaterialSku ->
                ProductMaster productMaster = inventoryReader.getProductMasterBySku()?.getAt(sku)
                materialUsageDataByMaterialSku.each { String materialSku, Map materialUsageData ->
                    ProductMaster materialMaster = inventoryReader.getProductMasterBySku()?.getAt(materialSku)

                    Integer rowNum = currentDataRowBySheetName[ NORM_USAGE_SHEET_NAME ] as Integer
                    writeNormUsageData(rowNum, productMaster, materialMaster, materialUsageData)

                    currentDataRowBySheetName[ NORM_USAGE_SHEET_NAME ] = ++rowNum
                }
            }
        }
    }

    private void writeNormUsageData(Integer rowNum,
                                    ProductMaster productMaster,
                                    ProductMaster materialMaster,
                                    Map materialUsageData) {
        String materialUom = materialMaster.getUom()
        boolean isWholeMaterialUom = checkIfIsWholeUom(materialUom)

        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_IDX_CELL_NUM, WRAP_STYLE,
                rowNum - DATA_START_ROW_NUM + 1,
                HEADER_ROW_HEIGHT
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_VOUCHER_ID_CELL_NUM, WRAP_STYLE,
                materialUsageData[ MD_VOUCHER_ID ] as String)
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_DATE_CELL_NUM, WRAP_STYLE,
                DateTimeResolver.getDateString(materialUsageData[ MD_VOUCHER_DATE ] as ZonedDateTime)
        )

        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_SKU_CELL_NUM, WRAP_STYLE,
                productMaster.getSku()
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_ITEM_LABEL_CELL_NUM, WRAP_STYLE,
                productMaster.getLabel()
        )

        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_MATERIAL_SKU_CELL_NUM, WRAP_STYLE,
                materialMaster.getSku()
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_MATERIAL_LABEL_CELL_NUM, WRAP_STYLE,
                materialMaster.getLabel()
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_MATERIAL_UOM_CELL_NUM, WRAP_STYLE,
                materialMaster.getUom()
        )

        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_NORM_AMOUNT_CELL_NUM, ROSE_FILL_NUMBER_BOLD_STYLE,
                materialUsageData[ MD_NORM_AMOUNT ]
        )

        BigDecimal openQuantity = (materialUsageData[ MD_INVENTORY_QUANTITY ] as BigDecimal) ?: 0.0
        BigDecimal openValue = (materialUsageData[ MD_INVENTORY_VALUE ] as BigDecimal) ?: 0.0
        BigDecimal outQuantity = (materialUsageData[ MD_OUT_QUANTITY ] as BigDecimal) ?: 0.0
        BigDecimal outValue = (materialUsageData[ MD_OUT_VALUE ] as BigDecimal) ?: 0.0
        BigDecimal closeQuantity = openQuantity - outQuantity
        BigDecimal closeValue = openValue - outValue
        BigDecimal closeUnitPrice = closeValue && closeQuantity ? closeValue / closeQuantity : 0.0

        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_INVENTORY_QUANTITY_CELL_NUM, isWholeMaterialUom ?  GREY_FILL_STYLE : GREY_FILL_NUMBER_STYLE,
                openQuantity
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_INVENTORY_UNIT_PRICE_CELL_NUM, GREY_FILL_STYLE,
                materialUsageData[ MD_INVENTORY_UNIT_PRICE ]
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_INVENTORY_VALUE_CELL_NUM, GREY_FILL_STYLE,
                openValue
        )


        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_OUT_QUANTITY_CELL_NUM, isWholeMaterialUom ? ORANGE_FILL_STYLE : ORANGE_FILL_NUMBER_STYLE,
                outQuantity
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_OUT_UNIT_PRICE_CELL_NUM, ORANGE_FILL_STYLE,
                materialUsageData[ MD_OUT_UNIT_PRICE ]
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_OUT_VALUE_CELL_NUM, ORANGE_FILL_STYLE,
                outValue
        )

        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_CLOSE_QUANTITY_CELL_NUM, isWholeMaterialUom ? YELLOW_FILL_STYLE : YELLOW_FILL_NUMBER_STYLE,
                closeQuantity
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_CLOSE_UNIT_PRICE_CELL_NUM, YELLOW_FILL_STYLE,
                closeUnitPrice
        )
        XSSFTools.setCellValue(
                NORM_USAGE_SHEET, rowNum, NU_CLOSE_VALUE_CELL_NUM, YELLOW_FILL_STYLE,
                closeValue
        )
    }

    //////////////////////////// SUM ROWS
    //////////////////////////// SUM ROWS
    //////////////////////////// SUM ROWS
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

    //////////////////// HEADER ID ROWS
    //////////////////// HEADER ID ROWS
    //////////////////// HEADER ID ROWS
    private void addHeaderIdRows() {
        addHeaderIdRows(TRADE_GOODS_IN_OUT_SHEET)
        addHeaderIdRows(TRADE_GOODS_INVENTORY_SHEET)
        addHeaderIdRows(MATERIAL_IN_OUT_SHEET)
        addHeaderIdRows(MATERIAL_INVENTORY_SHEET)
        addHeaderIdRows(NORM_USAGE_SHEET)
    }

    private void addHeaderIdRows(Sheet sheet) {
        Row headerRow = sheet.getRow(HEADER_SECOND_ROW_NUM)
        for (Cell cell : headerRow) {
            Integer cellNum = cell.getColumnIndex()
            XSSFTools.setCellValue(sheet, HEADER_ID_ROW_NUM, cellNum, WRAP_ITALIC_STYLE, "(${cellNum + 1})")
        }
    }

    //////////////////// FILTERS
    //////////////////// FILTERS
    //////////////////// FILTERS
    private void addFilters() {
        TRADE_GOODS_IN_OUT_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_ID_ROW_NUM, currentDataRowBySheetName[ TRADE_GOODS_IN_OUT_SHEET_NAME ] as Integer,
                        VOUCHER_ID_CELL_NUM, ACCUMULATE_INVENTORY_VALUE_CELL_NUM
                )
        )

        MATERIAL_IN_OUT_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_ID_ROW_NUM, currentDataRowBySheetName[ MATERIAL_IN_OUT_SHEET_NAME ] as Integer,
                        VOUCHER_ID_CELL_NUM, ACCUMULATE_INVENTORY_VALUE_CELL_NUM
                )
        )

        TRADE_GOODS_INVENTORY_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_ID_ROW_NUM, currentDataRowBySheetName[ TRADE_GOODS_INVENTORY_SHEET_NAME ] as Integer,
                        IV_SKU_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM
                )
        )

        MATERIAL_INVENTORY_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_ID_ROW_NUM, currentDataRowBySheetName[ MATERIAL_INVENTORY_SHEET_NAME ] as Integer,
                        IV_SKU_CELL_NUM, IV_CLOSE_UNIT_PRICE_CELL_NUM
                )
        )

        NORM_USAGE_SHEET.setAutoFilter(
                new CellRangeAddress(
                        HEADER_ID_ROW_NUM, currentDataRowBySheetName[ NORM_USAGE_SHEET_NAME ] as Integer,
                        NU_IDX_CELL_NUM, NU_CLOSE_VALUE_CELL_NUM
                )
        )
    }

    ZonedDateTime getFromDate() {
        return fromDate
    }

    ZonedDateTime getToDate() {
        return toDate
    }
}
