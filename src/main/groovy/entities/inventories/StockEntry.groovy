package entities.inventories

import java.time.ZonedDateTime

class StockEntry {
    static TYPE_STOCK_IN = "StockIn"
    static TYPE_STOCK_OUT = "StockOut"

    private String voucherId
    private ZonedDateTime date
    private TreeMap<String, StockContent> stockContentByItemLabel
    private String type

    private BigDecimal totalValue
    private BigDecimal totalTax
    private BigDecimal totalValueWithTax

    private CustomerMaster customerMaster

    StockEntry(String voucherId,
               String type,
               ZonedDateTime date,
               CustomerMaster customerMaster) {
        this.voucherId = voucherId
        this.type = type
        this.date = date
        stockContentByItemLabel = new TreeMap<>()
        this.customerMaster = customerMaster

        this.totalValue = 0.0
        this.totalTax = 0.0
        this.totalValueWithTax = 0.0
    }

    StockContent addStockContent(ProductMaster productMaster,
                                 String uom,
                                 BigDecimal quantity,
                                 BigDecimal unitPrice,
                                 BigDecimal value,
                                 BigDecimal tax,
                                 BigDecimal valueWithTax,
                                 BigDecimal taxRate) {
        StockContent stockContent = getStockContent(productMaster.getLabel())
        if (stockContent) {
            return stockContent
        }

        stockContent = new StockContent(
                productMaster, uom, quantity, unitPrice, value, tax, valueWithTax, taxRate, this
        )

        stockContentByItemLabel[ productMaster.getLabel() ] = stockContent

        this.totalValue += value
        this.totalTax += tax
        this.totalValueWithTax += valueWithTax

        return stockContent
    }

    StockContent getStockContent(String itemLabel) {
        return stockContentByItemLabel?.getAt(itemLabel)
    }

    String getVoucherId() {
        return voucherId
    }

    TreeMap<String, StockContent> getStockContentByItemLabel() {
        return stockContentByItemLabel
    }

    String getType() {
        return type
    }

    ZonedDateTime getDate() {
        return date
    }

    BigDecimal getTotalValue() {
        return totalValue
    }

    BigDecimal getTotalTax() {
        return totalTax
    }

    BigDecimal getTotalValueWithTax() {
        return totalValueWithTax
    }

    CustomerMaster getCustomerMaster() {
        return customerMaster
    }

    @Override
    public String toString() {
        return "StockEntry{" +
                "voucherId='" + voucherId + '\'' +
                ", date=" + date +
                ", stockContentByItemLabel=" + stockContentByItemLabel +
                ", type='" + type + '\'' +
                ", totalValue=" + totalValue +
                ", totalTax=" + totalTax +
                ", totalValueWithTax=" + totalValueWithTax +
                '}';
    }
}
