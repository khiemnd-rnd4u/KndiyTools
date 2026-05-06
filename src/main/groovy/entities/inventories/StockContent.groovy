package entities.inventories

class StockContent {
    private ProductMaster productMaster
    private String uom
    private BigDecimal quantity
    private BigDecimal unitPrice
    private BigDecimal value
    private BigDecimal tax
    private BigDecimal valueWithTax
    private BigDecimal taxRate

    private StockEntry stockEntry

    StockContent(ProductMaster productMaster,
                 String uom,
                 BigDecimal quantity,
                 BigDecimal unitPrice,
                 BigDecimal value,
                 BigDecimal tax,
                 BigDecimal valueWithTax,
                 BigDecimal taxRate,
                 StockEntry stockEntry) {
        this.productMaster = productMaster
        this.uom = uom
        this.quantity = quantity
        this.unitPrice = unitPrice
        this.value = value
        this.tax = tax
        this.valueWithTax = valueWithTax
        this.taxRate = taxRate
        this.stockEntry = stockEntry
    }

    ProductMaster getProductMaster() {
        return productMaster
    }

    String getUom() {
        return uom
    }

    BigDecimal getQuantity() {
        return quantity
    }

    BigDecimal getUnitPrice() {
        return unitPrice
    }

    BigDecimal getValue() {
        return value
    }

    BigDecimal getTax() {
        return tax
    }

    BigDecimal getValueWithTax() {
        return valueWithTax
    }

    BigDecimal getTaxRate() {
        return taxRate
    }

    StockEntry getStockEntry() {
        return stockEntry
    }

    @Override
    public String toString() {
        return "StockContent{" +
                "productMaster=" + productMaster +
                ", uom='" + uom + '\'' +
                ", quantity=" + quantity +
                ", unitPrice=" + unitPrice +
                ", value=" + value +
                ", tax=" + tax +
                ", valueWithTax=" + valueWithTax +
                ", taxRate=" + taxRate +
                '}';
    }
}
