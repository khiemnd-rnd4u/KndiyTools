package entities.inventories

class ProductMaster {
    private String sku
    private String label
    private String itemFamily
    private String uom
    private String processingTarget

    static final String FAMILY_NVL = "Nguyên vật liệu"
    static final String FAMILY_TRADE_GOODS = "Thương mại"
    static final String FAMILY_PROCESSED_GOODS = "Phí gia công"

    ProductMaster(String sku,
                  String label,
                  String itemFamily,
                  String uom,
                  String processingTarget = null) {
        this.sku = sku
        this.label = label
        this.itemFamily = itemFamily
        this.uom = uom
        this.processingTarget = processingTarget
    }

    String getSku() {
        return sku
    }

    String getLabel() {
        return label
    }

    String getItemFamily() {
        return itemFamily
    }

    String getUom() {
        return uom
    }

    String getProcessingTarget() {
        return processingTarget
    }

    @Override
    public String toString() {
        return "ProductMaster{" +
                "sku='" + sku + '\'' +
                ", label='" + label + '\'' +
                ", itemFamily='" + itemFamily + '\'' +
                ", uom='" + uom + '\'' +
                '}'
    }
}
