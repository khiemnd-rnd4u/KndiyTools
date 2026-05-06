package entities.inventories

class ProductMaster {
    private String sku
    private String label
    private String itemFamily
    private String uom

    static final String FAMILY_NVL = "Nguyên vật liệu"
    static final String FAMILY_TRADE_GOODS = "Thương mại"
    static final String FAMILY_PROCESSED_GOODS = "Phí gia công"

    ProductMaster(String sku,
                  String label,
                  String itemFamily,
                  String uom) {
        this.sku = sku
        this.label = label
        this.itemFamily = itemFamily
        this.uom = uom
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
