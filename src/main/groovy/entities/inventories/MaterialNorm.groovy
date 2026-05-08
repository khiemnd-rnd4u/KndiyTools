package entities.inventories

class MaterialNorm {
    private ProductMaster targetProductMaster
    private Map<String, BigDecimal> normAmountByMaterialSku
    private Map<String, ProductMaster> materialProductMasterByMaterialSku

    MaterialNorm(ProductMaster targetProductMaster) {
        this.targetProductMaster = targetProductMaster

        normAmountByMaterialSku = [ : ]
        materialProductMasterByMaterialSku = [ : ]
    }

    void addMaterialComponent(ProductMaster materialProductMaster,
                              BigDecimal normAmount) {
        String sku = materialProductMaster.getSku()
        normAmountByMaterialSku[ sku ] = normAmount
        materialProductMasterByMaterialSku[ sku ] = materialProductMaster
    }

    ProductMaster getTargetProductMaster() {
        return targetProductMaster
    }

    Map<String, BigDecimal> getNormAmountByMaterialSku() {
        return normAmountByMaterialSku
    }

    Map<String, ProductMaster> getMaterialProductMasterByMaterialSku() {
        return materialProductMasterByMaterialSku
    }
}
