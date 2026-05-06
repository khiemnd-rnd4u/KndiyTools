package entities.inventories

import java.time.ZonedDateTime

class CustomerMaster {
    private String customerId
    private String customer
    private String address

    private TreeMap<String, TreeMap<String, StockEntry>> stockEntryByVoucherIdByType

    CustomerMaster(String customerId,
                   String customer,
                   String address) {
        this.customerId = customerId
        this.customer = customer
        this.address = address

        stockEntryByVoucherIdByType = new TreeMap<>()
    }

    StockEntry getOrSetStockEntry(String voucherId,
                                  String type,
                                  ZonedDateTime date) {
        StockEntry stockEntry = stockEntryByVoucherIdByType?.getAt(type)?.getAt(voucherId)
        if (!stockEntry) {
            stockEntry = new StockEntry(voucherId, type, date, this)
            TreeMap<String, StockEntry> stockEntryById = stockEntryByVoucherIdByType.get(type, new TreeMap<>())
            stockEntryById[ voucherId ] = stockEntry
        }

        return stockEntry
    }

    String getCustomerId() {
        return customerId
    }

    String getCustomer() {
        return customer
    }

    String getAddress() {
        return address
    }


    @Override
    public String toString() {
        return "CustomerMaster{" +
                "customerId='" + customerId + '\'' +
                ", customer='" + customer + '\'' +
                ", address='" + address + '\'' +
                ", stockEntryByVoucherIdByType=" + stockEntryByVoucherIdByType +
                '}';
    }
}
