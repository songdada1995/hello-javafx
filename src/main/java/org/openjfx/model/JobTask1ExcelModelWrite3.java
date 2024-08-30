package org.openjfx.model;

import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.converters.bigdecimal.BigDecimalNumberConverter;
import lombok.Data;

import java.math.BigDecimal;

/**
 * @author songbo
 * @version 1.0
 * @date 2023/11/10 17:10
 */
@ExcelIgnoreUnannotated
@Data
public class JobTask1ExcelModelWrite3 {

    @ExcelProperty(index = 0, value = "Invoice ID")
    private String invoiceID;

    @ExcelProperty(index = 1, value = "Order Date")
    private String orderDate;

    @ExcelProperty(index = 2, value = "Ship Date")
    private String shipDate;

    @ExcelProperty(index = 3, value = "Return Date")
    private String returnDate;

    @ExcelProperty(index = 4, value = "Cost Date")
    private String costDate;

    @ExcelProperty(index = 5, value = "Transaction Type")
    private String transactionType;

    @ExcelProperty(index = 6, value = "Quantity")
    private Integer quantity;

    @NumberFormat("0.00")
    @ExcelProperty(index = 7, value = "Net Sales", converter = BigDecimalNumberConverter.class)
    private BigDecimal netSales;

    @ExcelProperty(index = 8, value = "Net Sales Currency")
    private String netSalesCurrency;

    @NumberFormat("0.00")
    @ExcelProperty(index = 9, value = "List Price", converter = BigDecimalNumberConverter.class)
    private BigDecimal listPrice;

    @ExcelProperty(index = 10, value = "List Price Currency")
    private String listPriceCurrency;

    @NumberFormat("0.00")
    @ExcelProperty(index = 11, value = "Rebate In Agreement Currency", converter = BigDecimalNumberConverter.class)
    private BigDecimal rebateInAgreementCurrency;

    @ExcelProperty(index = 12, value = "Agreement Currency")
    private String agreementCurrency;

    @NumberFormat("0.00")
    @ExcelProperty(index = 13, value = "Rebate In Purchase Order Currency", converter = BigDecimalNumberConverter.class)
    private BigDecimal rebateInPurchaseOrderCurrency;

    @ExcelProperty(index = 14, value = "Purchase Order Currency")
    private String purchaseOrderCurrency;

    @ExcelProperty(index = 15, value = "Purchase Order")
    private String purchaseOrder;

    @ExcelProperty(index = 16, value = "Asin")
    private String asin;

    @ExcelProperty(index = 17, value = "UPC")
    private String upc;

    @ExcelProperty(index = 18, value = "EAN")
    private String ean;

    @ExcelProperty(index = 19, value = "Manufacturer")
    private String manufacturer;

    @ExcelProperty(index = 20, value = "Distributor")
    private String distributor;

    @ExcelProperty(index = 21, value = "Product Group")
    private String productGroup;

    @ExcelProperty(index = 22, value = "Category")
    private String category;

    @ExcelProperty(index = 23, value = "Subcategory")
    private String subcategory;

    @ExcelProperty(index = 24, value = "Title")
    private String title;

    @ExcelProperty(index = 25, value = "Product Description")
    private String productDescription;

    @ExcelProperty(index = 26, value = "Binding")
    private String binding;

    @ExcelProperty(index = 27, value = "Promotion Id")
    private String promotionId;

    @ExcelProperty(index = 28, value = "Cost Type")
    private String costType;

    @ExcelProperty(index = 29, value = "Order Country")
    private String orderCountry;

    @ExcelProperty(index = 30, value = "Multi-Country Parent Agreement ID")
    private String multiCountryParentAgreementId;

    @ExcelProperty(index = 31, value = "Marketplace")
    private String marketplace;

    @ExcelProperty(index = 32, value = "Merchant Id")
    private String merchantId;

    @ExcelProperty(index = 33, value = "Our Price (Per Unit)")
    private String ourPricePerUnit;
}
