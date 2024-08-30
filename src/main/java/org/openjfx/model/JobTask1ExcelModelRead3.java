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
public class JobTask1ExcelModelRead3 {

    @ExcelProperty("Order Date")
    private String orderDate;

    @ExcelProperty("Ship Date")
    private String shipDate;

    @ExcelProperty("Return Date")
    private String returnDate;

    @ExcelProperty("Cost Date")
    private String costDate;

    @ExcelProperty("Transaction Type")
    private String transactionType;

    @ExcelProperty("Quantity")
    private Integer quantity;

    @NumberFormat("0.00")
    @ExcelProperty(value = "Net Sales", converter = BigDecimalNumberConverter.class)
    private BigDecimal netSales;

    @ExcelProperty("Net Sales Currency")
    private String netSalesCurrency;

    @NumberFormat("0.00")
    @ExcelProperty(value = "List Price", converter = BigDecimalNumberConverter.class)
    private BigDecimal listPrice;

    @ExcelProperty("List Price Currency")
    private String listPriceCurrency;

    @NumberFormat("0.00")
    @ExcelProperty(value = "Rebate In Agreement Currency", converter = BigDecimalNumberConverter.class)
    private BigDecimal rebateInAgreementCurrency;

    @ExcelProperty("Agreement Currency")
    private String agreementCurrency;

    @NumberFormat("0.00")
    @ExcelProperty(value = "Rebate In Purchase Order Currency", converter = BigDecimalNumberConverter.class)
    private BigDecimal rebateInPurchaseOrderCurrency;

    @ExcelProperty("Purchase Order Currency")
    private String purchaseOrderCurrency;

    @ExcelProperty("Purchase Order")
    private String purchaseOrder;

    @ExcelProperty("Asin")
    private String asin;

    @ExcelProperty("UPC")
    private String upc;

    @ExcelProperty("EAN")
    private String ean;

    @ExcelProperty("Manufacturer")
    private String manufacturer;

    @ExcelProperty("Distributor")
    private String distributor;

    @ExcelProperty("Product Group")
    private String productGroup;

    @ExcelProperty("Category")
    private String category;

    @ExcelProperty("Subcategory")
    private String subcategory;

    @ExcelProperty("Title")
    private String title;

    @ExcelProperty("Product Description")
    private String productDescription;

    @ExcelProperty("Binding")
    private String binding;

    @ExcelProperty("Promotion Id")
    private String promotionId;

    @ExcelProperty("Cost Type")
    private String costType;

    @ExcelProperty("Order Country")
    private String orderCountry;

    @ExcelProperty("Multi-Country Parent Agreement ID")
    private String multiCountryParentAgreementId;

    @ExcelProperty("Marketplace")
    private String marketplace;

    @ExcelProperty("Merchant Id")
    private String merchantId;

    @ExcelProperty("Our Price (Per Unit)")
    private String ourPricePerUnit;

}
