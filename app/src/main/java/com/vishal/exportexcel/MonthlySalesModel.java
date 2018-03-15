package com.vishal.exportexcel;

/**
 * Created by vishal.halani on 15-Mar-18.
 */

public class MonthlySalesModel {
    private String monthName;
    private  int sales;

    public MonthlySalesModel(String monthName, int sales) {
        this.monthName = monthName;
        this.sales = sales;
    }

    public String getMonthName() {
        return monthName;
    }

    public void setMonthName(String monthName) {
        this.monthName = monthName;
    }

    public int getSales() {
        return sales;
    }

    public void setSales(int sales) {
        this.sales = sales;
    }
}
