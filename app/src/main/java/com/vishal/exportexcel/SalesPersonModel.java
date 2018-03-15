package com.vishal.exportexcel;

import java.util.ArrayList;

/**
 * Created by vishal.halani on 15-Mar-18.
 */

public class SalesPersonModel {
    private String name;
    private ArrayList<MonthlySalesModel> monthlySaleslist = new ArrayList<>();

    public SalesPersonModel(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ArrayList<MonthlySalesModel> getMonthlySaleslist() {
        return monthlySaleslist;
    }

    public void setMonthlySaleslist(ArrayList<MonthlySalesModel> monthlySaleslist) {
        this.monthlySaleslist = monthlySaleslist;
    }
}
