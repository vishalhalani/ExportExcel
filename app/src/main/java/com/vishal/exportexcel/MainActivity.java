package com.vishal.exportexcel;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.RelativeLayout;
import android.widget.Toast;


import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import pub.devrel.easypermissions.AfterPermissionGranted;
import pub.devrel.easypermissions.EasyPermissions;

public class MainActivity extends AppCompatActivity implements EasyPermissions.PermissionCallbacks {

    private RelativeLayout layout;
    private static final int RC_EXPORT = 123;
    private ArrayList<SalesPersonModel> modelList;
    private ArrayList<String> monthList;
    private WritableCellFormat timesRegular;
    private WritableCellFormat timesCenterRegular;
    private WritableCellFormat timesHeader;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        layout = (RelativeLayout) findViewById(R.id.parent_layout);
        addDummyData();
        Button btn = (Button) findViewById(R.id.button);
        btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                excelTask();
            }
        });


    }

    private void addDummyData() {
        modelList = new ArrayList<>();
        monthList = new ArrayList<>();
        monthList.add("Jan");
        monthList.add("Feb");
        monthList.add("March");


        ArrayList<String> salesPersonName = new ArrayList<>();
        salesPersonName.add("Vishal Halani");
        salesPersonName.add("Jay Patel");
        salesPersonName.add("Akshay Shah");

        int personCounter = 0;
        for (String personName : salesPersonName) {
            SalesPersonModel model = new SalesPersonModel(personName);
            for (int i = 0; i < monthList.size(); i++) {

                if (personCounter == 0) {
                    MonthlySalesModel monthlySalesModel;
                    if (i == 0) {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 100000);
                    } else if (i == 1) {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 120000);
                    } else {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 150000);
                    }

                    model.getMonthlySaleslist().add(monthlySalesModel);
                } else if (personCounter == 1) {
                    MonthlySalesModel monthlySalesModel;
                    if (i == 0) {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 80000);
                    } else if (i == 1) {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 100000);
                    } else {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 120000);
                    }

                    model.getMonthlySaleslist().add(monthlySalesModel);
                } else {
                    MonthlySalesModel monthlySalesModel;
                    if (i == 0) {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 90000);
                    } else if (i == 1) {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 120000);
                    } else {
                        monthlySalesModel = new MonthlySalesModel(monthList.get(i), 140000);
                    }

                    model.getMonthlySaleslist().add(monthlySalesModel);
                }

            }
            personCounter++;

            modelList.add(model);
        }


    }

    @AfterPermissionGranted(RC_EXPORT)
    private void excelTask() {
        if (EasyPermissions.hasPermissions(MainActivity.this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
            // Have permission, do the thing!
            createExcelFile(modelList);

//            createCSVFile(modelList);

        } else {
            // Request one permission
            EasyPermissions.requestPermissions(MainActivity.this, getResources().getString(R.string.rationale_storage),
                    RC_EXPORT, Manifest.permission.WRITE_EXTERNAL_STORAGE);
        }
    }


    @Override
    public void onPermissionsGranted(int requestCode, @NonNull List<String> perms) {

    }

    @Override
    public void onPermissionsDenied(int requestCode, @NonNull List<String> perms) {

    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == RC_EXPORT) {
            if (grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                excelTask();
            } else {
                Toast.makeText(this, "Storage Permission Denied", Toast.LENGTH_SHORT).show();
            }

        }
    }

    private void createExcelFile(ArrayList<SalesPersonModel> modelList) {


        if (modelList != null && modelList.size() > 0) {

            ArrayList<Integer> totalList = new ArrayList<>();


            for (String month : monthList) {
                int total = 0;
                for (int i = 0; i < modelList.size(); i++) {

                    for (MonthlySalesModel m : modelList.get(i).getMonthlySaleslist()) {
                        if (m.getMonthName().equals(month)) {
                            total += m.getSales();
                        }

                    }
                }
                totalList.add(total);
            }

            String fileName = "Monthly Sales Report.xls";


            File directory = new File(Environment.getExternalStorageDirectory(), "Export Reports");

            //create directory if not exist
            if (!directory.exists()) {
                directory.mkdirs();
            }

            try {

                //file path
                File file = new File(directory, fileName);
                WorkbookSettings wbSettings = new WorkbookSettings();
                wbSettings.setLocale(new Locale("en", "EN"));
                WritableWorkbook workbook;
                workbook = Workbook.createWorkbook(file, wbSettings);
                //Excel sheet name. 0 represents first sheet
                WritableSheet sheet = workbook.createSheet("Monthly Sales Report", 0);

//                sheet.addCell(new Label(0, 0, " Sample Excel Sheet")); // column and row
                createLabel();
                //Merge col[0-3] and row[1]
                sheet.addCell(new Label(0, 0, "SR. No.", timesHeader));
                sheet.mergeCells(0, 0, 0, 1);
                sheet.addCell(new Label(1, 0, "Sales Person Name", timesHeader));
                sheet.mergeCells(1, 0, 1, 1);
                sheet.addCell(new Label(2, 0, "Monthly Sales", timesHeader));
                sheet.mergeCells(2, 0, 4, 0);

                int counter = 2;

                for (String month : monthList) {

                    sheet.addCell(new Label(counter, 1, month, timesHeader));
                    counter++;
                }


                int rowCounter = 2;
                int serialNoCounter = 1;


                for (int j = 0; j < modelList.size(); j++) {


                    int columnCounter = 0;
                    Number serialnum = new Number(columnCounter, rowCounter, serialNoCounter, timesCenterRegular);
                    sheet.addCell(serialnum);
                    columnCounter++;
                    sheet.addCell(new Label(columnCounter, rowCounter, modelList.get(j).getName(), timesRegular));
                    columnCounter++;

                    for (String month : monthList) {
                        for (MonthlySalesModel m : modelList.get(j).getMonthlySaleslist()) {

                            if (m.getMonthName().equals(month)) {
                                Number onsite = new Number(columnCounter, rowCounter, m.getSales(), timesCenterRegular);
                                sheet.addCell(onsite);
                                columnCounter++;

                            }
                        }
                    }
                    rowCounter++;
                    serialNoCounter++;
                }

                sheet.addCell(new Label(0, rowCounter, "Total", timesHeader));
                sheet.mergeCells(0, rowCounter, 1, rowCounter);

                Number totalJan = new Number(2, rowCounter, totalList.get(0), timesHeader);
                sheet.addCell(totalJan);
                Number totalFeb = new Number(3, rowCounter, totalList.get(1), timesHeader);
                sheet.addCell(totalFeb);
                Number totalMarch = new Number(4, rowCounter, totalList.get(2), timesHeader);
                sheet.addCell(totalMarch);


                workbook.write();
                workbook.close();


                if (file.getParent() != null) {
                    String filepath = file.getParent().concat("/").concat(fileName);
                    Log.d("FILE PATH", "createExcelFile: =>" + filepath);
                    Toast.makeText(this, "Report Exported at =>" + filepath, Toast.LENGTH_SHORT).show();

                }


            } catch (Exception e) {
                e.printStackTrace();
            }

        }
    }


    // Define the cell format
    private void createLabel()
            throws WriteException {
        // Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);

        timesRegular = new WritableCellFormat(times10pt);

        // Lets automatically wrap the cells
        timesRegular.setWrap(true);
        timesRegular.setVerticalAlignment(VerticalAlignment.CENTRE);
        timesRegular.setBorder(Border.ALL, BorderLineStyle.THIN);


        timesCenterRegular = new WritableCellFormat(times10pt);
        timesCenterRegular.setWrap(true);
        timesCenterRegular.setAlignment(Alignment.CENTRE);
        timesCenterRegular.setVerticalAlignment(VerticalAlignment.CENTRE);
        timesCenterRegular.setBorder(Border.ALL, BorderLineStyle.THIN);


        // create create a bold font
        WritableFont times10ptBoldUnderline = new WritableFont(
                WritableFont.TIMES, 10, WritableFont.BOLD, false);
        timesHeader = new WritableCellFormat(times10ptBoldUnderline);
        timesHeader.setBackground(Colour.GRAY_25);
        timesHeader.setBorder(Border.ALL, BorderLineStyle.THIN);
        // Lets automatically wrap the cells
        timesHeader.setWrap(true);


        timesHeader.setAlignment(Alignment.CENTRE);
        timesHeader.setVerticalAlignment(VerticalAlignment.CENTRE);


        CellView cv = new CellView();
        cv.setFormat(timesRegular);
        cv.setFormat(timesHeader);
        cv.setFormat(timesCenterRegular);

        cv.setAutosize(true);

    }
}
