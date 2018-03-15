# ExportExcel
Example have export data in excel file like below image.


  <img src="https://github.com/vishalhalani/ExportExcel/blob/master/sampleExcel.png"  width=400 height=100/>
  
  # Download Library
  Download library jar file from  [here](https://github.com/vishalhalani/ExportExcel/blob/master/app/libs/jxl.jar) and put inside libs folder of application
  
 # Create workbook
 ```
 //Create file
                File file = new File(directory, fileName);
                // get workbook setting
                WorkbookSettings wbSettings = new WorkbookSettings();
                //Set Workbook settings
                wbSettings.setLocale(new Locale("en", "EN"));
                WritableWorkbook workbook;
                // create workbook
                workbook = Workbook.createWorkbook(file, wbSettings);
 ```
 # Create Worksheet
 ```
   WritableSheet sheet = workbook.createSheet("Monthly Sales Report", 0);
 ```
 # Format Cell
 ```
        WritableCellFormat timesRegular;
        WritableCellFormat timesCenterRegular;
        WritableCellFormat timesHeader;
        
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);

        timesRegular = new WritableCellFormat(times10pt);

        // Lets automatically wrap the cells
        timesRegular.setWrap(true);
        timesRegular.setVerticalAlignment(VerticalAlignment.CENTRE);
        timesRegular.setBorder(Border.ALL, BorderLineStyle.THIN);
        
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

        // set format to cell
        CellView cv = new CellView();
        cv.setFormat(timesRegular);
        cv.setFormat(timesHeader);
        cv.setFormat(timesCenterRegular);

        cv.setAutosize(true);

 ```
 # Create Cell
  Worksheet have method of add cell to add cell in worksheet<br>
   <b>Label</b> use for add string value in cell<br>
   <b>Number</b> user for add Number value in cell<br>
  ```
  // add header String
  sheet.addCell(new Label(0, 0, "SR. No.", timesHeader));
                sheet.mergeCells(0, 0, 0, 1);
  // add value in cell
  
  Number serialnum = new Number(columnnumber, rownumber, value, cellformat);
                    sheet.addCell(serialnum);
 ```
 # Merge Cell
 you can merge cell using <b>mergeCells</b> method it contains 4 parameter startColumnPosition, startRowPosition, endColumnPosition, endRowPosition
 
 ```
  sheet.addCell(new Label(2, 0, "Monthly Sales", timesHeader));
  sheet.mergeCells(2, 0, 4, 0);
 ```

                
