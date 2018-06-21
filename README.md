# ExcelUtil
Excel Utility in C#

   Excel Utility (C) 2017 NNNIC 

   File : ExcelUtil.cs
   
   The vs project is for test. 


    Futures

    #1. Use Microsoft soft original dll only
    #2. Quick Access
    #3. Capable to attach to the excel window
    #4. Simple Interface

    <<Sammary>>

    BookCtr   book = ExcelUtil.OpenBook(path);
    SheetCtr  sheet = app.GetSheet(sheetname);

    object[,] values =  sheet.GetValues(numofrow=null,numofcol=null); //default : Used Range
    sheet.SetValues(values); 

    book.Close(); 
    book.Write();
    book.WriteAs(path)

    BookCtr book = book.OpenBookInApp(path);  //Open a book ont the excel window

    BookCtr book = ExcelUtil.AttachBook(path); //Attach to a existing excelwindow

If you want to use raw apis only, see https://gist.github.com/NNNIC/6df2c06b2f8985334e9b9a267300e441 
