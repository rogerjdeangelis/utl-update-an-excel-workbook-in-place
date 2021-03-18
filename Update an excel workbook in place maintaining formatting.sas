Update an excel workbook in place maintaining formatting

GitHub
https://tinyurl.com/btr4ws3c
https://github.com/rogerjdeangelis/utl-update-an-excel-workbook-in-place

*You do not need to do this for this solution but may need for other solution or different excel defaults
Open the workbook turn on sharing and save the workbook
Also in some ca es you need to force a recalculation by typing 'cntl+alt+shift+f9'

  Problem
     Update age in an existing sheet and recalculate the the sum of ages
     while maintaining the existing formulas and the coloring


   Two Solutions

       a. SAS
       b. R (also


Related Repos (there are other ways)

https://tinyurl.com/sg5ohbp
https://github.com/rogerjdeangelis/utl-update-an-existing-excel-named-range-R-python-sas

https://tinyurl.com/3kvyfetm
https://github.com/rogerjdeangelis?tab=repositories&q=update+excel+in+place&type=&language=

https://github.com/rogerjdeangelis?tab=repositories&q=update+excel&type=&language=
https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python


*_                   _
(_)_ __  _ __  _   _| |_
| | '_ \| '_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
;


* create a workbook with cell column colored and with a formula in column B;

* incase you rerun;
%utlfkil(d:\xls\formulas.xlsx);

ods excel file='d:\xls\formulas.xlsx' options(sheet_name="inplace");

proc report data=have nowd missing;

column age agesqr;

define age    / "Age" display;
define agesqr / computed "Age_Squared" format=3. ;

compute agesqr;
   call define(_col_, "Style", "Style = [background = yellow tagattr='formula:sum(A2:A3)']");
endcomp;
run;quit;

ods excel close;


 SHEET INPLACE IN WORKBOOK IN D:\XLS\FORMULAS.XLSX

                ____________
               | Column B   |
               | is Yellow  |
  +-------------------------+
  |     A      |    B       |
  +-------------------------+
1 |    AGE     | SUM_AGE    |
  +------------+------------+
2 |    11      |    23      | =SUM(A2:A3)
  +------------+------------+
3 |    12      |    23      | =SUM(A2:A3)
  +------------+------------+

[INPLACE]


*
 _ __  _ __ ___   ___ ___  ___ ___
| '_ \| '__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
  __ _     ___  __ _ ___
 / _` |   / __|/ _` / __|
| (_| |_  \__ \ (_| \__ \
 \__,_(_) |___/\__,_|___/

;

proc sql dquote=ansi;
    connect to excel as excel(Path="d:\xls\formulas.xlsx");
    execute(
      update [inplace$]
      set age=88
      where AGE=12
    ) by excel;
    disconnect from excel;
Quit;

*_        ____
| |__    |  _ \
| '_ \   | |_) |
| |_) |  |  _ <
|_.__(_) |_| \_\

;

* Note setStyleAction(wb,XLC$"STYLE_ACTION.NONE");

%utl_submit_r64('
library(XLConnect);
wb <- loadWorkbook("d:/xls/formulas.xlsx", create=TRUE);
setStyleAction(wb,XLC$"STYLE_ACTION.NONE");
Data <- data.frame(AGE=88);
writeWorksheet(wb,Data,"inplace",startRow=2,startCol=1,header=FALSE);
saveWorkbook(wb);
');

*            _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| '_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
;

 SHEET INPLACE IN WORKBOOK IN D:\XLS\FORMULAS.XLSX

Note the sum of ages has changed and column is still yellow

                ____________
               | Column B   |
               | is Yellow  |
  +-------------------------+
  |     A      |    B       |
  +-------------------------+
1 |    AGE     | SUM_AGE    |
  +------------+------------+
2 |    11      |   100      | =SUM(A2:A3)
  +------------+------------+
3 |    88      |   100      | =SUM(A2:A3)
  +------------+------------+

[INPLACE]
