# jnr-jenner-analytics-part-VIII-excel-update-sheet-in-place-using-sql-keeping-all-formatting
Jenner Analytica part VIII excel update sheet in place using sql keeping all formatting
    /*--
    c:/utl/jnr-jenner-analytics-part-VIII-excel-update-sheet-in-place-using-sql-keeping-all-formatting.sas
    ---*/

    Jenner Analytica part VIII excel update sheet in place using sql keeping all formatting

    Too log to post, see
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-VIII-excel-update-sheet-in-place-using-sql-keeping-all-formatting

    One of the challenges with excel is updating a worksheets in place
    without losing any formatting.

    The followingpost converts column weight from pounds to kilograms
    while keeping the bold yellow formatting.

    d:/xls/class.xlsx (weight updated to kilograms in place)

                BOLD                   BOLD
                YELLOW                 YELLOW
                POUNDS                 KILOGRAMS
    name        weight     name        weight
    Alfred       112.5     Alfred      51.029
    Alice           84     Alice       38.102
    Barbara         98     Barbara     44.462

    CONTENTS

       1 Crreate a substanially formated excel worksheet
       2 Update in place keeping all worksheet fromatting.
       3 Related excel repositories

    SOAPBOX ON
     Both Jenner and the SLC seem to honor more formatting with ODS HTML than
     with ODS EXCEL
     Python xlwings seems to do a remarkable job converting html to excel.
     xlwings is the best option for inplace sql processing of excel workbooks?
    SOAPBOX OFF

    Other inline languages
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-I-inline-drop-down-to-duckdb-sql
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-II-drop-down-to-open-source-spss-run-regression
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-III-inline-drop-down-to-open-source-matlab-summarize-data
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-IV-inline-drop-down-to-poweshell-add-column-to-excell-sheet
    https://github.com/rogerjdeangelis/nr-jenner-analytics-part-V-proc-R-macro-wrapper-and-resolving-macro-triggers-within-R
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-VI-proc-R-macro-wrapper-and-resolving-macro-triggers-within-R
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-VII-inline-drop-down-to-perl
    https://github.com/rogerjdeangelis/jnr-jenner-analytics-part-VIII-excel-update-sheet-in-place-using-sql-keeping-all-formatting

    /******************************************************************************************************************/
    /* 1  Create a substanially formated excel worksheet                                                              */
    /******************************************************************************************************************/

    /*--- CREATE BOLD YELLOW COLUMN ---*/
    %utlfkil(d:/htm/class.html);
    %utlfkil(d:/xls/class.xlsx);
    options nolabel;
    ods html file="d:/htm/class.html";
    proc report data=workx.class(Obs=3);
    define weight / display
       style(column)=[
          font_weight=bold
          background=yellow
       ];
    run;
    ods html close;

    /*--- CONVERTING HTML TO EXCEL WITH FORMATTING ---*/
    options set=PROC_PYPATH "D:\py314\python.exe";
    proc python;
    submit;
    import xlwings as xw
    with xw.App(visible=False, add_book=False) as app:
        book = app.books.open(r"d:/htm/class.html")
        book.save(r"d:/xls/class.xlsx")
        book.close()
    endsubmit;
    run;

    d:/xls/class.xlsx
    =================                BOLD
                                     YELLOW
                                     POUNDS
    name      sex       age  height  weight
    Alfred    M          14      69   112.5
    Alice     F          13    56.5      84
    Barbara   F          13    65.3      98


    LOG:  13:24:07
    ==============

    NOTE: Copyright (c) 2026 Jenner Analytics Ltd., London, England.
    NOTE: Jenner v1.5.102 (build v1.5.102+67ce60c837.20261020T035203Z)
          Licensed to Roger DeAngelis, Serial 3BF7B3E6.
    NOTE: DATA _null_

    autexec started.

    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library WORKX assigned path=d:\wpswrkx.
    NOTE: DATA workx.class

    NOTE: Processing inline DATALINES (19 lines)

    NOTE: Read 19 rows from DATALINES.
    NOTE: The data set WORKX.CLASS has 19 observations and 5 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_

    LOG:  13:24:07
    NOTE: DATA _null_ completed. Output written to FILE PRINT
    NOTE: Option SASAUTOS changed to c:/otojnr.
    NOTE: DATA _null_

    NOTE: Reading from fileref c:/jnr/runsas_selection.sas (c:/jnr/runsas_selection.sas)
        1 %utlfkil(d:/htm/class.html);
        2 %utlfkil(d:/xls/class.xlsx);
        3 options nolabel;
        4 ods html file="d:/htm/class.html";
        5 proc report data=workx.class(Obs=3);
        6 define weight / display
        7    style(column)=[
        8       font_weight=bold
        9       background=yellow
       10    ];
       11 run;
       12 ods html close;
       13
       14 /*--- CONVERTING HTML TO EXCEL WITH FORMATTING ---*/
       15 options set=PROC_PYPATH "D:\py314\python.exe";
       16 proc python;
       17 submit;
       18 import xlwings as xw
       19 with xw.App(visible=False, add_book=False) as app:
       20     book = app.books.open(r"d:/htm/class.html")
       21     book.save(r"d:/xls/class.xlsx")
       22     book.close()
       23 endsubmit;
       24 run;

    NOTE: Read 24 rows from c:/jnr/runsas_selection.sas.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_

    autoexec completed  13:24:07

    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: ODS HTML: Destination opened (file=d:/htm/class.html).
    NOTE: PROC REPORT data=workx.class

    NOTE: PROC REPORT output 3 observations with 5 variables.

    NOTE: Read 3 rows from workx.class.
    NOTE: The data set (console) has 3 observations and 5 variables.
    NOTE: PROC REPORT elapsed:
      wall  0.00 seconds
      cpu   0.02 seconds
    NOTE: ODS HTML: Output written to d:/htm/class.html.
    NOTE: Option SET changed to PROC_PYPATH D:\py314\python.exe.
    NOTE: PROC PYTHON step executing.
    NOTE: PROC PYTHON step completed.

    /******************************************************************************************************************/
    /* 2  Update in place keeping all worksheet fromatting                                                            */
    /******************************************************************************************************************/

    /*--- UPDATE THE BOLD YELLOW COLUMN WITH KILOGRAMS---*/
    options set=PROC_PYPATH "D:\py314\python.exe";
    proc python;
    submit;
    import sqlite3
    import xlwings as xw

    with xw.App(visible=False, add_book=False) as app:
        book = app.books.open(r"d:/xls/class.xlsx")
        sheet = book.sheets["class"]

        data = sheet.range("A1").expand("table").options(ndim=2).value
        col = [str(h).lower() for h in data[0]].index("weight") + 1

        db = sqlite3.connect(":memory:")
        db.execute("CREATE TABLE class (rownum INTEGER, Weight REAL)")
        db.executemany(
            "INSERT INTO class VALUES (?, ?)",
            [(rownum, row[col - 1])
             for rownum, row in enumerate(data[1:], start=2)]
        )

        for rownum, kilograms in db.execute(
            "SELECT rownum, Weight * 0.45359237 FROM class"
        ):
            sheet.cells(rownum, col).value = kilograms

        db.close()
        book.save()
        book.close()

    endsubmit;
    run;

    d:/xls/class.xlsx           BOLD
                                YELLOW
                                KILOGRAMS
    name      sex  age  height  weight
    Alfred    M     14      69  51.029
    Alice     F     13    56.5  38.102
    Barbara   F     13    65.3  44.462

    LOG:  13:01:50
    ==============

    NOTE: Copyright (c) 2026 Jenner Analytics Ltd., London, England.
    NOTE: Jenner v1.5.102 (build v1.5.102+67ce60c837.20261020T035203Z)
          Licensed to Roger DeAngelis, Serial 3BF7B3E6.
    NOTE: DATA _null_

    autexec started.

    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Library WORKX assigned path=d:\wpswrkx.
    NOTE: DATA workx.class

    NOTE: Processing inline DATALINES (19 lines)

    NOTE: Read 19 rows from DATALINES.
    NOTE: The data set WORKX.CLASS has 19 observations and 5 variables.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_

    LOG:  13:01:50
    NOTE: DATA _null_ completed. Output written to FILE PRINT
    NOTE: Option SASAUTOS changed to c:/otojnr.
    NOTE: DATA _null_

    NOTE: Reading from fileref c:/jnr/runsas_selection.sas (c:/jnr/runsas_selection.sas)
        1 options set=PROC_PYPATH "D:\py314\python.exe";
        2 proc python;
        3 submit;
        4 import sqlite3
        5 import xlwings as xw
        6
        7 with xw.App(visible=False, add_book=False) as app:
        8     book = app.books.open(r"d:/xls/class.xlsx")
        9     sheet = book.sheets["class"]
       10
       11     data = sheet.range("A1").expand("table").options(ndim=2).value
       12     col = [str(h).lower() for h in data[0]].index("weight") + 1
       13
       14     db = sqlite3.connect(":memory:")
       15     db.execute("CREATE TABLE class (rownum INTEGER, Weight REAL)")
       16     db.executemany(
       17         "INSERT INTO class VALUES (?, ?)",
       18         [(rownum, row[col - 1])
       19          for rownum, row in enumerate(data[1:], start=2)]
       20     )
       21
       22     for rownum, kilograms in db.execute(
       23         "SELECT rownum, Weight * 0.45359237 FROM class"
       24     ):
       25         sheet.cells(rownum, col).value = kilograms
       26
       27     db.close()
       28     book.save()
       29     book.close()
       30
       31 endsubmit;
       32 run;

    NOTE: Read 32 rows from c:/jnr/runsas_selection.sas.
    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: DATA _null_

    autoexec completed  13:01:50

    NOTE: DATA elapsed:
      wall  0.00 seconds
      cpu   0.00 seconds
    NOTE: Option SET changed to PROC_PYPATH D:\py314\python.exe.
    NOTE: PROC PYTHON step executing.
    NOTE: PROC PYTHON step completed.

    /******************************************************************************************************************/
    /* 3 Related repositories                                                                                         */
    /******************************************************************************************************************/

    https://github.com/rogerjdeangelis/utl-altair-slc-drop-down-to-powershell-and-programatically-add-and-execute-an-excel-macro
    https://github.com/rogerjdeangelis/utl-highlite-sas-dataset-and-view-the-table-in-excel-without-sas-access
    https://github.com/rogerjdeangelis/utl-altair-slc-get-the-color-of-a-cell-in-excel-xlsx
    https://github.com/rogerjdeangelis/utl-convert-excel-to-csv-by-dropping-down-to-r-or-python
    https://github.com/rogerjdeangelis/utl-select-type-and-length-using-odbc-excel-passthru-query
    https://github.com/rogerjdeangelis/altair-slc-create-excel-gpa-sheet-as-the-avg-of-math-and-music-grade-sheets
    https://github.com/rogerjdeangelis/utl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates
    https://github.com/rogerjdeangelis/utl-formatting-ai-seacrh-output-in-pdf-rtf-and-excel-format-perplexity-chatGPT-results
    https://github.com/rogerjdeangelis/utl-importing-multiple-excel-files-which-names-are-defined-by-state
    https://github.com/rogerjdeangelis/utl-roll-up-adverse-events-by-patient-and-date-using-sqlite-groupcat-slc-r-python-and-excel
    https://github.com/rogerjdeangelis/utl-select-fruits-purchased-and-not-purchased--from-old-macdonalds-farm-excel-r-sql
    https://github.com/rogerjdeangelis/utl-substituting-name-and-label-to-column-headings-in-excel
    https://github.com/rogerjdeangelis/utl-total-weekly-cost-for-breakfast-lunch-and-dinner-for-mary-jane-and-mike-sas-r-sql-python-excel
    https://github.com/rogerjdeangelis/utl_excel_combining_sheets_without_common_names_types_lengths
    https://github.com/rogerjdeangelis/utl_excel_import_entire_directory
    https://github.com/rogerjdeangelis/utl_ods_excel_font_size_and_justification_proc_report_titles_formatting
    https://github.com/rogerjdeangelis/utl-add-sequence-numbers-to-each-group-using-sas-and-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-adding-a-sheet-to-an-existing-open-and-on-screen-or-saved-excel-workbook
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-creating-and-analyzing-tables-created-from-multiple-pdf-and-excel-files
    https://github.com/rogerjdeangelis/utl-altair-slc-determining-type-and-length-of-excel-columns-before-importing
    https://github.com/rogerjdeangelis/utl-altair-slc-excel-using-undocumented-excel-vba-and-sql-functions-to-prepare-a-sheet-for-import
    https://github.com/rogerjdeangelis/utl-for-each-row-return-the-column-name-of-the-largest-value-base-and-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-force-excel-to-read-all-the-columns-as-numeric-or-character
    https://github.com/rogerjdeangelis/utl-import-all-excel-workskkets-and-named-ranges--in-all-workbooks-in-a-directory
    https://github.com/rogerjdeangelis/utl-matching-tables-based-on-a-hierachy-of-rules-sas-hash-sql-r-pytho-excel
    https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-convert-many-pdf-files-with-tables-to-excel-workbooks-by-date
    https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-create-excel-report-with-subtotals-and-grand-total
    https://github.com/rogerjdeangelis/utl-move-even-numbered-columns-to-the-end-of-the-sas-pdv-datastep-and-sql-r-excel
    https://github.com/rogerjdeangelis/utl-place-animal-behavior-into-four-categories-using-sql-sas-r-python-excel-mapping-lookups
    https://github.com/rogerjdeangelis/utl-splitting-rows-based-on-variable-suffixes-with-and-without-sql-arrays-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-subset-a-database-table-based-on-a-list-of-names-in-excel
    https://github.com/rogerjdeangelis/utl_excel_Import_and_transpose_range_A9-Y97_using_only_one_procedure
    https://github.com/rogerjdeangelis/utl_joining_and_updating_excel_sheets_without_importing_data
    https://github.com/rogerjdeangelis/utl-add-sequence-number-by-group-sas-sql-with-without-monotonic-and-in-sql-r-python-excel
    https://github.com/rogerjdeangelis/utl-apply-transactions-to-master-to-flag-male-and-non-male-patients-sas-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-automating-excel-data-preperation-monarch-and-altair-percsonal-slc
    https://github.com/rogerjdeangelis/utl-calculate-percentage-by-group-in-wps-r-python-excel-sql-no-sql
    https://github.com/rogerjdeangelis/utl-create-new-variable-by-multiplying-corresponding-variables-and-summing-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc
    https://github.com/rogerjdeangelis/utl-export-excel-sheet-to-sas-table-using-one-line-macro-invocation-without-sas-access-powershell
    https://github.com/rogerjdeangelis/utl-find-airlines-that-use-mutiple-vendors-for-in-flight-services-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-If-over-sixty-days-between-rx-then-new-rx-else-refill-rx-sas-and-sql-r-sas-python-excel
    https://github.com/rogerjdeangelis/utl-importing-excel-datetime-values-in-xlsx-and-xlsx-workbooks
    https://github.com/rogerjdeangelis/utl-programatically-search-all-cells-in-an-excel-sheet-for-an-arbitrary-string-python-openxl
    https://github.com/rogerjdeangelis/utl-r-open-closed-excel-workbook-an-update-master-sheet-using-transaction-sheet
    https://github.com/rogerjdeangelis/utl-roll-up-adverse-events-by-patient-and-date-using-sql-groupcat-r-python-and-excel
    https://github.com/rogerjdeangelis/utl-round-trip-sas-time-values-to-excel-and-back-sas-and-r-openxlsx
    https://github.com/rogerjdeangelis/utl-sas-ods-excel-to-create-excel-report-and-separate-png-graph-finally-r-for-layout-in-excel
    https://github.com/rogerjdeangelis/utl-sas-to-and-from-sqllite-excel-ms-access-spss-stata-using-r-packages-without-sas
    https://github.com/rogerjdeangelis/utl-sum-the-three-largest-yoy-growth-by-company-sql-partitioning-and-rank-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-temporarily-change-the-excel-base-font-to-a-blue-consolas-monospace-font-r-openxlsx
    https://github.com/rogerjdeangelis/utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart
    https://github.com/rogerjdeangelis/utl_1130pm_batch_SAS_job_that_imports_an_excel_if_modified_that_day
    https://github.com/rogerjdeangelis/utl_create_many_excel_workbooks_for_selected_cities_in_sashelp_zipcode_with_logging
    https://github.com/rogerjdeangelis/utl_excel_experimenting-with-the-new-ods-excel-destination
    https://github.com/rogerjdeangelis/utl_import_data_from_excel_sheet_with_headers_and_footers_without_specifying_range_option
    https://github.com/rogerjdeangelis/utl_maintaining_all_significant_digits_when_importing_excel_sheet
    https://github.com/rogerjdeangelis/utl-add-a-tab-to-excel-that-autmatically-impot-a-sas-datasets-www.colectica-com
    https://github.com/rogerjdeangelis/utl-apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r
    https://github.com/rogerjdeangelis/utl-avoid-storing-numbers-as-text-when-exporting-mixed-type-to-excel
    https://github.com/rogerjdeangelis/utl-calculating-three-day-moving-sum-by-city-using-r-python-excel-sql
    https://github.com/rogerjdeangelis/utl-complete-years-using-overall-max-and-min-applied-to-each-group-r-sas-and-sql-r-python-excel
    https://github.com/rogerjdeangelis/utl-concatenate-contents-from-the-same-column-base-sas-r-python-excel-sql
    https://github.com/rogerjdeangelis/utl-create-graphs-in-excel-using-excel-chart-templates
    https://github.com/rogerjdeangelis/utl-creating-variables-that-flag-changes-over-time-for-a-medication-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-get-the-color-of-a-cell-in-excel-xlsx
    https://github.com/rogerjdeangelis/utl-how-to-check-whether-a-student-is-in-the-Excel-sheet-class
    https://github.com/rogerjdeangelis/utl-how-to-import-the-ms-excel-sheet-names-in-actual-order
    https://github.com/rogerjdeangelis/utl-merge-cells-in-excel-and-rtf-output-using-proc-report-and-r-openxlsx-package
    https://github.com/rogerjdeangelis/utl-monarch-and-personal-altair-slc-odbc-appending-rows-to-an-existing-excel-sheet
    https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python
    https://github.com/rogerjdeangelis/utl-pivot-transpose-an-excel-sheet-with-columns-that-are-excel-dates
    https://github.com/rogerjdeangelis/utl-position-a-sas-table-arbitrarily-in-excel-without-column-names-using-r-and-python
    https://github.com/rogerjdeangelis/utl-select-the-top-n-and-bottom-n-by-group-sql-r-python-excel
    https://github.com/rogerjdeangelis/utl-side-by-side-sas-tables-in-one-excel-sheet
    https://github.com/rogerjdeangelis/utl-sql-arrays-to-simplfy-summarization-of-many-variables-in-sas-r-python-excel-sql-arrays
    https://github.com/rogerjdeangelis/utl-subset-sas-dataset-based-on-the-value-of-a-macro-variable-contiained-in-an-excel-worksheet
    https://github.com/rogerjdeangelis/utl-tables-to-specific-excel-cells
    https://github.com/rogerjdeangelis/utl-using-column-position-instead-of-excel-column-names-due-to-misspellings-sas-r-python
    https://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets
    https://github.com/rogerjdeangelis/utl_creating_www_hyperlinks_in_ods_excel
    https://github.com/rogerjdeangelis/utl_excel_create_a_sheet_for_each_table_with_variable_name_position_and_label
    https://github.com/rogerjdeangelis/utl_excel_exporting_data_with_leading_zeros
    https://github.com/rogerjdeangelis/utl_excel_highlight_individual_cells_based_on_indicator_variables
    https://github.com/rogerjdeangelis/utl_excel_importing_unicode_and_other_special_characters_without_changing_sas_encoding
    https://github.com/rogerjdeangelis/utl_excel_import_long_colnames
    https://github.com/rogerjdeangelis/utl_excel_import_two_excel_ranges_within_one_sheet
    https://github.com/rogerjdeangelis/utl_excel_import_xlsm_to_sas_dataset
    https://github.com/rogerjdeangelis/utl_excel_sas_wps_r_import_xlsx_without_sas_access_to_pc_files
    https://github.com/rogerjdeangelis/utl_maintaining_numeric_significance_when_exporting_and_importing_excel_workbooks
    https://github.com/rogerjdeangelis/utl_ods_excel_conditionaly_higlight_individua1_cells
    https://github.com/rogerjdeangelis/ods_excel_does_not_always_honor_start_at--bug
    https://github.com/rogerjdeangelis/utl-a-cursory-comparison-of-excel-alternatives-using-r-python-and-libre-office-calc
    https://github.com/rogerjdeangelis/utl-add-formula-to-an-existing-excel-worksheet-using-powershell
    https://github.com/rogerjdeangelis/utl-add-one-to-each-new-occurance-of-a-value-in-a-series-sas-hash-and--sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-adding-a-column-to-excelsheet-using-personal-altair-slc
    https://github.com/rogerjdeangelis/utl-adding-a-second-ods-excel-created-sheet-to-a-closed-ods-excel-workbook
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-comparing-two-excel-workbook-sheets
    https://github.com/rogerjdeangelis/utl-altair-slc-add-column-bmi-to-an-existing-excel-worksheet-using-powershell
    https://github.com/rogerjdeangelis/utl-altair-slc-convert-pdf-tables-to-excel-sheets
    https://github.com/rogerjdeangelis/utl-altair-slc-language6-rexcel-read-write-text-files-summarize-and-looping-with-sqlite
    https://github.com/rogerjdeangelis/utl-altair-slc-split-closed-workbook-excel-sheet-into-male-and-female-sheets-libname-excel
    https://github.com/rogerjdeangelis/utl-classic-pivot-wider-transpose-with-output-compound-column-names-using-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-combine-name-value-pairs-with-all-other-name-value-pairs-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-count-employees-that-have-made-at-leat-one-stock-ticker-purchase-r-sql-python-excel
    https://github.com/rogerjdeangelis/utl-create-a-pdf-excel-html-proc-report-with-greek-letters
    https://github.com/rogerjdeangelis/utl-creating-a-two-by-two-grid-of-reports-in-excel
    https://github.com/rogerjdeangelis/utl-excel-fixing-bad-formatting-using-passthru
    https://github.com/rogerjdeangelis/utl-excel-using-proc-report-workarea-columns-to-operate--on-arbitrary-row
    https://github.com/rogerjdeangelis/utl-extract-sheet-names-from-multiple-excel-versions-using-r
    https://github.com/rogerjdeangelis/utl-find-out-which-excel-columns-are-dates-and-assign-date-type
    https://github.com/rogerjdeangelis/utl-how-to-find-common-values-in-four-columns-r-and-sql-r-sas-python-excel
    https://github.com/rogerjdeangelis/utl-how-to-load-excel-sheets-with-sheet-names-with-31-characters
    https://github.com/rogerjdeangelis/utl-identify-missing-drug-presciption-days-by-patient-using-sql-recursion-r-python-and-excel
    https://github.com/rogerjdeangelis/utl-Import-excel-sheet-as-character-fixing-truncation-mixed-type-columns-and-appending-issues
    https://github.com/rogerjdeangelis/utl-importing-excel-string_of_32-thousand-characters-SAS-XLConnect
    https://github.com/rogerjdeangelis/utl-increment-a-counter-each-time-a-missing-value-is-encountered-usin-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-keeping-leading-and-trailing-zeros-in-character-fields-with-ods-excel-output
    https://github.com/rogerjdeangelis/utl-load-and-extract-ms-excel-document-properties-metadata
    https://github.com/rogerjdeangelis/utl-manipulate-excel-directly-using-passthru-microsoft-sql-wps-r-rodbc
    https://github.com/rogerjdeangelis/utl-monarch-altair-personal-slc-combine-two-tables-in-one-excel-sheet
    https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-import-excel-sheet-with-many-tabs
    https://github.com/rogerjdeangelis/utl-monarch-and-personal-altair-slc-reading-excel-pivot-tables
    https://github.com/rogerjdeangelis/utl-monarch-and-personal-slc-appending-all-sheets-to-first-sheet-Excel97-2003-workbooks
    https://github.com/rogerjdeangelis/utl-ninetyfifth-percentiles-of-home-runs-for-lf-and-rf-baseball-players-sas-and-sql-r-excel
    https://github.com/rogerjdeangelis/utl-ods-excel-color-code-every-other-column-in-a-specified-row
    https://github.com/rogerjdeangelis/utl-ods-excel-hilite-diagonal-cells
    https://github.com/rogerjdeangelis/utl-Password-protect-an-EXCEL-file-in-sas-without-X-command
    https://github.com/rogerjdeangelis/utl-pivot-excel-columns-and-output-a-database-table
    https://github.com/rogerjdeangelis/utl-pivot-wide-when-variable-names-contain-values-sql-and-base-r-sas-oython-excel-postgreSQL
    https://github.com/rogerjdeangelis/utl-posting-your-problem-with-an-ascii-image-that-looks-just-like-an-excel-shee
    https://github.com/rogerjdeangelis/utl-rogers-ods-excel-lineprinter-style-formchar-gridlines-sort-of-sharebuffers
    https://github.com/rogerjdeangelis/utl-sas-ods-bidirectional-hyperlinked-table-of-contents-in-ods-pdf-html-and-excel
    https://github.com/rogerjdeangelis/utl-select-phase2-and-phase3-trials-with-both-treatment-and-control-arms-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-select-the-top-five-ages-using-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-select-the-top-two-largest-values-by-group-using-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-send-all-tables-in-a-sas-library-to-excel
    https://github.com/rogerjdeangelis/utl-seven-algorithms-to-convert-a-sas-dataset-to-an-excel-workbook
    https://github.com/rogerjdeangelis/utl-skilled-nursing-cost-reports-2011-2019-in-excel
    https://github.com/rogerjdeangelis/utl-sql-workaround-fix-for-excel-bug-when-counting-distinct-values
    https://github.com/rogerjdeangelis/utl-stack-tables-when-tables-are-not-conguent-base-r-and-sql-r-python-excel
    https://github.com/rogerjdeangelis/utl-string-functions-to-parse-messy-four-word--string-into-name-grade-and-weight-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-sum-all-pairs-of-columns-using-by-row-base-r-sql-sas-r-excel
    https://github.com/rogerjdeangelis/utl-sum-purchases-before-and-after-sale-dates-sql-sas-r-python-and-excel
    https://github.com/rogerjdeangelis/utl-update-in-place-sheet2-by-adding-dinner-costs-from-sheet1-preserving-excel-formatting-r
    https://github.com/rogerjdeangelis/utl_excel_add_formulas
    https://github.com/rogerjdeangelis/utl_excel_create_sql_insert_and_value_statements_to_update_databases
    https://github.com/rogerjdeangelis/utl_excel_fix_type_length_on_import
    https://github.com/rogerjdeangelis/utl_excel_import_sas_functions_fail_on_cells_with_mutiple_line_breaks
    https://github.com/rogerjdeangelis/utl_excel_import_sub_rectangle
    https://github.com/rogerjdeangelis/utl_excel_reading_a_single_cell
    https://github.com/rogerjdeangelis/utl_excel_update_inplace
    https://github.com/rogerjdeangelis/utl_import_excel_column_names_that_contain_a_dollar_sign_and_rename_without
    https://github.com/rogerjdeangelis/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet
    https://github.com/rogerjdeangelis/utl_programatically_execute_excel_macro_using_wps_proc_python
    https://github.com/rogerjdeangelis/utl_sas_v5_transport_file_to_excel
    https://github.com/rogerjdeangelis/utl_stacking-strings-in-one-excel-cell-using-ods-excel-newline-carriage-return-line-feed-tags
    https://github.com/rogerjdeangelis/utl_table_of_contents_with_excel_links_to_sheets
    https://github.com/rogerjdeangelis/excel-how-do-I-remove-troublesome-characters-before-importing
    https://github.com/rogerjdeangelis/utl-adding-a-password-to-an-existing-excel-workbook
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-issue-importing-excel-sheet-where-column-names-are-in-the-third-row
    https://github.com/rogerjdeangelis/utl-altair-slc-appending-mutiple-excel-ranges-within-one-sheet
    https://github.com/rogerjdeangelis/utl-altair-slc-summarizing-excel-bank-statements
    https://github.com/rogerjdeangelis/utl-altair-slc-three-hour-moving-sum-of-pings-by-http-date-status-using-slc-r-python-excel
    https://github.com/rogerjdeangelis/utl-applying-meta-data-and-importing-data-from-an-excel-named-range
    https://github.com/rogerjdeangelis/utl-combine-text-in-an-excel-column-down-multiple-rows-by-group
    https://github.com/rogerjdeangelis/utl-convert-excel-sheet-to-sas-dataset-without-sas-access
    https://github.com/rogerjdeangelis/utl-converting-r-sas-and-excel-serial-dates-to-sqllite-serial-dates-and-posix-iso8601-clinical
    https://github.com/rogerjdeangelis/utl-copy-all-sas-datasets-in-work-library-to-tabs-in-one-excel-workbook
    https://github.com/rogerjdeangelis/utl-create-r-dataframe-and-sas-dataset-from-second-sheet-in-multisheet-excel-workbook-r-openxlsx
    https://github.com/rogerjdeangelis/utl-creating-multiple-odbc-tables-in-a-one-excel-sheet
    https://github.com/rogerjdeangelis/utl-cumulative-sum-by-group-in-the-order-of-rank-variable-sas-and-sql-r-python-octave-excel
    https://github.com/rogerjdeangelis/utl-cumulative-sums-by-group-using-base-sas-and-sql-sas-r-python-excel
    https://github.com/rogerjdeangelis/utl-do-not-add-data-transformations-to-create-csv-files-from-excel-or-any-other-data-structure
    https://github.com/rogerjdeangelis/utl-does-the-excel-sheet-exist
    https://github.com/rogerjdeangelis/utl-example-of-perl-regex-in-sas-r-python-and-excel-sql
    https://github.com/rogerjdeangelis/utl-excel-hiperlinks-click-on-your-favorite-baseball-player-and-a-google-search-will-pop-up
    https://github.com/rogerjdeangelis/utl-excel-import-number-strings-with-spaces
    https://github.com/rogerjdeangelis/utl-fix-excel-columns-with-mutiple-datatypes-on-the-excel-side-using-ms-sql-and-passthru
    https://github.com/rogerjdeangelis/utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value
    https://github.com/rogerjdeangelis/utl-identify-sequential-maximums-in-a-series-of-numbers-using-base-sas-and-sql-r-python-excel
    https://github.com/rogerjdeangelis/utl-import-a-messy-excel-file
    https://github.com/rogerjdeangelis/utl-import-excel-when-column-names-are-excel-dates
    https://github.com/rogerjdeangelis/utl-import-excel-workbooks-in-all-folders-and-subfolders
    https://github.com/rogerjdeangelis/utl-importing-excel-when-sheetname-has-spaces
    https://github.com/rogerjdeangelis/utl-importing-multiple-excel-worksheets-without-access-to-pc-files
    https://github.com/rogerjdeangelis/utl-join-a-sas-table-with-an-excel-table-when-column-names-that-are-dates
    https://github.com/rogerjdeangelis/utl-keep-orginal-SAS-table-but-mask-excel-output
    https://github.com/rogerjdeangelis/utl-layout-ods-excel-reports-in-a-grid
    https://github.com/rogerjdeangelis/utl-no-need-for-sql-or-sort-merge-use-a-elegant-hash-excel-vlookup
    https://github.com/rogerjdeangelis/utl-ods-export-sas-table-to-excel-with-rotated-column-headers
    https://github.com/rogerjdeangelis/utl-pivot-long--excel-sheet-and-run-a-regression-in-r-and-python
    https://github.com/rogerjdeangelis/utl-safe-way-import-excel-time-value
    https://github.com/rogerjdeangelis/utl-select-the-top-ten-rows-from-excel-table-without-importing-to-sas
    https://github.com/rogerjdeangelis/utl-side-by-side-reports-within-arbitrary-positions-in-one-excel-sheet-wps-r
    https://github.com/rogerjdeangelis/utl-simple-example-of-a-three-table-full-outer-join-by-ids-sas-r-python-excel-sql
    https://github.com/rogerjdeangelis/utl-simple-r-code-to-covert-excel-to-sas-and-sas-to-excel
    https://github.com/rogerjdeangelis/utl-simple-three-letter-commands-to-format-perplexity-AI-results-for-word-pdf-text-and-excel
    https://github.com/rogerjdeangelis/utl-single-click-and-eight-excel-tabs-are-converted-to-csv-files
    https://github.com/rogerjdeangelis/utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf
    https://github.com/rogerjdeangelis/utl-update-an-excel-workbook-in-place
    https://github.com/rogerjdeangelis/utl-update-an-existing-excel-named-range-R-python-sas
    https://github.com/rogerjdeangelis/utl-using-excel-to-get-a-usefull-proc-tabulate-output-table
    https://github.com/rogerjdeangelis/utl-using-proc-odstext-to-add-documentation-tabs-to-your-excel-workbook
    https://github.com/rogerjdeangelis/utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc
    https://github.com/rogerjdeangelis/utl-within-sql-join-a-a-text-or-excel-file-to-a-sas-or-foreign-table-without-proc-import
    https://github.com/rogerjdeangelis/utl_convert_a_sas_dataset_to_excel_without_sas_or_excel_only_need_R
    https://github.com/rogerjdeangelis/utl_excel-copying-a-worlbook-with-many-named-ranges-into-sas-tables
    https://github.com/rogerjdeangelis/utl_excel_add_formula_inplace
    https://github.com/rogerjdeangelis/utl_excel_add_sheet
    https://github.com/rogerjdeangelis/utl_excel_determine_type_length
    https://github.com/rogerjdeangelis/utl_excel_import_all_columns_as_character_and_preserve_long_variable_names
    https://github.com/rogerjdeangelis/utl_excel_import_data_from_a_xlsx_file_where_first_2_rows_are_header
    https://github.com/rogerjdeangelis/utl_excel_update_xlsm_workbook_using_SAS_dataset
    https://github.com/rogerjdeangelis/utl_excel_using_a_cell_value_for_the_name_of_sas_dataset
    https://github.com/rogerjdeangelis/utl_importing_three_excel_tables_that_are_in_one_sheet
    https://github.com/rogerjdeangelis/utl_import_all_excel_workbooks_created_in_the_previous_seven_days
    https://github.com/rogerjdeangelis/utl_import_excel_unicode
    https://github.com/rogerjdeangelis/utl_proc_import_columns_as_character_from_excel_linux_or_windows
    https://github.com/rogerjdeangelis/utl_put_excel_sheetnames_into_sas_macro_variable
    https://github.com/rogerjdeangelis/utl_side_by_side_excel_reports
    https://github.com/rogerjdeangelis/utl-add-monthly-worksheets-to-an-existing-yearly-excel-workbook
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-digitize-images-and-output-points-to-excel
    https://github.com/rogerjdeangelis/utl-altair-personal-slc-usng-an-excel-database-to-analyze-healthcare-labor-costs
    https://github.com/rogerjdeangelis/utl-appending-records-to-an-existing-excel-sheet
    https://github.com/rogerjdeangelis/utl-clear-named-and-unnamed-cell-ranges-in-excel
    https://github.com/rogerjdeangelis/utl-concatenating-thirty-seven-excel-tabs-while-correcting-column-types-and-using-longest-length
    https://github.com/rogerjdeangelis/utl-Delete-all-files-in-a-directory-with-a-specified-extension-ie-delete-excel-files
    https://github.com/rogerjdeangelis/utl-excel-changing-cell-contents-inside-proc-report
    https://github.com/rogerjdeangelis/utl-excel-hiding-columns-height-and-weight-in-sheet-class
    https://github.com/rogerjdeangelis/utl-excel-report-with-two-side-by-side-graphs-below_python
    https://github.com/rogerjdeangelis/utl-excel-use-the-name-of-the-last-variable-in-the-pdv-for-sheet-name
    https://github.com/rogerjdeangelis/utl-extracting-hyperlinks-from-an-excel-sheet-python
    https://github.com/rogerjdeangelis/utl-fix-excel-date-fields-on-the-excel-side-using-ms-sql-and-passthru
    https://github.com/rogerjdeangelis/utl-import-all-excel-columns-as-character
    https://github.com/rogerjdeangelis/utl-import-all-excel-columns-as-character-three-solutions
    https://github.com/rogerjdeangelis/utl-import-all-excel-dates-as-character-strings-and-convert-back-to-SAS-dates
    https://github.com/rogerjdeangelis/utl-import-csv-file-to-excel-with-leading-zeros
    https://github.com/rogerjdeangelis/utl-in-palce-updates-to-an-existing-shared-excel-workbook
    https://github.com/rogerjdeangelis/utl-preserving-excel-formatting-when-writing-to-an-existing-worksheet
    https://github.com/rogerjdeangelis/utl-programatically-downlaod-an-excel-file-from-the-web
    https://github.com/rogerjdeangelis/utl-programatically-execute-excel-vba-macro-using-sas-python
    https://github.com/rogerjdeangelis/utl-remove-sheet-from-excel-workbook
    https://github.com/rogerjdeangelis/utl-remove-sheet-from-existing-excel-worksheet-unix-and-windows-R
    https://github.com/rogerjdeangelis/utl-rename-excel-columns-to-common-names-before-creating-sas-tables
    https://github.com/rogerjdeangelis/utl-renaming-duplicate-excel-column-names-before-importing
    https://github.com/rogerjdeangelis/utl-select-the-diagonal-values-from-a-dataset-in-excel-r-wps-python
    https://github.com/rogerjdeangelis/utl-sending-a-formula-to-excel-to-reference-a-cell-in-another-sheet
    https://github.com/rogerjdeangelis/utl-update-existing-excel-sheet-in-place-using-r-dcom-client
    https://github.com/rogerjdeangelis/utl-using-only-r-openxlsx-to-add-excel-formulas-to-an-existing-sheet
    https://github.com/rogerjdeangelis/utl_excel_add_to_sheet
    https://github.com/rogerjdeangelis/utl_excel_import_only_female_students
    https://github.com/rogerjdeangelis/utl_excel_merge_two-sheets
    https://github.com/rogerjdeangelis/utl_excel_update_rectangle
    https://github.com/rogerjdeangelis/utl_excel_updating-named-ranged-cells
    https://github.com/rogerjdeangelis/utl_fix_excel_column_names_before_import
    https://github.com/rogerjdeangelis/utl_how_to_get_data_from_excel_file_into_wps_sas_procedure
    https://github.com/rogerjdeangelis/utl_ods_excel_merging_cells_after_column_header_and_before_column_names
    https://github.com/rogerjdeangelis/utl_passthru_to_excel_to_fix_column_names
    https://github.com/rogerjdeangelis/utl_renaming_duplicate_excel_columns_to_avoid_name_collisions_when_importing
    https://github.com/rogerjdeangelis/utl-casting-and-reformatting-excel-data-before-importing
    https://github.com/rogerjdeangelis/utl-does-the-excel-named-range-table-exist
    https://github.com/rogerjdeangelis/utl-excel-database-schema-and-tables-as-workbooks-sheets-and-named-ranges-update-existing-workbooks
    https://github.com/rogerjdeangelis/utl-excel-grid-of-four-reports-in-one-sheet
    https://github.com/rogerjdeangelis/utl-excel-import-individual-cells
    https://github.com/rogerjdeangelis/utl-Import-the-datepart-of-an-excel-datetime-formatted-columns
    https://github.com/rogerjdeangelis/utl-importing-inconsistently-formatted-excel-dates-numeric-and-charater-in-the-same-column
    https://github.com/rogerjdeangelis/utl-safely-sending-dates-or-datetimes-back-and-forth-to-excel
    https://github.com/rogerjdeangelis/utl-side-by-side-proc-report-output-in-pdf-html-and-excel
    https://github.com/rogerjdeangelis/utl-very-fast-summation-of-ll6-columns-in-excel-without-importing-to-sheet
    https://github.com/rogerjdeangelis/utl_excel_using_byval_sex_and_sheet_interval_bygroups_to_create_multiple_worksheets

    /*--- END ---*/
