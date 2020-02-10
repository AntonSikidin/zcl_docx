#Tool to create Microsoft Word docx from abap.


Installation
Install package via ABAPGIT https://docs.abapgit.org/guide-install.html
  
![alt_text](images/z001_1.png "image_tooltip")  
  
![alt_text](images/z001_2.png "image_tooltip")  
  
![alt_text](images/z001_3.png "image_tooltip")    
![alt_text](images/z001_4.png "image_tooltip")    
![alt_text](images/z001_5.png "image_tooltip")  

Where there are mistakes, just click Pull_zip twice.

For example, the following document should be created:
  
![alt_text](images/z001_6.png "image_tooltip")  

Initially, define variables and repeated parts  
![alt_text](images/z001_7.png "image_tooltip")  
 
Variable – simple text.
Table line – contains table row that may consist of several or zero variables and text.
Document – contain several or zero texts, variables, table row.

We have something like this   
![alt_text](images/z001_8.png "image_tooltip")  

Let reduce our document
  
![alt_text](images/z001_9.png "image_tooltip")  
 
Structure of data for our document shoudt be like 
├── name  
├── date  
├── time  
├── document1  
│   ├── data_row  
│   │   ├── CARRID  
│   │   ├── CLASS  
│   │   ├── FORCURAM  
│   │   ├── LOCCURAM  
│   │   ├── LOCCURKEY  
│   │   └── ORDER_DATE  
│   │    .  
│   │    .  
│   ├── data_row  
│   │   ├── CARRID  
│   │   ├── CLASS  
│   │   ├── FORCURAM  
│   │   ├── LOCCURAM  
│   │   ├── LOCCURKEY  
│   │   └── ORDER_DATE  
│   └── subtotal  
│       ├── FORCURKEY  
│       └── LOCCURAM  
│    .  
│    .  
├── document1  
│   ├── data_row  
│   │   ├── CARRID  
│   │   ├── CLASS  
│   │   ├── FORCURAM  
│   │   ├── LOCCURAM  
│   │   ├── LOCCURKEY  
│   │   └── ORDER_DATE  
│   │    .  
│   │    .  
│   ├── data_row  
│   │   ├── CARRID  
│   │   ├── CLASS  
│   │   ├── FORCURAM  
│   │   ├── LOCCURAM  
│   │   ├── LOCCURKEY  
│   │   └── ORDER_DATE  
│   └── subtotal  
│       ├── FORCURKEY  
│       └── LOCCURAM  
├── total   
│   ├── FORCURKEY  
│   └── LOCCURAM  
├── sign  
│   ├── NAME_FIRST   
│   └── NAME_LAST  
│   .  
│   .  
└── sign      
    ├── NAME_FIRST  
    └── NAME_LAST  
  
   
Or simple

├── name  (value)  
├── date  (value)  
├── time  (value)  
├── document1   (document repeated)  
│   ├── data_row  (table repeated)  
│   │   ├── CARRID  
│   │   ├── CLASS  
│   │   ├── FORCURAM  
│   │   ├── LOCCURAM  
│   │   ├── LOCCURKEY  
│   │   └── ORDER_DATE  
│   └── subtotal (table of 1 row)  
│       ├── FORCURKEY  
│       └── LOCCURAM  
├── total (table of 1 row)  
│   ├── FORCURKEY  
│   └── LOCCURAM  
└── sign    (table repeated)  
    ├── NAME_FIRST  
    └── NAME_LAST  
  
  
Let sign variable placeholder.

At first toggle developer toolbar. https://www.google.com/search?q=microsoft+office+16+toggle+developer+toolbar

Select  
![alt_text](images/z001_10.png "image_tooltip")  
Make tag  
![alt_text](images/z001_11.png "image_tooltip")  


Click properties   
![alt_text](images/z001_12.png "image_tooltip")  

Enter tag name  
![alt_text](images/z001_13.png "image_tooltip")  

Tag name is case insensitive; all the way, it will be converted to uppercase

Value can have any tag name, value inside table row must have name of field of row structure 
For example,  tag name for sign row    
![alt_text](images/z001_14.png "image_tooltip")  

Name all variable placeholder
In design time it must look like this    
![alt_text](images/z001_15.png "image_tooltip")  


At second we mark placeholder for table row(data_row, subtotal, total, sign)
Place mouse cursor to the left of row   
![alt_text](images/z001_16.png "image_tooltip")  

Click   
![alt_text](images/z001_17.png "image_tooltip")    
![alt_text](images/z001_18.png "image_tooltip")  

Properties, data_row  
![alt_text](images/z001_19.png "image_tooltip")  

Subtotal  
![alt_text](images/z001_20.png "image_tooltip")  

Total  
![alt_text](images/z001_21.png "image_tooltip")  

Sign  
![alt_text](images/z001_22.png "image_tooltip")  

Then template in design mode must look like  
![alt_text](images/z001_23.png "image_tooltip")  

Now, we move on to the task with an asterisk. In 99% cases you do not need this. I just show opportunity how to make more complex document.
If you want, you can make infinite depth of your document.
We need join 2 rows in one placeholder. Unfortunately, Microsoft Office cannot make placeholder for 2 rows. It can make placeholder for one row or for a whole table.

There we have 2 way:
1)	Cheap and wrong
2)	More complex and True

First way:

Copy your table 3 times   
![alt_text](images/z001_24.png "image_tooltip")  

Cut unnecessary part from each table   
![alt_text](images/z001_25.png "image_tooltip")  

Make placeholder for whole second table  
![alt_text](images/z001_26.png "image_tooltip")  


Remove spaces between tables
In my case, it look like  
![alt_text](images/z001_27.png "image_tooltip")  

I don’t like this, maybe, I cannot work with tables.

Second way I prefer:
Select two rows, make placeholder. Office create placeholder for first row. It is ok.  
![alt_text](images/z001_28.png "image_tooltip")  
  
![alt_text](images/z001_29.png "image_tooltip")  

Now, save the template, close office.

Rename template.docx to template.zip  
![alt_text](images/z001_30.png "image_tooltip")    
![alt_text](images/z001_31.png "image_tooltip")  

Unpack to subfolder  
![alt_text](images/z001_32.png "image_tooltip")  

Navigate inside, then in subfolder ‘word’   
![alt_text](images/z001_33.png "image_tooltip")  

We need notepad++ https://notepad-plus-plus.org/downloads/

Open document.xml with notepad++   
![alt_text](images/z001_34.png "image_tooltip")  

Navigate plugin->plugins admin..  
![alt_text](images/z001_35.png "image_tooltip")  

Search “Xml tools”  
![alt_text](images/z001_36.png "image_tooltip")  

Install
Now you can pretty print xml document  
![alt_text](images/z001_37.png "image_tooltip")  
  
![alt_text](images/z001_38.png "image_tooltip")  

Find placeholder we created. In our case it document2  
![alt_text](images/z001_39.png "image_tooltip")  
Ctrl+f
  
![alt_text](images/z001_40.png "image_tooltip")  

We can see our placeholder
<w:tag w:val="document2"/>  
![alt_text](images/z001_41.png "image_tooltip")  

It starts in line 1293 with tag <w:sdt>

Collapse tree to see where it end  
![alt_text](images/z001_42.png "image_tooltip")  

It end near line 1621  
![alt_text](images/z001_43.png "image_tooltip")  

Now collapse next placeholder to see where it ends  
![alt_text](images/z001_44.png "image_tooltip")  

It ends near line  1757   
![alt_text](images/z001_45.png "image_tooltip")  

Expand all.
Go to line 1621
Cut 2 lines  1619, 1620   
![alt_text](images/z001_46.png "image_tooltip")  

        </w:sdtContent>
      </w:sdt>  
![alt_text](images/z001_47.png "image_tooltip")  

Navigate to line 1757  
![alt_text](images/z001_48.png "image_tooltip")  

Insert 2 lines before line 1757  
![alt_text](images/z001_49.png "image_tooltip")  

Yellow – inserted lines.

Save, close.

Navigate 1 level up   
![alt_text](images/z001_50.png "image_tooltip")    
![alt_text](images/z001_51.png "image_tooltip")  

Select all file at this level, add to zip archive   
![alt_text](images/z001_52.png "image_tooltip")    
![alt_text](images/z001_53.png "image_tooltip")  

Rename template_docx.zip to template_docx.docx

Now we can see placeholder hold 2 rows  
![alt_text](images/z001_54.png "image_tooltip")  

Go to transaction smw0
Select Binary data, enter  
![alt_text](images/z001_55.png "image_tooltip")  
Object name                     Z_TEST_DOCX2
  
![alt_text](images/z001_56.png "image_tooltip")  
Create  
![alt_text](images/z001_57.png "image_tooltip")    
![alt_text](images/z001_58.png "image_tooltip")  
  
![alt_text](images/z001_59.png "image_tooltip")  

Now let’s create test program.

You can read this program with comments or see program Z_TEST_DOCX2 that already exists in the package.
Template Z_TEST_DOCX2 also exists in the package

```
*&---------------------------------------------------------------------*
*& Report Z_TEST_DOCX_2
*&---------------------------------------------------------------------*
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

REPORT z_test_docx_2.

START-OF-SELECTION.


  TYPES
"  structure to hold our data
  : BEGIN OF t_data
  ,   carrid  TYPE s_carr_id
  ,   class  TYPE s_class
  ,   forcuram  TYPE s_f_cur_pr
  ,   forcurkey  TYPE s_curr
  ,   loccuram  TYPE s_l_cur_pr
  ,   loccurkey  TYPE s_currcode
  ,   order_date  TYPE s_bdate
  , END OF t_data
  .



  DATA
        : lt_carrid TYPE TABLE OF s_carr_id
        , lt_total TYPE TABLE OF t_data
        , lt_sub_total TYPE TABLE OF t_data
        , lt_sub_total_tmp TYPE TABLE OF t_data
        , lt_data TYPE TABLE OF t_data
        , lt_tmp TYPE TABLE OF t_data

        , lt_adrp TYPE TABLE OF adrp

        .

  " select data for "sign" placeholder
  SELECT * INTO TABLE lt_adrp FROM adrp UP TO 5 ROWS.

*  REFRESH lt_adrp.
*
*
*  APPEND INITIAL LINE TO lt_adrp ASSIGNING FIELD-SYMBOL(<fs_adrp>).
*  <fs_adrp>-name_first = 'Renee'.
*  <fs_adrp>-name_last = 'Villegas'.
*  APPEND INITIAL LINE TO lt_adrp ASSIGNING <fs_adrp>.
*
*  <fs_adrp>-name_first = 'Meerab'.
*  <fs_adrp>-name_last = 'Finnegan'.
*  APPEND INITIAL LINE TO lt_adrp ASSIGNING <fs_adrp>.
*
*  <fs_adrp>-name_first = 'Jozef'.
*  <fs_adrp>-name_last = 'Beil'.
*  APPEND INITIAL LINE TO lt_adrp ASSIGNING <fs_adrp>.
*
*  <fs_adrp>-name_first = 'Leonard'.
*  <fs_adrp>-name_last = 'Yates'.
*  APPEND INITIAL LINE TO lt_adrp ASSIGNING <fs_adrp>.
*
*  <fs_adrp>-name_first = 'Kyron'.
*  <fs_adrp>-name_last = 'Stevens'.





  " select data to display in main table,
  " may be it may create in easy and in proper way,
  " but now it just data for proof of concept

  SELECT *
    INTO CORRESPONDING FIELDS OF TABLE lt_tmp
    FROM sbook
    .

  lt_data = lt_tmp. "make backup, because our data corrupted, while calculate total and subtotal

  LOOP AT lt_tmp ASSIGNING FIELD-SYMBOL(<fs_tmp>).
    COLLECT <fs_tmp>-carrid INTO lt_carrid.  " every document1 or document2 hold data by single carrid

    CLEAR
    : <fs_tmp>-class
    , <fs_tmp>-forcurkey
    , <fs_tmp>-loccurkey
    , <fs_tmp>-order_date
    .

    COLLECT <fs_tmp> INTO lt_sub_total . "calculate  subtotal

    CLEAR
    : <fs_tmp>-carrid
    .

    COLLECT <fs_tmp> INTO lt_total. " calculate total.

  ENDLOOP.


  lt_tmp = lt_data.

  REFRESH lt_data.

  LOOP AT lt_sub_total ASSIGNING FIELD-SYMBOL(<fs_sub_total>).

    DATA
          : lv_i TYPE i
          .
    CLEAR lv_i.

    LOOP AT lt_tmp ASSIGNING <fs_tmp> WHERE carrid = <fs_sub_total>-carrid.
      ADD 1 TO lv_i.

      IF lv_i > 4 . " four row for proof of concept is enought
        EXIT.
      ENDIF.

      APPEND <fs_tmp> TO lt_data.

    ENDLOOP.

  ENDLOOP.


  " craete main wariable to hold our data
  DATA(lr_data) = zcl_docx2=>create_data( ).


  " fill place holder for simple variable, iv_key case insensitive
  lr_data->append_key_value(  iv_key = 'name' iv_value = sy-uname ).
  lr_data->append_key_value(  iv_key = 'date' iv_value = |{ sy-datum DATE = ENVIRONMENT }| ).
  lr_data->append_key_value(  iv_key = 'time' iv_value = |{ sy-uzeit TIME = ENVIRONMENT }| ).

  " fill place holder for total ( table of 1 row)
  lr_data->append_key_table( iv_key = 'total' iv_table = lt_total ).

  "fill  placeholder for sign, iv_key always case insensitive
  lr_data->append_key_table( iv_key = 'sign' iv_table = lt_adrp ).


  " fill our main table

  LOOP AT lt_sub_total ASSIGNING <fs_sub_total>.

    REFRESH
    : lt_tmp
    , lt_sub_total_tmp
    .

    APPEND <fs_sub_total> TO lt_sub_total_tmp. " subtotal table of 1 row
    LOOP AT lt_data ASSIGNING FIELD-SYMBOL(<fs_data>) WHERE carrid = <fs_sub_total>-carrid.
      APPEND <fs_data> TO lt_tmp.  " table for placeholder "data_row"
    ENDLOOP.


    " create  variable to hold data as subchild of our main variable

    DATA(lr_document1) = lr_data->create_document( 'document1' ).

    " fill placeholder "data_row" in document "document1"
    lr_document1->append_key_table( iv_key = 'data_row' iv_table = lt_tmp ).

    " fill placeholder "subtotal" in document "document1"
    lr_document1->append_key_table( iv_key = 'subtotal' iv_table = lt_sub_total_tmp ).


    " our temlate contain 2 variant of main table: easy and wrong, more complex and true

    "create variable to hold data for document2
    DATA(lr_document2) = lr_data->create_document( 'document2' ).

    " fill placeholder "data_row" in document "document2"
    lr_document2->append_key_table( iv_key = 'data_row' iv_table = lt_tmp ).

    " fill placeholder "subtotal" in document "document2"
    lr_document2->append_key_table( iv_key = 'subtotal' iv_table = lt_sub_total_tmp ).

  ENDLOOP.



*final moment get document

  DATA
        : lv_document TYPE xstring  " variable to hold generated document, can be omitted
        .



  lv_document = zcl_docx2=>get_document(
      iv_w3objid    = 'Z_TEST_DOCX2' " name of our template, obligatory
*      iv_on_desktop = 'X'           " by default save document on desktop
*      iv_folder     = 'report'      " in folder by default 'report'
*      iv_path       = ''            " IF iv_path IS INITIAL  save on desctop or sap_tmp folder
*      iv_file_name  = 'report.docx' " file name by default
*      iv_no_execute = ''            " if filled -- just get document no run office
*      iv_protect    = ''            " if filled protect document from editing, but not protect from sequence
                                     " ctrl+a, ctrl+c, ctrl+n, ctrl+v, edit
      ir_data       = lr_data        " root of our data, obligatory
*      iv_no_save    = ''            " just get binary data not save on disk
      ).
```
	  
Run program and get something like this
  
![alt_text](images/z001_60.png "image_tooltip")    
![alt_text](images/z001_61.png "image_tooltip")  



Both variants seem acceptable.  For the record, second gives you the most options.
Use it to create whatever you want.

