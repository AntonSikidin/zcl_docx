*&---------------------------------------------------------------------*
*& Report ZDOCX_EXAMPLE
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdocx_example_checkbox.

*describe types #asd

TYPES
: BEGIN OF t_table1
,       row_disappear TYPE string
,       row_disappear1 TYPE string
,       row_disappear2 TYPE string
,       row_disappear3 TYPE string
, END OF t_table1

, tt_table1 TYPE TABLE OF t_table1 WITH EMPTY KEY


, BEGIN OF t_table2
,       chk3 TYPE string
,       text3 TYPE string
, END OF t_table2

, tt_table2 TYPE TABLE OF t_table2 WITH EMPTY KEY


, BEGIN OF t_data
,       checked TYPE string
,       unchecked TYPE string
, table1 TYPE tt_table1
, table2 TYPE tt_table2
, END OF t_data

, tt_data TYPE TABLE OF t_data WITH EMPTY KEY


.



DATA
      : gs_templ_data TYPE t_data
      .

gs_templ_data-checked = 'Y'.
gs_templ_data-unchecked = ''.

APPEND INITIAL LINE TO gs_templ_data-table2 ASSIGNING field-symbol(<fs_2>).
<fs_2>-chk3 = 'X'.
<fs_2>-text3 = 'row1'.

APPEND INITIAL LINE TO gs_templ_data-table2 ASSIGNING <fs_2>.
<fs_2>-chk3 = ''.
<fs_2>-text3 = 'row2'.

APPEND INITIAL LINE TO gs_templ_data-table2 ASSIGNING <fs_2>.
<fs_2>-chk3 = 'x'.
<fs_2>-text3 = 'row3'.

APPEND INITIAL LINE TO gs_templ_data-table2 ASSIGNING <fs_2>.
<fs_2>-chk3 = ''.
<fs_2>-text3 = 'row4'.

*second case: save on desctop and open document

zcl_docx3=>get_document(
    iv_w3objid    = 'ZDOCX_EXAMPLE_CHECKBOX'   " name of our template
*      iv_template   = ''            " you can feed template as xstring instead of load from smw0
*      iv_on_desktop = 'X'           " by default save document on desktop
*      iv_folder     = 'report'      " in folder by default 'report'
*      iv_path       = ''            " IF iv_path IS INITIAL  save on desctop or sap_tmp folder
*      iv_file_name  = 'report.docx' " file name by default
*      iv_no_execute = ''            " if filled -- just get document no run office
*      iv_protect    = ''            " if filled protect document from editing, but not protect from sequence
                                   " ctrl+a, ctrl+c, ctrl+n, ctrl+v, edit
    iv_data       = gs_templ_data  " root of our data, obligatory
*      iv_no_save    = ''            " just get binary data not save on disk
    ).
