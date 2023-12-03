*&---------------------------------------------------------------------*
*& Report ZDOCX_EXAMPLE
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdocx_example_checkbox.

*describe types #asd

TYPES:
 begin of t_TABLE_1,
       ROW_DISAPPEAR type string,
       ROW_DISAPPEAR1 type string,
       ROW_DISAPPEAR2 type string,
       ROW_DISAPPEAR3 type string,
 end of t_TABLE_1,

 t_t_TABLE_1 type table of t_TABLE_1 with empty key,


 begin of t_TABLE_2,
       CHK3 type string,
       TEXT3 type string,
 end of t_TABLE_2,

 t_t_TABLE_2 type table of t_TABLE_2 with empty key,




 begin of t_data,
       CHECKED type string,
       UNCHECKED type string,
 TABLE_1 type t_t_TABLE_1,
 TABLE_2 type t_t_TABLE_2,
 end of t_data,

 t_t_data type table of t_data with empty key.




DATA
      : gs_templ_data TYPE t_data
      .

gs_templ_data-checked = 'Y'.
gs_templ_data-unchecked = ''.

APPEND INITIAL LINE TO gs_templ_data-table_2 ASSIGNING field-symbol(<fs_2>).
<fs_2>-chk3 = 'X'.
<fs_2>-text3 = 'row1'.

APPEND INITIAL LINE TO gs_templ_data-table_2 ASSIGNING <fs_2>.
<fs_2>-chk3 = ''.
<fs_2>-text3 = 'row2'.

APPEND INITIAL LINE TO gs_templ_data-table_2 ASSIGNING <fs_2>.
<fs_2>-chk3 = 'x'.
<fs_2>-text3 = 'row3'.

APPEND INITIAL LINE TO gs_templ_data-table_2 ASSIGNING <fs_2>.
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
