*&---------------------------------------------------------------------*
*& Report  ZDOCX_IMAGE_TEST
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
REPORT zdocx_image_test.


TYPES:
  BEGIN OF t_table_1,
    asd        TYPE string,
    img        TYPE zst_docx_image,
    name       TYPE string,
    other_text TYPE string,
  END OF t_table_1,

  t_t_table_1 TYPE TABLE OF t_table_1 WITH DEFAULT KEY,


  BEGIN OF t_data,
    table_1    TYPE t_t_table_1,
    second_img TYPE zst_docx_image,
    third_img  TYPE zst_docx_image,
  END OF t_data,

  t_t_data TYPE TABLE OF t_data WITH DEFAULT KEY.
DATA
      : gs_data TYPE t_data
      , lv_img1 TYPE xstring
      , lv_img2 TYPE xstring
      , lv_img3 TYPE xstring
      , lv_img4 TYPE xstring
      , lv_img5 TYPE xstring
      , lv_img6 TYPE xstring
      , template TYPE xstring
      , lv_bin_content TYPE xstring
      .

FIELD-SYMBOLS
               : <fs_line> TYPE t_table_1
               .


START-OF-SELECTION.



  PERFORM load_image USING 'ZDOCX_IMAGE_1' lv_img1.
  PERFORM load_image USING 'ZDOCX_IMAGE_2' lv_img2.
  PERFORM load_image USING 'ZDOCX_IMAGE_3' lv_img3.
  PERFORM load_image USING 'ZDOCX_IMAGE_4' lv_img4.
  PERFORM load_image USING 'ZDOCX_IMAGE_5' lv_img5.
  PERFORM load_image USING 'ZDOCX_IMAGE_6' lv_img6.

  gs_data-second_img-image = lv_img5.
  gs_data-third_img-image = lv_img6.

  DO 2 TIMES.

    APPEND INITIAL LINE TO gs_data-table_1 ASSIGNING <fs_line>.
    <fs_line>-img-image = lv_img1.
    <fs_line>-name = |NAME 1 index{ sy-index }|.
    <fs_line>-other_text = |other_text 1  index { sy-index }|.


    APPEND INITIAL LINE TO gs_data-table_1 ASSIGNING <fs_line>.
    <fs_line>-img-image = lv_img2.
    <fs_line>-img-cx = 8.
    <fs_line>-img-cy = 8.

    <fs_line>-name = |NAME 2 index{ sy-index }|.
    <fs_line>-other_text = |other_text 2  index { sy-index }|.

    APPEND INITIAL LINE TO gs_data-table_1 ASSIGNING <fs_line>.
    <fs_line>-asd = |asd 1 index{ sy-index }|.
    <fs_line>-img-image = lv_img3.
    <fs_line>-name = |NAME 3 index{ sy-index }|.
    <fs_line>-other_text = |other_text 3  index { sy-index }|.

    APPEND INITIAL LINE TO gs_data-table_1 ASSIGNING <fs_line>.
    <fs_line>-img-image = lv_img4.
    <fs_line>-img-cx = 6.
    <fs_line>-img-cy = 6.
    <fs_line>-name = |NAME 4 index{ sy-index }|.
    <fs_line>-other_text = |other_text 4  index { sy-index }|.

  ENDDO.


  zcl_docx3=>get_document(
    iv_w3objid    = 'ZDOCX_TEST_IMAGE'   " name of our template
*        iv_template   = lv_bin_content            " you can feed template as xstring instead of load from smw0
*      iv_on_desktop = 'X'           " by default save document on desktop
*      iv_folder     = 'report'      " in folder by default 'report'
*      iv_path       = ''            " IF iv_path IS INITIAL  save on desctop or sap_tmp folder
*      iv_file_name  = 'report.docx' " file name by default
*      iv_no_execute = 'X'            " if filled -- just get document no run office
*      iv_protect    = ''            " if filled protect document from editing, but not protect from sequence
                                     " ctrl+a, ctrl+c, ctrl+n, ctrl+v, edit
      iv_data       = gs_data  " root of our data, obligatory
*      iv_no_save    = ''            " just get binary data not save on disk
      ).



FORM load_image USING iv_w3objid TYPE  w3objid
                      iv_image TYPE xstring .

  DATA
           : lv_templ_xstr TYPE xstring
           , lt_mime TYPE TABLE OF w3mime
           .

  DATA(ls_key) = VALUE wwwdatatab(
  relid = 'MI'
  objid = iv_w3objid ).

  CALL FUNCTION 'WWWDATA_IMPORT'
    EXPORTING
      key    = ls_key
    TABLES
      mime   = lt_mime
    EXCEPTIONS
      OTHERS = 1.
  IF sy-subrc <> 0.
    RETURN.
  ENDIF.

  TRY.
      iv_image = cl_bcs_convert=>xtab_to_xstring( lt_mime ).
    CATCH cx_bcs.
      RETURN.
  ENDTRY.

ENDFORM.
