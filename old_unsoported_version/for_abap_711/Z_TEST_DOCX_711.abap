*&---------------------------------------------------------------------*
*& Report  Z_TEST_DOCX
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
REPORT z_test_docx.
** ### #####
INCLUDE zcl_docx_class_711.

START-OF-SELECTION.

  DATA
        : lo_docx TYPE REF TO lcl_docx
        , lr_data TYPE REF TO lcl_recursive_data
        , lv_str TYPE string
        .


  CREATE OBJECT lr_data.

  lr_data->append_key_value( iv_key = 'name' iv_value = sy-uname ).
  CONCATENATE sy-datum+6(2) '.' sy-datum+4(2) '.'  sy-datum(4) INTO lv_str.

  lr_data->append_key_value( iv_key = 'date' iv_value = lv_str ).

  CONCATENATE sy-uzeit(2) ':' sy-uzeit+2(2) ':' sy-uzeit+4(2)  INTO lv_str.
  lr_data->append_key_value( iv_key = 'time' iv_value = lv_str ).







  DATA
        : lt_sflight TYPE TABLE OF sflight

        .


  SELECT *
    INTO TABLE lt_sflight
    UP TO 5 ROWS
    FROM sflight.



  lr_data->append_key_table( iv_key = 'line1' iv_table  = lt_sflight ).





  DATA
        :  lt_carrid  TYPE TABLE OF s_carr_id
        ,  lt_sflight1 TYPE TABLE OF sflight
        ,  lr_doc TYPE REF TO lcl_recursive_data
        .

  SELECT DISTINCT carrid INTO TABLE lt_carrid
    UP TO 5 ROWS
    FROM sflight.

    FIELD-SYMBOLs
          :  <fs_carrid> type s_carr_id
          .

  LOOP AT lt_carrid ASSIGNING <fs_carrid>.

    REFRESH lt_sflight1.

    SELECT * INTO TABLE lt_sflight1
      FROM  sflight
      UP TO 5 ROWS
      WHERE carrid = <fs_carrid>.

    lr_doc = lr_data->create_document( 'DOC1' ).

    lr_doc->append_key_value( iv_key = 'carrid1' iv_value = <fs_carrid> ).
    lr_doc->append_key_table( iv_key = 'line2' iv_table = lt_sflight1 ).


  ENDLOOP.


  CREATE OBJECT lo_docx .

  lo_docx->load_smw0( 'Z_TEST_DOCX' ).

  lo_docx->map_data( EXPORTING ir_data = lr_data ).

  CALL METHOD lo_docx->save
    EXPORTING
      on_desktop   = 'X'
      iv_folder    = 'report'
      iv_file_name = 'report.docx'
      no_execute   = ''.
