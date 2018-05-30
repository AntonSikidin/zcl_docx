PROGRAM xz1 .

INCLUDE zcl_docx1.
*INCLUDE zcl_docx_class.


START-OF-SELECTION.


  TYPES

  : BEGIN OF t_data
  ,   carrid  TYPE s_carr_id
  ,   class	TYPE s_class
  ,   forcuram  TYPE s_f_cur_pr
  ,   forcurkey	TYPE s_curr
  ,   loccuram  TYPE s_l_cur_pr
  ,   loccurkey	TYPE s_currcode
  ,   order_date  TYPE s_bdate
  , END OF t_data
  .



  DATA
        : lo_docx TYPE REF TO lcl_docx
        , lr_data TYPE REF TO lcl_recursive_data
        , lr_tmp_data TYPE REF TO lcl_recursive_data

        , lt_carrid TYPE TABLE OF s_carr_id
        , lt_total TYPE TABLE OF t_data
        , lt_sub_total TYPE TABLE OF t_data
        , lt_sub_total_tmp TYPE TABLE OF t_data
        , lt_data TYPE TABLE OF t_data
        , lt_tmp TYPE TABLE OF t_data

        , lt_adrp TYPE TABLE OF adrp

        .



  CREATE OBJECT lr_data.



  SELECT *
    INTO CORRESPONDING FIELDS OF TABLE lt_tmp
    FROM sbook
*    where carrid in ('AZ', 'DL')
    .

  lt_data = lt_tmp.

  LOOP AT lt_tmp ASSIGNING FIELD-SYMBOL(<fs_tmp>).
    COLLECT <fs_tmp>-carrid INTO lt_carrid.

    CLEAR
    : <fs_tmp>-class
    , <fs_tmp>-forcurkey
    , <fs_tmp>-loccurkey
    , <fs_tmp>-order_date
    .

    COLLECT <fs_tmp> INTO lt_sub_total .


    CLEAR
    : <fs_tmp>-carrid
    .

    COLLECT <fs_tmp> INTO lt_total.

  ENDLOOP.



  SELECT * INTO TABLE lt_adrp FROM adrp UP TO 5 ROWS.
  lt_tmp = lt_data.

  REFRESH lt_data.

  LOOP AT lt_sub_total ASSIGNING FIELD-SYMBOL(<fs_sub_total>).

    DATA
          : lv_i TYPE i
          .
    CLEAR lv_i.

    LOOP AT lt_tmp ASSIGNING <fs_tmp> WHERE carrid = <fs_sub_total>-carrid.
      ADD 1 TO lv_i.

      IF lv_i > 4 .
        EXIT.
      ENDIF.

      APPEND <fs_tmp> TO lt_data.

    ENDLOOP.

  ENDLOOP.

*doc1
*podpis
*total
*subtotal

  CREATE OBJECT lr_data.

  lr_data->append_key_value(  iv_key = 'name' iv_value = sy-uname ).
  lr_data->append_key_value(  iv_key = 'date' iv_value = |{ sy-datum DATE = ENVIRONMENT }| ).
  lr_data->append_key_value(  iv_key = 'time' iv_value = |{ sy-uzeit TIME = ENVIRONMENT }| ).

  LOOP AT lt_sub_total ASSIGNING <fs_sub_total>.

    REFRESH
    : lt_tmp
    , lt_sub_total_tmp
    .

    APPEND <fs_sub_total> TO lt_sub_total_tmp.
    LOOP AT lt_data ASSIGNING FIELD-SYMBOL(<fs_data>) WHERE carrid = <fs_sub_total>-carrid.
      APPEND <fs_data> TO lt_tmp.
    ENDLOOP.


    lr_tmp_data = lr_data->create_document( 'doc1' ).

    lr_tmp_data->append_key_table( iv_key = 'data' iv_table = lt_tmp ).
    lr_tmp_data->append_key_table( iv_key = 'subtotal' iv_table = lt_sub_total_tmp ).

  ENDLOOP.


  lr_data->append_key_table( iv_key = 'total' iv_table = lt_total ).
  lr_data->append_key_table( iv_key = 'podpis' iv_table = lt_adrp ).



  CREATE OBJECT lo_docx .

  lo_docx->load_smw0( 'Z_TEST_DOCX2' ).

  lo_docx->map_data( EXPORTING ir_data = lr_data ).

  CALL METHOD lo_docx->save
    EXPORTING
      on_desktop   = 'X'
      iv_folder    = 'report'
      iv_file_name = 'report.docx'
      no_execute   = ''.