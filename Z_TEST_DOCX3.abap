*&---------------------------------------------------------------------*
*& Report Z_TEST_DOCX_3
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
report z_test_docx_3.


* тут класс
include zcl_docx_class.

start-of-selection.


  types

  : begin of t_data
  ,   carrid  type s_carr_id
  ,   class	type s_class
  ,   forcuram  type s_f_cur_pr
  ,   forcurkey	type s_curr
  ,   loccuram  type s_l_cur_pr
  ,   loccurkey	type s_currcode
  ,   order_date  type s_bdate
  , end of t_data
  .



  data
        : lo_docx type ref to lcl_docx
        , lr_data type ref to lcl_recursive_data
        , lr_tmp_data type ref to lcl_recursive_data

        , lt_carrid type table of s_carr_id
        , lt_total type table of t_data
        , lt_sub_total type table of t_data
        , lt_sub_total_tmp type table of t_data
        , lt_data type table of t_data
        , lt_tmp type table of t_data

        , lt_adrp type table of adrp

        .



  create object lr_data.



  select *
    into corresponding fields of table lt_tmp
    from sbook
*    where carrid in ('AZ', 'DL')
    .

  lt_data = lt_tmp.

  loop at lt_tmp assigning field-symbol(<fs_tmp>).
    collect <fs_tmp>-carrid into lt_carrid.

    clear
    : <fs_tmp>-class
    , <fs_tmp>-forcurkey
    , <fs_tmp>-loccurkey
    , <fs_tmp>-order_date
    .

    collect <fs_tmp> into lt_sub_total .


    clear
    : <fs_tmp>-carrid
    .

    collect <fs_tmp> into lt_total.

  endloop.



  select * into table lt_adrp from adrp up to 5 rows.
  lt_tmp = lt_data.

  refresh lt_data.

  loop at lt_sub_total assigning field-symbol(<fs_sub_total>).

    data
          : lv_i type i
          .
    clear lv_i.

    loop at lt_tmp assigning <fs_tmp> where carrid = <fs_sub_total>-carrid.
      add 1 to lv_i.

      if lv_i > 4 .
        exit.
      endif.

      append <fs_tmp> to lt_data.

    endloop.

  endloop.

*doc1
*podpis
*total
*subtotal

  create object lr_data.

  lr_data->append_key_value(  iv_key = 'NAME' iv_value = sy-uname ).
  lr_data->append_key_value(  iv_key = 'DATE' iv_value = |{ sy-datum date = environment }| ).
  lr_data->append_key_value(  iv_key = 'TIME' iv_value = |{ sy-uzeit time = environment }| ).

  loop at lt_sub_total assigning <fs_sub_total>.

    refresh
    : lt_tmp
    , lt_sub_total_tmp
    .

    append <fs_sub_total> to lt_sub_total_tmp.
    loop at lt_data assigning field-symbol(<fs_data>) where carrid = <fs_sub_total>-carrid.
      append <fs_data> to lt_tmp.
    endloop.


    lr_tmp_data = lr_data->create_document( 'DOC1' ).

    lr_tmp_data->append_key_table( iv_key = 'DATA' iv_table = lt_tmp ).
    lr_tmp_data->append_key_table( iv_key = 'SUBTOTAL' iv_table = lt_sub_total_tmp ).

  endloop.


  lr_data->append_key_table( iv_key = 'TOTAL' iv_table = lt_total ).
  lr_data->append_key_table( iv_key = 'PODPIS' iv_table = lt_adrp ).






  create object lo_docx .

  lo_docx->load_smw0( 'Z_TEST_DOCX3' ).


  lo_docx->map_data( exporting ir_data = lr_data ).


  call method lo_docx->save
    exporting
      on_desktop   = 'X'
      iv_folder    = 'report'
      iv_file_name = 'Z_TEST_DOCX3.docx'
      no_execute   = ''.