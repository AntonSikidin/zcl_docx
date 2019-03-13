*&---------------------------------------------------------------------*
*& Report Z_TEST_DOCX1
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
report z_test_docx1.


* тут класс
include zcl_docx_class.

START-OF-SELECTION.

data
      : lo_docx type ref to lcl_docx
      , lr_data type ref to lcl_recursive_data
      .


create object lr_data.

lr_data->append_key_value(  iv_key = 'NAME' iv_value = sy-uname ).
lr_data->append_key_value(  iv_key = 'DATE' iv_value = |{ sy-datum date = environment }| ).
lr_data->append_key_value(  iv_key = 'TIME' iv_value = |{ sy-uzeit time = environment }| ).


create object lo_docx .

lo_docx->load_smw0( 'Z_TEST_DOCX1' ).


lo_docx->map_data( exporting ir_data = lr_data ).


call method lo_docx->save
  exporting
    on_desktop   = 'X'
    iv_folder    = 'report'
    iv_file_name = 'Z_TEST_DOCX1.docx'
    no_execute   = ''.