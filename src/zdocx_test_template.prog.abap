*&---------------------------------------------------------------------*
*& Report  ZDOCX_TEST_TEMPLATE
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*
REPORT zdocx_test_template.
INCLUDE zdocx_test_template_f01.

PARAMETERS
: p_obj TYPE  w3objid
, p_file TYPE string LOWER CASE
, p_type TYPE strukname
.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.  " Обробщик події F4
  PERFORM get_file_path CHANGING p_file.


START-OF-SELECTION.

  DATA
        : dref TYPE REF TO data

        .

  FIELD-SYMBOLS
                 : <fs_data> TYPE any
                 .


  CREATE DATA dref TYPE (p_type).

  ASSIGN dref->* TO <fs_data>.

  PERFORM test_template USING <fs_data> p_obj p_file.
