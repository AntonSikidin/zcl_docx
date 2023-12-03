*&---------------------------------------------------------------------*
*& Report ZDOCX_GET_TYPES
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zdocx_get_types.

TYPES
: t_t_text TYPE TABLE OF text80 WITH DEFAULT KEY
.

DATA
      : gv_data_model_xml TYPE string
      .

PARAMETERS: p_fpath TYPE string OBLIGATORY LOWER CASE.

PARAMETERS: p_normal RADIOBUTTON GROUP rad1 DEFAULT 'X'
          , p_other RADIOBUTTON GROUP rad1
          .

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fpath.  " Обробщик події F4
  PERFORM get_file_path CHANGING p_fpath. " Вызов подпрограммы с передачей параметра f_path

START-OF-SELECTION.

  PERFORM get_data_model.
  PERFORM check_save_name.
  PERFORM show_data_model.


FORM check_save_name.
*variable
  DATA
        : lv_regex TYPE string VALUE '<([^>/]+)/?>'
        , lt_result_tab TYPE match_result_tab
        , lt_names TYPE TABLE OF string
        , lv_names TYPE string
        , lt_bad_names TYPE TABLE OF string
        , lv_tabname TYPE tabname
        , lr_text       TYPE REF TO cl_demo_text
        , lv_text TYPE char80
        .
*get all names
  FIND ALL OCCURRENCES OF REGEX lv_regex IN gv_data_model_xml RESULTS lt_result_tab.


  LOOP AT lt_result_tab ASSIGNING FIELD-SYMBOL(<fs_result>).
    LOOP AT <fs_result>-submatches ASSIGNING FIELD-SYMBOL(<fs_submatch>).
      lv_names = gv_data_model_xml+<fs_submatch>-offset(<fs_submatch>-length).
      TRANSLATE lv_names TO UPPER CASE.
      COLLECT lv_names INTO lt_names.
    ENDLOOP.
  ENDLOOP.

  SORT lt_names.

  DELETE ADJACENT DUPLICATES FROM lt_names.
*check names

  LOOP AT lt_names ASSIGNING FIELD-SYMBOL(<fs_names>).

    SELECT SINGLE tabname INTO lv_tabname
      FROM dd02l
      WHERE tabname = <fs_names>.

    IF sy-subrc = 0.
      APPEND <fs_names> TO lt_bad_names.
    ENDIF.

  ENDLOOP.

*show bad names
  IF lt_bad_names IS NOT INITIAL.


    lr_text = cl_demo_text=>get_handle( ).

    lr_text->add_line( 'This names exist in table DD02l, thats why they cannot be used as tag names' ).
    lr_text->add_line( '' ).


    LOOP AT lt_bad_names ASSIGNING <fs_names>.
      lv_text = <fs_names>.
      lr_text->add_line( lv_text ).
*    WRITE : / <ls_source>.

    ENDLOOP.

    lr_text->display( ).

    MESSAGE 'Bad names in template' TYPE 'E'.

  ENDIF.





ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
FORM get_file_path CHANGING cv_path TYPE string.
  CLEAR cv_path.

  DATA:
    lv_rc          TYPE  i,
    lv_user_action TYPE  i,
    lt_file_table  TYPE  filetable,
    ls_file_table  LIKE LINE OF lt_file_table.

  cl_gui_frontend_services=>file_open_dialog(
  EXPORTING
    window_title        = 'select template  docx'
    multiselection      = ''
    default_extension   = '*.docx'
    file_filter         = 'Text file (*.docx)|*.docx|All (*.*)|*.*'
  CHANGING
    file_table          = lt_file_table
    rc                  = lv_rc
    user_action         = lv_user_action
  EXCEPTIONS
    OTHERS              = 1
    ).
  IF sy-subrc = 0.
    IF lv_user_action = cl_gui_frontend_services=>action_ok.
      IF lt_file_table IS NOT INITIAL.
        READ TABLE lt_file_table INTO ls_file_table INDEX 1.
        IF sy-subrc = 0.
          cv_path = ls_file_table-filename.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM.                    " Get_file_path
*&---------------------------------------------------------------------*
*&      Form  GET_DATA_MODEL
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM get_data_model .



  DATA
        : lt_file_data TYPE solix_tab
        , lv_filelength TYPE i
        .
  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename                = p_fpath          " Name of file
      filetype                = 'BIN'                 " File Type (ASCII, Binary)
    IMPORTING
      filelength              = lv_filelength   " File Length
    CHANGING
      data_tab                = lt_file_data          " Transfer table for file contents
    EXCEPTIONS
      file_open_error         = 1                " File does not exist and cannot be opened
      file_read_error         = 2                " Error when reading file
      no_batch                = 3                " Cannot execute front-end function in background
      gui_refuse_filetransfer = 4                " Incorrect front end or error on front end
      invalid_type            = 5                " Incorrect parameter FILETYPE
      no_authority            = 6                " No upload authorization
      unknown_error           = 7                " Unknown error
      bad_data_format         = 8                " Cannot Interpret Data in File
      header_not_allowed      = 9                " Invalid header
      separator_not_allowed   = 10               " Invalid separator
      header_too_long         = 11               " Header information currently restricted to 1023 bytes
      unknown_dp_error        = 12               " Error when calling data provider
      access_denied           = 13               " Access to File Denied
      dp_out_of_memory        = 14               " Not enough memory in data provider
      disk_full               = 15               " Storage medium is full.
      dp_timeout              = 16               " Data provider timeout
      not_supported_by_gui    = 17               " GUI does not support this
      error_no_gui            = 18               " GUI not available
      OTHERS                  = 19 ).



  DATA
        : lv_bin_content TYPE xstring
        .

  lv_bin_content = cl_bcs_convert=>solix_to_xstring(

       it_solix = lt_file_data               " Input data
       iv_size  = lv_filelength ).


  DATA
        : lo_zip TYPE REF TO cl_abap_zip
        .

  CREATE OBJECT lo_zip.

  lo_zip->load( lv_bin_content ).

  DATA
        : lv_content TYPE xstring
        .

  lo_zip->get( EXPORTING  name =  'word/document.xml'
                IMPORTING content = lv_content ).

  DATA
        : lv_data_model_xml TYPE string
        .

  CALL TRANSFORMATION zdocx_data_model
  SOURCE XML lv_content
  RESULT XML gv_data_model_xml.


ENDFORM.


FORM show_data_model.


  DATA
        : lo_ixml_document TYPE REF TO if_ixml_document
        .

  DATA: lv_content       TYPE xstring,
        lo_ixml          TYPE REF TO if_ixml,
        lo_streamfactory TYPE REF TO if_ixml_stream_factory,
        lo_istream       TYPE REF TO if_ixml_istream,
        lo_parser        TYPE REF TO if_ixml_parser.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*

  lo_ixml           = cl_ixml=>create( ).
  lo_streamfactory  = lo_ixml->create_stream_factory( ).
  lo_istream        = lo_streamfactory->create_istream_string( gv_data_model_xml ).
  lo_ixml_document            = lo_ixml->create_document( ).
  lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                              istream        = lo_istream
                                              document       = lo_ixml_document ).
*    lo_parser->set_normalizing( 'X' ).
  lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
  lo_parser->parse( ).


  DATA
        : lt_source TYPE t_t_text
        , lr_node TYPE REF TO if_ixml_node
        , lr_nodes TYPE REF TO if_ixml_node_list
        .

  lr_node = lo_ixml_document.
  lr_nodes = lr_node->get_children( ).
  lr_node = lr_nodes->get_item( 0 ).

  PERFORM parse_tree USING lr_node  'data' 'X' CHANGING lt_source.

  FIELD-SYMBOLS
                 : <ls_source> TYPE text80
                 .


  IF p_normal IS NOT INITIAL.
    INSERT INITIAL LINE INTO lt_source ASSIGNING <ls_source> INDEX 1.
    <ls_source> = 'TYPES:'.

    DATA
          : lv_lines TYPE i
          .

    lv_lines = lines( lt_source ) - 2.

    READ TABLE lt_source ASSIGNING <ls_source> INDEX lv_lines.
    REPLACE ALL OCCURRENCES OF ',' IN <ls_source> WITH '.'.


  ELSE.
    APPEND INITIAL LINE TO lt_source ASSIGNING <ls_source>.
    <ls_source> = '.'.
    READ TABLE lt_source ASSIGNING <ls_source> INDEX 1.
    REPLACE ALL OCCURRENCES OF ',' IN <ls_source> WITH ':'.
    INSERT INITIAL LINE INTO lt_source ASSIGNING <ls_source> INDEX 1.
    <ls_source> = 'TYPES'.
  ENDIF.

  DATA
        : lr_text       TYPE REF TO cl_demo_text
        .

  lr_text = cl_demo_text=>get_handle( ).


  LOOP AT lt_source ASSIGNING <ls_source>.
    lr_text->add_line( <ls_source> ).
*    WRITE : / <ls_source>.

  ENDLOOP.

  lr_text->display( ).




ENDFORM.

FORM parse_tree USING p_r_node TYPE REF TO if_ixml_node  p_name p_add_flag TYPE string CHANGING ct_source TYPE t_t_text.

*data
  DATA
         : lv_node_name TYPE string
         , lv_flag_no_add TYPE c
         .


  DATA: lr_nodes       TYPE REF TO if_ixml_node_list,
        lr_child       TYPE REF TO if_ixml_node,
        lv_index       TYPE i,
        lv_child_count TYPE i.


  lv_node_name = p_r_node->get_name( ).

  lr_nodes = p_r_node->get_children( ).
  DATA
        : lt_tmp_source TYPE t_t_text
        .

*добавляємо аочато опису типу
  FIELD-SYMBOLS
                 : <fs_source> TYPE text80
                 .
  APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.

  IF p_normal IS NOT INITIAL.
    <fs_source> = | begin of t_{ p_name },|.
  ELSE.
    <fs_source> = |, begin of t_{ p_name }|.
  ENDIF.

*обробляємо поля

  WHILE lv_index < lr_nodes->get_length( ).

    lr_child = lr_nodes->get_item( lv_index ).

    lv_node_name = lr_child->get_name( ).

    TRANSLATE lv_node_name TO UPPER CASE.

    IF lv_node_name = 'ITEM' .
      lv_flag_no_add = 'X'.

      PERFORM parse_tree USING lr_child  p_name ''  CHANGING ct_source.

    ELSE.

      DATA
            : lr_child_of_child TYPE REF TO if_ixml_node_list
            .

      lr_child_of_child = lr_child->get_children( ).
      lv_child_count = lr_child_of_child->get_length( ).

      APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.
      IF lv_child_count = 0.

        IF p_normal IS NOT INITIAL.
          <fs_source> = |       { lv_node_name } type string,|.
        ELSE.
          <fs_source> = |,       { lv_node_name } type string |.
        ENDIF.
      ELSE.
        IF p_normal IS NOT INITIAL.
          <fs_source> = | { lv_node_name } type t_t_{ lv_node_name }, |.
        ELSE.
          <fs_source> = |, { lv_node_name } type t_t_{ lv_node_name } |.
        ENDIF.

        PERFORM parse_tree USING lr_child  lv_node_name '' CHANGING ct_source.
      ENDIF.

    ENDIF.

    lv_index = lv_index + 1.

  ENDWHILE.

*кінець опису типу
  APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.
  IF p_normal IS NOT INITIAL.
    <fs_source> = | end of t_{ p_name },|.
  ELSE.
    <fs_source> = |, end of t_{ p_name }|.
  ENDIF.

  APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.
  APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.
  IF p_normal IS NOT INITIAL.
    <fs_source> = | t_t_{ p_name } type table of t_{ p_name } with DEFAULT key,|.
  ELSE.
    <fs_source> = |, t_t_{ p_name } type table of t_{ p_name } with DEFAULT key|.
  ENDIF.

  APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.
  APPEND INITIAL LINE TO lt_tmp_source ASSIGNING <fs_source>.


  IF lv_flag_no_add IS INITIAL OR p_add_flag IS NOT INITIAL.
    APPEND LINES OF lt_tmp_source TO ct_source.
  ENDIF.

ENDFORM.
