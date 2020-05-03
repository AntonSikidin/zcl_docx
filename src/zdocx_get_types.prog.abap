*&---------------------------------------------------------------------*
*& Report ZDOCX_GET_TYPES
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
report zdocx_get_types.

types
: tt_text type table of text80 with empty key
.

data
      : gv_data_model_xml type string
      .

parameters: p_fpath type string obligatory lower case.

parameters: p_normal radiobutton group rad1 default 'X'
          , p_other radiobutton group rad1
          .

at selection-screen on value-request for p_fpath.  " Обробщик події F4
  perform get_file_path changing p_fpath. " Вызов подпрограммы с передачей параметра f_path

start-of-selection.

  perform get_data_model.
  perform show_data_model.


*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
form get_file_path changing cv_path type string.
  clear cv_path.

  data:
    lv_rc          type  i,
    lv_user_action type  i,
    lt_file_table  type  filetable,
    ls_file_table  like line of lt_file_table.

  cl_gui_frontend_services=>file_open_dialog(
  exporting
    window_title        = 'select template  docx'
    multiselection      = ''
    default_extension   = '*.docx'
    file_filter         = 'Text file (*.docx)|*.docx|All (*.*)|*.*'
  changing
    file_table          = lt_file_table
    rc                  = lv_rc
    user_action         = lv_user_action
  exceptions
    others              = 1
    ).
  if sy-subrc = 0.
    if lv_user_action = cl_gui_frontend_services=>action_ok.
      if lt_file_table is not initial.
        read table lt_file_table into ls_file_table index 1.
        if sy-subrc = 0.
          cv_path = ls_file_table-filename.
        endif.
      endif.
    endif.
  endif.
endform.                    " Get_file_path
*&---------------------------------------------------------------------*
*&      Form  GET_DATA_MODEL
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
form get_data_model .


  data(lt_file_data) = value solix_tab( ).
  cl_gui_frontend_services=>gui_upload(
    exporting
      filename                = p_fpath          " Name of file
      filetype                = 'BIN'                 " File Type (ASCII, Binary)
    importing
      filelength              = data(lv_filelength)   " File Length
    changing
      data_tab                = lt_file_data          " Transfer table for file contents
    exceptions
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
      others                  = 19 ).



  data
        : lv_bin_content type xstring
        .

  lv_bin_content = cl_bcs_convert=>solix_to_xstring(
     exporting
       it_solix = lt_file_data               " Input data
       iv_size  = lv_filelength ).


  data
        : lo_zip type ref to cl_abap_zip
        .

  create object lo_zip.

  lo_zip->load( lv_bin_content ).

  data
        : lv_content type xstring
        .

  lo_zip->get( exporting  name =  'word/document.xml'
                importing content = lv_content ).

  data
        : lv_data_model_xml type string
        .

  call transformation zdocx_data_model
  source xml lv_content
  result xml gv_data_model_xml.


endform.


form show_data_model.


  data
        : ro_ixml type ref to if_ixml_document
        .

  data: lv_content       type xstring,
        lo_ixml          type ref to if_ixml,
        lo_streamfactory type ref to if_ixml_stream_factory,
        lo_istream       type ref to if_ixml_istream,
        lo_parser        type ref to if_ixml_parser.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*

  lo_ixml           = cl_ixml=>create( ).
  lo_streamfactory  = lo_ixml->create_stream_factory( ).
  lo_istream        = lo_streamfactory->create_istream_string( gv_data_model_xml ).
  ro_ixml            = lo_ixml->create_document( ).
  lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                              istream        = lo_istream
                                              document       = ro_ixml ).
*    lo_parser->set_normalizing( 'X' ).
  lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
  lo_parser->parse( ).


  data
        : lt_source type tt_text
        , lr_node type ref to if_ixml_node
        , nodes type ref to if_ixml_node_list
        .

  lr_node = ro_ixml.
  nodes = lr_node->get_children( ).
  lr_node = nodes->get_item( 0 ).

  perform parse_tree using lr_node  'data' changing lt_source.



  if p_normal is not initial.
    insert initial line into lt_source assigning field-symbol(<ls_source>) index 1.
    <ls_source> = 'TYPES:'.

    data
          : lv_lines type i
          .

    lv_lines = lines( lt_source ) - 2.

    read table lt_source assigning <ls_source> index lv_lines.
    replace all occurrences of ',' in <ls_source> with '.'.


  else.
    append initial line to lt_source assigning <ls_source>.
    <ls_source> = '.'.
    read table lt_source assigning <ls_source> index 1.
    replace all occurrences of ',' in <ls_source> with ':'.
    insert initial line into lt_source assigning <ls_source> index 1.
    <ls_source> = 'TYPES'.
  endif.





  cl_demo_output=>new( 'TEXT'
    )->display( lt_source ).
endform.

form parse_tree using ir_node type ref to if_ixml_node  name type string changing source type tt_text.


  data
         : lv_node_name type string
         , lv_flag_no_add type c
         .


  data: nodes       type ref to if_ixml_node_list,
        child       type ref to if_ixml_node,
        index       type i,
        child_count type i.


  lv_node_name = ir_node->get_name( ).

  nodes = ir_node->get_children( ).
  data
        : tmp_source type tt_text
        .
  append initial line to tmp_source assigning field-symbol(<fs_source>).

  if p_normal is not initial.
    <fs_source> = | begin of t_{ name },|.
  else.
    <fs_source> = |, begin of t_{ name }|.
  endif.



  while index < nodes->get_length( ).

    child = nodes->get_item( index ).

    lv_node_name = child->get_name( ).

    translate lv_node_name to upper case.

    if lv_node_name = 'ITEM' .
      lv_flag_no_add = 'X'.

      perform parse_tree using child  name changing source.

    else.

      data
            : child_of_child type ref to if_ixml_node_list
            .

      child_of_child = child->get_children( ).
      child_count = child_of_child->get_length( ).

      append initial line to tmp_source assigning <fs_source>.
      if child_count = 0.

        if p_normal is not initial.
          <fs_source> = |       { lv_node_name } type string,|.
        else.
          <fs_source> = |,       { lv_node_name } type string |.
        endif.
      else.
        if p_normal is not initial.
          <fs_source> = | { lv_node_name } type tt_{ lv_node_name }, |.
        else.
          <fs_source> = |, { lv_node_name } type tt_{ lv_node_name } |.
        endif.

        perform parse_tree using child  lv_node_name changing source.
      endif.

    endif.

    index = index + 1.

  endwhile.


  append initial line to tmp_source assigning <fs_source>.
  if p_normal is not initial.
    <fs_source> = | end of t_{ name },|.
  else.
    <fs_source> = |, end of t_{ name }|.
  endif.

  append initial line to tmp_source assigning <fs_source>.
  append initial line to tmp_source assigning <fs_source>.
  if p_normal is not initial.
    <fs_source> = | tt_{ name } type table of t_{ name } with empty key,|.
  else.
    <fs_source> = |, tt_{ name } type table of t_{ name } with empty key|.
  endif.

  append initial line to tmp_source assigning <fs_source>.
  append initial line to tmp_source assigning <fs_source>.


  if lv_flag_no_add is initial.
    append lines of tmp_source to source.
  endif.

endform.
