*&---------------------------------------------------------------------*
*&  Include           ZDOCX_TEST_TEMPLATE_F01
*&---------------------------------------------------------------------*


FORM test_template USING p_data TYPE any
                         p_obj TYPE w3objid
                         p_file TYPE string .

  DATA: lv_bin_content TYPE xstring
        .
  zcl_docx3=>make_some_data( CHANGING cv_data = p_data ).


  IF p_obj IS INITIAL.
    DATA(lt_file_data) = VALUE solix_tab( ).
    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename                = p_file          " Name of file
        filetype                = 'BIN'                 " File Type (ASCII, Binary)
      IMPORTING
        filelength              = DATA(lv_filelength)   " File Length
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





    lv_bin_content = cl_bcs_convert=>solix_to_xstring(
       EXPORTING
         it_solix = lt_file_data               " Input data
         iv_size  = lv_filelength ).

  ENDIF.

  zcl_docx3=>get_document(
    iv_w3objid    = p_obj    " name of our template
       iv_template   = lv_bin_content            " you can feed template as xstring instead of load from smw0
      iv_on_desktop = 'X'           " by default save document on desktop
      iv_folder     = 'test_docx'      " in folder by default 'report'
*      iv_path       = 'C:\Users\o-sikidin-ap\Desktop\zcl_docx\222\'            " IF iv_path IS INITIAL  save on desctop or sap_tmp folder
      iv_file_name  = |template_{ sy-datum DATE = ENVIRONMENT }_{ sy-uzeit }.docx| " file name by default
*      iv_no_execute = 'X'            " if filled -- just get document no run office
*      iv_protect    = ''            " if filled protect document from editing, but not protect from sequence
                                    " ctrl+a, ctrl+c, ctrl+n, ctrl+v, edit
     iv_data       = p_data  " root of our data, obligatory
*      iv_no_save    = ''            " just get binary data not save on disk
     ).
ENDFORM.

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
