class ZCL_DOCX3 definition
  public
  final
  create public .

public section.

  class-methods GET_DOCUMENT
    importing
      !IV_W3OBJID type W3OBJID optional
      !IV_TEMPLATE type XSTRING optional
      !IV_ON_DESKTOP type XFELD default 'X'
      !IV_FOLDER type STRING default 'report'
      !IV_PATH type STRING default ''
      !IV_FILE_NAME type STRING default 'report.docx'
      !IV_NO_EXECUTE type XFELD default ''
      !IV_PROTECT type XFELD default ''
      !IV_NO_SAVE type XFELD default ''
      value(IV_DATA) type DATA
    returning
      value(RV_DOCUMENT) type XSTRING .
  class-methods MAKE_SOME_DATA
    importing
      value(IV_LEVEL) type I default 1
    changing
      !CV_DATA type DATA .
protected section.

  constants MC_DOCUMENT type STRING value 'word/document.xml' ##NO_TEXT.

  methods LOAD_SMW0
    importing
      !IV_W3OBJID type W3OBJID .
  methods CONVERT_DATA_TO_NC
    importing
      !IV_DATA type DATA .
  methods PROTECT .
  methods PROTECT_SPACE2
    importing
      !IV_CONTENT type XSTRING
    returning
      value(RV_CONTENT) type XSTRING .
  methods RESTORE_SPACE2
    importing
      !IV_CONTENT type XSTRING
    returning
      value(RV_CONTENT) type XSTRING .
  methods PROTECT_SPACE3
    importing
      !IV_CONTENT type XSTRING
    returning
      value(RV_CONTENT) type XSTRING .
  methods RESTORE_SPACE3
    importing
      !IV_CONTENT type XSTRING
    returning
      value(RV_CONTENT) type XSTRING .
private section.

  data MO_ZIP type ref to CL_ABAP_ZIP .
  data MO_TEMPL_DATA_NC type ref to IF_IXML_NODE_COLLECTION .
  data MT_IMAGES type ZTT_DOCX_IMAGE .

  methods GET_XCODE
    importing
      !IV_VALUE type CHAR1
    returning
      value(RV_VALUE) type CSI_BYTE .
  methods STR_TO_XSTR
    importing
      !IV_STR type STRING
    returning
      value(RV_XSTR) type XSTRING .
  methods CREATE_IMAGES
    changing
      !CV_DATA type DATA .
  methods COLLECT_IMAGES
    changing
      !CV_DATA type DATA .
  methods SIGN_IMAGES
    changing
      !CV_DATA type DATA .
  methods WRITE_IMAGES .
ENDCLASS.



CLASS ZCL_DOCX3 IMPLEMENTATION.


  METHOD collect_images.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--

    DATA
          : l_r_typedescr      TYPE REF TO cl_abap_typedescr
          , l_r_structdescr      TYPE REF TO cl_abap_structdescr
          , lv_len TYPE i
          , ls_docx_image TYPE zst_docx_image
          .

    FIELD-SYMBOLS
                   : <fs_image> TYPE zst_docx_image
                   , <fs_table> TYPE ANY TABLE
                   , <fs_component> TYPE abap_compdescr
                   , <fs_tmp> TYPE any
                   .

    l_r_typedescr ?= cl_abap_typedescr=>describe_by_data( cv_data ).

    IF l_r_typedescr->type_kind = cl_abap_typedescr=>typekind_struct2.
      l_r_structdescr ?= l_r_typedescr.
      IF l_r_structdescr->absolute_name = '\TYPE=ZST_DOCX_IMAGE'.
        ASSIGN cv_data TO <fs_image>.

        lv_len = xstrlen( <fs_image>-image ).

        CALL FUNCTION 'CALCULATE_HASH_FOR_RAW'
          EXPORTING
            alg            = 'MD5'
            data           = <fs_image>-image
            length         = lv_len
          IMPORTING
            hash           = <fs_image>-hash
          EXCEPTIONS
            unknown_alg    = 1
            param_error    = 2
            internal_error = 3
            OTHERS         = 4.

        ls_docx_image-hash = <fs_image>-hash.
        ls_docx_image-image = <fs_image>-image.
        clear <fs_image>-image.

        COLLECT ls_docx_image INTO mt_images.

        IF <fs_image>-cx is not INITIAL or <fs_image>-cy is not INITIAL.
          <fs_image>-cx_emus = <fs_image>-cx * 360000.
          <fs_image>-cy_emus = <fs_image>-cy * 360000.
          <fs_image>-USE_SIZE = 'X'.
        ENDIF.


      ELSE.
        LOOP AT l_r_structdescr->components ASSIGNING <fs_component>.
          ASSIGN COMPONENT <fs_component>-name OF STRUCTURE cv_data TO <fs_tmp>.
          collect_images( CHANGING cv_data = <fs_tmp> ).
        ENDLOOP.
      ENDIF.

    ENDIF.

    IF l_r_typedescr->type_kind = cl_abap_typedescr=>typekind_table.

      ASSIGN cv_data TO <fs_table>.

      LOOP AT <fs_table> ASSIGNING <fs_tmp> .
        collect_images( CHANGING cv_data = <fs_tmp> ).
      ENDLOOP.

    ENDIF.


  ENDMETHOD.


  METHOD convert_data_to_nc.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

    DATA
          :lv_data_xml_str TYPE string
          , lv_regex TYPE string  VALUE '<([a-zA-Z0-9_]*)>'
          , lt_result_tab TYPE match_result_tab
          , lt_types TYPE TABLE OF string
          , lv_type TYPE string
          , lv_tabname TYPE tabname
          , lv_search_str TYPE string
          .

    FIELD-SYMBOLS
                   : <fs_result> TYPE match_result
                   , <fs_submatch> TYPE submatch_result
                   .

    CALL TRANSFORMATION id
    SOURCE data = iv_data
          RESULT XML lv_data_xml_str.


    FIND ALL OCCURRENCES OF REGEX lv_regex IN lv_data_xml_str RESULTS lt_result_tab.

    LOOP AT lt_result_tab ASSIGNING <fs_result>.

      LOOP AT <fs_result>-submatches ASSIGNING <fs_submatch>.
        lv_type = lv_data_xml_str+<fs_submatch>-offset(<fs_submatch>-length).
        COLLECT lv_type INTO lt_types.
      ENDLOOP.
    ENDLOOP.
    LOOP AT lt_types INTO lv_type.

      SELECT SINGLE tabname INTO lv_tabname
        FROM dd02l
        WHERE tabname = lv_type.

      CHECK sy-subrc = 0.

      lv_search_str = |<{ lv_type }>|.
      REPLACE ALL OCCURRENCES OF lv_search_str IN lv_data_xml_str WITH '<item>'.

      lv_search_str = |</{ lv_type }>|.
      REPLACE ALL OCCURRENCES OF lv_search_str IN lv_data_xml_str WITH '</item>'.

    ENDLOOP.



    DATA
          : lo_ixml  TYPE REF TO if_ixml
          .
    lo_ixml = cl_ixml=>create( ).
    DATA
          :lo_stream_factory TYPE REF TO if_ixml_stream_factory
          .

    lo_stream_factory = lo_ixml->create_stream_factory( ).

    DATA
          : lo_istream TYPE REF TO  if_ixml_istream
          .
    lo_istream = lo_stream_factory->create_istream_cstring( lv_data_xml_str ).

    DATA
          : lo_document TYPE REF TO if_ixml_document
          .
    lo_document = lo_ixml->create_document( ).

    DATA
          :lo_parser TYPE REF TO if_ixml_parser
          .
    lo_parser = lo_ixml->create_parser(
    stream_factory = lo_stream_factory
    istream = lo_istream
    document = lo_document ).
    lo_parser->parse( ).
    mo_templ_data_nc = lo_document->get_elements_by_tag_name_ns( name = 'DATA' ).

  ENDMETHOD.


  METHOD create_images.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--

    collect_images( CHANGING cv_data = cv_data ).

    write_images( ).

    sign_images( CHANGING cv_data = cv_data ).



  ENDMETHOD.


  METHOD get_document.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

*variable
    DATA
          : lo_docx TYPE REF TO zcl_docx3
          , lv_content TYPE xstring
          , lv_res_xml_xstr TYPE xstring
          , lv_doc_xml_xstr TYPE xstring
          .

    CREATE OBJECT lo_docx .
    CREATE OBJECT lo_docx->mo_zip.



*get template
    IF iv_template IS INITIAL.
      lo_docx->load_smw0( iv_w3objid ).
    ELSE.
      lo_docx->mo_zip->load( iv_template ).
    ENDIF.

    lo_docx->create_images( CHANGING cv_data = iv_data ).

    lo_docx->convert_data_to_nc( iv_data ).


    lo_docx->mo_zip->get( EXPORTING  name =  lo_docx->mc_document
                    IMPORTING content = lv_content ).
*protect tspaces
    lo_docx->protect_space2(
      EXPORTING
        iv_content = lv_content
      RECEIVING
        rv_content = lv_doc_xml_xstr )   .

*call transformation
    CALL TRANSFORMATION zdocx_del_repeated_text
          SOURCE XML lv_doc_xml_xstr
          RESULT XML lv_doc_xml_xstr.


    CALL TRANSFORMATION zdocx_fill_template
          SOURCE XML lv_doc_xml_xstr
          PARAMETERS data = lo_docx->mo_templ_data_nc
          RESULT XML lv_res_xml_xstr.


    CALL TRANSFORMATION zdocx_del_wsdt
      SOURCE XML lv_res_xml_xstr
      RESULT XML lv_res_xml_xstr.

*restore spaces
    lo_docx->restore_space2(
      EXPORTING
        iv_content = lv_res_xml_xstr
      RECEIVING
        rv_content = lv_content )  .


    IF iv_protect IS NOT INITIAL .
      lo_docx->protect( ).
    ENDIF.

*save template
    lo_docx->mo_zip->delete( name = lo_docx->mc_document ).
    lo_docx->mo_zip->add( name    = lo_docx->mc_document
               content = lv_content ).

    rv_document = lo_docx->mo_zip->save( ).

    IF  iv_no_save IS NOT INITIAL.
      RETURN.
    ENDIF.

    DATA
          : lt_file_tab  TYPE solix_tab
          , lv_bytecount TYPE i
          , lv_path      TYPE string
          .

    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = rv_document ).
    lv_bytecount = xstrlen( rv_document ).


    IF iv_path IS INITIAL.
      IF iv_on_desktop IS NOT INITIAL.
        cl_gui_frontend_services=>get_desktop_directory( CHANGING desktop_directory = lv_path ).
      ELSE.
        cl_gui_frontend_services=>get_temp_directory( CHANGING temp_dir = lv_path ).
      ENDIF.
      cl_gui_cfw=>flush( ).
    ELSE.
      lv_path = iv_path.
    ENDIF.

    CONCATENATE lv_path '\' iv_folder '\'  iv_file_name  INTO lv_path.

    cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                       filename     = lv_path
                                                       filetype     = 'BIN'
                                              CHANGING data_tab     = lt_file_tab
                                                     EXCEPTIONS
                                                      OTHERS = 1
        ).

    IF sy-subrc NE 0.
      RETURN.
    ENDIF.


    IF iv_no_execute IS  INITIAL.
      cl_gui_frontend_services=>execute(  document  =  lv_path ).
    ENDIF.


  ENDMETHOD.


  METHOD get_xcode.

    DATA
      : lv_url_code TYPE url_code
      .
    FIELD-SYMBOLS
                 : <fs_x> TYPE x
                 .

    CALL FUNCTION 'URL_ASCII_CODE_GET'
      EXPORTING
        trans_char = iv_value
      IMPORTING
        char_code  = lv_url_code.

    ASSIGN rv_value TO <fs_x>  .
    <fs_x> = lv_url_code.

  ENDMETHOD.


  method LOAD_SMW0.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*
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
        lv_templ_xstr = cl_bcs_convert=>xtab_to_xstring( lt_mime ).
      CATCH cx_bcs.
        RETURN.
    ENDTRY.

    mo_zip->load( lv_templ_xstr ).

  endmethod.


  METHOD make_some_data.
*variable
    FIELD-SYMBOLS
                   : <fs_any> TYPE any
                   , <fs_any_table> TYPE  TABLE
                   .

    DATA
          : lv_new_level TYPE i
          .

*get components

    DATA lt_comp    TYPE abap_compdescr_tab.
    DATA lo_struct  TYPE REF TO cl_abap_structdescr.
    lo_struct ?= cl_abap_typedescr=>describe_by_data( cv_data ).
    lt_comp = lo_struct->components[].

    LOOP AT lt_comp ASSIGNING FIELD-SYMBOL(<fs_comp>).
      ASSIGN COMPONENT <fs_comp>-name OF STRUCTURE cv_data TO <fs_any>.

      CASE <fs_comp>-type_kind.
        WHEN 'b'.
          <fs_any> = iv_level.
        WHEN 'C'.
          <fs_any> = |{ <fs_comp>-name }_{ iv_level }|.
        WHEN 'D'.
          <fs_any> = sy-datum.
        WHEN 'g'.
          <fs_any> = |{ <fs_comp>-name }_{ iv_level }|.
        WHEN 'h'.
          ASSIGN <fs_any> TO <fs_any_table>.
*add lines
          DO 5 TIMES.
            lv_new_level  = iv_level  * 10 + sy-index.
            APPEND INITIAL LINE TO  <fs_any_table> ASSIGNING FIELD-SYMBOL(<fs_line>).
*            recursive call for table line
            make_some_data( EXPORTING iv_level = lv_new_level
                                     CHANGING cv_data = <fs_line> ).
          ENDDO.

        WHEN 'I'.
          <fs_any> = iv_level.
        WHEN 'N'.
          <fs_any> = iv_level.
        WHEN 'P'.
          <fs_any> = iv_level.
        WHEN 's'.
          <fs_any> = iv_level.
        WHEN 'v'.
*          recursive call for struct
          make_some_data( EXPORTING iv_level = iv_level
                          CHANGING cv_data = <fs_any> ).
      ENDCASE.

    ENDLOOP.



  ENDMETHOD.


  method PROTECT.

*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

    DATA: lv_content       TYPE xstring,
          lo_ixml          TYPE REF TO if_ixml,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_istream       TYPE REF TO if_ixml_istream,
          lo_parser        TYPE REF TO if_ixml_parser,
          lr_element       TYPE REF TO if_ixml_element
          .


    DATA
          : lr_settings TYPE REF TO if_ixml_document
          .

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*

    mo_zip->get( EXPORTING  name =  'word/settings.xml'
                        IMPORTING content = lv_content ).

    lo_ixml           = cl_ixml=>create( ).
    lo_streamfactory  = lo_ixml->create_stream_factory( ).
    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
    lr_settings            = lo_ixml->create_document( ).
    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                istream        = lo_istream
                                                document       = lr_settings ).
*    lo_parser->set_normalizing( 'X' ).
    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
    lo_parser->parse( ).


    DATA(lr_iterator) = lr_settings->create_iterator( ).

    DO.
      DATA(lr_node) = lr_iterator->get_next( ).
      IF lr_node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK lr_node->get_type( ) = if_ixml_node=>co_node_element.

      DATA
            : lv_node_name TYPE string
            .

      lv_node_name = lr_node->get_name( ).

      CHECK  lv_node_name = 'settings'.

      lr_element = lr_settings->create_element_ns( name = 'documentProtection'
                                                prefix = 'w' ).

      lr_element->set_attribute_ns( name = 'edit'
                                 prefix = 'w'
                                 value = 'readOnly' ).

      lr_element->set_attribute_ns( name = 'enforcement'
                                 prefix = 'w'
                                 value = '1' ).

      lr_element->set_attribute_ns( name = 'cryptProviderType'
                           prefix = 'w'
                           value = 'rsaFull' ).

      lr_element->set_attribute_ns( name = 'cryptAlgorithmClass'
                          prefix = 'w'
                          value = 'hash' ).

      lr_element->set_attribute_ns( name = 'cryptAlgorithmType'
                               prefix = 'w'
                               value = 'typeAny' ).

      lr_element->set_attribute_ns( name = 'cryptAlgorithmSid'
                                 prefix = 'w'
                                 value = '4' ).

      lr_element->set_attribute_ns( name = 'cryptSpinCount'
                                prefix = 'w'
                                value = '100000' ).

      lr_element->set_attribute_ns( name = 'hash'
                                prefix = 'w'
                                value = 'zApjmCLBrWyDDRNiRAmxszP+gYc=' ).

      lr_element->set_attribute_ns( name = 'salt'
                                 prefix = 'w'
                                 value = 'erqLP912rJurhcV4a1bb8A==' ).


      lr_node->append_child( lr_element ) .

      EXIT.

    ENDDO.


    DATA
          : lo_ostream        TYPE REF TO if_ixml_ostream
          , lo_renderer       TYPE REF TO if_ixml_renderer

          .

* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_ixml = cl_ixml=>create( ).

    CLEAR lv_content.


    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = lv_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lr_settings ).
    lo_renderer->render( ).

    mo_zip->delete( name = 'word/settings.xml' ).
    mo_zip->add( name    = 'word/settings.xml'
               content = lv_content ).


  endmethod.


  METHOD protect_space2.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*
*    variable
    DATA : lv_pos TYPE i


          , lv_left TYPE x
          , lv_w TYPE x
          , lv_semi TYPE x
          , lv_t TYPE x
          , lv_space TYPE x
          , lv_right TYPE x

          , lv_x TYPE x
          , lv_tmp TYPE x

          , lv_search_open_tag TYPE flag
          , lv_search_first_close_tag TYPE flag
          , lv_check_x5 TYPE flag
          , lv_content_in TYPE flag
          , lv_step_count TYPE i
          , lt_stack TYPE TABLE OF x

          , lv_space_protected TYPE xstring
          , lv_len TYPE i
          .
*get some other values
    lv_left   = get_xcode( '<' ).
    lv_w      = get_xcode( 'w' ).
    lv_semi   = get_xcode( ':' ).
    lv_t      = get_xcode( 't' ).
    lv_space  = get_xcode( ' ' ).
    lv_right  = get_xcode( '>' ).


    APPEND lv_left TO lt_stack.
    APPEND lv_w TO lt_stack.
    APPEND lv_semi TO lt_stack.
    APPEND lv_t TO lt_stack.

    lv_x = get_xcode( '[' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'S' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'P' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'A' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'C' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'E' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( ']' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .

    lv_step_count = 1.
    lv_search_open_tag = 'X'.

*count iteration

    lv_len = xstrlen( iv_content ).

*    запускаем конечный автомат

    DO lv_len TIMES.

      lv_pos = sy-index - 1.

      lv_x = iv_content+lv_pos(1).

      IF  lv_content_in IS NOT INITIAL.
        IF lv_x = lv_left.
          CLEAR lv_content_in.
          lv_search_open_tag = 'X'.
        ENDIF.

        IF lv_x NE lv_space.
          CONCATENATE rv_content lv_x INTO rv_content IN BYTE MODE.
        ELSE.
          CONCATENATE rv_content lv_space_protected INTO rv_content IN BYTE MODE.
        ENDIF.
      ELSE.

        CONCATENATE rv_content lv_x INTO rv_content IN BYTE MODE.

        IF lv_search_open_tag IS NOT INITIAL.
          READ TABLE lt_stack INDEX lv_step_count INTO lv_tmp.
          IF lv_x = lv_tmp.
            lv_step_count = lv_step_count + 1.
          ELSE.
            lv_step_count = 1.

          ENDIF.

          IF lv_step_count = 5.
            lv_step_count = 1.
            CLEAR lv_search_open_tag.
            lv_check_x5 = 'X'.
          ENDIF.

        ELSEIF lv_check_x5 IS NOT INITIAL.
          IF lv_x = lv_space.
            lv_search_first_close_tag = 'X'.
          ELSEIF lv_x = lv_right.
            lv_content_in = 'X'.
          ELSE.
            lv_search_open_tag = 'X'.
          ENDIF.

          CLEAR lv_check_x5.

        ELSEIF lv_search_first_close_tag IS NOT INITIAL.
          IF lv_x = lv_right.
            lv_content_in = 'X'.
            CLEAR lv_search_first_close_tag.
          ENDIF.
        ENDIF.

      ENDIF.

    ENDDO.

  ENDMETHOD.


  METHOD PROTECT_SPACE3.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*


data
      : lt_except TYPE TABLE OF string
      , lv_from(3) TYPE x
      , lv_to(8) TYPE x
      .

lt_except = value #(
                     (  |E29892| )
                     (  |E29890| )
                     (  |E28093| )
                     (  |E28099| )
                     (  |E2809C| )
                     (  |E2809D| )
                   ).


    DATA
          : lv_string TYPE string
          , lt_content_source TYPE TABLE OF string
          , lt_content_dest TYPE TABLE OF string
          , lv_content TYPE xstring

          , lt_result TYPE match_result_tab
          .

    lv_content = iv_content.



    LOOP AT lt_except ASSIGNING FIELD-SYMBOL(<fs_except>).
      lv_from =  <fs_except>.
      lv_to = str_to_xstr( |[{  <fs_except>  }]| ).
      REPLACE ALL OCCURRENCES OF lv_from IN lv_content WITH lv_to IN BYTE MODE .

    ENDLOOP.



    lv_string =  cl_abap_codepage=>convert_from( lv_content ).


*    CALL FUNCTION 'CRM_IC_XML_XSTRING2STRING'
*      EXPORTING
*        inxstring = lv_content
*      IMPORTING
*        outstring = lv_string.


    SPLIT lv_string AT '<' INTO TABLE lt_content_source.

    LOOP AT lt_content_source ASSIGNING FIELD-SYMBOL(<ls_content>).
      IF <ls_content> IS NOT INITIAL.
        <ls_content> = '<' && <ls_content>.
      ENDIF.

    ENDLOOP.

    DATA
          : lv_str1 TYPE string
          , lv_str2 TYPE string
          .
    LOOP AT lt_content_source ASSIGNING FIELD-SYMBOL(<lv_source>).
      FIND REGEX '<w:t(\s|>)' IN   <lv_source>.
      IF sy-subrc NE 0.
        APPEND <lv_source> TO lt_content_dest.
      ELSE.

        SPLIT <lv_source> AT '>' INTO lv_str1 lv_str2.
        lv_str1 = lv_str1 && '>'.
        APPEND lv_str1 TO lt_content_dest.
        REPLACE ALL OCCURRENCES OF REGEX '\s' IN lv_str2 WITH '[SPACE]'.

        REPLACE ALL OCCURRENCES OF  '&#171;' IN lv_str2 WITH '[171]'.
        REPLACE ALL OCCURRENCES OF  '&#187;' IN lv_str2 WITH '[187]'.
        REPLACE ALL OCCURRENCES OF  '&#39;' IN lv_str2 WITH '[39]'.

        APPEND lv_str2 TO lt_content_dest.


      ENDIF.

    ENDLOOP.


    CLEAR lv_string.
    LOOP AT lt_content_dest ASSIGNING FIELD-SYMBOL(<lv_dest>).
      lv_string = lv_string && <lv_dest>.

    ENDLOOP.

    rv_content =  cl_abap_codepage=>convert_to( lv_string ).


*    CALL FUNCTION 'CRM_IC_XML_STRING2XSTRING'
*      EXPORTING
*        instring   = lv_string
*      IMPORTING
*        outxstring = rv_content.
  ENDMETHOD.


  METHOD restore_space2.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*
    DATA
          : lv_space TYPE x
          , lv_x TYPE x
          , lv_space_protected TYPE xstring
          .

    lv_space = get_xcode( ' ' ).
    lv_x = get_xcode( '[' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'S' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'P' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'A' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'C' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( 'E' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .


    lv_x = get_xcode( ']' ).
    lv_space_protected = |{ lv_space_protected }{ lv_x }|  .

    rv_content = iv_content.
    REPLACE ALL OCCURRENCES OF lv_space_protected IN rv_content WITH lv_space IN BYTE MODE .

data
      : lt_except TYPE TABLE OF string
      , lv_from(3) TYPE x
      , lv_to(8) TYPE x
      .

lt_except = value #(
                     (  |E29892| )
                     (  |E29890| )

                   ).


    LOOP AT lt_except ASSIGNING FIELD-SYMBOL(<fs_except>).
      lv_from =  <fs_except>.
      lv_to = str_to_xstr( |[{  <fs_except>  }]| ).
      REPLACE ALL OCCURRENCES OF lv_to IN rv_content WITH lv_from IN BYTE MODE .

    ENDLOOP.

  ENDMETHOD.


  METHOD RESTORE_SPACE3.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

    DATA
       : lv_string TYPE string
       , lt_string TYPE TABLE OF string
       , lt_data TYPE TABLE OF text255
       .

*convert
    lv_string =  cl_abap_codepage=>convert_from( iv_content ).


*replace

    REPLACE ALL OCCURRENCES OF '[171]'  IN lv_string WITH  '&#171;'.
    REPLACE ALL OCCURRENCES OF  '[187]' IN lv_string WITH '&#187;' .
    REPLACE ALL OCCURRENCES OF  '[39]' IN lv_string WITH '&#39;' .

    SPLIT lv_string AT '[SPACE]' INTO TABLE lt_string.

    DATA
          : lv_len TYPE i
          , lv_buf TYPE text255
          , lv_pos TYPE i
          , lv_rem TYPE i
          , lv_254 TYPE i VALUE 254
          , lv_255 TYPE i VALUE 255
          .


    lv_pos = 0.
    lv_rem = lv_255.
    LOOP AT lt_string ASSIGNING FIELD-SYMBOL(<lv_str>).

      lv_len = strlen( <lv_str> ).

      WHILE lv_len > 0.
        IF lv_len > lv_rem.
          lv_buf+lv_pos(lv_rem) = <lv_str>(lv_rem).
          APPEND lv_buf TO lt_data.

          <lv_str> = <lv_str>+lv_rem.
          lv_pos = 0.
          lv_rem = lv_255.
          CLEAR lv_buf.
        ELSE.
          lv_buf+lv_pos = <lv_str>.
          lv_pos = lv_pos + lv_len.
          lv_rem = lv_rem - lv_len.
          CLEAR <lv_str>.

        ENDIF.
        lv_len = strlen( <lv_str> ).



      ENDWHILE.




      IF lv_pos < lv_254.
        lv_pos = lv_pos + 1.
        lv_rem = lv_rem - 1.
      ELSEIF lv_pos  = lv_254.
        APPEND lv_buf TO lt_data.
        lv_pos = 0.
        lv_rem = lv_255.
        CLEAR lv_buf.
      ELSE.
        APPEND lv_buf TO lt_data.
        lv_pos = 1.
        lv_rem = lv_254.
        CLEAR lv_buf.


      ENDIF.


    ENDLOOP.

    APPEND lv_buf TO lt_data.


    CLEAR lv_string.

    LOOP AT lt_data ASSIGNING FIELD-SYMBOL(<lv_data>).

      CONCATENATE lv_string <lv_data> INTO lv_string RESPECTING BLANKS.

    ENDLOOP.


    rv_content =  cl_abap_codepage=>convert_to( lv_string ).

*restore some masked symbols
data
      : lt_except TYPE TABLE OF string
      , lv_from(3) TYPE x
      , lv_to(8) TYPE x
      .

lt_except = value #(
                     (  |E29892| )
                     (  |E29890| )
                     (  |E28093| )
                     (  |E28099| )
                     (  |E2809C| )
                     (  |E2809D| )
                   ).


    LOOP AT lt_except ASSIGNING FIELD-SYMBOL(<fs_except>).
      lv_from =  <fs_except>.
      lv_to = str_to_xstr( |[{  <fs_except>  }]| ).
      REPLACE ALL OCCURRENCES OF lv_to IN rv_content WITH lv_from IN BYTE MODE .

    ENDLOOP.



  ENDMETHOD.


  METHOD sign_images.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

    CHECK mt_images IS NOT INITIAL.

    DATA
          : l_r_typedescr      TYPE REF TO cl_abap_typedescr
          , l_r_structdescr      TYPE REF TO cl_abap_structdescr
          , ls_docx_image TYPE zst_docx_image
          .

    FIELD-SYMBOLS
                   : <fs_image> TYPE zst_docx_image
                   , <fs_table> TYPE ANY TABLE
                   , <fs_component> TYPE abap_compdescr
                   , <fs_tmp> TYPE any
                   .

    l_r_typedescr ?= cl_abap_typedescr=>describe_by_data( cv_data ).

    IF l_r_typedescr->type_kind = cl_abap_typedescr=>typekind_struct2.
      l_r_structdescr ?= l_r_typedescr.
      IF l_r_structdescr->absolute_name = '\TYPE=ZST_DOCX_IMAGE'.
        ASSIGN cv_data TO <fs_image>.

        READ TABLE mt_images INTO ls_docx_image WITH KEY hash = <fs_image>-hash.
        <fs_image>-name = ls_docx_image-name.

      ELSE.
        LOOP AT l_r_structdescr->components ASSIGNING <fs_component>.
          ASSIGN COMPONENT <fs_component>-name OF STRUCTURE cv_data TO <fs_tmp>.
          sign_images( CHANGING cv_data = <fs_tmp> ).
        ENDLOOP.
      ENDIF.

    ENDIF.

    IF l_r_typedescr->type_kind = cl_abap_typedescr=>typekind_table.

      ASSIGN cv_data TO <fs_table>.

      LOOP AT <fs_table> ASSIGNING <fs_tmp> .
        sign_images( CHANGING cv_data = <fs_tmp> ).
      ENDLOOP.

    ENDIF.


  ENDMETHOD.


  METHOD str_to_xstr.

    DATA
          : lv_len TYPE i
          , lv_x TYPE x
          , lv_pos TYPE i
          , lv_c TYPE char1
          .

    lv_len  = strlen( iv_str ).

    DO lv_len  TIMES.
      lv_pos = sy-index - 1.

      lv_c = iv_str+lv_pos(1).

      lv_x = get_xcode( lv_c ).

      CONCATENATE rv_xstr lv_x INTO rv_xstr IN BYTE MODE.

    ENDDO.

  ENDMETHOD.


  METHOD write_images.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*
    CHECK mt_images IS NOT INITIAL.

    DATA: lv_content       TYPE xstring,
          lo_ixml          TYPE REF TO if_ixml,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_istream       TYPE REF TO if_ixml_istream,
          lo_parser        TYPE REF TO if_ixml_parser,
          lr_element       TYPE REF TO if_ixml_element
          .

    DATA
          : lr_document_rels TYPE REF TO if_ixml_document
          , lv_i TYPE i
          , lv_max TYPE i
          , lv_target TYPE string

          , lr_iterator TYPE REF TO IF_IXML_NODE_ITERATOR
          , lr_node TYPE REF TO IF_IXML_NODE

          .
*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*

    mo_zip->get( EXPORTING  name =  'word/_rels/document.xml.rels'
                        IMPORTING content = lv_content ).

    lo_ixml           = cl_ixml=>create( ).
    lo_streamfactory  = lo_ixml->create_stream_factory( ).
    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
    lr_document_rels            = lo_ixml->create_document( ).
    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                istream        = lo_istream
                                                document       = lr_document_rels ).
*    lo_parser->set_normalizing( 'X' ).
    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
    lo_parser->parse( ).


    lr_iterator = lr_document_rels->create_iterator( ).

    DO.
      lr_node = lr_iterator->get_next( ).
      IF lr_node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK lr_node->get_type( ) = if_ixml_node=>co_node_element.

      DATA
            : lv_node_name TYPE string
            , lr_attributes TYPE REF TO IF_IXML_NAMED_NODE_MAP
            , lr_id TYPE REF TO IF_IXML_NODE
            , lv_value TYPE string
            .

      lv_node_name = lr_node->get_name( ).

      CHECK  lv_node_name = 'Relationship'.

      lr_attributes = lr_node->get_attributes( ).
      lr_id = lr_attributes->get_named_item_ns( 'Id' ).
      lv_value = lr_id->get_value( ).
      lv_i = lv_value+3.

      IF lv_i > lv_max.
        lv_max = lv_i.
      ENDIF.


    ENDDO.

    lr_iterator = lr_document_rels->create_iterator( ).

    DO.
      lr_node = lr_iterator->get_next( ).
      IF lr_node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK lr_node->get_type( ) = if_ixml_node=>co_node_element.


      lv_node_name = lr_node->get_name( ).

      CHECK  lv_node_name = 'Relationships'.
      EXIT.

    ENDDO.

    FIELD-SYMBOLS
                   : <fs_image> TYPE ZST_DOCX_IMAGE
                   .


    LOOP AT mt_images ASSIGNING <fs_image>.
      lv_max = lv_max + 1.
      <fs_image>-name = |rId{ lv_max }|.

      lv_target = |media/image{ lv_max }.png|.

      lr_element = lr_document_rels->create_element_ns( name = 'Relationship'  ).
      lv_value = <fs_image>-name.

      lr_element->set_attribute_ns( name = 'Id'  value = lv_value ).
      lr_element->set_attribute_ns( name = 'Type'  value = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' ).
      lr_element->set_attribute_ns( name = 'Target'  value = lv_target ).

      lr_node->append_child( lr_element ) .

      lv_target = |word/media/image{ lv_max }.png|.

      mo_zip->add( name    = lv_target
               content = <fs_image>-image ).

    ENDLOOP.

    DATA
      : lo_ostream        TYPE REF TO if_ixml_ostream
      , lo_renderer       TYPE REF TO if_ixml_renderer

      .

* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_ixml = cl_ixml=>create( ).

    CLEAR lv_content.


    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = lv_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = lr_document_rels ).
    lo_renderer->render( ).

    mo_zip->delete( name = 'word/_rels/document.xml.rels' ).
    mo_zip->add( name    = 'word/_rels/document.xml.rels'
               content = lv_content ).


  ENDMETHOD.
ENDCLASS.
