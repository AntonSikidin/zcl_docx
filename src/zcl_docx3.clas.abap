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
      !IV_DATA type DATA
    returning
      value(RV_DOCUMENT) type XSTRING .
protected section.

  constants MC_DOCUMENT type STRING value 'word/document.xml' ##NO_TEXT.

  methods LOAD_SMW0
    importing
      !IV_W3OBJID type W3OBJID .
  methods PROTECT_SPACE
    importing
      !IV_CONTENT type XSTRING
    returning
      value(RV_CONTENT) type XSTRING .
  methods RESTORE_SPACE
    importing
      !IV_CONTENT type XSTRING
    returning
      value(RV_CONTENT) type XSTRING .
  methods CONVERT_DATA_TO_NC
    importing
      !IV_DATA type DATA .
  methods PROTECT .
private section.

  data MO_ZIP type ref to CL_ABAP_ZIP .
  data MO_TEMPL_DATA_NC type ref to IF_IXML_NODE_COLLECTION .
ENDCLASS.



CLASS ZCL_DOCX3 IMPLEMENTATION.


  METHOD convert_data_to_nc.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

    DATA(lv_data_xml_str) = VALUE string( ).
    CALL TRANSFORMATION id
    SOURCE data = iv_data
          RESULT XML lv_data_xml_str.


    DATA(lo_ixml) = cl_ixml=>create( ).
    DATA(lo_stream_factory) = lo_ixml->create_stream_factory( ).
    DATA(lo_istream) = lo_stream_factory->create_istream_cstring( lv_data_xml_str ).
    DATA(lo_document) = lo_ixml->create_document( ).
    DATA(lo_parser) = lo_ixml->create_parser(
          stream_factory = lo_stream_factory
          istream = lo_istream
          document = lo_document ).
    lo_parser->parse( ).
    mo_templ_data_nc = lo_document->get_elements_by_tag_name_ns( name = 'DATA' ).

  ENDMETHOD.


  method get_document.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*


    data
          : lo_docx type ref to zcl_docx3
          , lv_content type xstring
          , lv_res_xml_xstr type xstring
          , lv_doc_xml_xstr type xstring
          .

    create object lo_docx .
    create object lo_docx->mo_zip.

    lo_docx->convert_data_to_nc( iv_data ).

    if iv_template is initial.
      lo_docx->load_smw0( iv_w3objid ).
    else.
      lo_docx->mo_zip->load( iv_template ).
    endif.


    lo_docx->mo_zip->get( exporting  name =  lo_docx->mc_document
                    importing content = lv_content ).

    lo_docx->protect_space(
      exporting
        iv_content = lv_content
      receiving
        rv_content = lv_doc_xml_xstr )   .


    call transformation zdocx_del_repeated_text
          source xml lv_doc_xml_xstr
          result xml lv_doc_xml_xstr.


    call transformation zdocx_fill_template
          source xml lv_doc_xml_xstr
          parameters data = lo_docx->mo_templ_data_nc
          result xml lv_res_xml_xstr.


    call transformation zdocx_del_wsdt
      source xml lv_res_xml_xstr
      result xml lv_res_xml_xstr.


    lo_docx->restore_space(
      exporting
        iv_content = lv_res_xml_xstr
      receiving
        rv_content = lv_content )  .


    if iv_protect is not initial .
      lo_docx->protect( ).
    endif.


    lo_docx->mo_zip->delete( name = lo_docx->mc_document ).
    lo_docx->mo_zip->add( name    = lo_docx->mc_document
               content = lv_content ).

    rv_document = lo_docx->mo_zip->save( ).

    if  iv_no_save is not initial.
      return.
    endif.

    data
          : lt_file_tab  type solix_tab
          , lv_bytecount type i
          , lv_path      type string
          .

    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = rv_document ).
    lv_bytecount = xstrlen( rv_document ).


    if iv_path is initial.
      if iv_on_desktop is not initial.
        cl_gui_frontend_services=>get_desktop_directory( changing desktop_directory = lv_path ).
      else.
        cl_gui_frontend_services=>get_temp_directory( changing temp_dir = lv_path ).
      endif.
      cl_gui_cfw=>flush( ).
    else.
      lv_path = iv_path.
    endif.

    concatenate lv_path '\' iv_folder '\'  iv_file_name  into lv_path.

    cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                       filename     = lv_path
                                                       filetype     = 'BIN'
                                              changing data_tab     = lt_file_tab
                                                     exceptions
                                                      others = 1
        ).

    if sy-subrc ne 0.
      return.
    endif.


    if iv_no_execute is  initial.
      cl_gui_frontend_services=>execute(  document  =  lv_path ).
    endif.


  endmethod.


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


  method PROTECT_SPACE.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*


    DATA
          : lv_string TYPE string
          , lt_content_source TYPE TABLE OF string
          , lt_content_dest TYPE TABLE OF string
          .

   lv_string =  cl_abap_codepage=>convert_from( iv_content ).


*    CALL FUNCTION 'CRM_IC_XML_XSTRING2STRING'
*      EXPORTING
*        inxstring = iv_content
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
  endmethod.


  method restore_space.
*--------------------------------------------------------------------*
*  Autor: Anton.Sikidin@gmail.com
*--------------------------------------------------------------------*

    data
       : lv_string type string
       , lt_string type table of string
       , lt_data type table of text255
       .


    lv_string =  cl_abap_codepage=>convert_from( iv_content ).


*    CALL FUNCTION 'CRM_IC_XML_XSTRING2STRING'
*      EXPORTING
*        inxstring = iv_content
*      IMPORTING
*        outstring = lv_string.



    replace all occurrences of '[171]'  in lv_string with  '&#171;'.
    replace all occurrences of  '[187]' in lv_string with '&#187;' .
    replace all occurrences of  '[39]' in lv_string with '&#39;' .

    split lv_string at '[SPACE]' into table lt_string.

    data
          : lv_len type i
          , lv_buf type text255
          , lv_pos type i
          , lv_rem type i
          , lv_254 type i value 254
          , lv_255 type i value 255
          .


    lv_pos = 0.
    lv_rem = lv_255.
    loop at lt_string assigning field-symbol(<lv_str>).

      lv_len = strlen( <lv_str> ).

      while lv_len > 0.
        if lv_len > lv_rem.
          lv_buf+lv_pos(lv_rem) = <lv_str>(lv_rem).
          append lv_buf to lt_data.

          <lv_str> = <lv_str>+lv_rem.
          lv_pos = 0.
          lv_rem = lv_255.
          clear lv_buf.
        else.
          lv_buf+lv_pos = <lv_str>.
          lv_pos = lv_pos + lv_len.
          lv_rem = lv_rem - lv_len.
          clear <lv_str>.

        endif.
        lv_len = strlen( <lv_str> ).



      endwhile.




      if lv_pos < lv_254.
        lv_pos = lv_pos + 1.
        lv_rem = lv_rem - 1.
      elseif lv_pos  = lv_254.
        append lv_buf to lt_data.
        lv_pos = 0.
        lv_rem = lv_255.
        clear lv_buf.
      else.
        append lv_buf to lt_data.
        lv_pos = 1.
        lv_rem = lv_254.
        clear lv_buf.


      endif.


    endloop.

    append lv_buf to lt_data.


    clear lv_string.

    loop at lt_data assigning field-symbol(<lv_data>).

      concatenate lv_string <lv_data> into lv_string respecting blanks.

    endloop.


    rv_content =  cl_abap_codepage=>convert_to( lv_string ).


*    CALL FUNCTION 'CRM_IC_XML_STRING2XSTRING'
*      EXPORTING
*        instring   = lv_string
*      IMPORTING
*        outxstring = rv_content.

  endmethod.
ENDCLASS.
