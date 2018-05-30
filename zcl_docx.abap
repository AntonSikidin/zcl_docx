*&---------------------------------------------------------------------*
*&  Include           ZCL_DOCX_CLASS
*&
*&    Author: Sikidin A.P.  anton.sikidin@gmail.com
*&
*&    ZCL_DOCX is replasement of zwww_openform for docx
*&
*&---------------------------------------------------------------------*

CLASS lcl_recursive_data DEFINITION DEFERRED.


TYPES
: BEGIN OF t_key_value
,  key TYPE string
,  value TYPE string
, END OF t_key_value

, tt_key_value TYPE TABLE OF t_key_value

, BEGIN OF t_key_table
,   key TYPE string
,   value TYPE REF TO data
, END OF   t_key_table

, tt_key_table TYPE TABLE OF t_key_table

, t_lcl TYPE REF TO lcl_recursive_data
, tt_lcl TYPE TABLE OF t_lcl

, tt_keys TYPE TABLE OF string
.

CLASS lcl_recursive_data DEFINITION.
  PUBLIC SECTION.


    DATA
          : key_value TYPE tt_key_value
          , key_table TYPE tt_key_table
          , key TYPE string
          , keys TYPE tt_keys
          , key_lcl TYPE tt_lcl
          .
    METHODS append_key_value
      IMPORTING
        VALUE(iv_key) TYPE string
        !iv_value     TYPE any .


    METHODS append_key_table
      IMPORTING
        VALUE(iv_key) TYPE string
        !iv_table     TYPE ANY TABLE.

    METHODS create_document
      IMPORTING
        VALUE(iv_key) TYPE string
      RETURNING
        VALUE(r_data) TYPE REF TO lcl_recursive_data   .



ENDCLASS.

CLASS lcl_recursive_data IMPLEMENTATION.
  METHOD append_key_value.


    TRANSLATE iv_key TO UPPER CASE.

    APPEND INITIAL LINE TO key_value ASSIGNING FIELD-SYMBOL(<fs_key_value>).
    <fs_key_value>-key = |\{{ iv_key }\}| .
    <fs_key_value>-value = iv_value.


  ENDMETHOD.

  METHOD append_key_table.

    TRANSLATE iv_key TO UPPER CASE.

    APPEND INITIAL LINE TO key_table ASSIGNING FIELD-SYMBOL(<fs_key_table>).
    <fs_key_table>-key = iv_key.

    CREATE DATA <fs_key_table>-value LIKE iv_table.
    ASSIGN <fs_key_table>-value->* TO FIELD-SYMBOL(<fs_any_table>).
    <fs_any_table> = iv_table.

  ENDMETHOD.

  METHOD create_document .

    TRANSLATE iv_key TO UPPER CASE.

    CREATE OBJECT r_data.
    r_data->key = iv_key.

    COLLECT iv_key INTO keys.

    APPEND r_data TO key_lcl.


  ENDMETHOD.

ENDCLASS.

CLASS lcl_docx DEFINITION.

  PUBLIC SECTION.

    TYPES
    : t_ref_node TYPE REF TO if_ixml_node
    , tt_node TYPE STANDARD TABLE OF t_ref_node

    , BEGIN OF t_bookmark
    , id TYPE string
    , start_node TYPE   t_ref_node
    , end_node TYPE   t_ref_node
    , in_table TYPE xfeld
    , name TYPE string
    , start_height   TYPE i
    , end_height TYPE i
    , END OF t_bookmark

    , BEGIN OF t_stack_data
    , start TYPE xfeld
    , id TYPE string
    , node TYPE t_ref_node
    , position TYPE i
    , length TYPE i
    , collision TYPE i
    , END OF t_stack_data

    , BEGIN OF t_collision
    , collision TYPE string
    , count TYPE i
    , END OF  t_collision

    .


    METHODS load_smw0
      IMPORTING
        !i_w3objid TYPE w3objid .


    METHODS add_file
      IMPORTING
        !iv_path TYPE string
        !iv_data TYPE xstring.
    METHODS save
      IMPORTING
        !on_desktop   TYPE xfeld DEFAULT 'X'
        !iv_folder    TYPE string DEFAULT 'report'
        !iv_path      TYPE string DEFAULT ''
        !iv_file_name TYPE string DEFAULT 'report.docx'
        !no_execute   TYPE xfeld DEFAULT '' .





    METHODS map_data
      IMPORTING
        !ir_xml_node TYPE REF TO if_ixml_document OPTIONAL
        !ir_data     TYPE REF TO lcl_recursive_data.

    METHODS check_flag
      IMPORTING
        !it_keys TYPE tt_keys.



  PROTECTED SECTION.

    CONSTANTS c_document TYPE string VALUE 'word/document.xml' ##NO_TEXT.

    METHODS map_values
      IMPORTING
        !ir_xml_node  TYPE REF TO if_ixml_node OPTIONAL
        !it_key_value TYPE tt_key_value .


    METHODS map_table
      IMPORTING
        !ir_xml_node TYPE REF TO if_ixml_document OPTIONAL
        !iv_key      TYPE string
        !it_data     TYPE ANY TABLE .

    METHODS append_node
      IMPORTING
        !ir_source TYPE REF TO if_ixml_node OPTIONAL
        !ir_dest   TYPE REF TO if_ixml_node OPTIONAL
        !iv_key    TYPE string.

    METHODS get_fragment
      IMPORTING
        !ir_xml_node       TYPE REF TO if_ixml_document OPTIONAL
        !iv_key            TYPE string
      RETURNING
        VALUE(rr_fragment) TYPE REF TO if_ixml_document.

    METHODS normalize_key.
    METHODS align_bookmark.

    METHODS get_from_zip_archive
      IMPORTING
        !i_filename      TYPE string
      RETURNING
        VALUE(r_content) TYPE xstring .
    METHODS get_ixml_from_zip_archive
      IMPORTING
        !i_filename   TYPE string
      RETURNING
        VALUE(r_ixml) TYPE REF TO if_ixml_document .

  PRIVATE SECTION.

    DATA zip TYPE REF TO cl_abap_zip .
    DATA document TYPE REF TO if_ixml_document .

    METHODS upper_case
      IMPORTING
        !iv_str       TYPE string
      RETURNING
        VALUE(rv_str) TYPE string.

    METHODS dump_data
      IMPORTING
        !node      TYPE REF TO if_ixml_document OPTIONAL
        !node_node TYPE REF TO if_ixml_node OPTIONAL
        !fname     TYPE string.

    METHODS create_document
      RETURNING
        VALUE(rp_content) TYPE xstring .

    METHODS map_line
      IMPORTING
        !node       TYPE REF TO if_ixml_document
        !components TYPE abap_compdescr_tab
        !data       TYPE any .
ENDCLASS.

CLASS lcl_docx IMPLEMENTATION.
  METHOD load_smw0.
    DATA
          : lv_templ_xstr TYPE xstring
          , lt_mime TYPE TABLE OF w3mime
          .

    DATA(ls_key) = VALUE wwwdatatab(
    relid = 'MI'
    objid = i_w3objid ).

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

    IF zip IS INITIAL.
      CREATE OBJECT zip.
    ENDIF.

    zip->load( lv_templ_xstr ).

    document = me->get_ixml_from_zip_archive( me->c_document ).
    normalize_key( ).
    align_bookmark( ).
  ENDMETHOD.

  METHOD normalize_key.
    DATA
          : lt_nodes TYPE TABLE OF t_ref_node
          , lv_in TYPE c
          , lv_regex_open TYPE string VALUE '\{[^\}]*$'
          , lv_regex_close TYPE string VALUE '^\}[^\{]*'
          , lv_tmp_str TYPE string
          .

    DATA(iterator) = document->create_iterator( ).

    DO.
      DATA(node) = iterator->get_next( ).
      IF node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK node->get_type( ) = if_ixml_node=>co_node_element.

      CHECK  node->get_name( ) = 'p'.

      REFRESH lt_nodes.
      CLEAR lv_in.

      DATA(nodes) = node->get_children( ).
      DO nodes->get_length( ) TIMES.

        DATA(child) = nodes->get_item( sy-index - 1 ).


        DATA(nodes_2) = child->get_children( ).

        DO  nodes_2->get_length( ) TIMES.
          DATA(child_2) = nodes_2->get_item( sy-index - 1 ).

          CHECK child_2->get_name( ) = 't'.

          DATA(child_2_value) = child_2->get_value( ).

          IF lv_in IS NOT INITIAL.

            APPEND child_2 TO lt_nodes.

            FIND REGEX lv_regex_close  IN child_2_value.

            CHECK sy-subrc = 0.

            FIND REGEX lv_regex_open  IN child_2_value.

            CHECK sy-subrc NE 0.

            CLEAR lv_tmp_str .

            LOOP AT lt_nodes ASSIGNING FIELD-SYMBOL(<fs_node>).
              child_2_value = <fs_node>->get_value( ).
              lv_tmp_str = |{ lv_tmp_str }{ child_2_value }|.
              <fs_node>->set_value( '' ).

            ENDLOOP.

            READ TABLE lt_nodes ASSIGNING <fs_node> INDEX 1.
            <fs_node>->set_value( upper_case( lv_tmp_str ) ).

            CLEAR  lv_in.

          ELSE.

            FIND REGEX lv_regex_open  IN child_2_value.

            CHECK sy-subrc = 0.

            lv_in = 'X'.

            APPEND child_2 TO lt_nodes.

          ENDIF.

        ENDDO.

      ENDDO.

    ENDDO.

  ENDMETHOD.
  METHOD align_bookmark.
    DATA
           : lt_bokkmarks TYPE TABLE OF t_bookmark
           .

    DATA(iterator) = document->create_iterator( ).

    DO  .
      DATA(node) = iterator->get_next( ).

      IF node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK node->get_type( ) = if_ixml_node=>co_node_element.

      DATA(name) = node->get_name( ).

      CASE name.
        WHEN 'bookmarkStart'.
        WHEN 'bookmarkEnd'.
        WHEN OTHERS.
          CONTINUE.
      ENDCASE.

      DATA(attributes) = node->get_attributes( ).

      DATA
            : lv_name TYPE string
            , lv_id TYPE string
            .


      DO attributes->get_length( ) TIMES.

        DATA(attribute) = attributes->get_item( sy-index - 1 ).

        CASE attribute->get_name( ).
          WHEN 'id'.
            lv_id = attribute->get_value( ).
            READ TABLE lt_bokkmarks ASSIGNING FIELD-SYMBOL(<fs_bookmark>) WITH KEY id = lv_id.
            IF sy-subrc NE 0.
              APPEND INITIAL LINE TO lt_bokkmarks ASSIGNING <fs_bookmark>.
              <fs_bookmark>-id = lv_id.
            ENDIF.

            CASE name.
              WHEN 'bookmarkStart'.
                <fs_bookmark>-start_node = node.
                <fs_bookmark>-start_height = node->get_height( ).
              WHEN 'bookmarkEnd'.
                <fs_bookmark>-end_node = node.
                <fs_bookmark>-end_height = node->get_height( ).
            ENDCASE.


          WHEN 'name'.
            lv_name = attribute->get_value( ) .
            TRANSLATE lv_name TO UPPER CASE.
            attribute->set_value( lv_name ) .
            <fs_bookmark>-name = lv_name.


        ENDCASE.

      ENDDO.

    ENDDO.


*    table or not

    DATA
          : lv_node TYPE t_ref_node
          , lv_ref_node TYPE t_ref_node
          , lv_ref_node_2 TYPE t_ref_node

          .

    LOOP AT lt_bokkmarks ASSIGNING <fs_bookmark>.


      lv_node = <fs_bookmark>-start_node.
      DO .

        IF lv_node->is_root( ) IS NOT   INITIAL.
          EXIT.
        ENDIF.

        DATA(node_name) = lv_node->get_name( ).

        IF node_name = 'tr'.
          <fs_bookmark>-in_table = 'X'.
          EXIT.
        ENDIF.

        lv_node = lv_node->get_parent( ).


      ENDDO.

    ENDLOOP.


    SORT lt_bokkmarks BY id ASCENDING .

    LOOP AT lt_bokkmarks ASSIGNING <fs_bookmark>.

      IF <fs_bookmark>-in_table = 'X'.

        lv_ref_node = <fs_bookmark>-start_node.

        DO .
          node_name = lv_ref_node->get_name( ).

          IF node_name = 'tr'.
            lv_node = lv_ref_node->get_parent( ).
            EXIT.
          ENDIF.
          lv_ref_node = lv_ref_node->get_parent( ).
        ENDDO.

        <fs_bookmark>-start_node->remove_node( ).
        lv_node->insert_child( new_child = <fs_bookmark>-start_node
                                  ref_child = lv_ref_node ).

        CLEAR
        : lv_ref_node_2
        .

        lv_ref_node_2 = <fs_bookmark>-end_node.

        DO .
          IF lv_ref_node_2 IS INITIAL.
            EXIT.
          ENDIF.
          node_name = lv_ref_node_2->get_name( ).

          IF node_name = 'tr'.
            EXIT.
          ENDIF.
          lv_ref_node_2 = lv_ref_node_2->get_parent( ).
        ENDDO.

        <fs_bookmark>-end_node->remove_node( ).

        IF lv_ref_node_2 IS NOT INITIAL.
          lv_node->insert_child( new_child = <fs_bookmark>-end_node
                          ref_child = lv_ref_node_2 ).
        ELSE.
          lv_node->append_child( new_child = <fs_bookmark>-end_node ).
        ENDIF.

      ELSE.

        lv_ref_node = <fs_bookmark>-start_node->get_parent( ).
        <fs_bookmark>-start_node->remove_node( ).
        lv_node = lv_ref_node->get_parent( ).

        lv_node->insert_child( new_child = <fs_bookmark>-start_node
                              ref_child = lv_ref_node ).


        DATA
              : lv_height_start TYPE i
              , lv_height_end TYPE i
              .

        lv_height_start = <fs_bookmark>-start_node->get_height( ).

        lv_ref_node = <fs_bookmark>-end_node.

        DO .

          IF lv_ref_node IS INITIAL.
            EXIT.
          ENDIF.

          lv_height_end = lv_ref_node->get_height( ).

          IF lv_height_end = lv_height_start.
            EXIT.
          ENDIF.


          lv_ref_node = lv_ref_node->get_parent( ).


        ENDDO.

        IF lv_ref_node NE <fs_bookmark>-end_node.
          <fs_bookmark>-end_node->remove_node( ).

          lv_node = lv_ref_node->get_parent( ).

          lv_node->insert_child( new_child = <fs_bookmark>-end_node
                                ref_child = lv_ref_node ).
        ENDIF.

      ENDIF.

    ENDLOOP.



    DATA
          : lv_position TYPE i
          , lt_stack_data TYPE TABLE OF t_stack_data
          , lt_stack_data_sorted TYPE TABLE OF t_stack_data
          , lt_id TYPE TABLE OF string

          , lt_collision TYPE TABLE OF t_collision
          , ls_collision TYPE t_collision

          , lt_old TYPE TABLE OF t_ref_node
          , lt_sorted TYPE TABLE OF t_ref_node
          .

    iterator = document->create_iterator( ).


    DO  .
      node = iterator->get_next( ).
      IF node IS INITIAL.
        EXIT.
      ENDIF.
      CHECK node->get_type( ) = if_ixml_node=>co_node_element.

      name = node->get_name( ).

      CASE name.
        WHEN 'bookmarkStart'.
        WHEN 'bookmarkEnd'.
        WHEN OTHERS.
          ADD 1 TO lv_position.

          CONTINUE.
      ENDCASE.

      attributes = node->get_attributes( ).


      APPEND INITIAL LINE TO lt_stack_data ASSIGNING FIELD-SYMBOL(<fs_stack_data>).

      IF name = 'bookmarkStart'.
        <fs_stack_data>-start = 'X'.
      ENDIF.

      <fs_stack_data>-position = lv_position.
      <fs_stack_data>-collision = lv_position.
      <fs_stack_data>-node = node.

      ls_collision-collision = lv_position.
      ls_collision-count  = 1.

      COLLECT ls_collision INTO lt_collision.

      DO attributes->get_length( ) TIMES.
        attribute = attributes->get_item( sy-index - 1 ).
        CHECK attribute->get_name( ) = 'id'.

        <fs_stack_data>-id = attribute->get_value( ).
        COLLECT <fs_stack_data>-id INTO lt_id.

      ENDDO.

    ENDDO.

    lt_stack_data_sorted = lt_stack_data.

    DATA
          : lv_length TYPE i
          .

    LOOP AT lt_id ASSIGNING FIELD-SYMBOL(<fs_id>).

      CLEAR lv_length.

      LOOP AT lt_stack_data_sorted ASSIGNING <fs_stack_data> WHERE id = <fs_id>.

        IF lv_length IS INITIAL.
          lv_length = <fs_stack_data>-position.
        ELSE.
          lv_length = <fs_stack_data>-position - lv_length .
        ENDIF.

      ENDLOOP.

      LOOP AT lt_stack_data_sorted ASSIGNING <fs_stack_data> WHERE id = <fs_id>.
        CASE 'X'.
          WHEN <fs_stack_data>-start.
            <fs_stack_data>-length =  lv_length.
          WHEN OTHERS.
            <fs_stack_data>-length =  lv_length * -1 .
        ENDCASE.
      ENDLOOP.

    ENDLOOP.

    SORT lt_stack_data_sorted BY collision  start length DESCENDING.


    LOOP AT lt_collision ASSIGNING FIELD-SYMBOL(<fs_collision>) WHERE count > 1.

      REFRESH
      : lt_old
      , lt_sorted
      .

      LOOP AT lt_stack_data ASSIGNING <fs_stack_data> WHERE collision = <fs_collision>-collision.
        APPEND <fs_stack_data>-node TO lt_old.
      ENDLOOP.

      LOOP AT lt_stack_data_sorted ASSIGNING <fs_stack_data> WHERE collision = <fs_collision>-collision.
        APPEND <fs_stack_data>-node TO lt_sorted.
      ENDLOOP.

      CHECK lt_old NE lt_sorted.

      READ TABLE lt_old ASSIGNING FIELD-SYMBOL(<fs_old>) INDEX <fs_collision>-count.

      CLEAR lv_ref_node.

      lv_node = <fs_old>->get_parent( ).
      lv_ref_node = <fs_old>->get_next( ).

      LOOP AT lt_old ASSIGNING <fs_old>.
        <fs_old>->remove_node( ).
      ENDLOOP.


      LOOP AT lt_sorted ASSIGNING FIELD-SYMBOL(<fs_sorted>).

        IF lv_ref_node IS NOT INITIAL.
          lv_node->insert_child( new_child = <fs_sorted>
                            ref_child = lv_ref_node ).
        ELSE.

          lv_node->append_child( new_child = <fs_sorted> ).
        ENDIF.

      ENDLOOP.

    ENDLOOP.

  ENDMETHOD.

  METHOD add_file.

    zip->delete( name = iv_path ).

    zip->add( name    = iv_path
               content = iv_data ).

  ENDMETHOD.
  METHOD save.
    DATA
             : lv_content         TYPE xstring
             .
    lv_content = me->create_document( ).

    zip->delete( name = me->c_document ).
    zip->add( name    = me->c_document
               content = lv_content ).

    lv_content = zip->save( ).

    DATA
          : lt_file_tab  TYPE solix_tab
          , lv_bytecount TYPE i
          , lv_path      TYPE string
          .

    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_content ).
    lv_bytecount = xstrlen( lv_content ).


    IF iv_path IS INITIAL.
      IF on_desktop IS NOT INITIAL.
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
    file_write_error          = 1
    no_batch                  = 2
    gui_refuse_filetransfer   = 3
    invalid_type              = 4
    no_authority              = 5
    unknown_error             = 6
    header_not_allowed        = 7
    separator_not_allowed     = 8
    filesize_not_allowed      = 9
    header_too_long           = 10
    dp_error_create           = 11
    dp_error_send             = 12
    dp_error_write            = 13
    unknown_dp_error          = 14
    access_denied             = 15
    dp_out_of_memory          = 16
    disk_full                 = 17
    dp_timeout                = 18
    file_not_found            = 19
    dataprovider_exception    = 20
    control_flush_error       = 21
    not_supported_by_gui      = 22
    error_no_gui              = 23
    OTHERS                    = 24
        ).


    CHECK no_execute IS  INITIAL.
    cl_gui_frontend_services=>execute(  document  =  lv_path ).

  ENDMETHOD.
  METHOD map_values.

    CHECK it_key_value IS NOT INITIAL.

    IF ir_xml_node IS BOUND.
      DATA(iterator) = ir_xml_node->create_iterator( ).
    ELSE.
      iterator = document->create_iterator( ).
    ENDIF.


    DO.
      DATA(node) = iterator->get_next( ).
      IF node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK node->get_type( ) = if_ixml_node=>co_node_element.

      CHECK node->get_name( ) = 't'.

      DATA(value) = node->get_value( ).
      LOOP AT it_key_value ASSIGNING FIELD-SYMBOL(<fs_key_value>).
        REPLACE ALL OCCURRENCES OF <fs_key_value>-key IN value WITH <fs_key_value>-value.
        CHECK sy-subrc = 0.
        node->set_value( value  ).

      ENDLOOP.

    ENDDO.
  ENDMETHOD.


  METHOD get_fragment.
    DATA
                      : lv_found TYPE c
                      , lv_id  TYPE string
                      , lr_start TYPE REF TO if_ixml_node


                      , lt_node TYPE tt_node
                      , lv_first_run TYPE c

                      , lv_start TYPE c
                      .



    DATA(iterator) = ir_xml_node->create_iterator( ).


    DATA(ixmlfactory) = cl_ixml=>create( ).
    rr_fragment = ixmlfactory->create_document( ).


    DO.
      DATA(node) = iterator->get_next( ).

      IF node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK node->get_type( ) = if_ixml_node=>co_node_element.

      CHECK node->get_name( ) = 'bookmarkStart'.

      DATA(attributes) = node->get_attributes( ).


      DO attributes->get_length( ) TIMES.
        DATA(attribute) = attributes->get_item( sy-index - 1 ).
        CASE attribute->get_name( ).
          WHEN 'id'.
            lv_id = attribute->get_value( ).
          WHEN 'name'.
            IF attribute->get_value( ) = iv_key .
              lv_found = 'X'.
              lv_first_run = 'X'.
              lr_start = node.
            ENDIF.

        ENDCASE.

      ENDDO.

      CHECK  lv_first_run IS NOT INITIAL.
      CLEAR lv_first_run .
      EXIT.

    ENDDO.

    CHECK lr_start IS NOT INITIAL.
    DATA(parent) = lr_start->get_parent( ).
    DATA(children) = parent->get_children( ).

    DO  children->get_length( ) TIMES.
      DATA(child) = children->get_item( sy-index - 1 ).
      DATA(name) = child->get_name( ).

      IF name = 'bookmarkEnd' AND lv_found IS NOT INITIAL.
        attributes = child->get_attributes( ).
        DO attributes->get_length( ) TIMES.

          attribute = attributes->get_item( sy-index - 1 ).

          CHECK attribute->get_name( ) = 'id'.
          CHECK lv_id = attribute->get_value( ).

          CLEAR lv_start.

        ENDDO.

      ENDIF.

      IF lv_start IS NOT INITIAL.
        APPEND child TO lt_node.
      ENDIF.


      IF name = 'bookmarkStart' AND lv_found IS NOT INITIAL.
        attributes = child->get_attributes( ).
        DO attributes->get_length( ) TIMES.
          attribute = attributes->get_item( sy-index - 1 ).
          CHECK attribute->get_name( ) = 'id'.
          CHECK lv_id = attribute->get_value( ).

          lv_start = 'X'.

        ENDDO.

      ENDIF.

    ENDDO.

    LOOP AT lt_node ASSIGNING FIELD-SYMBOL(<fs_node>).
      DATA(clone_node) = <fs_node>->clone( ) .

      name = clone_node->get_name( ).
      rr_fragment->append_child( clone_node ).
      <fs_node>->remove_node( ).

    ENDLOOP.

  ENDMETHOD.

  METHOD  append_node .

    DATA
          : lv_node_name TYPE string
          .

    DATA(iterator) = ir_dest->create_iterator( ).

    DO.
      DATA(node) = iterator->get_next( ).

      IF node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK  node->get_type( ) = if_ixml_node=>co_node_element.

      CHECK node->get_name( ) = 'bookmarkStart'.

      DATA(attributes) = node->get_attributes( ).


      DO attributes->get_length( ) TIMES.
        DATA(attribute) = attributes->get_item( sy-index - 1 ).

        CHECK attribute->get_name( ) = 'name'.

        CHECK attribute->get_value( ) = iv_key .

        DATA(parent) = node->get_parent( ).


        DATA(children) = ir_source->get_children( ).
        DATA(ch_count) = children->get_length( ).
        DO  ch_count TIMES.
          DATA(child) = children->get_item( sy-index - 1 ).

          DATA(clone_child) = child->clone( ).

          parent->insert_child( new_child = clone_child
                                ref_child = node ).

        ENDDO.

        RETURN.

      ENDDO.

    ENDDO.

  ENDMETHOD.

  METHOD  map_data .

    DATA
          : lr_node TYPE REF TO if_ixml_document

          , lr_document TYPE REF TO if_ixml_document
          , lr_clone TYPE REF TO if_ixml_document
          .

    FIELD-SYMBOLS
                   : <fs_any_table> TYPE ANY TABLE
                   .

    IF ir_xml_node IS BOUND.
      lr_node = ir_xml_node.
    ELSE.
      lr_node = document.
    ENDIF.

    map_values(       ir_xml_node  = lr_node
                      it_key_value = ir_data->key_value  ) .



    LOOP AT ir_data->key_table ASSIGNING FIELD-SYMBOL(<fs_key_table>).

      ASSIGN <fs_key_table>-value->* TO <fs_any_table>.

      map_table(  ir_xml_node  = lr_node
                  iv_key       = <fs_key_table>-key
                  it_data      = <fs_any_table> ).

    ENDLOOP.



    LOOP AT ir_data->keys ASSIGNING FIELD-SYMBOL(<fs_key>).

      lr_document = get_fragment( ir_xml_node = lr_node iv_key = <fs_key> ).

      LOOP AT ir_data->key_lcl ASSIGNING FIELD-SYMBOL(<fs_key_lcl>).

        CHECK <fs_key_lcl>->key = <fs_key>.

        CHECK <fs_key_lcl>->key_value IS NOT INITIAL OR
              <fs_key_lcl>->key_table IS NOT INITIAL OR
              <fs_key_lcl>->keys      IS NOT INITIAL.

        lr_clone ?= lr_document->clone( ).

        map_data( ir_xml_node  = lr_clone
                  ir_data      = <fs_key_lcl> ).

        append_node(  ir_source = lr_clone
                      ir_dest   = lr_node
                      iv_key    = <fs_key> ).


      ENDLOOP.

    ENDLOOP.

  ENDMETHOD.

  METHOD check_flag.
    DATA(iterator) = document->create_iterator( ).

    DO.
      DATA(node) = iterator->get_next( ).
      IF node IS INITIAL.
        EXIT.
      ENDIF.

      IF node->get_type( ) <> if_ixml_node=>co_node_element.
        CONTINUE.
      ENDIF.
      DATA(name) = node->get_name( ).
      IF name = 'fldChar'.

        DATA(checkbox_iterator) = node->create_iterator( ).

        DATA
              : lv_found TYPE c
              .

        CLEAR lv_found.

        DO .
          DATA(check_box_node) = checkbox_iterator->get_next( ).
          IF check_box_node IS INITIAL.
            EXIT.
          ENDIF.

          IF check_box_node->get_type( ) <> if_ixml_node=>co_node_element.
            CONTINUE.
          ENDIF.

          DATA(check_box_name) = check_box_node->get_name( ).

          IF check_box_name = 'name'.
            DATA(attributes) = check_box_node->get_attributes( ).

            DO attributes->get_length( ) TIMES.
              DATA(attribute) = attributes->get_item( sy-index - 1 ).
              IF attribute->get_name( ) = 'val'.
                DATA(atribute_value) =  attribute->get_value( ) .

                READ TABLE it_keys TRANSPORTING NO FIELDS WITH KEY table_line = atribute_value.

                IF sy-subrc = 0.
                  lv_found = 'X'.
                  EXIT.
                ENDIF.

              ENDIF.

            ENDDO.



          ENDIF.

          IF lv_found IS NOT INITIAL AND check_box_name = 'default'.
            attributes = check_box_node->get_attributes( ).

            DO attributes->get_length( ) TIMES.
              attribute = attributes->get_item( sy-index - 1 ).
              IF attribute->get_name( ) = 'val'.
                attribute->set_value( '1' ) .
                EXIT.
              ENDIF.

            ENDDO.

          ENDIF.

        ENDDO.


      ENDIF.


    ENDDO.

  ENDMETHOD.

  METHOD map_table.

    DATA
      : lr_node TYPE REF TO if_ixml_document
      , lr_document TYPE REF TO if_ixml_document
      , lr_clone TYPE REF TO if_ixml_document

      , l_r_structdescr      TYPE REF TO cl_abap_structdescr
      .


    IF ir_xml_node IS BOUND.
      lr_node = ir_xml_node.
    ELSE.
      lr_node = document.
    ENDIF.



    lr_document = get_fragment( ir_xml_node = lr_node
                                iv_key      = iv_key ).

    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<fs_data>).
      lr_clone  ?= lr_document->clone( ).

      IF l_r_structdescr IS NOT BOUND.
        l_r_structdescr ?= cl_abap_structdescr=>describe_by_data( <fs_data> ).
      ENDIF.

      map_line(    node       = lr_clone
                   components = l_r_structdescr->components
                   data       = <fs_data>              ).

      append_node(  ir_source = lr_clone
                    ir_dest   = lr_node
                    iv_key    = iv_key ).
    ENDLOOP.

  ENDMETHOD.


  METHOD get_from_zip_archive.
    ASSERT zip IS BOUND. " zip object has to exist at this point

    zip->get( EXPORTING  name =  i_filename
                        IMPORTING content = r_content ).

  ENDMETHOD.
  METHOD get_ixml_from_zip_archive.
    DATA: lv_content       TYPE xstring,
          lo_ixml          TYPE REF TO if_ixml,
          lo_streamfactory TYPE REF TO if_ixml_stream_factory,
          lo_istream       TYPE REF TO if_ixml_istream,
          lo_parser        TYPE REF TO if_ixml_parser.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*
    lv_content        = me->get_from_zip_archive( i_filename ).
    lo_ixml           = cl_ixml=>create( ).
    lo_streamfactory  = lo_ixml->create_stream_factory( ).
    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
    r_ixml            = lo_ixml->create_document( ).
    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                istream        = lo_istream
                                                document       = r_ixml ).
    lo_parser->set_normalizing( 'X' ).
    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
    lo_parser->parse( ).
  ENDMETHOD.


  METHOD upper_case.
    DATA
          : result_tab TYPE match_result_tab
          , lv_search TYPE string
          , lv_replace TYPE string
          .
    rv_str = iv_str.

    FIND ALL OCCURRENCES OF REGEX '\{[^\}]*\}' IN iv_str RESULTS result_tab.
    LOOP AT result_tab ASSIGNING FIELD-SYMBOL(<fs_result>).
      lv_search = iv_str+<fs_result>-offset(<fs_result>-length).
      lv_replace = lv_search.

      TRANSLATE lv_replace TO UPPER CASE.

      REPLACE ALL OCCURRENCES OF lv_search IN rv_str WITH lv_replace.

    ENDLOOP.

  ENDMETHOD.

  METHOD dump_data .

    DATA
      : lo_ixml           TYPE REF TO if_ixml
      , lo_streamfactory  TYPE REF TO if_ixml_stream_factory
      , lo_ostream        TYPE REF TO if_ixml_ostream
      , lo_renderer       TYPE REF TO if_ixml_renderer
      , lv_xstring TYPE xstring
      .

* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_ixml = cl_ixml=>create( ).


    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = lv_xstring ).
    IF node_node IS SUPPLIED.
      DATA(document) = lo_ixml->create_document( ).
      document->append_child( node_node ).
      lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = document ).
    ELSE.
      lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = node ).
    ENDIF.

    lo_renderer->render( ).


    DATA
          : lt_file_tab  TYPE solix_tab
          , lv_bytecount TYPE i
          , lv_path      TYPE string
          .

    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_xstring ).
    lv_bytecount = xstrlen( lv_xstring ).

    cl_gui_frontend_services=>get_desktop_directory( CHANGING desktop_directory = lv_path ).

    cl_gui_cfw=>flush( ).

    CONCATENATE lv_path '\report\'  fname '.txt'  INTO lv_path.

    cl_gui_frontend_services=>gui_download( EXPORTING bin_filesize = lv_bytecount
                                                       filename     = lv_path
                                                       filetype     = 'BIN'
                                              CHANGING data_tab     = lt_file_tab
                                                     EXCEPTIONS
    file_write_error          = 1
    no_batch                  = 2
    gui_refuse_filetransfer   = 3
    invalid_type              = 4
    no_authority              = 5
    unknown_error             = 6
    header_not_allowed        = 7
    separator_not_allowed     = 8
    filesize_not_allowed      = 9
    header_too_long           = 10
    dp_error_create           = 11
    dp_error_send             = 12
    dp_error_write            = 13
    unknown_dp_error          = 14
    access_denied             = 15
    dp_out_of_memory          = 16
    disk_full                 = 17
    dp_timeout                = 18
    file_not_found            = 19
    dataprovider_exception    = 20
    control_flush_error       = 21
    not_supported_by_gui      = 22
    error_no_gui              = 23
    OTHERS                    = 24
        ).


  ENDMETHOD.
  METHOD create_document.
    DATA
          : lo_ixml           TYPE REF TO if_ixml
          , lo_streamfactory  TYPE REF TO if_ixml_stream_factory
          , lo_ostream        TYPE REF TO if_ixml_ostream
          , lo_renderer       TYPE REF TO if_ixml_renderer
          .

* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_ixml = cl_ixml=>create( ).


    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = rp_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = me->document ).
    lo_renderer->render( ).
  ENDMETHOD.
  METHOD map_line.

    DATA
      : lv_key TYPE string
      , lv_value TYPE string
      , lv_date TYPE datum
      , lv_uzeit TYPE sy-uzeit
      , lv_type TYPE c
      , lv_node_value TYPE string
      .

    DATA(iterator) = node->create_iterator( ).

    DO.
      DATA(lr_node) = iterator->get_next( ).

      IF lr_node IS INITIAL.
        EXIT.
      ENDIF.

      CHECK lr_node->get_type( ) = if_ixml_node=>co_node_element.
      CHECK lr_node->get_name( ) = 't'.

      lv_node_value = lr_node->get_value( ).
      LOOP AT components ASSIGNING FIELD-SYMBOL(<fs_component>).
        ASSIGN COMPONENT <fs_component>-name OF STRUCTURE data TO FIELD-SYMBOL(<fs_value>).

        DESCRIBE FIELD <fs_value> TYPE lv_type.

        lv_key = |\{{ <fs_component>-name }\}|.


        CASE lv_type.
          WHEN 'D'.
            lv_date = <fs_value>.
            lv_value = |{ lv_date DATE = ENVIRONMENT }|.
          WHEN 'T'.
            lv_uzeit = <fs_value>.
            lv_value = |{ lv_uzeit TIME = ENVIRONMENT }|.
          WHEN OTHERS.
            lv_value = <fs_value>.
        ENDCASE.


        REPLACE ALL OCCURRENCES OF lv_key IN lv_node_value WITH lv_value.
        CHECK sy-subrc = 0.

        lr_node->set_value( lv_node_value  ).

      ENDLOOP.

    ENDDO.

  ENDMETHOD.

ENDCLASS.