*&---------------------------------------------------------------------*
*&  Include           ZCL_DOCX_CLASS
*&
*&    Author: Sikidin A.P.  anton.sikidin@gmail.com
*&
*&    ZCL_DOCX is replasement of zwww_openform for docx
*&
*&---------------------------------------------------------------------*

class lcl_recursive_data definition deferred.


types
: begin of t_key_value
,  key type string
,  value type string
, end of t_key_value

, tt_key_value type table of t_key_value

, begin of t_key_table
,   key type string
,   value type ref to data
, end of   t_key_table

, tt_key_table type table of t_key_table

, t_lcl type ref to lcl_recursive_data
, tt_lcl type table of t_lcl

, tt_keys type table of string
.

class lcl_recursive_data definition.
  public section.


    data
          : key_value type tt_key_value
          , key_table type tt_key_table
          , key type string
          , keys type tt_keys
          , key_lcl type tt_lcl
          .
    methods append_key_value
      importing
        value(iv_key) type string
        !iv_value     type any .


    methods append_key_table
      importing
        value(iv_key) type string
        !iv_table     type any table.

    methods create_document
      importing
        value(iv_key) type string
      returning
        value(r_data) type ref to lcl_recursive_data   .



endclass.

class lcl_recursive_data implementation.
  method append_key_value.


    translate iv_key to upper case.

    append initial line to key_value assigning field-symbol(<fs_key_value>).
    <fs_key_value>-key = |\{{ iv_key }\}| .
    <fs_key_value>-value = iv_value.


  endmethod.

  method append_key_table.

    translate iv_key to upper case.

    append initial line to key_table assigning field-symbol(<fs_key_table>).
    <fs_key_table>-key = iv_key.

    create data <fs_key_table>-value like iv_table.
    assign <fs_key_table>-value->* to field-symbol(<fs_any_table>).
    <fs_any_table> = iv_table.

  endmethod.

  method create_document .

    translate iv_key to upper case.

    create object r_data.
    r_data->key = iv_key.

    collect iv_key into keys.

    append r_data to key_lcl.


  endmethod.

endclass.

class lcl_docx definition.

  public section.

    types
    : t_ref_node type ref to if_ixml_node
    , tt_node type standard table of t_ref_node

    , begin of t_bookmark
    , id type string
    , start_node type   t_ref_node
    , end_node type   t_ref_node
    , in_table type xfeld
    , name type string
    , start_height   type i
    , end_height type i
    , end of t_bookmark

    , begin of t_stack_data
    , start type xfeld
    , id type string
    , node type t_ref_node
    , position type i
    , length type i
    , collision type i
    , end of t_stack_data

    , begin of t_collision
    , collision type string
    , count type i
    , end of  t_collision

    .


    methods load_smw0
      importing
        !i_w3objid type w3objid .


    methods add_file
      importing
        !iv_path type string
        !iv_data type xstring.
    methods save
      importing
        !on_desktop   type xfeld default 'X'
        !iv_folder    type string default 'report'
        !iv_path      type string default ''
        !iv_file_name type string default 'report.docx'
        !no_execute   type xfeld default '' .





    methods map_data
      importing
        !ir_xml_node type ref to if_ixml_document optional
        !ir_data     type ref to lcl_recursive_data.

    methods check_flag
      importing
        !it_keys type tt_keys.



  protected section.

    constants c_document type string value 'word/document.xml' ##NO_TEXT.

    methods map_values
      importing
        !ir_xml_node  type ref to if_ixml_node optional
        !it_key_value type tt_key_value .


    methods map_table
      importing
        !ir_xml_node type ref to if_ixml_document optional
        !iv_key      type string
        !it_data     type any table .

    methods append_node
      importing
        !ir_source type ref to if_ixml_node optional
        !ir_dest   type ref to if_ixml_node optional
        !iv_key    type string.

    methods get_fragment
      importing
        !ir_xml_node       type ref to if_ixml_document optional
        !iv_key            type string
      returning
        value(rr_fragment) type ref to if_ixml_document.

    methods normalize_key.
    methods align_bookmark.

    methods get_from_zip_archive
      importing
        !i_filename      type string
      returning
        value(r_content) type xstring .
    methods get_ixml_from_zip_archive
      importing
        !i_filename   type string
      returning
        value(r_ixml) type ref to if_ixml_document .

    methods normalize_space
      importing
        !iv_content      type xstring
      returning
        value(r_content) type xstring.


  private section.

    data zip type ref to cl_abap_zip .
    data document type ref to if_ixml_document .

    methods upper_case
      importing
        !iv_str       type string
      returning
        value(rv_str) type string.

    methods dump_data
      importing
        !node      type ref to if_ixml_document optional
        !node_node type ref to if_ixml_node optional
        !fname     type string.

    methods create_document
      returning
        value(rp_content) type xstring .

    methods map_line
      importing
        !node       type ref to if_ixml_document
        !components type abap_compdescr_tab
        !data       type any .
endclass.

class lcl_docx implementation.
  method load_smw0.
    data
          : lv_templ_xstr type xstring
          , lt_mime type table of w3mime
          .

    data(ls_key) = value wwwdatatab(
    relid = 'MI'
    objid = i_w3objid ).

    call function 'WWWDATA_IMPORT'
      exporting
        key    = ls_key
      tables
        mime   = lt_mime
      exceptions
        others = 1.
    if sy-subrc <> 0.
      return.
    endif.

    try.
        lv_templ_xstr = cl_bcs_convert=>xtab_to_xstring( lt_mime ).
      catch cx_bcs.
        return.
    endtry.

    if zip is initial.
      create object zip.
    endif.

    zip->load( lv_templ_xstr ).

    document = me->get_ixml_from_zip_archive( me->c_document ).


    normalize_key( ).
*    align_bookmark( ).
  endmethod.

  method normalize_key.
    data
          : lt_nodes type table of t_ref_node
          , lv_in type c
          , lv_regex_open type string value '\{[^\}]*$'
          , lv_regex_close type string value '\}[^\{]*'
          , lv_tmp_str type string
          .

*    dump_data(  node = document
*               fname = 'before' ).

    data(iterator) = document->create_iterator( ).

    do.
      data(node) = iterator->get_next( ).
      if node is initial.
        exit.
      endif.

      check node->get_type( ) = if_ixml_node=>co_node_element.

      check  node->get_name( ) = 'p'.

      refresh lt_nodes.
      clear lv_in.

      data(nodes) = node->get_children( ).
      do nodes->get_length( ) times.

        data(child) = nodes->get_item( sy-index - 1 ).


        data(nodes_2) = child->get_children( ).

        do  nodes_2->get_length( ) times.
          data(child_2) = nodes_2->get_item( sy-index - 1 ).

          check child_2->get_name( ) = 't'.

          data(child_2_value) = child_2->get_value( ).

          child_2_value = child_2->get_value( ).
          child_2->set_value( upper_case( child_2_value ) ).


          if lv_in is not initial.

            append child_2 to lt_nodes.


            find regex lv_regex_close  in child_2_value.

            check sy-subrc = 0.

            find regex lv_regex_open  in child_2_value.

            check sy-subrc ne 0.

            clear lv_tmp_str .

            loop at lt_nodes assigning field-symbol(<fs_node>).
              child_2_value = <fs_node>->get_value( ).
              lv_tmp_str = |{ lv_tmp_str }{ child_2_value }|.
              <fs_node>->set_value( '' ).

            endloop.

            read table lt_nodes assigning <fs_node> index 1.
            <fs_node>->set_value( upper_case( lv_tmp_str ) ).

            data
                  :  lo_element    type ref to if_ixml_element
                  .
            lo_element ?= <fs_node>.

            lo_element->set_attribute( name = 'space'
                                       namespace = 'xml'
                                       value = 'preserve' ).




            clear  lv_in.

            refresh lt_nodes.

          else.

            find regex lv_regex_open  in child_2_value.

            check sy-subrc = 0.

            lv_in = 'X'.

            append child_2 to lt_nodes.

          endif.

        enddo.

      enddo.

    enddo.

*    dump_data( node = document
*               fname = 'after' ).

  endmethod.
  method align_bookmark.
    data
           : lt_bokkmarks type table of t_bookmark
           .

    data(iterator) = document->create_iterator( ).

    do  .
      data(node) = iterator->get_next( ).

      if node is initial.
        exit.
      endif.

      check node->get_type( ) = if_ixml_node=>co_node_element.

      data(name) = node->get_name( ).

      case name.
        when 'bookmarkStart'.
        when 'bookmarkEnd'.
        when others.
          continue.
      endcase.

      data(attributes) = node->get_attributes( ).

      data
            : lv_name type string
            , lv_id type string
            .


      do attributes->get_length( ) times.

        data(attribute) = attributes->get_item( sy-index - 1 ).

        case attribute->get_name( ).
          when 'id'.
            lv_id = attribute->get_value( ).
            read table lt_bokkmarks assigning field-symbol(<fs_bookmark>) with key id = lv_id.
            if sy-subrc ne 0.
              append initial line to lt_bokkmarks assigning <fs_bookmark>.
              <fs_bookmark>-id = lv_id.
            endif.

            case name.
              when 'bookmarkStart'.
                <fs_bookmark>-start_node = node.
                <fs_bookmark>-start_height = node->get_height( ).
              when 'bookmarkEnd'.
                <fs_bookmark>-end_node = node.
                <fs_bookmark>-end_height = node->get_height( ).
            endcase.


          when 'name'.
            lv_name = attribute->get_value( ) .
            translate lv_name to upper case.
            attribute->set_value( lv_name ) .
            <fs_bookmark>-name = lv_name.


        endcase.

      enddo.

    enddo.


*    table or not

    data
          : lv_node type t_ref_node
          , lv_ref_node type t_ref_node
          , lv_ref_node_2 type t_ref_node

          .

    loop at lt_bokkmarks assigning <fs_bookmark>.


      lv_node = <fs_bookmark>-start_node.
      do .

        if lv_node->is_root( ) is not   initial.
          exit.
        endif.

        data(node_name) = lv_node->get_name( ).

        if node_name = 'tr'.
          <fs_bookmark>-in_table = 'X'.
          exit.
        endif.

        lv_node = lv_node->get_parent( ).


      enddo.

    endloop.


    sort lt_bokkmarks by id ascending .

    loop at lt_bokkmarks assigning <fs_bookmark>.

      if <fs_bookmark>-in_table = 'X'.

        lv_ref_node = <fs_bookmark>-start_node.

        do .
          node_name = lv_ref_node->get_name( ).

          if node_name = 'tr'.
            lv_node = lv_ref_node->get_parent( ).
            exit.
          endif.
          lv_ref_node = lv_ref_node->get_parent( ).
        enddo.

        <fs_bookmark>-start_node->remove_node( ).
        lv_node->insert_child( new_child = <fs_bookmark>-start_node
                                  ref_child = lv_ref_node ).

        clear
        : lv_ref_node_2
        .

        lv_ref_node_2 = <fs_bookmark>-end_node.

        do .
          if lv_ref_node_2 is initial.
            exit.
          endif.
          node_name = lv_ref_node_2->get_name( ).

          if node_name = 'tr'.
            exit.
          endif.
          lv_ref_node_2 = lv_ref_node_2->get_parent( ).
        enddo.



        if lv_ref_node_2 is not initial.
          <fs_bookmark>-end_node->remove_node( ).
          lv_node->insert_child( new_child = <fs_bookmark>-end_node
                          ref_child = lv_ref_node_2 ).
*        ELSE.
*          lv_node->append_child( new_child = <fs_bookmark>-end_node ).
        endif.

      else.

        lv_ref_node = <fs_bookmark>-start_node->get_parent( ).
        <fs_bookmark>-start_node->remove_node( ).
        lv_node = lv_ref_node->get_parent( ).

        lv_node->insert_child( new_child = <fs_bookmark>-start_node
                              ref_child = lv_ref_node ).


        data
              : lv_height_start type i
              , lv_height_end type i
              .

        lv_height_start = <fs_bookmark>-start_node->get_height( ).

        lv_ref_node = <fs_bookmark>-end_node.

        do .

          if lv_ref_node is initial.
            exit.
          endif.

          lv_height_end = lv_ref_node->get_height( ).

          if lv_height_end = lv_height_start.
            exit.
          endif.


          lv_ref_node = lv_ref_node->get_parent( ).


        enddo.

        if lv_ref_node ne <fs_bookmark>-end_node.
          <fs_bookmark>-end_node->remove_node( ).

          lv_node = lv_ref_node->get_parent( ).

          lv_node->insert_child( new_child = <fs_bookmark>-end_node
                                ref_child = lv_ref_node ).
        endif.

      endif.

    endloop.



    data
          : lv_position type i
          , lt_stack_data type table of t_stack_data
          , lt_stack_data_sorted type table of t_stack_data
          , lt_id type table of string

          , lt_collision type table of t_collision
          , ls_collision type t_collision

          , lt_old type table of t_ref_node
          , lt_sorted type table of t_ref_node
          .

    iterator = document->create_iterator( ).


    do  .
      node = iterator->get_next( ).
      if node is initial.
        exit.
      endif.
      check node->get_type( ) = if_ixml_node=>co_node_element.

      name = node->get_name( ).

      case name.
        when 'bookmarkStart'.
        when 'bookmarkEnd'.
        when others.
          add 1 to lv_position.

          continue.
      endcase.

      attributes = node->get_attributes( ).


      append initial line to lt_stack_data assigning field-symbol(<fs_stack_data>).

      if name = 'bookmarkStart'.
        <fs_stack_data>-start = 'X'.
      endif.

      <fs_stack_data>-position = lv_position.
      <fs_stack_data>-collision = lv_position.
      <fs_stack_data>-node = node.

      ls_collision-collision = lv_position.
      ls_collision-count  = 1.

      collect ls_collision into lt_collision.

      do attributes->get_length( ) times.
        attribute = attributes->get_item( sy-index - 1 ).
        check attribute->get_name( ) = 'id'.

        <fs_stack_data>-id = attribute->get_value( ).
        collect <fs_stack_data>-id into lt_id.

      enddo.

    enddo.

    lt_stack_data_sorted = lt_stack_data.

    data
          : lv_length type i
          .

    loop at lt_id assigning field-symbol(<fs_id>).

      clear lv_length.

      loop at lt_stack_data_sorted assigning <fs_stack_data> where id = <fs_id>.

        if lv_length is initial.
          lv_length = <fs_stack_data>-position.
        else.
          lv_length = <fs_stack_data>-position - lv_length .
        endif.

      endloop.

      loop at lt_stack_data_sorted assigning <fs_stack_data> where id = <fs_id>.
        case 'X'.
          when <fs_stack_data>-start.
            <fs_stack_data>-length =  lv_length.
          when others.
            <fs_stack_data>-length =  lv_length * -1 .
        endcase.
      endloop.

    endloop.

    sort lt_stack_data_sorted by collision  start length descending.


    loop at lt_collision assigning field-symbol(<fs_collision>) where count > 1.

      refresh
      : lt_old
      , lt_sorted
      .

      loop at lt_stack_data assigning <fs_stack_data> where collision = <fs_collision>-collision.
        append <fs_stack_data>-node to lt_old.
      endloop.

      loop at lt_stack_data_sorted assigning <fs_stack_data> where collision = <fs_collision>-collision.
        append <fs_stack_data>-node to lt_sorted.
      endloop.

      check lt_old ne lt_sorted.

      read table lt_old assigning field-symbol(<fs_old>) index <fs_collision>-count.

      clear lv_ref_node.

      lv_node = <fs_old>->get_parent( ).
      lv_ref_node = <fs_old>->get_next( ).

      loop at lt_old assigning <fs_old>.
        <fs_old>->remove_node( ).
      endloop.


      loop at lt_sorted assigning field-symbol(<fs_sorted>).

        if lv_ref_node is not initial.
          lv_node->insert_child( new_child = <fs_sorted>
                            ref_child = lv_ref_node ).
        else.

          lv_node->append_child( new_child = <fs_sorted> ).
        endif.

      endloop.

    endloop.

  endmethod.

  method add_file.

    zip->delete( name = iv_path ).

    zip->add( name    = iv_path
               content = iv_data ).

  endmethod.
  method save.

    data
             : lv_content         type xstring
             , lv_content2         type xstring
             .

*    dump_data( node = document
*                   fname = 'before save' ).

    lv_content = me->create_document( ).



    data
          : lv_string type string
          , lt_content_source type table of string
          , lt_content_dest type table of string
          , lt_string type table of string
          , lt_data type table of text255
          .

    call function 'CRM_IC_XML_XSTRING2STRING'
      exporting
        inxstring = lv_content
      importing
        outstring = lv_string.


    split lv_string at '[SPACE]' into table lt_string.

    data
          : lv_len type i
          , lv_buf type text255
          , lv_pos type i
          , lv_rem type i
          .
    lv_pos = 0.
    lv_rem = 255.
    loop at lt_string assigning field-symbol(<fs_str>).

      lv_len = strlen( <fs_str> ).

      while lv_len > 0.
        if lv_len > lv_rem.
          lv_buf+lv_pos(lv_rem) = <fs_str>(lv_rem).
          append lv_buf to lt_data.

          <fs_str> = <fs_str>+lv_rem.
          lv_pos = 0.
          lv_rem = 255.
          clear lv_buf.
        else.
          lv_buf+lv_pos = <fs_str>.
          lv_pos = lv_pos + lv_len.
          lv_rem = lv_rem - lv_len.
          clear <fs_str>.

        endif.
        lv_len = strlen( <fs_str> ).



      endwhile.

      if lv_pos < 254.
        lv_pos = lv_pos + 1.
        lv_rem = lv_rem - 1.
      elseif lv_pos  = 254.
        append lv_buf to lt_data.
        lv_pos = 0.
        lv_rem = 255.
        clear lv_buf.
      else.
        append lv_buf to lt_data.
        lv_pos = 1.
        lv_rem = 254.
        clear lv_buf.


      endif.


    endloop.

    append lv_buf to lt_data.


    field-symbols
                   : <xstr> type x
                   .

    data
          : lv_x1(1) type x
          , lv_i type i
          , lv_i2 type i
          , lv_str TYPE string
          .

    clear lv_content.
*    loop at lt_data assigning <xstr> casting .
*      concatenate lv_content <xstr> into lv_content in byte mode.
*    endloop.



 LOOP AT lt_data ASSIGNING FIELD-SYMBOL(<fs_data>).

   CONCATENATE lv_str <fs_data> INTO lv_str RESPECTING BLANKS.

 ENDLOOP.

 DATA
          : lr_conv_out TYPE REF TO cl_abap_conv_out_ce
          , lv_echo_xstring TYPE xstring
          .

    lr_conv_out = cl_abap_conv_out_ce=>create(
*      encoding    = 'UTF-8'               " Кодировка в которую будем преобразовывать
    ).


    lr_conv_out->convert( EXPORTING data = lv_str IMPORTING buffer = lv_content ).


    zip->delete( name = me->c_document ).
    zip->add( name    = me->c_document
               content = lv_content ).

    lv_content = zip->save( ).

    data
          : lt_file_tab  type solix_tab
          , lv_bytecount type i
          , lv_path      type string
          .

    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_content ).
    lv_bytecount = xstrlen( lv_content ).


    if iv_path is initial.
      if on_desktop is not initial.
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
    others                    = 24
        ).


    check no_execute is  initial.
    cl_gui_frontend_services=>execute(  document  =  lv_path ).

  endmethod.
  method map_values.

    check it_key_value is not initial.

    if ir_xml_node is bound.
      data(iterator) = ir_xml_node->create_iterator( ).
    else.
      iterator = document->create_iterator( ).
    endif.


    do.
      data(node) = iterator->get_next( ).
      if node is initial.
        exit.
      endif.

      check node->get_type( ) = if_ixml_node=>co_node_element.

      check node->get_name( ) = 't'.

      data(value) = node->get_value( ).
      loop at it_key_value assigning field-symbol(<fs_key_value>).
        replace all occurrences of <fs_key_value>-key in value with <fs_key_value>-value.
        check sy-subrc = 0.
        node->set_value( value  ).

      endloop.

    enddo.
  endmethod.


  method get_fragment.
    data
                      : lv_found type c
                      , lv_id  type string
                      , lr_start type ref to if_ixml_node


                      , lt_node type tt_node
                      , lv_first_run type c

                      , lv_start type c
                      .



    data(iterator) = ir_xml_node->create_iterator( ).


    data(ixmlfactory) = cl_ixml=>create( ).
    rr_fragment = ixmlfactory->create_document( ).


    do.
      data(node) = iterator->get_next( ).

      if node is initial.
        exit.
      endif.

      check node->get_type( ) = if_ixml_node=>co_node_element.

      check node->get_name( ) = 'bookmarkStart'.

      data(attributes) = node->get_attributes( ).


      do attributes->get_length( ) times.
        data(attribute) = attributes->get_item( sy-index - 1 ).
        case attribute->get_name( ).
          when 'id'.
            lv_id = attribute->get_value( ).
          when 'name'.
            if attribute->get_value( ) = iv_key .
              lv_found = 'X'.
              lv_first_run = 'X'.
              lr_start = node.
            endif.

        endcase.

      enddo.

      check  lv_first_run is not initial.
      clear lv_first_run .
      exit.

    enddo.

    check lr_start is not initial.
    data(parent) = lr_start->get_parent( ).
    data(children) = parent->get_children( ).

    do  children->get_length( ) times.
      data(child) = children->get_item( sy-index - 1 ).
      data(name) = child->get_name( ).

      if name = 'bookmarkEnd' and lv_found is not initial.
        attributes = child->get_attributes( ).
        do attributes->get_length( ) times.

          attribute = attributes->get_item( sy-index - 1 ).

          check attribute->get_name( ) = 'id'.
          check lv_id = attribute->get_value( ).

          clear lv_start.

        enddo.

      endif.

      if lv_start is not initial.
        append child to lt_node.
      endif.


      if name = 'bookmarkStart' and lv_found is not initial.
        attributes = child->get_attributes( ).
        do attributes->get_length( ) times.
          attribute = attributes->get_item( sy-index - 1 ).
          check attribute->get_name( ) = 'id'.
          check lv_id = attribute->get_value( ).

          lv_start = 'X'.

        enddo.

      endif.

    enddo.

    loop at lt_node assigning field-symbol(<fs_node>).
      data(clone_node) = <fs_node>->clone( ) .

      name = clone_node->get_name( ).
      rr_fragment->append_child( clone_node ).
      <fs_node>->remove_node( ).

    endloop.

  endmethod.

  method  append_node .

    data
          : lv_node_name type string
          .

    data(iterator) = ir_dest->create_iterator( ).

    do.
      data(node) = iterator->get_next( ).

      if node is initial.
        exit.
      endif.

      check  node->get_type( ) = if_ixml_node=>co_node_element.

      check node->get_name( ) = 'bookmarkStart'.

      data(attributes) = node->get_attributes( ).


      do attributes->get_length( ) times.
        data(attribute) = attributes->get_item( sy-index - 1 ).

        check attribute->get_name( ) = 'name'.

        check attribute->get_value( ) = iv_key .

        data(parent) = node->get_parent( ).


        data(children) = ir_source->get_children( ).
        data(ch_count) = children->get_length( ).
        do  ch_count times.
          data(child) = children->get_item( sy-index - 1 ).

          data(clone_child) = child->clone( ).

          parent->insert_child( new_child = clone_child
                                ref_child = node ).

        enddo.

        return.

      enddo.

    enddo.

  endmethod.

  method  map_data .

    data
          : lr_node type ref to if_ixml_document

          , lr_document type ref to if_ixml_document
          , lr_clone type ref to if_ixml_document
          .

    field-symbols
                   : <fs_any_table> type any table
                   .

    if ir_xml_node is bound.
      lr_node = ir_xml_node.
    else.
      lr_node = document.
    endif.

    map_values(       ir_xml_node  = lr_node
                      it_key_value = ir_data->key_value  ) .



    loop at ir_data->key_table assigning field-symbol(<fs_key_table>).

      assign <fs_key_table>-value->* to <fs_any_table>.

      map_table(  ir_xml_node  = lr_node
                  iv_key       = <fs_key_table>-key
                  it_data      = <fs_any_table> ).

    endloop.



    loop at ir_data->keys assigning field-symbol(<fs_key>).

      lr_document = get_fragment( ir_xml_node = lr_node iv_key = <fs_key> ).

      loop at ir_data->key_lcl assigning field-symbol(<fs_key_lcl>).

        check <fs_key_lcl>->key = <fs_key>.

        check <fs_key_lcl>->key_value is not initial or
              <fs_key_lcl>->key_table is not initial or
              <fs_key_lcl>->keys      is not initial.

        lr_clone ?= lr_document->clone( ).

        map_data( ir_xml_node  = lr_clone
                  ir_data      = <fs_key_lcl> ).

        append_node(  ir_source = lr_clone
                      ir_dest   = lr_node
                      iv_key    = <fs_key> ).


      endloop.

    endloop.

  endmethod.

  method check_flag.
    data(iterator) = document->create_iterator( ).

    do.
      data(node) = iterator->get_next( ).
      if node is initial.
        exit.
      endif.

      if node->get_type( ) <> if_ixml_node=>co_node_element.
        continue.
      endif.
      data(name) = node->get_name( ).
      if name = 'fldChar'.

        data(checkbox_iterator) = node->create_iterator( ).

        data
              : lv_found type c
              .

        clear lv_found.

        do .
          data(check_box_node) = checkbox_iterator->get_next( ).
          if check_box_node is initial.
            exit.
          endif.

          if check_box_node->get_type( ) <> if_ixml_node=>co_node_element.
            continue.
          endif.

          data(check_box_name) = check_box_node->get_name( ).

          if check_box_name = 'name'.
            data(attributes) = check_box_node->get_attributes( ).

            do attributes->get_length( ) times.
              data(attribute) = attributes->get_item( sy-index - 1 ).
              if attribute->get_name( ) = 'val'.
                data(atribute_value) =  attribute->get_value( ) .

                read table it_keys transporting no fields with key table_line = atribute_value.

                if sy-subrc = 0.
                  lv_found = 'X'.
                  exit.
                endif.

              endif.

            enddo.



          endif.

          if lv_found is not initial and check_box_name = 'default'.
            attributes = check_box_node->get_attributes( ).

            do attributes->get_length( ) times.
              attribute = attributes->get_item( sy-index - 1 ).
              if attribute->get_name( ) = 'val'.
                attribute->set_value( '1' ) .
                exit.
              endif.

            enddo.

          endif.

        enddo.


      endif.


    enddo.

  endmethod.

  method map_table.

    data
      : lr_node type ref to if_ixml_document
      , lr_document type ref to if_ixml_document
      , lr_clone type ref to if_ixml_document

      , l_r_structdescr      type ref to cl_abap_structdescr
      .


    if ir_xml_node is bound.
      lr_node = ir_xml_node.
    else.
      lr_node = document.
    endif.



    lr_document = get_fragment( ir_xml_node = lr_node
                                iv_key      = iv_key ).

    loop at it_data assigning field-symbol(<fs_data>).
      lr_clone  ?= lr_document->clone( ).

      if l_r_structdescr is not bound.
        l_r_structdescr ?= cl_abap_structdescr=>describe_by_data( <fs_data> ).
      endif.

      map_line(    node       = lr_clone
                   components = l_r_structdescr->components
                   data       = <fs_data>              ).

      append_node(  ir_source = lr_clone
                    ir_dest   = lr_node
                    iv_key    = iv_key ).
    endloop.

  endmethod.


  method get_from_zip_archive.
    assert zip is bound. " zip object has to exist at this point

    zip->get( exporting  name =  i_filename
                        importing content = r_content ).

  endmethod.
  method normalize_space.


    data
          : lv_string type string
          , lt_content_source type table of string
          , lt_content_dest type table of string
          .


    call function 'CRM_IC_XML_XSTRING2STRING'
      exporting
        inxstring = iv_content
      importing
        outstring = lv_string.


    split lv_string at '<' into table lt_content_source.

    loop at lt_content_source assigning field-symbol(<fs_content>).
      if <fs_content> is not initial.
        <fs_content> = '<' && <fs_content>.
      endif.

    endloop.




*    DATA
*          : moff TYPE i
*          , lv_start TYPE string
*          , lv_end TYPE string
*          , lt_buf TYPE TABLE OF string
*
*          , lv_head TYPE string
*
*          , lv_need_append(1)
*          , lv_flag_in(1)
*          , lv_flag_first_was(1)
*
*          , lv_flag_joined(1)
*          .
*
*    LOOP AT lt_content_source ASSIGNING <fs_content>.
*
**      FIND '<w:p ' IN <fs_content>.
**      IF sy-subrc = 0.
**
**      ENDIF.
*
*      IF <fs_content> = '</w:p>'.
*        APPEND lv_start TO lt_content_dest.
*
*
*        APPEND LINES OF lt_buf TO lt_content_dest.
*        REFRESH lt_buf.
*
*        APPEND <fs_content> TO lt_content_dest.
*
*        CLEAR
*         : lv_need_append
*         , lv_flag_in
*         , lv_start
*         , lv_flag_first_was
*         .
*
*        CONTINUE .
*      ENDIF.
*
*      FIND REGEX '<w:t(\s|>)' IN   <fs_content>.
*      IF sy-subrc NE 0.
*
*        IF lv_flag_in IS INITIAL and lv_start is INITIAL.
*          APPEND <fs_content> TO lt_content_dest.
*        ELSE.
*          APPEND <fs_content> TO lt_buf.
*        ENDIF.
*
*        CONTINUE .
*      ENDIF.
*
*      IF lv_flag_first_was IS INITIAL.
*        lv_flag_first_was = 'X'.
*      ELSE.
*        lv_flag_in = 'X'.
*      ENDIF.
*
*      SPLIT <fs_content> AT '>' INTO lv_head lv_end.
*
*      lv_head = lv_head && '>'.
*
*      IF lv_flag_in IS INITIAL .
*        APPEND lv_head TO lt_content_dest.
*      ELSE.
*        APPEND lv_head TO lt_buf.
*      ENDIF.
*
*      CLEAR
*      : lv_flag_joined
*      .
*
*      IF lv_flag_in IS NOT INITIAL.
*
*        IF  lv_need_append IS NOT INITIAL.
*          lv_start = lv_start && lv_end.
*          lv_flag_joined = 'X'.
*        ELSE.
*          FIND REGEX '^\s' IN lv_end.
*
*          IF sy-subrc = 0.
*            lv_start = lv_start && lv_end.
*            lv_flag_joined = 'X'.
*          ENDIF.
*
*        ENDIF.
*
*      ENDIF.
*
*      IF lv_flag_joined IS INITIAL.
*        APPEND lv_start TO lt_content_dest.
*        APPEND LINES OF lt_buf TO lt_content_dest.
*        REFRESH lt_buf.
*
*        lv_start = lv_end.
*
*      ENDIF.
*
*
*      FIND REGEX '\s$' IN lv_start .
*      IF sy-subrc = 0.
*        lv_need_append = 'X'.
*      ELSE.
*        CLEAR lv_need_append.
*      ENDIF.
*
*
*    ENDLOOP.

    data
          : lv_str1 type string
          , lv_str2 type string
          .
    loop at lt_content_source assigning field-symbol(<fs_source>).
      find regex '<w:t(\s|>)' in   <fs_source>.
      if sy-subrc ne 0.
        append <fs_source> to lt_content_dest.
      else.

        split <fs_source> at '>' into lv_str1 lv_str2.
        lv_str1 = lv_str1 && '>'.
        append lv_str1 to lt_content_dest.
        replace all occurrences of regex '\s' in lv_str2 with '[SPACE]'.
        append lv_str2 to lt_content_dest.


      endif.

    endloop.


*    lt_content_dest = lt_content_source.

    clear lv_string.
    loop at lt_content_dest assigning field-symbol(<fs_dest>).
      lv_string = lv_string && <fs_dest>.

    endloop.

    call function 'CRM_IC_XML_STRING2XSTRING'
      exporting
        instring   = lv_string
      importing
        outxstring = r_content.



  endmethod.

  method get_ixml_from_zip_archive.
    data: lv_content       type xstring,
          lo_ixml          type ref to if_ixml,
          lo_streamfactory type ref to if_ixml_stream_factory,
          lo_istream       type ref to if_ixml_istream,
          lo_parser        type ref to if_ixml_parser.

*--------------------------------------------------------------------*
* Load XML file from archive into an input stream,
* and parse that stream into an ixml object
*--------------------------------------------------------------------*
    lv_content        = me->get_from_zip_archive( i_filename ).
    lv_content = normalize_space( lv_content ).
    lo_ixml           = cl_ixml=>create( ).
    lo_streamfactory  = lo_ixml->create_stream_factory( ).
    lo_istream        = lo_streamfactory->create_istream_xstring( lv_content ).
    r_ixml            = lo_ixml->create_document( ).
    lo_parser         = lo_ixml->create_parser( stream_factory = lo_streamfactory
                                                istream        = lo_istream
                                                document       = r_ixml ).
*    lo_parser->set_normalizing( 'X' ).
    lo_parser->set_validating( mode = if_ixml_parser=>co_no_validation ).
    lo_parser->parse( ).
  endmethod.


  method upper_case.
    data
          : result_tab type match_result_tab
          , lv_search type string
          , lv_replace type string
          .
    rv_str = iv_str.

*    replace all occurrences of regex '^\{' in rv_str with ' {'  .
*    find regex '\}$' in rv_str.
*    if sy-subrc = 0.
*      concatenate rv_str space into rv_str respecting blanks.
*    endif.


    find all occurrences of regex '\{[^\}]*\}' in iv_str results result_tab.
    loop at result_tab assigning field-symbol(<fs_result>).
      lv_search = iv_str+<fs_result>-offset(<fs_result>-length).
      lv_replace = lv_search.

      translate lv_replace to upper case.

      replace all occurrences of lv_search in rv_str with lv_replace.

    endloop.

  endmethod.

  method dump_data .

    data
      : lo_ixml           type ref to if_ixml
      , lo_streamfactory  type ref to if_ixml_stream_factory
      , lo_ostream        type ref to if_ixml_ostream
      , lo_renderer       type ref to if_ixml_renderer
      , lv_xstring type xstring
      .

* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_ixml = cl_ixml=>create( ).


    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = lv_xstring ).
    if node_node is supplied.
      data(document) = lo_ixml->create_document( ).
      document->append_child( node_node ).
      lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = document ).
    else.
      lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = node ).
    endif.

    lo_renderer->render( ).


    data
          : lt_file_tab  type solix_tab
          , lv_bytecount type i
          , lv_path      type string
          .

    lt_file_tab = cl_bcs_convert=>xstring_to_solix( iv_xstring  = lv_xstring ).
    lv_bytecount = xstrlen( lv_xstring ).

    cl_gui_frontend_services=>get_desktop_directory( changing desktop_directory = lv_path ).

    cl_gui_cfw=>flush( ).

    concatenate lv_path '\report\'  fname '.txt'  into lv_path.

    cl_gui_frontend_services=>gui_download( exporting bin_filesize = lv_bytecount
                                                       filename     = lv_path
                                                       filetype     = 'BIN'
                                              changing data_tab     = lt_file_tab
                                                     exceptions
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
    others                    = 24
        ).


  endmethod.
  method create_document.
    data
          : lo_ixml           type ref to if_ixml
          , lo_streamfactory  type ref to if_ixml_stream_factory
          , lo_ostream        type ref to if_ixml_ostream
          , lo_renderer       type ref to if_ixml_renderer
          .

* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_ixml = cl_ixml=>create( ).


    lo_streamfactory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_streamfactory->create_ostream_xstring( string = rp_content ).
    lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream document = me->document ).
    lo_renderer->render( ).
  endmethod.
  method map_line.

    data
      : lv_key type string
      , lv_value type string
      , lv_date type datum
      , lv_uzeit type sy-uzeit
      , lv_type type c
      , lv_node_value type string
      .

    data(iterator) = node->create_iterator( ).

    do.
      data(lr_node) = iterator->get_next( ).

      if lr_node is initial.
        exit.
      endif.

      check lr_node->get_type( ) = if_ixml_node=>co_node_element.
      check lr_node->get_name( ) = 't'.

      lv_node_value = lr_node->get_value( ).
      loop at components assigning field-symbol(<fs_component>).
        assign component <fs_component>-name of structure data to field-symbol(<fs_value>).

        describe field <fs_value> type lv_type.

        lv_key = |\{{ <fs_component>-name }\}|.


        case lv_type.
          when 'D'.
            lv_date = <fs_value>.
            lv_value = |{ lv_date date = environment }|.
          when 'T'.
            lv_uzeit = <fs_value>.
            lv_value = |{ lv_uzeit time = environment }|.
          when others.
            lv_value = <fs_value>.
        endcase.


        replace all occurrences of lv_key in lv_node_value with lv_value.
        check sy-subrc = 0.

        lr_node->set_value( lv_node_value  ).

      endloop.

    enddo.

  endmethod.

endclass.