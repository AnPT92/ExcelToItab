class ZCL_CA_DATA_TO_ITAB definition
  public
  final
  create public .

public section.

  types:
    BEGIN OF ty_decimal_map,
        fieldname TYPE fieldname,
        decimals  TYPE decimals,
      END OF ty_decimal_map .
  types:
    tt_decimal_map TYPE SORTED TABLE OF ty_decimal_map WITH UNIQUE KEY fieldname .

  constants MC_SOURCE_LOCAL type CHAR1 value 'L' ##NO_TEXT. "Local PC
  constants MC_SOURCE_SERVER type CHAR1 value 'S' ##NO_TEXT. "Server

  class-methods DATA_TO_ITAB
    importing
      !IV_FILENAME type CLIKE
      !IV_WORKSHEET type CLIKE optional
      !IV_BEGIN_ROW type I default 1
      !IV_BEGIN_COL type I default 1
      !IV_END_ROW type I optional
      !IV_END_COL type I optional
      !IV_MAPPING_LINE type I optional
      !IT_DECIMAL_MAP type TT_DECIMAL_MAP optional
      !IV_STRUCT type STRUKNAME optional
      !IV_DELIMETER type CHAR1 optional
    exporting
      !ET_ITAB type STANDARD TABLE
      !ET_RETURN type BAPIRET2_T
    exceptions
      ERROR .
  class-methods CSV_TO_ITAB
    importing
      !IV_FILENAME type CLIKE
      !IV_DELIMETER type CHAR1 default ','
      !IV_BEGIN_ROW type I default 1
      !IV_BEGIN_COL type I default 1
      !IV_END_ROW type I optional
      !IV_END_COL type I optional
      !IV_MAPPING_LINE type I optional
      !IT_DECIMAL_MAP type TT_DECIMAL_MAP optional
      !IV_STRUCT type STRUKNAME optional
    exporting
      !ET_ITAB type STANDARD TABLE
      !ET_RETURN type BAPIRET2_T
    exceptions
      ERROR .
  class-methods EXCEL_TO_ITAB
    importing
      !IV_FILENAME type CLIKE
      !IV_WORKSHEET type CLIKE optional
      !IV_SOURCE type CHAR1 default 'L'
      !IV_BEGIN_ROW type I default 1
      !IV_BEGIN_COL type I default 1
      !IV_END_ROW type I optional
      !IV_END_COL type I optional
      !IV_MAPPING_LINE type I optional
      !IT_DECIMAL_MAP type TT_DECIMAL_MAP optional
      !IV_STRUCT type STRUKNAME optional
    exporting
      !ET_ITAB type STANDARD TABLE
      !ET_RETURN type BAPIRET2_T
    exceptions
      ERROR .
  class-methods DOWNLOAD_EXCEL
    importing
      !IV_FORMNAME type CLIKE
      !IV_FILENAME type CLIKE optional
      !IV_CONTEXT_REF type ANY optional
    exceptions
      ERROR .
  class-methods GET_FILE_PATH
    importing
      !IV_FILTER type CLIKE default '*.xlsx|*.XLSX|*.csv|*.CSV|*.txt|*.TXT'
      !IV_LOCAL type FLAG default 'X'
      !IV_SERVER type FLAG default ''
    exporting
      !EV_FILE type ANY
    exceptions
      ERROR .
  class-methods FLUSH .
  class-methods TEXT_TO_MSGID
    importing
      !IV_MSG type CLIKE optional
      !IV_TABIX type SY-TABIX optional
      !IV_PASS_TO_SYSTEM type FLAG default 'X'
      !IV_MSGTYPE type SYST_MSGTY default 'E'
    exporting
      !ES_RETURN type BAPIRET2
      !EV_REMAIN type STRING .
PROTECTED SECTION.

  TYPES:
    BEGIN OF mty_field_attr,
      fieldname TYPE dd03l-fieldname,
      rollname  TYPE dd03l-rollname,
      domname   TYPE dd04l-domname,
      datatype  TYPE dd01l-datatype,
      convexit  TYPE dd01l-convexit,
      lowercase TYPE dd01l-lowercase,
      reffield  TYPE dd03l-reffield,
    END OF mty_field_attr .
  TYPES:
    mtt_field_attr TYPE SORTED TABLE OF mty_field_attr WITH UNIQUE KEY fieldname .

  CLASS-DATA mt_convexit TYPE mtt_field_attr .
  CLASS-DATA mt_uppercase TYPE mtt_field_attr .
  CLASS-DATA mt_cuky TYPE mtt_field_attr .
private section.

  class-data MV_FILENAME type STRING .
  class-data MT_WORKSHEET type STRINGTAB .
  class-data MV_HEADER type XSTRING .
  class-data MV_EXTENSION type CHAR10 .
  class-data MO_ITAB_UPL type ref to DATA .
  class-data MT_COMPONENTS type ABAP_COMPDESCR_TAB .
  class-data MO_ITAB_LINE_OUT type ref to DATA .
  class-data MT_RETURN type BAPIRET2_T .

  class-methods RAW_DATA_TO_INPUT
    importing
      !IV_BEGIN_ROW type I default 1
      !IV_BEGIN_COL type I default 1
      !IV_END_ROW type I optional
      !IV_END_COL type I optional
      !IV_MAPPING_LINE type I optional
      !IT_DECIMAL_MAP type TT_DECIMAL_MAP optional
      !IV_STRUCT type STRUKNAME optional
    exporting
      !ET_ITAB type STANDARD TABLE
    changing
      !CT_RAW_ITAB type STANDARD TABLE .
  class-methods CREATE_ITAB_UPLOAD
    importing
      !IT_ITAB type STANDARD TABLE
      !IT_DECIMAL_MAP type TT_DECIMAL_MAP optional
      !IV_STRUCT type STRUKNAME optional .
  class-methods CREATE_MAPPING_DATA
    importing
      !IT_ITAB type STANDARD TABLE
      !IV_MAPPING_LINE type I
    returning
      value(RT_MAP_ITAB) type STRINGTAB
    exceptions
      ERROR .
  class-methods GET_EXTENSION
    importing
      !IV_FILENAME type CLIKE .
  class-methods GET_SAVE_DIRECTORY
    importing
      !IV_FILENAME type STRING default 'Template'
    exporting
      value(EV_CANCEL) type FLAG
      value(EV_FULLPATH) type STRING .
  class-methods CONVERT_BY_COMPONENT
    importing
      !IS_COMP type ABAP_COMPDESCR
      !IV_TABIX type SY-TABIX optional
    changing
      !CV_VALUE type ANY .
  class-methods CONVERT_EXIT_INPUT
    importing
      !IV_TABIX type SY-TABIX optional
    changing
      !CS_ITAB_LINE type ANY .
ENDCLASS.



CLASS ZCL_CA_DATA_TO_ITAB IMPLEMENTATION.


  METHOD convert_by_component.
    DATA lo_decimal TYPE REF TO data.
    DATA lv_float   TYPE float.
    DATA lv_uzeit   TYPE uzeit.
    DATA lv_datum   TYPE datum.
    DATA lv_lengh   TYPE int4.
    DATA lo_elem    TYPE REF TO cl_abap_elemdescr.

    IF cv_value IS INITIAL.
      RETURN.
    ENDIF.

    CASE is_comp-type_kind.
      WHEN cl_abap_typedescr=>typekind_time.
        " Time > Accept SAP format (HHMMSS) or Windows format (HH:MM:SS)
        lv_lengh = strlen( cv_value ).
        IF lv_lengh = 6.
          "Use as it is
        ELSEIF lv_lengh = 8.
          cv_value = cv_value+0(2) && cv_value+3(2) && cv_value+6(2).
        ELSE.
          text_to_msgid( iv_msg   = |{ cv_value } is invalid Hours Format|
                         iv_tabix = iv_tabix ).
          CLEAR cv_value.
        ENDIF.

        IF cv_value IS NOT INITIAL.
          lv_uzeit = cv_value.
          CALL FUNCTION 'TIME_CHECK_PLAUSIBILITY'
            EXPORTING
              time                      = lv_uzeit
            EXCEPTIONS
              plausibility_check_failed = 1
              OTHERS                    = 2.
          IF sy-subrc <> 0.
            text_to_msgid( iv_tabix = iv_tabix ).
            CLEAR cv_value.
          ENDIF.
        ENDIF.


      WHEN cl_abap_typedescr=>typekind_date.
        " Date > Accept SAP format YYYYMMDD
        "               or Windows format DD.MM.YYYY or DD/MM/YYYY
        "               Add format YYYY.MM.DD or YYYY/MM/DD or YYYY-MM-DD
        lv_lengh = strlen( cv_value ).
        IF lv_lengh = 8.
          "Use as it is
        ELSEIF lv_lengh = 10.
          IF cv_value+0(4) CO '0123456789'.
            cv_value = cv_value+0(4) && cv_value+5(2) && cv_value+8(2).
          ELSE.
            cv_value = cv_value+6(4) && cv_value+3(2) && cv_value+0(2).
          ENDIF.
        ELSE.
          text_to_msgid( iv_msg   = |{ cv_value } is invalid Date Format|
                         iv_tabix = iv_tabix ).
          CLEAR cv_value.
        ENDIF.

        IF cv_value IS NOT INITIAL.
          lv_datum = cv_value.
          CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
            EXPORTING
              date                      = lv_datum
            EXCEPTIONS
              plausibility_check_failed = 1
              OTHERS                    = 2.
          IF sy-subrc <> 0.
            text_to_msgid( iv_tabix = iv_tabix ).
            CLEAR cv_value.
          ENDIF.
        ENDIF.

      WHEN cl_abap_typedescr=>typekind_packed.
        " Decimal number > Convert to float first to avoid Excel error
        TRY.
            lv_float = cv_value.
          CATCH cx_sy_conversion_no_number.
            text_to_msgid( iv_msg   = |{ cv_value } is invalid Number Format|
                           iv_tabix = iv_tabix ).
            CLEAR cv_value.
            RETURN.
        ENDTRY.
        ASSIGN COMPONENT is_comp-name OF STRUCTURE mo_itab_line_out->* TO FIELD-SYMBOL(<fs_value_out>).
        IF sy-subrc = 0.
          lo_elem ?= cl_abap_elemdescr=>describe_by_data( p_data = <fs_value_out> ).
          CREATE DATA lo_decimal TYPE HANDLE lo_elem.
          ASSIGN lo_decimal->* TO FIELD-SYMBOL(<fs_decimal>).
          <fs_decimal> = lv_float.
          cv_value = |{ <fs_decimal> }|.
        ENDIF.

      WHEN OTHERS.

    ENDCASE.
  ENDMETHOD.


  METHOD convert_exit_input.
    DATA lv_function        TYPE funcname.
    DATA lv_currency        TYPE waers.
    DATA lv_amount_external TYPE bapicurr_d.
    DATA lv_amount_internal TYPE dmbtr.
    DATA ls_return          TYPE bapireturn.

    DATA lv_eu_lname        TYPE eu_lname.
    DATA lt_func_parameter  TYPE abap_func_parmbind_tab.
    DATA ls_func_parameter  LIKE LINE OF lt_func_parameter.
    DATA lt_func_exception  TYPE abap_func_excpbind_tab.
    DATA ls_func_exception  LIKE LINE OF lt_func_exception.

    " --
    IF     mt_convexit  IS INITIAL
       AND mt_uppercase IS INITIAL
       AND mt_cuky      IS INITIAL.
      RETURN.
    ENDIF.

    " Convert Uppercase
    LOOP AT mt_uppercase INTO DATA(ls_uppercase).
      ASSIGN COMPONENT ls_uppercase-fieldname OF STRUCTURE cs_itab_line TO FIELD-SYMBOL(<fs_any>).
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      IF <fs_any> IS INITIAL.
        CONTINUE.
      ENDIF.

      IF     ls_uppercase-datatype  = 'CHAR'
         AND ls_uppercase-convexit IS INITIAL.
        CONDENSE <fs_any>.
        <fs_any> = to_upper( <fs_any> ).
      ENDIF.
    ENDLOOP.

    " Convert Input Function
    LOOP AT mt_convexit INTO DATA(ls_convexit).
      ASSIGN COMPONENT ls_convexit-fieldname OF STRUCTURE cs_itab_line TO <fs_any>.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      IF <fs_any> IS INITIAL.
        CONTINUE.
      ENDIF.

      IF ls_convexit-convexit IS NOT INITIAL.
        lv_function = 'CONVERSION_EXIT_' && ls_convexit-convexit && '_INPUT'.
        SELECT SINGLE funcname
          FROM v_fdir
          WHERE funcname  = @lv_function
          AND   generated = ''
          AND   active    = 'X'
        INTO @DATA(lv_funcname).
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        lv_eu_lname = lv_funcname.
        cl_fb_function_utility=>meth_get_interface(
          EXPORTING
            im_name             = lv_eu_lname
          IMPORTING
            ex_interface        = DATA(ls_interface)
          EXCEPTIONS
            error_occured       = 1
            object_not_existing = 2
            OTHERS              = 3
        ).
        IF sy-subrc <> 0.
          CONTINUE.
        ENDIF.

        REFRESH lt_func_parameter.
        REFRESH lt_func_exception.

        LOOP AT ls_interface-import INTO DATA(ls_import).
          IF ls_import-parameter CS 'INPUT'.
            IF ls_import-structure IS NOT INITIAL.
              CREATE DATA ls_func_parameter-value TYPE (ls_import-structure).
            ELSE.
              CREATE DATA ls_func_parameter-value LIKE <fs_any>.
            ENDIF.
            ls_func_parameter-name      = ls_import-parameter.
            ls_func_parameter-kind      = abap_func_exporting.
            ls_func_parameter-value->*  = <fs_any>.
            INSERT ls_func_parameter INTO TABLE lt_func_parameter.
          ENDIF.
        ENDLOOP.
        LOOP AT ls_interface-export INTO DATA(ls_export).
          IF ls_export-parameter CS 'OUTPUT'.
            IF ls_export-structure IS NOT INITIAL.
              CREATE DATA ls_func_parameter-value TYPE (ls_export-structure).
            ELSE.
              CREATE DATA ls_func_parameter-value TYPE (ls_convexit-domname).
            ENDIF.
            ls_func_parameter-name   = ls_export-parameter.
            ls_func_parameter-kind   = abap_func_importing.
            INSERT ls_func_parameter INTO TABLE lt_func_parameter.
          ENDIF.
        ENDLOOP.
        LOOP AT ls_interface-except INTO DATA(ls_except).
          ls_func_exception-value = sy-tabix.
          ls_func_exception-name  = ls_except-parameter.
          INSERT ls_func_exception INTO TABLE lt_func_exception.
        ENDLOOP.

        TRY.
            CALL FUNCTION lv_funcname
              PARAMETER-TABLE lt_func_parameter
              EXCEPTION-TABLE lt_func_exception.
            IF sy-subrc <> 0.
              text_to_msgid( iv_tabix = iv_tabix ).
              CLEAR <fs_any>.
              CONTINUE.
            ENDIF.
          CATCH cx_sy_dyn_call_illegal_func.
            CLEAR <fs_any>.
            CONTINUE.
        ENDTRY.

        LOOP AT lt_func_parameter INTO ls_func_parameter.
          IF ls_func_parameter-name CS 'OUTPUT'.
            <fs_any> = ls_func_parameter-value->*.
          ENDIF.
        ENDLOOP.
      ENDIF.
      UNASSIGN <fs_any>.
    ENDLOOP.

    " Convert for amount field
    LOOP AT mt_cuky INTO DATA(ls_cuky).
      ASSIGN COMPONENT ls_cuky-fieldname OF STRUCTURE cs_itab_line TO <fs_any>.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.

      IF sy-subrc = 0.
        ASSIGN COMPONENT ls_cuky-reffield OF STRUCTURE cs_itab_line TO FIELD-SYMBOL(<fs_waers>).
        IF sy-subrc = 0.
          IF <fs_waers> IS NOT INITIAL.
            lv_currency        = <fs_waers>.
            lv_amount_external = <fs_any>.
            CALL FUNCTION 'BAPI_CURRENCY_CONV_TO_INTERNAL'
              EXPORTING
                currency             = lv_currency
                amount_external      = lv_amount_external
                max_number_of_digits = 23     " Max
              IMPORTING
                amount_internal      = lv_amount_internal
                return               = ls_return.
            IF ls_return-type = 'E'.
              sy-msgid = ls_return-code       .
              sy-msgno = ls_return-log_no     .
              sy-msgv1 = ls_return-message_v1 .
              sy-msgv2 = ls_return-message_v2 .
              sy-msgv3 = ls_return-message_v3 .
              sy-msgv4 = ls_return-message_v4 .
              text_to_msgid( iv_tabix = iv_tabix ).
            ENDIF.
            <fs_any> = lv_amount_internal.
            UNASSIGN <fs_waers>.
          ENDIF.
        ENDIF.
      ENDIF.

      UNASSIGN <fs_any>.
    ENDLOOP.

  ENDMETHOD.


  METHOD csv_to_itab.
    DATA lt_data_tab TYPE stringtab.
    DATA lt_value    TYPE stringtab.
    DATA lv_numc6    TYPE posnr_vl.
    DATA lv_index    TYPE sy-index.

    DATA lo_struct    TYPE REF TO cl_abap_structdescr.
    DATA lo_table     TYPE REF TO cl_abap_tabledescr.
    DATA lt_comp_dyn  TYPE abap_component_tab.
    DATA ls_comp_dyn  LIKE LINE OF lt_comp_dyn.
    DATA lr_csv_data  TYPE REF TO data.

    FIELD-SYMBOLS <lt_data_raw> TYPE STANDARD TABLE.
    REFRESH et_itab.
    REFRESH et_return.
    REFRESH mt_return.

    " ---------------------------------------------------------------------
    IF iv_filename IS INITIAL.
      text_to_msgid( iv_msg = 'File Name is required' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.

    " ---------------------------------------------------------------------
    " Check to avoid Upload Multiple Sheets with the same File
    IF mv_filename <> iv_filename.
      mv_filename = iv_filename.

      " Check File Extension
      IF mv_extension IS INITIAL.
        get_extension( iv_filename = iv_filename ).
      ENDIF.
      IF  mv_extension <> 'CSV'
      AND mv_extension <> 'TXT'.
        text_to_msgid( iv_msg = 'Wrong File Extension' ).
        et_return = mt_return.
        RAISE error.
      ENDIF.

      " ---------------------------------------------------------------------
      " Upload Excel
      cl_gui_frontend_services=>gui_upload( EXPORTING  filename                = mv_filename      " Name of file
                                                       filetype                = 'ASC'            " File Type (ASC or BIN)
                                            IMPORTING  header                  = mv_header        " File Header in Case of Binary Upload
                                            CHANGING   data_tab                = lt_data_tab      " Transfer table for file contents
                                            EXCEPTIONS file_open_error         = 1                " File does not exist and cannot be opened
                                                       file_read_error         = 2                " Error when reading file
                                                       no_batch                = 3                " Front-End Function Cannot Be Executed in Backgrnd
                                                       gui_refuse_filetransfer = 4                " Incorrect front end or error on front end
                                                       invalid_type            = 5                " Incorrect parameter FILETYPE
                                                       no_authority            = 6                " No Authorization for Upload
                                                       unknown_error           = 7
                                                       bad_data_format         = 8                " Cannot Interpret Data in File
                                                       header_not_allowed      = 9                " Invalid header
                                                       separator_not_allowed   = 10               " Invalid separator
                                                       header_too_long         = 11               " The header information is limited to 1023 bytes at present
                                                       unknown_dp_error        = 12               " Error when calling data provider
                                                       access_denied           = 13               " Access to File Denied
                                                       dp_out_of_memory        = 14               " Not Enough Memory in Data Provider
                                                       disk_full               = 15               " Storage Medium full
                                                       dp_timeout              = 16               " Timeout of Data Provider
                                                       OTHERS                  = 17 ).
      IF sy-subrc <> 0.
        text_to_msgid( ).
        et_return = mt_return.
        RAISE error.
      ENDIF.
    ENDIF.

    " Check First Line
    READ TABLE lt_data_tab INTO DATA(ls_first) INDEX 1.
    IF sy-subrc <> 0.
      text_to_msgid( iv_msg = 'No Data' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.

    " Get Number of Column for Output
    REFRESH lt_value.
    CALL FUNCTION 'RSDS_CONVERT_CSV'
      EXPORTING
        i_data_sep       = iv_delimeter
        i_esc_char       = '"'
        i_record         = ls_first
        i_field_count    = 9999
      IMPORTING
        e_t_data         = lt_value
      EXCEPTIONS
        escape_no_close  = 1
        escape_improper  = 2
        conversion_error = 3
        OTHERS           = 4.
    IF sy-subrc <> 0.
      text_to_msgid( ).
      et_return = mt_return.
      RAISE error.
    ENDIF.
    IF lt_value IS INITIAL.
      text_to_msgid( iv_msg = 'No Data' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.

    "Create dynamic table to contain data
    DATA(lv_lines) = lines( lt_value ).
    REFRESH lt_comp_dyn.
    CLEAR lv_numc6.
    DO lv_lines TIMES.
      lv_numc6 += 1.
      ls_comp_dyn-name =  'COL_' && lv_numc6.
      ls_comp_dyn-type ?= cl_abap_elemdescr=>describe_by_name( 'DSTRING' ).
      APPEND ls_comp_dyn TO lt_comp_dyn.
      CLEAR ls_comp_dyn.
    ENDDO.
    lo_struct = cl_abap_structdescr=>create( p_components = lt_comp_dyn ).
    lo_table  = cl_abap_tabledescr=>create( lo_struct ).
    CREATE DATA lr_csv_data TYPE HANDLE lo_table.
    ASSIGN lr_csv_data->* TO <lt_data_raw>.
    LOOP AT lt_data_tab INTO DATA(ls_data_tab).
      REFRESH lt_value.
      CALL FUNCTION 'RSDS_CONVERT_CSV'
        EXPORTING
          i_data_sep       = iv_delimeter
          i_esc_char       = '"'
          i_record         = ls_data_tab
          i_field_count    = 9999
        IMPORTING
          e_t_data         = lt_value
        EXCEPTIONS
          escape_no_close  = 1
          escape_improper  = 2
          conversion_error = 3
          OTHERS           = 4.
      IF sy-subrc <> 0.
        text_to_msgid( ).
        et_return = mt_return.
        RAISE error.
      ENDIF.

      CLEAR lv_index.
      APPEND INITIAL LINE TO <lt_data_raw> ASSIGNING FIELD-SYMBOL(<ls_data_raw>).
      WHILE 1 = 1.    "Will exit if not assign
        lv_index += 1.
        ASSIGN COMPONENT lv_index OF STRUCTURE <ls_data_raw> TO FIELD-SYMBOL(<fs_value>).
        READ TABLE lt_value INTO DATA(ls_value) INDEX lv_index.
        IF sy-subrc <> 0
        OR <fs_value> IS NOT ASSIGNED.
          EXIT.
        ENDIF.

        <fs_value> = ls_value.
        UNASSIGN <fs_value>.
        CLEAR ls_value.
      ENDWHILE.
    ENDLOOP.

**********************************************************************
    raw_data_to_input(
      EXPORTING
        iv_begin_row    = iv_begin_row
        iv_begin_col    = iv_begin_col
        iv_end_row      = iv_end_row
        iv_end_col      = iv_end_col
        iv_mapping_line = iv_mapping_line
        it_decimal_map  = it_decimal_map
        iv_struct       = iv_struct
      IMPORTING
        et_itab         = et_itab
      CHANGING
        ct_raw_itab     = <lt_data_raw> ).
    et_return = mt_return.

    IF line_exists( et_return[ type = 'A' ] )
    OR line_exists( et_return[ type = 'E' ] ).
      RAISE error.
    ENDIF.

  ENDMETHOD.


  METHOD data_to_itab.

    REFRESH et_itab.
    IF iv_filename IS INITIAL.
      text_to_msgid( iv_msg = 'File Name is required' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.

    " Check File Extension
    get_extension( iv_filename = iv_filename ).

    CASE mv_extension.
      WHEN 'CSV' OR 'TXT'.
        csv_to_itab( EXPORTING  iv_filename     = iv_filename
                                iv_delimeter    = iv_delimeter
                                iv_begin_row    = iv_begin_row
                                iv_begin_col    = iv_begin_col
                                iv_end_row      = iv_end_row
                                iv_end_col      = iv_end_col
                                iv_mapping_line = iv_mapping_line
                                it_decimal_map  = it_decimal_map
                                iv_struct       = iv_struct
                     IMPORTING  et_itab         = et_itab
                                et_return       = et_return
                     EXCEPTIONS error           = 1 ).
      WHEN 'XLSX'.
        excel_to_itab( EXPORTING  iv_filename     = iv_filename
                                  iv_worksheet    = iv_worksheet
                                  iv_begin_row    = iv_begin_row
                                  iv_begin_col    = iv_begin_col
                                  iv_end_row      = iv_end_row
                                  iv_end_col      = iv_end_col
                                  iv_mapping_line = iv_mapping_line
                                  it_decimal_map  = it_decimal_map
                                  iv_struct       = iv_struct
                       IMPORTING  et_itab         = et_itab
                                  et_return       = et_return
                       EXCEPTIONS error           = 1 ).
      WHEN OTHERS.
        text_to_msgid( iv_msg = 'Invalid File Extension' ).
        sy-subrc = 0.
    ENDCASE.
    IF sy-subrc <> 0.
      et_return = mt_return.
    ENDIF.

  ENDMETHOD.


  METHOD download_excel.
    DATA lv_filename TYPE string.
    DATA lv_fullpath TYPE string.
    DATA lv_cancel   TYPE flag.
    DATA lv_dummy    TYPE dummy.

    " --
    lv_filename = iv_filename.
    IF lv_filename IS INITIAL.
      lv_filename = sy-tcode.
    ENDIF.

    " --
    IF iv_context_ref IS SUPPLIED.
      ASSIGN iv_context_ref TO FIELD-SYMBOL(<fs_context_ref>).
    ELSE.
      ASSIGN lv_dummy TO <fs_context_ref>.
    ENDIF.

    " --
    get_save_directory( EXPORTING iv_filename = lv_filename
                        IMPORTING ev_cancel   = lv_cancel
                                  ev_fullpath = lv_fullpath ).
    IF lv_cancel IS NOT INITIAL.
      RETURN.
    ENDIF.

    " Call Excel Workbench Function Module
    TRY.  " Avoid dump Class
        CALL FUNCTION 'ZXLWB_CALLFORM'
          EXPORTING
            iv_formname        = iv_formname
            iv_context_ref     = <fs_context_ref>
            iv_viewer_inplace  = ''
            iv_save_as         = lv_fullpath
            iv_viewer_suppress = 'X'
          EXCEPTIONS
            process_terminated = 1
            OTHERS             = 2.
        IF sy-subrc <> 0.
          MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                  WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
        ENDIF.
      CATCH cx_sy_dyn_call_illegal_func.
        MESSAGE 'Excel Workbench is required' TYPE 'E' RAISING error.
      CATCH cx_sy_dyn_call_param_missing.
        MESSAGE 'Parameter for Excel Workbench is missing' TYPE 'E' RAISING error.
      CATCH cx_sy_dyn_call_illegal_type.
        MESSAGE 'Parameter for Excel Workbench is illegal type' TYPE 'E' RAISING error.
    ENDTRY.
  ENDMETHOD.


  METHOD excel_to_itab.
    DATA lv_filelength TYPE i.
    DATA lt_data_tab   TYPE solix_tab.
    FIELD-SYMBOLS <lt_data_raw> TYPE STANDARD TABLE.

    " ---------------------------------------------------------------------
    REFRESH et_itab.
    REFRESH et_return.
    REFRESH mt_return.

    " ---------------------------------------------------------------------
    IF iv_filename IS INITIAL.
      text_to_msgid( iv_msg = 'File Name is required' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.

    " ---------------------------------------------------------------------
    IF mv_filename <> iv_filename.
      mv_filename = iv_filename.

      " Check File Extension
      IF mv_extension IS INITIAL.
        get_extension( iv_filename = iv_filename ).
      ENDIF.
      IF  mv_extension <> 'XLSX'.
        text_to_msgid( iv_msg = 'Wrong File Extension' ).
        et_return = mt_return.
        RAISE error.
      ENDIF.

      IF iv_source = mc_source_local.
        " Upload Excel from Local
        cl_gui_frontend_services=>gui_upload( EXPORTING  filename                = iv_filename      " Name of file
                                                         filetype                = 'BIN'            " File Type (ASC or BIN)
                                              IMPORTING  filelength              = lv_filelength    " File Length
                                                         header                  = mv_header        " File Header in Case of Binary Upload
                                              CHANGING   data_tab                = lt_data_tab      " Transfer table for file contents
                                              EXCEPTIONS file_open_error         = 1                " File does not exist and cannot be opened
                                                         file_read_error         = 2                " Error when reading file
                                                         no_batch                = 3                " Front-End Function Cannot Be Executed in Backgrnd
                                                         gui_refuse_filetransfer = 4                " Incorrect front end or error on front end
                                                         invalid_type            = 5                " Incorrect parameter FILETYPE
                                                         no_authority            = 6                " No Authorization for Upload
                                                         unknown_error           = 7
                                                         bad_data_format         = 8                " Cannot Interpret Data in File
                                                         header_not_allowed      = 9                " Invalid header
                                                         separator_not_allowed   = 10               " Invalid separator
                                                         header_too_long         = 11               " The header information is limited to 1023 bytes at present
                                                         unknown_dp_error        = 12               " Error when calling data provider
                                                         access_denied           = 13               " Access to File Denied
                                                         dp_out_of_memory        = 14               " Not Enough Memory in Data Provider
                                                         disk_full               = 15               " Storage Medium full
                                                         dp_timeout              = 16               " Timeout of Data Provider
                                                         OTHERS                  = 17 ).
        IF sy-subrc <> 0.
          text_to_msgid( ).
          et_return = mt_return.
          RAISE error.
        ENDIF.

        " Convert to xstring
        CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
          EXPORTING
            input_length = lv_filelength
          IMPORTING
            buffer       = mv_header
          TABLES
            binary_tab   = lt_data_tab
          EXCEPTIONS
            failed       = 1
            OTHERS       = 2.
        IF sy-subrc <> 0.
          text_to_msgid( ).
          et_return = mt_return.
          RAISE error.
        ENDIF.

      ELSEIF iv_source = mc_source_server.
        " Get Excel from Application Server
        OPEN DATASET mv_filename FOR INPUT IN BINARY MODE.
        IF sy-subrc <> 0.
          text_to_msgid( iv_msg = 'Invalid File Directory' ).
          et_return = mt_return.
          RAISE error.
        ELSE.
          READ DATASET mv_filename INTO mv_header.
          IF sy-subrc <> 0.
            CLEAR mv_header.
            text_to_msgid( iv_msg = 'Error Reading File' ).
            et_return = mt_return.
            RAISE error.
          ENDIF.
          CLOSE DATASET mv_filename.
        ENDIF.

      ELSE.
        text_to_msgid( iv_msg = 'Invalid Source Parameter for class Excel' ).
        et_return = mt_return.
        RAISE error.
      ENDIF.
    ENDIF.

    " Get List of Worksheets
    TRY.
        DATA(lr_excel_data) = NEW cl_fdt_xl_spreadsheet( document_name = iv_filename
                                                         xdocument     = mv_header ).
      CATCH cx_fdt_excel_core.
        text_to_msgid( iv_msg = 'No Input Found' ).
        et_return = mt_return.
        RAISE error.
    ENDTRY.
    lr_excel_data->if_fdt_doc_spreadsheet~get_worksheet_names( IMPORTING worksheet_names = mt_worksheet ).

    " Get Sheet data
    IF iv_worksheet IS INITIAL.
      READ TABLE mt_worksheet INTO DATA(ls_worksheet) INDEX 1.
    ELSE.
      READ TABLE mt_worksheet INTO ls_worksheet
           WITH TABLE KEY table_line = iv_worksheet.
    ENDIF.
    IF ls_worksheet IS INITIAL.
      text_to_msgid( iv_msg = 'No Sheet Found' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.

    " Read the Upload Sheet
    TRY.
        DATA(lo_data_ref) = lr_excel_data->if_fdt_doc_spreadsheet~get_itab_from_worksheet( ls_worksheet ).
      CATCH cx_root INTO DATA(ls_exception).
        DATA(lv_msg) = ls_exception->get_text( ).
        text_to_msgid( iv_msg = lv_msg ).
        et_return = mt_return.
        RAISE error.
    ENDTRY.
    IF lo_data_ref IS NOT BOUND.
      text_to_msgid( iv_msg = 'Missing dimension. Try open Excel and save to generate before uploading' ).
      et_return = mt_return.
      RAISE error.
    ENDIF.
    ASSIGN lo_data_ref->* TO <lt_data_raw>.

**********************************************************************
    raw_data_to_input(
      EXPORTING
        iv_begin_row    = iv_begin_row
        iv_begin_col    = iv_begin_col
        iv_end_row      = iv_end_row
        iv_end_col      = iv_end_col
        iv_mapping_line = iv_mapping_line
        it_decimal_map  = it_decimal_map
        iv_struct       = iv_struct
      IMPORTING
        et_itab         = et_itab
      CHANGING
        ct_raw_itab     = <lt_data_raw> ).
    et_return = mt_return.

    IF line_exists( et_return[ type = 'A' ] )
    OR line_exists( et_return[ type = 'E' ] ).
      LOOP AT et_return INTO DATA(ls_return)
           WHERE type CA 'AE'.
        EXIT.
      ENDLOOP.
      MESSAGE ID ls_return-id TYPE 'E' NUMBER ls_return-number
          WITH  ls_return-message_v1
                ls_return-message_v2
                ls_return-message_v3
                ls_return-message_v4
          RAISING error.
    ENDIF.

  ENDMETHOD.


  METHOD flush.

    FREE mo_itab_upl.
    FREE mo_itab_line_out.

    CLEAR mv_filename.
    CLEAR mv_header.
    CLEAR mv_extension.

    REFRESH mt_worksheet.
    REFRESH mt_components.
    REFRESH mt_convexit.
    REFRESH mt_cuky.

  ENDMETHOD.


  METHOD get_extension.
    DATA lv_filename TYPE char1024.

    lv_filename = iv_filename.
    CALL FUNCTION 'TRINT_FILE_GET_EXTENSION'
      EXPORTING
        filename  = lv_filename
        uppercase = 'X'
      IMPORTING
        extension = mv_extension.
  ENDMETHOD.


  METHOD get_file_path.
    DATA lt_filetable   TYPE filetable.
    DATA lv_rc          TYPE i.
    DATA lv_directory   TYPE string.
    DATA lv_user_action TYPE i.

    IF ev_file IS NOT SUPPLIED.
      MESSAGE 'No Output Set Up' TYPE 'E' RAISING error.
    ENDIF.

    IF     iv_local  IS INITIAL
       AND iv_server IS INITIAL.
      MESSAGE 'Invalid Action' TYPE 'E' RAISING error.

    ELSEIF iv_local IS NOT INITIAL.
      cl_gui_frontend_services=>file_open_dialog( EXPORTING  file_filter             = iv_filter
                                                  CHANGING   file_table              = lt_filetable
                                                             rc                      = lv_rc
                                                             user_action             = lv_user_action
                                                  EXCEPTIONS file_open_dialog_failed = 1
                                                             cntl_error              = 2
                                                             error_no_gui            = 3
                                                             not_supported_by_gui    = 4
                                                             OTHERS                  = 5 ).
      IF sy-subrc = 0.
        READ TABLE lt_filetable INTO DATA(ls_filetable) INDEX 1.
        IF sy-subrc = 0.
          ev_file = ls_filetable-filename.
        ENDIF.
      ELSE.
        CLEAR ev_file.
        MESSAGE ID sy-msgid TYPE 'S' NUMBER sy-msgno
                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
                DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
      IF lv_user_action = 9.
        CLEAR ev_file.
        MESSAGE 'Cancelled' TYPE 'S' DISPLAY LIKE 'W'.
      ENDIF.

    ELSEIF iv_server IS NOT INITIAL.
      lv_directory = '/usr/sap/' && sy-sysid && '/D00'.
      CALL FUNCTION '/SAPDMC/LSM_F4_SERVER_FILE'
        EXPORTING
          directory        = lv_directory
          filemask         = '*'
        IMPORTING
          serverfile       = ev_file
        EXCEPTIONS
          canceled_by_user = 1
          OTHERS           = 2.
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE 'S' NUMBER sy-msgno
                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
                DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD get_save_directory.
    DATA lv_filename    TYPE string.
    DATA lv_directory   TYPE string.
    DATA lv_path        TYPE string.
    DATA lv_user_action TYPE i.

    lv_filename = iv_filename.
    IMPORT lv_path TO lv_path FROM MEMORY ID 'ZCL_EXCEL_DIRECTORY'.
    IF lv_path IS INITIAL.    " If empty = Default Desktop Directory
      cl_gui_frontend_services=>get_desktop_directory( CHANGING   desktop_directory    = lv_directory
                                                       EXCEPTIONS cntl_error           = 1                " Control error
                                                                  error_no_gui         = 2                " No GUI available
                                                                  not_supported_by_gui = 3                " GUI does not support this
                                                                  OTHERS               = 4 ).
      IF sy-subrc = 0.
        cl_gui_cfw=>update_view( ).
      ELSE.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.
    ENDIF.

    " --
    cl_gui_frontend_services=>file_save_dialog( EXPORTING  default_extension         = 'XLSX'                 " Default Extension
                                                           default_file_name         = lv_filename            " Default File Name
                                                           file_filter               = 'XLSX'                 " File Type Filter Table
                                                           initial_directory         = lv_directory           " Initial Directory
                                                CHANGING   filename                  = lv_filename            " File Name to Save
                                                           path                      = lv_path                " Path to File
                                                           fullpath                  = ev_fullpath            " Path + File Name
                                                           user_action               = lv_user_action
                                                EXCEPTIONS cntl_error                = 1                " Control error
                                                           error_no_gui              = 2                " No GUI available
                                                           not_supported_by_gui      = 3                " GUI does not support this
                                                           invalid_default_file_name = 4                " Invalid default file name
                                                           OTHERS                    = 5 ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE 'S' NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
              DISPLAY LIKE 'E'.
    ENDIF.
    IF lv_path IS NOT INITIAL.
      EXPORT lv_path FROM lv_path TO MEMORY ID 'ZCL_EXCEL_DIRECTORY'.
    ENDIF.
    IF lv_user_action = 9.
      MESSAGE 'Cancelled' TYPE 'W'.
      ev_cancel = 'X'.
    ENDIF.
  ENDMETHOD.


  METHOD create_itab_upload.

    DATA lo_struct    TYPE REF TO cl_abap_structdescr.
    DATA lo_table     TYPE REF TO cl_abap_tabledescr.
    DATA lt_comp_dyn  TYPE abap_component_tab.
    DATA ls_comp_dyn  LIKE LINE OF lt_comp_dyn.

**********************************************************************
    " Get Number of Column for Output
    lo_table  ?= cl_abap_tabledescr=>describe_by_data( p_data = it_itab ).
    lo_struct ?= lo_table->get_table_line_type( ).
    CREATE DATA mo_itab_line_out TYPE HANDLE lo_struct.
    mt_components = lo_struct->components.
    IF it_decimal_map IS NOT INITIAL.
      LOOP AT mt_components ASSIGNING FIELD-SYMBOL(<ls_components>).
        READ TABLE it_decimal_map INTO DATA(ls_decimal_map) BINARY SEARCH
             WITH KEY fieldname = <ls_components>-name.
        IF sy-subrc = 0.
          <ls_components>-decimals = ls_decimal_map-decimals.
        ENDIF.
      ENDLOOP.
    ENDIF.

**********************************************************************
    IF iv_struct IS NOT INITIAL.
      SELECT dd03l~fieldname,
             dd03l~rollname,
             dd04l~domname,
             dd01l~datatype,
             dd01l~convexit,
             dd01l~lowercase
        FROM dd03l
             INNER JOIN dd04l
               ON  dd04l~rollname = dd03l~rollname
               AND dd04l~as4local = dd03l~as4local
             INNER JOIN dd01l
               ON  dd01l~domname  = dd04l~domname
               AND dd01l~as4local = dd03l~as4local
        WHERE dd03l~tabname   = @iv_struct
          AND dd03l~as4local  = 'A'     "active
          AND dd01l~convexit <> @space
          AND dd01l~datatype <> 'CURR'  "Do not convert input for Currency
      INTO CORRESPONDING FIELDS OF TABLE @mt_convexit.

      " --
      SELECT dd03l~fieldname,
             dd03l~rollname,
             dd04l~domname,
             dd01l~datatype,
             dd01l~convexit,
             dd01l~lowercase
        FROM dd03l
             INNER JOIN dd04l
               ON  dd04l~rollname = dd03l~rollname
               AND dd04l~as4local = dd03l~as4local
             INNER JOIN dd01l
               ON  dd01l~domname  = dd04l~domname
               AND dd01l~as4local = dd03l~as4local
        WHERE dd03l~tabname   = @iv_struct
          AND dd03l~as4local  = 'A'    " active
          AND dd01l~lowercase = @space
          AND dd01l~datatype  = 'CHAR'
      INTO CORRESPONDING FIELDS OF TABLE @mt_uppercase.

      " --
      SELECT fieldname,
             reffield
        FROM dd03l
        WHERE datatype = 'CURR'
          AND tabname  = @iv_struct
          AND as4local = 'A'
      INTO CORRESPONDING FIELDS OF TABLE @mt_cuky.
    ENDIF.

**********************************************************************
    " Create output table with all fields are string type
    REFRESH lt_comp_dyn.
    LOOP AT mt_components INTO DATA(ls_components).
      ls_comp_dyn-name =  ls_components-name.
      ls_comp_dyn-type ?= cl_abap_elemdescr=>describe_by_name( 'DSTRING' ).
      APPEND ls_comp_dyn TO lt_comp_dyn.
      CLEAR ls_comp_dyn.
    ENDLOOP.
    lo_struct = cl_abap_structdescr=>create( p_components = lt_comp_dyn ).
    lo_table  = cl_abap_tabledescr=>create( lo_struct ).
    IF mo_itab_upl IS BOUND.
      FREE mo_itab_upl.
    ENDIF.
    CREATE DATA mo_itab_upl TYPE HANDLE lo_table.

  ENDMETHOD.


  METHOD create_mapping_data.

    DATA lv_index_in TYPE sy-tabix.

    READ TABLE it_itab ASSIGNING FIELD-SYMBOL(<ls_data>)
         INDEX iv_mapping_line.
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    CLEAR   lv_index_in .
    REFRESH rt_map_itab  .
    WHILE 1 = 1.    "Will exit if not assign
      lv_index_in += 1.
      ASSIGN COMPONENT lv_index_in OF STRUCTURE <ls_data> TO FIELD-SYMBOL(<fs_value>).
      IF sy-subrc = 0.
        APPEND INITIAL LINE TO rt_map_itab ASSIGNING FIELD-SYMBOL(<ls_map_tab>).
        <ls_map_tab> = <fs_value>.
        CONDENSE <ls_map_tab>.
        UNASSIGN <fs_value>.
      ELSE.
        EXIT.
      ENDIF.
    ENDWHILE.

  ENDMETHOD.


  METHOD raw_data_to_input.
    DATA lv_index_in   TYPE sy-index.
    DATA lv_index_out  TYPE sy-index.
    FIELD-SYMBOLS <lt_data_upl> TYPE STANDARD TABLE.

    " Read Mapping Line
    IF iv_mapping_line IS NOT INITIAL.
      DATA(lt_map_tab) = create_mapping_data( it_itab         = ct_raw_itab
                                              iv_mapping_line = iv_mapping_line ).
      IF lt_map_tab IS INITIAL.
        text_to_msgid( iv_msg = 'No Mapping Line Found' ).
        RETURN.
      ENDIF.
    ENDIF.

    " Create Static Object for Data
    create_itab_upload( it_itab        = et_itab
                        it_decimal_map = it_decimal_map
                        iv_struct      = iv_struct ).
    ASSIGN mo_itab_upl->* TO <lt_data_upl>.
    IF sy-subrc <> 0.
      text_to_msgid( iv_msg = 'Error generating Upload Data' ).
      RETURN.
    ENDIF.

    " Line from - to
    IF iv_end_row IS NOT INITIAL.
      DELETE ct_raw_itab FROM iv_end_row.
    ENDIF.
    IF iv_begin_row > 1.
      DELETE ct_raw_itab FROM 1 TO ( iv_begin_row - 1 ).
    ENDIF.

    " Transfer data
    LOOP AT ct_raw_itab ASSIGNING FIELD-SYMBOL(<ls_data_raw>).
      DATA(lv_row) = sy-tabix.

      APPEND INITIAL LINE TO <lt_data_upl> ASSIGNING FIELD-SYMBOL(<ls_data_upl>).

      IF lt_map_tab IS INITIAL.
        CLEAR lv_index_out.
        CLEAR lv_index_in.
        WHILE 1 = 1.    "Will exit if not assign
          lv_index_out += 1.
          ASSIGN COMPONENT lv_index_out OF STRUCTURE <ls_data_upl> TO FIELD-SYMBOL(<fs_val_out>).
          lv_index_in = lv_index_out + iv_begin_col - 1.
          ASSIGN COMPONENT lv_index_in OF STRUCTURE <ls_data_raw> TO FIELD-SYMBOL(<fs_val_in>).
          IF <fs_val_in>  IS NOT ASSIGNED
          OR <fs_val_out> IS NOT ASSIGNED.
            EXIT.
          ENDIF.

          " Convert by Component
          READ TABLE mt_components INTO DATA(ls_components)
               INDEX lv_index_out.
          IF sy-subrc = 0.
            convert_by_component( EXPORTING is_comp  = ls_components
                                            iv_tabix = lv_row
                                  CHANGING  cv_value = <fs_val_in> ).
          ENDIF.

          <fs_val_out> = <fs_val_in>.
          UNASSIGN <fs_val_in>.
          UNASSIGN <fs_val_out>.
        ENDWHILE.
      ELSE.
        LOOP AT lt_map_tab INTO DATA(ls_map_tab).
          lv_index_in = sy-tabix.
          ASSIGN COMPONENT lv_index_in OF STRUCTURE <ls_data_raw> TO <fs_val_in>.
          ASSIGN COMPONENT ls_map_tab  OF STRUCTURE <ls_data_upl> TO <fs_val_out>.
          IF NOT (     <fs_val_in>  IS ASSIGNED
                   AND <fs_val_out> IS ASSIGNED ).
            CONTINUE.
          ENDIF.

          " Convert by Component
          READ TABLE mt_components INTO ls_components
               WITH KEY name = ls_map_tab.
          IF sy-subrc = 0.
            convert_by_component( EXPORTING is_comp  = ls_components
                                            iv_tabix = lv_row
                                  CHANGING  cv_value = <fs_val_in> ).
          ENDIF.

          <fs_val_out> = <fs_val_in>.
          UNASSIGN <fs_val_in>.
          UNASSIGN <fs_val_out>.
        ENDLOOP.
      ENDIF.

      " Convert input
      convert_exit_input( EXPORTING iv_tabix     = lv_row
                          CHANGING  cs_itab_line = <ls_data_upl> ).

      UNASSIGN <ls_data_upl>.
    ENDLOOP.

**********************************************************************
    TRY.
        MOVE-CORRESPONDING <lt_data_upl> TO et_itab.
      CATCH  cx_sy_conversion_no_number.
        text_to_msgid( iv_msg = 'No Mapping Line Found' ).
    ENDTRY.

  ENDMETHOD.


  METHOD text_to_msgid.

    DATA lv_msg       TYPE text1024.
    DATA lv_line(160) TYPE c.
    DATA lt_out_lines TYPE TABLE OF text40.
    REFRESH lt_out_lines.
    CLEAR es_return.

    IF iv_msg IS SUPPLIED.
      IF iv_msg IS NOT INITIAL.
        lv_msg = iv_msg.
        CALL FUNCTION 'TEXT_SPLIT'
          EXPORTING
            length = 160
            text   = lv_msg
          IMPORTING
            line   = lv_line
            rest   = ev_remain.
        IF ev_remain IS NOT INITIAL.
          text_to_msgid( iv_msg     = `Input text is too long`
                         iv_tabix   = iv_tabix
                         iv_msgtype = 'W' ).  "Warning only
        ENDIF.

        CALL FUNCTION 'RKD_WORD_WRAP'
          EXPORTING
            textline            = lv_line             " Source text line
            delimiter           = space               " Indicator, which is used as a separator
            outputlen           = '40'                " Maximum output line width
          TABLES
            out_lines           = lt_out_lines        " All output lines as table
          EXCEPTIONS
            outputlen_too_large = 1
            OTHERS              = 2.
        IF sy-subrc <> 0.
          text_to_msgid( iv_tabix = iv_tabix ).
        ENDIF.
        LOOP AT lt_out_lines INTO DATA(ls_out_lines).
          CASE sy-tabix.
            WHEN 1. es_return-message_v1 = ls_out_lines.
            WHEN 2. es_return-message_v2 = ls_out_lines.
            WHEN 3. es_return-message_v3 = ls_out_lines.
            WHEN 4. es_return-message_v4 = ls_out_lines.
          ENDCASE.
        ENDLOOP.

        es_return-row    = iv_tabix.
        es_return-id     = '00'.
        es_return-number = '398'.
        es_return-type   = iv_msgtype.

        IF iv_pass_to_system IS NOT INITIAL.
          sy-msgid = es_return-id.
          sy-msgno = es_return-number.
          sy-msgty = es_return-type.
          sy-msgv1 = es_return-message_v1.
          sy-msgv2 = es_return-message_v2.
          sy-msgv3 = es_return-message_v3.
          sy-msgv4 = es_return-message_v4.
        ENDIF.
      ENDIF.

    ELSE.
      es_return-row        = iv_tabix.
      es_return-id         = sy-msgid.
      es_return-number     = sy-msgno.
      es_return-type       = sy-msgty.
      es_return-message_v1 = sy-msgv1.
      es_return-message_v2 = sy-msgv2.
      es_return-message_v3 = sy-msgv3.
      es_return-message_v4 = sy-msgv4.
    ENDIF.

    IF es_return IS NOT INITIAL.
      IF  es_return-id     IS NOT INITIAL
      AND es_return-number IS NOT INITIAL.
        CALL FUNCTION 'MESSAGE_TEXT_BUILD'
          EXPORTING
            msgid               = es_return-id
            msgnr               = es_return-number
            msgv1               = es_return-message_v1
            msgv2               = es_return-message_v2
            msgv3               = es_return-message_v3
            msgv4               = es_return-message_v4
          IMPORTING
            message_text_output = es_return-message.
      ENDIF.

      APPEND es_return TO mt_return.
    ENDIF.

  ENDMETHOD.
ENDCLASS.
