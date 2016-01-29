<%
'C_Like_Recordset
'-------------------------------------------------------------------------------
' Purpose:    Create and object to act like a recordset.
' PROPERTIES: Records - This returns a dictionary object (wich holds each individual record).
'                       - Key   - Position of the record.
'                       - Value - Record object (Refer to C_Like_Record).
'             Size    - Returns the size of the record set (total count).
' EXAMPLE OF USE:
'                 DIM oRecordset : oRecordset = NEW C_Like_Recordset
'                 DIM oItem
'
'                 CALL oRecordset.grabRecords("SELECT * FROM Product")
'
'                 FOR EACH oItem IN oRecordset.Records.ITEMS
'                   'what to do when looping through each record
'                 MEXT
'-------------------------------------------------------------------------------
CLASS C_Like_Recordset
  PUBLIC TopPage, BottomPage, CurrentPage, Controller, Key
  PUBLIC BelongsTo, HasMany
  PRIVATE m_DefaultAmountOnPage
  PRIVATE m_record

  'record varaibles
  PRIVATE m_collection, m_size

  'configuration
  PRIVATE m_current_page, m_page_size, m_table_name, m_search_by, m_sory_by

  PRIVATE SUB Class_Initialize()
    m_DefaultAmountOnPage = 10
    m_size                = 0
    SET m_collection      = SERVER.CREATEOBJECT("Scripting.Dictionary")
    TopPage               = 1
    BottomPage            = 1
    CurrentPage           = 1
  END SUB

  PRIVATE SUB Class_Terminate()
    m_DefaultAmountOnPage = 0
    m_size                = 0
    SET m_collection      = NOTHING
  END SUB

  'get collection and variables
  PUBLIC PROPERTY GET Records
    SET Records = m_collection
  END PROPERTY
  PUBLIC PROPERTY GET Size
    Size = m_size
  END PROPERTY

  PUBLIC SUB grabPageRecords(sSQL,iPage,iAmountOnPage)
    DIM bHaveRecords : bHaveRecords = FALSE
    DIM rsRecordSet
    DIM iCount, iTopPage, iLBound, iUBound
    DIM aRecords
    DIM aFields()
    DIM oTemp

    'create out recordset
    SET rsRecordSet = createReadonlyRecordset(sSQL)

    'make sure we have records
    IF ( NOT rsRecordSet.RS.EOF ) THEN
      'get total size
      m_size = rsRecordSet.Size

      'get all rows
      aRecords = rsRecordSet.RS.GETROWS

      'move to first record
      rsRecordSet.RS.MOVEFIRST

      'redim our fields arary (this stores all field info)
      REDIM aFields(UBound(aRecords, 1))

      'loop through all the fields
      FOR iCount = LBOUND(aRecords, 1) To UBound(aRecords, 1)
        'create a field property object (saves all our properties for later use)
        SET oTemp = NEW C_Like_Field_Properties

        'set the properties
        oTemp.field_name    = rsRecordSet.RS.FIELDS(iCount).NAME
        oTemp.field_type    = get_field_type(rsRecordSet.RS.FIELDS(iCount).TYPE)
        oTemp.required      = ((rsRecordSet.RS.FIELDS(iCount).ATTRIBUTES AND adFldIsNullable) = 0)
        oTemp.primary_key   = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")
        oTemp.unique        = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")

        SELECT CASE UCASE(oTemp.field_type)
          CASE "DATE_TIME"
            oTemp.max_size      = 28

          CASE ELSE
            oTemp.max_size      = rsRecordSet.RS.FIELDS(iCount).DEFINEDSIZE

        END SELECT

        IF ( oTemp.primary_key ) THEN Key = oTemp.field_name

        'save the field property object in our fields array
        SET aFields(iCount) = oTemp

        'release temp field property obejct
        SET oTemp = NOTHING
      NEXT

      'mark that we have records
      bHaveRecords    = rsRecordSet.HaveRecords
    END IF

    'release recordset
    SET rsRecordSet = NOTHING

    'make sure we have records
    IF ( bHaveRecords ) THEN
      'lets work out the total amount of pages
      IF ( iAmountOnPage < m_size ) THEN
        iTopPage = m_size/iAmountOnPage

        'lets make sure we have the correct amount of pages (if there is a remainder in the above math, add a page)
        IF ( iTopPage > INT(iTopPage) ) THEN iTopPage = INT(iTopPage)+1
      ELSE
        iTopPage = 1
      END IF

      'checks to see if we have gone too far
      IF ( iPage > iTopPage ) THEN
        bHaveRecords = FALSE
      END IF
    END IF

    'make sure we have records
    IF ( bHaveRecords ) THEN
      'set our page variables
      TopPage     = iTopPage
      CurrentPage = iPage

      'create hight and low range
      iLBound = ( (iPage-1) * iAmountOnPage )+1
      iUBound = iPage * iAmountOnPage

      'if our high range is greates than the size, set the high range to the size
      IF ( iUBound > m_size ) THEN
        iUBound = m_size
      END IF

      'negative 1 on both (becuse we are dealing with arrays)
      iLBound = iLBound - 1
      iUBound = iUBound - 1

      'add the rows
      CALL addRows(iLBound,iUBound,aFields,aRecords)
    END IF

    'release fields array and records array
    REDIM aFields(0)
    REDIM aRecords(0)
  END SUB

  PUBLIC SUB grabRecords(sSQL)
    DIM bHaveRecords : bHaveRecords = FALSE
    DIM rsRecordSet
    DIM iCount
    DIM sValue
    DIM aRecords
    DIM aFields()
    DIM oTemp

    SET rsRecordSet = createReadonlyRecordset(sSQL)

    IF ( NOT rsRecordSet.RS.EOF ) THEN
      rsRecordSet.RS.MOVELAST
      rsRecordSet.RS.MOVEFIRST

      'set size
      m_size = rsRecordSet.Size

      aRecords = rsRecordSet.RS.GETROWS
      rsRecordSet.RS.MOVEFIRST

      REDIM aFields(UBound(aRecords, 1))

      'loop through all the fields
      FOR iCount = LBOUND(aRecords, 1) To UBound(aRecords, 1)
        'create a field property object (saves all our properties for later use)
        SET oTemp = NEW C_Like_Field_Properties

        'set the properties
        oTemp.field_name    = rsRecordSet.RS.FIELDS(iCount).NAME
        oTemp.field_type    = get_field_type(rsRecordSet.RS.FIELDS(iCount).TYPE)
        oTemp.required      = ((rsRecordSet.RS.FIELDS(iCount).ATTRIBUTES AND adFldIsNullable) = 0)
        oTemp.primary_key   = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")
        oTemp.unique        = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")

        SELECT CASE UCASE(oTemp.field_type)
          CASE "DATE_TIME"
            oTemp.max_size      = 28

          CASE ELSE
            oTemp.max_size      = rsRecordSet.RS.FIELDS(iCount).DEFINEDSIZE

        END SELECT

        IF ( oTemp.primary_key ) THEN Key = oTemp.field_name

        'save the field property object in our fields array
        SET aFields(iCount) = oTemp

        'release temp field property obejct
        SET oTemp = NOTHING
      NEXT

      'mark that we have records
      bHaveRecords    = rsRecordSet.HaveRecords
    END IF

    'release recordset
    SET rsRecordSet = NOTHING

    'make sure we have records
    IF ( bHaveRecords ) THEN
      'add the rows
      CALL addRows(LBOUND(aRecords, 2),UBOUND(aRecords, 2),aFields,aRecords)
    END IF

    'release fields array and records array
    REDIM aFields(0)
    REDIM aRecords(0)
  END SUB

  PUBLIC SUB NewRecord(sSQL)
    DIM iCount : iCount = 0
    DIM rsRecordSet
    DIM oTemp, oField

    SET rsRecordSet = createReadonlyRecordset(sSQL)

    SET m_record = NEW C_Like_Record

    m_record.table    = Controller
    m_record.table_id = Key

    CALL m_record.set_field_amount(rsRecordset.RS.Fields.COUNT)

    'loop through all the fields
    FOR EACH oField IN rsRecordset.RS.Fields
      'create a field property object (saves all our properties for later use)
      SET oTemp = NEW C_Like_Field_Properties

      'set the properties
      oTemp.field_name    = rsRecordSet.RS.FIELDS(iCount).NAME
      oTemp.field_type    = get_field_type(rsRecordSet.RS.FIELDS(iCount).TYPE)
      oTemp.required      = ((rsRecordSet.RS.FIELDS(iCount).ATTRIBUTES AND adFldIsNullable) = 0)
      oTemp.primary_key   = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")
      oTemp.unique        = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")

      SELECT CASE UCASE(oTemp.field_type)
        CASE "DATE_TIME"
          oTemp.max_size      = 28

        CASE ELSE
          oTemp.max_size      = rsRecordSet.RS.FIELDS(iCount).DEFINEDSIZE

      END SELECT

      CALL m_record.addRSField(iCount,oTemp,NULL)

      SET oTemp = NOTHING

      iCount = iCount + 1
    NEXT

    CALL m_record.set_field_value(Key,0)

    m_collection.ADD CSTR(0), m_record

    'release recordset
    SET rsRecordSet = NOTHING
  END SUB

  SUB addRows(iStartPosition,iEndPosition,aFields,aRecords)
    DIM iRecordCount, iFieldCount, iCount, iArrayCount
    DIM oItem, oBelongsTo, oHasMany
    DIM sValue

    'defaults
    iCount = 0

    FOR iRecordCount = iStartPosition To iEndPosition
      SET m_record = NEW C_Like_Record

      m_record.table    = Controller
      m_record.table_id = Key

      CALL m_record.set_field_amount((UBOUND(aFields)+1))

      FOR iFieldCount = LBOUND(aRecords, 1) TO UBOUND(aRecords, 1)
        sValue      = aRecords(iFieldCount, iRecordCount)
        SET oItem   = aFields(iFieldCount)

        CALL m_record.addRSField(iFieldCount,oItem,sValue)

        SET oItem = NOTHING
      NEXT

      IF ( ISARRAY(BelongsTo) ) THEN
        IF ( UBOUND(BelongsTo) = 0 ) THEN

          IF ( bHaveInfo(BelongsTo(0)) ) THEN
            SET oBelongsTo = return_belongs_to(BelongsTo(0))

            m_record.parents.ADD CSTR(LCASE(BelongsTo(0))), oBelongsTo
          END IF
        ELSE
          FOR iArrayCount = 0 TO UBOUND(BelongsTo)
            SET oBelongsTo = return_belongs_to(BelongsTo(iArrayCount))

            m_record.parents.ADD CSTR(LCASE(BelongsTo(iArrayCount))), oBelongsTo
          NEXT
        END IF
      END IF

      IF ( ISARRAY(HasMany) ) THEN
        IF ( UBOUND(HasMany) = 0 ) THEN

          IF ( bHaveInfo(HasMany(0)) ) THEN
            SET oHasMany = return_has_many(HasMany(0))

            m_record.children.ADD CSTR(LCASE(HasMany(0))), oHasMany
          END IF
        ELSE
          FOR iArrayCount = 0 TO UBOUND(HasMany)
            SET oHasMany = return_has_many(HasMany(iArrayCount))

            m_record.children.ADD CSTR(LCASE(HasMany(iArrayCount))), oHasMany
          NEXT
        END IF
      END IF

      m_collection.ADD CSTR(iCount), m_record

      iCount = iCount + 1
    NEXT
  END SUB

  FUNCTION return_belongs_to(sTable)
    DIM oCollection

    IF ( bHaveInfo(sTable) ) THEN
      EXECUTE("SET oCollection = " & sTable & ".find_link(ARRAY(ARRAY(" & sTable & ".primary_key,""="",""" & m_record.value(LCASE(sTable) & "_id") & """)),TRUE,FALSE)")

      IF ( oCollection.Records.COUNT > 0 ) THEN
        SET return_belongs_to = oCollection.Records.ITEM(CSTR(0))
      ELSE
        SET return_belongs_to = NOTHING
      END IF
    ELSE
      SET return_belongs_to = NOTHING
    END IF
  END FUNCTION

  FUNCTION return_has_many(sTable)
    DIM oCollection

    IF ( bHaveInfo(sTable) ) THEN
      EXECUTE("SET oCollection = " & sTable & ".find_link(ARRAY(ARRAY(""" & Controller & "_id"",""="",""" & m_record.value(LCASE(sTable) & "_id") & """)),FALSE,TRUE)")

      IF ( oCollection.Records.COUNT > 0 ) THEN
        SET return_has_many = oCollection.Records.ITEM(CSTR(0))
      ELSE
        SET return_has_many = NOTHING
      END IF
    ELSE
      SET return_has_many = NOTHING
    END IF
  END FUNCTION
END CLASS

'C_Like_Record
'-------------------------------------------------------------------------------
' Purpose:    Creates an object wich holds specific information for each data field for a specific record.
'             It uses "record_array" property wich holds the field object in a multi dimensional array.
' PROPERTIES: set_field_amount   - Sets the array size, the array sie is multidimensional and based on the amount of fields you will be adding for the specific record
'             value              - Returns the value for a specific field name
'             field_type         - Returns the field type for a specific field name
'             required           - Returns wether or not the field is required
'             primary_key        - Returns wether or not the field is a primary key
'             max_size           - Returns the maximum size for a specific field name
'             find_param_value   - Returns the value of a selected parameter for a specific field name
'
'PUBLIC SUBS: addRSField         - adds a field object to a specific array location (specifically from a recordset).
'
'-------------------------------------------------------------------------------
CLASS C_Like_Record
  PUBLIC record_array, errors, table, table_id, parents, children

  'on object initialization
  PRIVATE SUB Class_Initialize()
    record_array      = NULL
    SET errors        = NEW C_Errors
    SET parents       = SERVER.CREATEOBJECT("Scripting.Dictionary")
    SET children      = SERVER.CREATEOBJECT("Scripting.Dictionary")
  END SUB

  'on object termination
  PRIVATE SUB Class_Terminate()
    record_array      = NULL
    SET errors        = NOTHING
    SET parents       = NOTHING
    SET children      = NOTHING
  END SUB

  PUBLIC SUB set_field_amount(amount)
    REDIM record_array(amount-1,1)
  END SUB

  PUBLIC FUNCTION value(name)
    value = find_param_value(name,"value")
  END FUNCTION

  PUBLIC FUNCTION field_type(name)
    field_type = find_param_value(name,"field_type")
  END FUNCTION

  PUBLIC FUNCTION required(name)
    required = find_param_value(name,"required")
  END FUNCTION

  PUBLIC FUNCTION primary_key(name)
    primary_key = find_param_value(name,"primary_key")
  END FUNCTION

  PUBLIC FUNCTION max_size(name)
    max_size = find_param_value(name,"max_size")
  END FUNCTION

  PUBLIC SUB setFromForm(ExcludeFields)
    DIM counter, field, error_item, temp

    FOR counter = 0 TO UBOUND(record_array)
      SET field = record_array(counter,1)

      IF ( NOT bFoundInArray(field.name,ExcludeFields) ) THEN
        SELECT CASE UCASE(field.field_type)
          CASE "DATE_TIME"
            temp = REQUEST(field.name & "_day") & "/"
            temp = temp & REQUEST(field.name & "_month") & "/"
            temp = temp & REQUEST(field.name & "_year") & " "
            temp = temp & REQUEST(field.name & "_hour") & ":"
            temp = temp & REQUEST(field.name & "_minute") & ""
            temp = temp & REQUEST(field.name & "_ampm")

            IF ( ISDATE(temp) ) THEN
              field.value = temp
            END IF

          CASE ELSE
            field.value = c_FormRequest(field.name)
        END SELECT
      END IF

      SET record_array(counter,1) = field
    NEXT
  END SUB

  PUBLIC FUNCTION valid()
    DIM counter, field, error_item

    FOR counter = 0 TO UBOUND(record_array)
      SET field = SingleItem.record_array(counter,1)

      IF ( NOT field.valid ) THEN
        FOR EACH error_item IN field.errors.Items
          CALL errors.addError(error_item.ID,error_item.Name,error_item.Description,NULL,NULL)
        NEXT
      END IF
    NEXT

    valid = ( errors.Size = 0 )
  END FUNCTION

  PUBLIC FUNCTION save
    DIM counter, field
    DIM rsRecordSet, rsUpdate

    'create out recordset
    SET rsRecordSet = createUpdateableRecordset(genegrateSQL(table,ARRAY(ARRAY(table_id,"=",value(table_id)))))
    SET rsUpdate    = rsRecordSet.RS

    IF ( valid ) THEN
      FOR counter = 0 TO UBOUND(record_array)
        SET field = record_array(counter,1)

        IF ( NOT field.primary_key ) THEN
          rsUpdate(field.name) = field.value
        END IF
      NEXT

      rsUpdate.UPDATE

      save = TRUE
    ELSE
      save = FALSE
    END IF

    'release recordset
    SET rsUpdate    = NOTHING
    SET rsRecordSet = NOTHING
  END FUNCTION

  PUBLIC FUNCTION delete
    DIM counter, field
    DIM rsRecordSet, rsUpdate

    'create out recordset
    SET rsRecordSet = createUpdateableRecordset(genegrateSQL(table,ARRAY(ARRAY(table_id,"=",value(table_id)))))
    SET rsUpdate    = rsRecordSet.RS

    IF ( NOT rsUpdate.EOF ) THEN
      rsUpdate.DELETE

      delete = TRUE
    ELSE
      delete = FALSE
    END IF

    'release recordset
    rsUpdate.CLOSE
    SET rsUpdate    = NOTHING
    SET rsRecordSet = NOTHING
  END FUNCTION

  PUBLIC FUNCTION fill_in_variables(data)
    DIM temp_data : temp_data = data
    DIM regexp, matches, match

    IF ( bHaveInfo(temp_data) ) THEN
      'Create our regular expression object
      SET regexp = NEW RegExp

      WITH regexp
        .PATTERN    = "(@.+?@)"
        .IGNORECASE = FALSE
        .GLOBAL     = TRUE
      END WITH

      'Strip out all invalid characters
      SET matches = regexp.EXECUTE(temp_data)

      'loop through all the matches
      FOR EACH match  IN matches
        temp_data = REPLACE(temp_data,match.VALUE,value(REPLACE(match.VALUE,"@","")))
      NEXT
    END IF

    fill_in_variables = temp_data
  END FUNCTION

  PUBLIC FUNCTION find_param_value(name,param)
    DIM iFieldCount : iFieldCount = 0
    DIM bEnd        : bEnd        = FALSE
    DIM field
    DIM value

    DO WHILE ( iFieldCount <= UBOUND(record_array) AND NOT bEnd )
      IF ( UCASE(name) = UCASE(record_array(iFieldCount,0)) ) THEN
        'set the object
        SET field = record_array(iFieldCount,1)

        'assign the param value
        EXECUTE("value = field." & param)

        'send correct value back
        find_param_value = value

        'release temp object
        SET field = NOTHING
      END IF

      iFieldCount = iFieldCount + 1
    LOOP
  END FUNCTION

  PUBLIC SUB set_field_value(name,value)
    CALL set_field_param(name,"value",value)
  END SUB

  PUBLIC SUB set_field_param(name,param,value)
    DIM iFieldCount : iFieldCount = 0
    DIM bEnd        : bEnd        = FALSE
    DIM field

    DO WHILE ( iFieldCount <= UBOUND(record_array) AND NOT bEnd )
      IF ( UCASE(name) = UCASE(record_array(iFieldCount,0)) ) THEN
        'set the object
        SET field = record_array(iFieldCount,1)

        'assign the param value
        EXECUTE("field." & param & " = value")

        SET record_array(iFieldCount,1) = field

        bEnd = TRUE
      END IF

      iFieldCount = iFieldCount + 1
    LOOP
  END SUB

  PUBLIC SUB addRSField(iArrayPosition,oItem,sValue)
    DIM field

    'set object
    SET field = NEW C_Like_Field

    'set field properties
    field.value       = sValue
    field.name        = oItem.field_name
    field.field_type  = oItem.field_type
    field.required_db = oItem.required
    field.primary_key = oItem.primary_key
    field.unique      = oItem.unique
    field.max_size    = oItem.max_size

    'save in field array
    record_array(iArrayPosition,0)      = oItem.field_name
    SET record_array(iArrayPosition,1)  = field
  END SUB
END CLASS

'C_Like_Field_Properties
'-------------------------------------------------------------------------------
' Purpose:    Creates an object wich olds properties for a specific field
' PROPERTIES: field_name    - Name of the field
'             field_type    - Type of field
'             primary_key   - Is the field a primary key
'             unique        - Does the field need to be unique
'             max_size      - The maximum size of the field
'             required      - Is the field required
'
' EXAMPLE OF USE:
'                 DIM oField = NEW C_Like_Field_Properties
'
'                 RESPONSE.WRITE oField.name
'-------------------------------------------------------------------------------
CLASS C_Like_Field_Properties
  PUBLIC field_name, field_type, required, primary_key, unique, max_size

END CLASS
%>
