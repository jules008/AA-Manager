Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 27 Apr 18
'===============================================================
Private Const StrMODULE As String = "ModGlobals"

Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Public Const PROJECT_FILE_NAME As String = "RDS Agreement Manager"
Public Const APP_NAME As String = "RDS Agreement Manager"
Public Const EXPORT_FILE_PATH As String = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\AA Manager\Library\Dev\"
Public Const IMPORT_FILE_PATH As String = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\AA Manager\Library\Dev\"
Public Const LIBRARY_FILE_PATH As String = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\AA Manager\Library\"
Public Const DB_FILE_NAME As String = "Rappel Data Pre-Live v0,04.accdb"
Public Const INI_FILE_PATH As String = "\System Files\"
Public Const INI_FILE_NAME As String = "System.ini"
Public Const PROTECT_ON As Boolean = True
Public Const STOP_FLAG As Boolean = False
Public Const MAINT_MSG As String = ""
Public Const SEND_ERR_MSG As Boolean = False
Public Const TEST_PREFIX As String = "TEST - "
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const RNG_WEEK_1 As String = "C4:AX10"
Public Const RNG_TIME_1 As String = "C3:AX3"
Public Const RNG_DAY_1 As String = "B4:B10"
Public Const RNG_WKNO_1 As String = "B3"
Public Const VERSION = "V0.0.0"
Public Const DB_VER = "V0.0.0"
Public Const VER_DATE = ""

' ===============================================================
' Error Constants
' ---------------------------------------------------------------
Public Const HANDLED_ERROR As Long = 9999
Public Const UNKNOWN_USER As Long = 1000
Public Const SYSTEM_RESTART As Long = 1001
Public Const NO_DATABASE_FOUND As Long = 1002
Public Const ACCESS_DENIED As Long = 1003
Public Const NO_INI_FILE As Long = 1004
Public Const DB_WRONG_VER As Long = 1005
Public Const GENERIC_ERROR As Long = 1006
Public Const FORM_INPUT_EMPTY As Long = 1007
Public Const NO_USER_SELECTED As Long = 1008

' ===============================================================
' Error Variables
' ---------------------------------------------------------------
Public FaultCount1002 As Integer
Public FaultCount1008 As Integer

' ===============================================================
' Global Variables
' ---------------------------------------------------------------
Public DEBUG_MODE As Boolean
Public SEND_EMAILS As Boolean
Public ENABLE_PRINT As Boolean
Public DB_PATH As String
Public DEV_MODE As Boolean
Public SYS_PATH As String

' ===============================================================
' Global Class Declarations
' ---------------------------------------------------------------
Public MailSystem As ClsMailSystem
Public CurrentUser As ClsPerson
Public Agreement As ClsAgreement

' ---------------------------------------------------------------
' Others
' ---------------------------------------------------------------

' ===============================================================
' Colours
' ---------------------------------------------------------------
Public Const COLOUR_1 As Long = 4846018
Public Const COLOUR_2 As Long = 5731910
Public Const COLOUR_3 As Long = 16574371
Public Const COLOUR_4 As Long = 8876295
Public Const COLOUR_5 As Long = 5721863
Public Const COLOUR_6 As Long = 0
Public Const COLOUR_7 As Long = 16777215
Public Const COLOUR_8 As Long = 5880731
Public Const COLOUR_9 As Long = 0
Public Const COLOUR_10 As Long = 0
Public Const COLOUR_11 As Long = 0

' ===============================================================
' Enum Declarations
' ---------------------------------------------------------------
Enum EnumContractType
    Under120hrs = 0
    Over120hrs = 1
End Enum

Enum EnumFormValidation
    FunctionalError = 0
    ValidationError = 1
    FormOK = 2
End Enum

Enum EnumRole
    WCS = 0
    FDS = 1
    Admin = 2
End Enum

Enum EnumTriState
    xTrue = 0
    xFalse = 1
    xError = 2
End Enum

' ===============================================================
' Type Declarations
' ---------------------------------------------------------------
Type TypeStation
    StationNo As Integer
    StationCallSign As String
    StationName As String
End Type

Type TypeCrewMember
    CrewNo As String
    Name As String
    Station1 As TypeStation
    Station2 As TypeStation
    Role As EnumRole
End Type
