VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HotelSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrVendor As String 'local copy
Private mintSegNum As Integer 'local copy
Private mdtmCheckInDate As Date 'local copy
Private mdtmCheckOutDate As Date 'local copy
Private mintNumNights As Integer 'local copy
Private mstrCityCode As String 'local copy
Private mstrPropertyNum As String 'local copy
Private mstrHARPCode As String 'local copy
Private mstrPropertyName As String 'local copy
Private mintNumPersons As Integer 'local copy
Private mintNumRooms As Integer 'local copy
Private mstrConfNum As String 'local copy
Private mstrRateStatus As String 'local copy
Private mstrRateCurrency As String 'local copy
Private mcurRateAmount As Currency 'local copy
Private mstrNegRateCode As String 'local copy
Private mstrFreqGuest As String 'local copy
Private mstrFreqTvlr As String 'local copy
Private mstrGuar As String 'local copy
Private mstrPassengerName As String 'local copy
Private mstrServInfo As String 'local copy
Private mstrRoomType As String 'local copy
Private mstrSegType As String 'local copy
Private mcurRefRate As Currency 'local copy
Private mcurLowRate As Currency 'local copy
Private mstrSavingsCodeRealized As String 'local copy
Private mstrSavingsCodeMissed As String 'local copy
Private mbolPrefVendor As Boolean 'local copy
Private mvarStatus As String 'local copy
Private mvarCorpDiscNum As String 'local copy
Private mvarRollaway As Long 'local copy
Private mvarCrib As Long 'local copy
Private mvarExtraPerson As Long 'local copy
Private mvarExtraCharge As Single 'local copy
Private mvarGuestName As String 'local copy
Private mvarDeposit As String 'local copy

Public Property Get TextHtlSeg() As String
    TextHtlSeg = UCase(Format(CStr(mintSegNum), "@@. ") & mstrVendor & " " & mstrCityCode & " " _
        & Format(mdtmCheckInDate, "ddmmm") & " " & mstrPropertyName)
    
End Property

Public Property Let Deposit(ByVal vData As String)
    mvarDeposit = vData
End Property

Public Property Get Deposit() As String
    Deposit = mvarDeposit
End Property

Public Property Let GuestName(ByVal vData As String)
    mvarGuestName = vData
End Property

Public Property Get GuestName() As String
    GuestName = mvarGuestName
End Property

Public Property Let ExtraCharge(ByVal vData As Single)
    mvarExtraCharge = vData
End Property

Public Property Get ExtraCharge() As Single
    ExtraCharge = mvarExtraCharge
End Property

Public Property Let ExtraPerson(ByVal vData As Long)
    mvarExtraPerson = vData
End Property

Public Property Get ExtraPerson() As Long
    ExtraPerson = mvarExtraPerson
End Property

Public Property Let Crib(ByVal vData As Long)
    mvarCrib = vData
End Property

Public Property Get Crib() As Long
    Crib = mvarCrib
End Property

Public Property Let Rollaway(ByVal vData As Long)
    mvarRollaway = vData
End Property

Public Property Get Rollaway() As Long
    Rollaway = mvarRollaway
End Property

Public Property Let CorpDiscNum(ByVal vData As String)
    mvarCorpDiscNum = vData
End Property

Public Property Get CorpDiscNum() As String
    CorpDiscNum = mvarCorpDiscNum
End Property

Public Property Let Status(ByVal vData As String)
    mvarStatus = vData
End Property

Public Property Get Status() As String
    Status = mvarStatus
End Property

Public Property Let PrefVendor(ByVal PeferredVendor As Boolean)
    mbolPrefVendor = PeferredVendor
End Property

Public Property Get PrefVendor() As Boolean
    PrefVendor = mbolPrefVendor
End Property

Public Property Let SavingsCodeMissed(ByVal MissedSaveCode As String)
    mstrSavingsCodeMissed = MissedSaveCode
End Property

Public Property Get SavingsCodeMissed() As String
    SavingsCodeMissed = mstrSavingsCodeMissed
End Property

Public Property Let SavingsCodeRealized(ByVal RealikzedSaveCode As String)
    mstrSavingsCodeRealized = RealikzedSaveCode
End Property

Public Property Get SavingsCodeRealized() As String
    SavingsCodeRealized = mstrSavingsCodeRealized
End Property

Public Property Let LowRate(ByVal LowestRate As Currency)
    mcurLowRate = LowestRate
End Property

Public Property Get LowRate() As Currency
    LowRate = mcurLowRate
End Property

Public Property Let RefRate(ByVal ReferenceRate As Currency)
    mcurRefRate = ReferenceRate
End Property

Public Property Get RefRate() As Currency
    RefRate = mcurRefRate
End Property

Public Property Let SegType(ByVal SegmentType As String)
    mstrSegType = SegmentType
End Property

Public Property Get SegType() As String
    SegType = mstrSegType
End Property

Public Property Let RoomType(ByVal RoomTypeCode As String)
    mstrRoomType = RoomTypeCode
End Property

Public Property Get RoomType() As String
    RoomType = mstrRoomType
End Property

Public Property Let ServInfo(ByVal ServiceInformation As String)
    mstrServInfo = ServiceInformation
End Property

Public Property Get ServInfo() As String
    ServInfo = mstrServInfo
End Property

Public Property Let PassengerName(ByVal NameOfPassenger As String)
    mstrPassengerName = NameOfPassenger
End Property

Public Property Get PassengerName() As String
    PassengerName = mstrPassengerName
End Property

Public Property Let Guar(ByVal Guarentee As String)
    mstrGuar = Guarentee
End Property

Public Property Get Guar() As String
    Guar = mstrGuar
End Property

Public Property Let FreqTvlr(ByVal FrequentTravelerNumber As String)
    mstrFreqTvlr = FrequentTravelerNumber
End Property

Public Property Get FreqTvlr() As String
    FreqTvlr = mstrFreqTvlr
End Property

Public Property Let FreqGuest(ByVal FreqGuestNumber As String)
    mstrFreqGuest = FreqGuestNumber
End Property

Public Property Get FreqGuest() As String
    FreqGuest = mstrFreqGuest
End Property

Public Property Let NegRateCode(ByVal NegotiatedRateCode As String)
    mstrNegRateCode = NegotiatedRateCode
End Property

Public Property Get NegRateCode() As String
    NegRateCode = mstrNegRateCode
End Property

Public Property Let RateAmount(ByVal RoomRate As Currency)
    mcurRateAmount = RoomRate
End Property

Public Property Get RateAmount() As Currency
    RateAmount = mcurRateAmount
End Property

Public Property Let RateCurrency(ByVal CurrencyCode As String)
    mstrRateCurrency = CurrencyCode
End Property

Public Property Get RateCurrency() As String
    RateCurrency = mstrRateCurrency
End Property

Public Property Let RateStatus(ByVal RateGuarStatus As String)
    mstrRateStatus = RateGuarStatus
End Property

Public Property Get RateStatus() As String
    RateStatus = mstrRateStatus
End Property

Public Property Let ConfNum(ByVal ConfirmationNumber As String)
    mstrConfNum = ConfirmationNumber
End Property

Public Property Get ConfNum() As String
    ConfNum = mstrConfNum
End Property

Public Property Let NumRooms(ByVal NumberOfRooms As Integer)
    mintNumRooms = NumberOfRooms
End Property

Public Property Get NumRooms() As Integer
    NumRooms = mintNumRooms
End Property

Public Property Let NumPersons(ByVal NumberOfPersons As Integer)
    mintNumPersons = NumberOfPersons
End Property

Public Property Get NumPersons() As Integer
    NumPersons = mintNumPersons
End Property

Public Property Let PropertyName(ByVal NameOfProperty As String)
    mstrPropertyName = NameOfProperty
End Property

Public Property Get PropertyName() As String
    PropertyName = mstrPropertyName
End Property

Public Property Let HARPCode(ByVal HARPNumber As String)
    mstrHARPCode = HARPNumber
End Property

Public Property Get HARPCode() As String
    HARPCode = mstrHARPCode
End Property

Public Property Let PropertyNum(ByVal PropertyNumber As String)
    mstrPropertyNum = PropertyNumber
End Property

Public Property Get PropertyNum() As String
    PropertyNum = mstrPropertyNum
End Property

Public Property Let CityCode(ByVal City As String)
    mstrCityCode = City
End Property

Public Property Get CityCode() As String
    CityCode = mstrCityCode
End Property

Public Property Let NumNights(ByVal NumberOfNights As Integer)
    mintNumNights = NumberOfNights
End Property

Public Property Get NumNights() As Integer
    NumNights = mintNumNights
End Property

Public Property Let CheckOutDate(ByVal OutDate As Date)
    mdtmCheckOutDate = OutDate
End Property

Public Property Get CheckOutDate() As Date
    CheckOutDate = mdtmCheckOutDate
End Property

Public Property Let CheckInDate(ByVal InDate As Date)
    mdtmCheckInDate = InDate
End Property

Public Property Get CheckInDate() As Date
    CheckInDate = mdtmCheckInDate
End Property

Public Property Let SegNum(ByVal SegmentNumber As Integer)
    mintSegNum = SegmentNumber
End Property

Public Property Get SegNum() As Integer
    SegNum = mintSegNum
End Property

Public Property Let Vendor(ByVal VendorCode As String)
    mstrVendor = VendorCode
End Property

Public Property Get Vendor() As String
    Vendor = mstrVendor
End Property


