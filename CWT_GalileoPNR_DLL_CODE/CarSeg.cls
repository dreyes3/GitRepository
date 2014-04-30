VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CarSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Vendor"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSegNum As Integer 'local copy
Private mvarStatus As String
Private mvarSegType As String 'local copy
Private mvarStartDtTime As Date 'local copy
Private mvarEndDtTime As Date
Private mvarVendor As String
Private mvarNumPersons As Integer  'local copy
Private mvarNumNight As Integer 'local copy
Private mvarStartPt As String 'local copy
Private mvarEndPt As String 'local copy
Private mvarSellType As String 'local copy
Private mvarOtherText As String 'local copy
'local variable(s) to hold property value(s)

Private mvarStartAirVendor As String 'local copy
Private mvarStartFltNum As String 'local copy
Private mvarActualStartPt As String 'local copy
Private mvarLocationCat As String 'local copy
Private mvarLocationNum As String 'local copy
Private mvarNumCars As Single 'local copy
Private mvarConfNum As String 'local copy
Private mvarRateType As String 'local copy
Private mvarRateCode As String 'local copy
Private mvarRateCat As String 'local copy
Private mvarCarType As String 'local copy
Private mvarRateCurrency As String 'local copy
Private mvarRateAmt As Single 'local copy
Private mvarRateGuarInd As String 'local copy
Private mvarExtraDayRateAmt As String 'local copy
Private mvarRG As String 'local copy
Private mvarBS As String 'local copy
Private mvarPaxName As String 'local copy
Public Property Let PaxName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaxName = 5
    mvarPaxName = vData
End Property


Public Property Get PaxName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaxName
    PaxName = mvarPaxName
End Property



Public Property Let BS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BS = 5
    mvarBS = vData
End Property


Public Property Get BS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BS
    BS = mvarBS
End Property



Public Property Let RG(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RG = 5
    mvarRG = vData
End Property


Public Property Get RG() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RG
    RG = mvarRG
End Property



Public Property Let ExtraDayRateAmt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExtraDayRateAmt = 5
    mvarExtraDayRateAmt = vData
End Property


Public Property Get ExtraDayRateAmt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExtraDayRateAmt
    ExtraDayRateAmt = mvarExtraDayRateAmt
End Property



Public Property Let RateGuarInd(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RateGuarInd = 5
    mvarRateGuarInd = vData
End Property


Public Property Get RateGuarInd() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RateGuarInd
    RateGuarInd = mvarRateGuarInd
End Property



Public Property Let RateAmt(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RateAmt = 5
    mvarRateAmt = vData
End Property


Public Property Get RateAmt() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RateAmt
    RateAmt = mvarRateAmt
End Property



Public Property Let RateCurrency(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RateCurrency = 5
    mvarRateCurrency = vData
End Property


Public Property Get RateCurrency() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RateCurrency
    RateCurrency = mvarRateCurrency
End Property



Public Property Let CarType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CarType = 5
    mvarCarType = vData
End Property


Public Property Get CarType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CarType
    CarType = mvarCarType
End Property



Public Property Let RateCat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RateCat = 5
    mvarRateCat = vData
End Property


Public Property Get RateCat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RateCat
    RateCat = mvarRateCat
End Property



Public Property Let RateCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RateCode = 5
    mvarRateCode = vData
End Property


Public Property Get RateCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RateCode
    RateCode = mvarRateCode
End Property



Public Property Let RateType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RateType = 5
    mvarRateType = vData
End Property


Public Property Get RateType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RateType
    RateType = mvarRateType
End Property



Public Property Let ConfNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConfNum = 5
    mvarConfNum = vData
End Property


Public Property Get ConfNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConfNum
    ConfNum = mvarConfNum
End Property



Public Property Let NumCars(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumCars = 5
    mvarNumCars = vData
End Property


Public Property Get NumCars() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumCars
    NumCars = mvarNumCars
End Property



Public Property Let LocationNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationNum = 5
    mvarLocationNum = vData
End Property


Public Property Get LocationNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationNum
    LocationNum = mvarLocationNum
End Property



Public Property Let LocationCat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocationCat = 5
    mvarLocationCat = vData
End Property


Public Property Get LocationCat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocationCat
    LocationCat = mvarLocationCat
End Property



Public Property Let ActualStartPt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ActualStartPt = 5
    mvarActualStartPt = vData
End Property


Public Property Get ActualStartPt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ActualStartPt
    ActualStartPt = mvarActualStartPt
End Property



Public Property Let StartFltNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartFltNum = 5
    mvarStartFltNum = vData
End Property


Public Property Get StartFltNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartFltNum
    StartFltNum = mvarStartFltNum
End Property



Public Property Let StartAirVendor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartAirVendor = 5
    mvarStartAirVendor = vData
End Property


Public Property Get StartAirVendor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartAirVendor
    StartAirVendor = mvarStartAirVendor
End Property




Public Property Get TextCarSeg() As String
    TextCarSeg = UCase(Format(CStr(mvarSegNum), "@@. ") & mvarVendor & " " & mvarStartPt & " " _
        & Format(mvarStartDtTime, "ddmmm") & " " & mvarCarType & " " & mvarStatus)
    
End Property
Public Property Let SegNum(ByVal vData As Integer)
    mvarSegNum = vData
End Property

Public Property Get SegNum() As Integer
    SegNum = mvarSegNum
End Property

Public Property Let Status(ByVal vData As String)
    mvarStatus = vData
End Property

Public Property Get Status() As String
    Status = mvarStatus
End Property

Public Property Let SegType(ByVal vData As String)
    mvarSegType = vData
End Property

Public Property Get SegType() As String
    SegType = mvarSegType
End Property

Public Property Let StartDtTime(ByVal vData As Date)
    mvarStartDtTime = vData
End Property

Public Property Get StartDtTime() As Date
    StartDtTime = mvarStartDtTime
End Property

Public Property Let EndDtTime(ByVal vData As Date)
    mvarEndDtTime = vData
End Property

Public Property Get EndDtTime() As Date
    EndDtTime = mvarEndDtTime
End Property

Public Property Let Vendor(ByVal vData As String)
    mvarVendor = vData
End Property

Public Property Get Vendor() As String
    Vendor = mvarVendor
End Property

Public Property Let NumPersons(ByVal vData As Integer)
    mvarNumPersons = vData
End Property

Public Property Get NumPersons() As Integer
    NumPersons = mvarNumPersons
End Property

Public Property Let NumNight(ByVal vData As Integer)
    mvarNumNight = vData
End Property

Public Property Get NumNight() As Integer
    NumNight = mvarNumNight
End Property

Public Property Let StartPt(ByVal vData As String)
    mvarStartPt = vData
End Property

Public Property Get StartPt() As String
    StartPt = mvarStartPt
End Property

Public Property Let EndPt(ByVal vData As String)
    mvarEndPt = vData
End Property

Public Property Get EndPt() As String
    EndPt = mvarEndPt
End Property

Public Property Let SellType(ByVal vData As String)
    mvarSellType = vData
End Property

Public Property Get SellType() As String
    SellType = mvarSellType
End Property

Public Property Let OtherText(ByVal vData As String)
    mvarOtherText = vData
End Property


Public Property Get OtherText() As String
    OtherText = mvarOtherText
End Property
