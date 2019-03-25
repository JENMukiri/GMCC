Attribute VB_Name = "CSA_calc"
Option Explicit

Function CSA_calculate(CroppingSystem_Range As Range, Season As Integer, Return_is As String) As Variant

'GrossMargin takes into account:
' - income (grain yield; biomass yield)
' - expenses (seeds, fertilizer, herbicides, pesticides and
'             labor hours for planting, weeding, pesticide application, harvesting)
' These are Season, Zone, Crop and Input-level specific


'Dim Crop As String 'both crops combined if intercropping
'Dim Previous_Crop As String
'Dim Crop_a As String 'crop-a of intercropping system
'Dim Crop_b As String 'crop-b of intercropping system
Dim Zone As String
Dim Crop As String
Dim Crop_of_Season(7) As String ' six possible crops
Dim CropA_of_Season(7) As String
Dim CropB_of_Season(7) As String
Dim ZoneCrop As String 'concatenating Zone and Crop
Dim ZoneCrop_of_Season(7) As String
Dim ZoneCropCrop(7) As String 'Zone plus current and last season crop
Dim ZonePrev_Crop As String
Dim ZoneCropA_of_Season(7) As String
Dim ZoneCropB_of_Season(7) As String
Dim Fert_type1 As String
Dim Fert_amount1 As Double
Dim Fert_type2 As String
Dim Fert_amount2 As Double
Dim Res_removal1 As Variant
Dim Res_removal2 As Variant
Dim Manure_amount As Double
Dim Pesticide_appl As String
Dim Herbicide_appl As String

Dim Yield1 As Double
Dim Yield2 As Double
Dim Biomass1 As Double
Dim Biomass2 As Double
Dim OLD_Y As Double 'to temporarily store old yield
Dim N_input As Double
Dim N_input_lastSeasons As Double

Dim AgronR As Range 'upper case R for indicating a range
Dim LaborR As Range 'Labor hour range
Dim CommPriceR As Range 'commodity price range
Dim FertNPriceR As Range 'mineral fertilizer N and price range
Dim InputAmountsR As Range 'Amount of inputs used
Dim AgronSubseqR As Range

Dim Price_Grain1 As Double
Dim Price_Grain2 As Double
Dim Price_Biom1 As Double
Dim Price_Biom2 As Double
Dim Price_Seed1 As Double
Dim Price_Seed2 As Double
Dim Price_Fertilizer1 As Double
Dim Price_Fertilizer2 As Double
Dim Price_Herbicide As Double
Dim Price_Pesticide As Double
Dim Price_Manure As Double
Dim Price_Labor As Double

Dim LH_plough As Single 'LH = labor hours
Dim LH_plant As Single
Dim LH_fert As Single
Dim LH_manure As Single
Dim LH_residue As Single
Dim LH_weed As Single
Dim LH_spray As Single
Dim LH_spray_accum As Single 'cumulative hours required for spraying
Dim LH_harvest As Single
Dim LH_postharv As Single
Dim Sum_LH As Single

Dim Seed1_kg As Double
Dim Seed2_kg As Double
Dim Herbicide_litre As Double
Dim Pesticide_litre As Double

Dim Income As Double
Dim Labor_costs As Double
Dim Input_costs As Double
Dim GrossMargin As Double
Dim NBalance As Double
Dim PBalance As Double

Dim Temp As Variant
Dim i As Integer

Dim Is_Intercrop(7) As Boolean
Dim Crop1_Is_Harvested(7) As Boolean
Dim Crop2_Is_Harvested(7) As Boolean

Dim Yield_change_factor As Double
Dim Biomass_change_factor As Double
Dim Monocrop_Seasons As Integer 'count number of mono-cropped seasons
Dim Stop_N_forward As Boolean

'Read Cropping System settings
Zone = CroppingSystem_Range(Season, 2).Value
Crop = CroppingSystem_Range(Season, 3).Value
Fert_type1 = CroppingSystem_Range(Season, 4).Value
Fert_amount1 = CroppingSystem_Range(Season, 5).Value
Res_removal1 = CroppingSystem_Range(Season, 6).Value
Res_removal2 = CroppingSystem_Range(Season, 7).Value
Manure_amount = CroppingSystem_Range(Season, 8).Value * 1000 'convert to kg/ha
Pesticide_appl = CroppingSystem_Range(Season, 9).Value
Herbicide_appl = CroppingSystem_Range(Season, 10).Value
Fert_type2 = CroppingSystem_Range(Season, 11).Value
Fert_amount2 = CroppingSystem_Range(Season, 12).Value
       
       
If Res_removal1 = "Burn" Then
    '100% is lost, hence comercial value of residues is lost
    Res_removal1 = 0
 Else
    Res_removal1 = Res_removal1 / 100
End If

If Res_removal2 = "Burn" Then
    '100% is lost, hence comercial value of residues is lost
    Res_removal2 = 0
 Else
    Res_removal2 = Res_removal2 / 100
End If
       
       
i = 0
Monocrop_Seasons = 1

'Read all previous crops, check whether they were harvested and whether they are intercrops
For i = Season To 1 Step -1
    Crop_of_Season(i) = CroppingSystem_Range(i, 3).Value
    'Check if selected crop is intercrop
    If InStr(Crop_of_Season(i), "/") > 0 Then
       Temp = Split(Crop_of_Season(i), "/")
       CropA_of_Season(i) = Temp(0)
       CropB_of_Season(i) = Temp(1)
       ZoneCropA_of_Season(i) = Zone & CropA_of_Season(i)
       ZoneCropB_of_Season(i) = Zone & CropB_of_Season(i)
       Is_Intercrop(i) = True
    Else
       ZoneCropA_of_Season(i) = Zone & Crop_of_Season(i)
       CropA_of_Season(i) = Crop_of_Season(i)
       Is_Intercrop(i) = False
    End If
    ZoneCrop_of_Season(i) = Zone & Crop_of_Season(i)
    Crop1_Is_Harvested(i) = CroppingSystem_Range(i, 13).Value
    Crop2_Is_Harvested(i) = CroppingSystem_Range(i, 14).Value
    'Find out whether previous crop is the same like current
    If CropA_of_Season(i + 1) = CropA_of_Season(i) Then
        If Crop1_Is_Harvested(i) Then
            Monocrop_Seasons = Monocrop_Seasons + 1 'number of mono-cropping seasons
            Stop_N_forward = True
         Else
            'Find out whether one of the previous crops continues growing into this season
            '(e.g. Sugarcane/Cowpea in season 1, then Cowpea harvested, and Sugarcane continue growing into season 2 and maybe 3)
            'Take forward N-input of last season
            If Stop_N_forward = False Then N_input_lastSeasons = N_input_lastSeasons + N_P_balance(CroppingSystem_Range, i, 0, 0, 0, 0, "N_input")
        End If
    Else
        If i < Season Then Exit For
   End If
Next
  
For i = 1 To Season
    ZoneCropCrop(i) = Zone & Crop_of_Season(i - 1) & Crop_of_Season(i)
Next i


'Assign worksheet name-ranges to VBA ranges
Set AgronR = ThisWorkbook.Sheets("Agronomy").Range("Agronomy_range")
Set LaborR = ThisWorkbook.Sheets("Labour").Range("Labor_range")
Set CommPriceR = ThisWorkbook.Sheets("Costs_Prices").Range("Price_Comm")
Set FertNPriceR = ThisWorkbook.Sheets("Costs_Prices").Range("Fertilizer_N_Price")
Set InputAmountsR = ThisWorkbook.Sheets("Inputs").Range("Input_List")
Set AgronSubseqR = ThisWorkbook.Sheets("Agronomy_subseq ").Range("Agron_subsequent")



'Read Base yield and biomass data; last number is column number of Agronomy range
If Crop1_Is_Harvested(Season) Then Yield1 = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop_of_Season(Season), AgronR.Columns(1), 0), 2)
If Crop2_Is_Harvested(Season) Then Yield2 = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop_of_Season(Season), AgronR.Columns(1), 0), 7)
If Crop1_Is_Harvested(Season) Then Biomass1 = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop_of_Season(Season), AgronR.Columns(1), 0), 10)
If Crop2_Is_Harvested(Season) Then Biomass2 = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop_of_Season(Season), AgronR.Columns(1), 0), 11)


If Season > 1 Then
    'Account for effect of crop rotation
    'read %reduction of yield and biomass from Agronomy_subseq worksheet
    Yield_change_factor = 1 + WorksheetFunction.Index(AgronSubseqR, WorksheetFunction.Match(ZoneCropCrop(Season), AgronSubseqR.Columns(1), 0), 2) / 100
    Biomass_change_factor = 1 + WorksheetFunction.Index(AgronSubseqR, WorksheetFunction.Match(ZoneCropCrop(Season), AgronSubseqR.Columns(1), 0), 3) / 100
    
   'Reduce yield and biomass change factors even further if mono-cropping goes on over many seasons
    If Monocrop_Seasons > 1 Then
        'For e.g. three seasons of cotton in a row, we are further reducing the yields by mupltiplying the yield response factor with itself
        '...but not for the first time planting the same crop (which is Monocrop_Seasons = 2), that is where the '-1' comes in ;)
        Yield_change_factor = Yield_change_factor ^ (Monocrop_Seasons - 1)
        Biomass_change_factor = Biomass_change_factor ^ (Monocrop_Seasons - 1)
    End If
    
    If Crop1_Is_Harvested(Season) Then Yield1 = Yield1 * Yield_change_factor
    If Crop2_Is_Harvested(Season) Then Yield2 = Yield2 * Yield_change_factor
    If Crop1_Is_Harvested(Season) Then Biomass1 = Biomass1 * Biomass_change_factor
    If Crop2_Is_Harvested(Season) Then Biomass2 = Biomass2 * Biomass_change_factor
End If

'Add fertilizer response on top of base yields to calculate final yield and biomass
N_input = N_input_lastSeasons + N_P_balance(CroppingSystem_Range, Season, 0, 0, 0, 0, "N_input") 'call this function to do the job ;)
OLD_Y = Yield1
If Crop1_Is_Harvested(Season) Then Yield1 = Yield1 + Fertilizer_response(ZoneCrop_of_Season(Season), N_input, 5)
If Crop1_Is_Harvested(Season) Then Biomass1 = Biomass1 * Yield1 / OLD_Y
OLD_Y = Yield2
If Crop2_Is_Harvested(Season) Then Yield2 = Yield2 + Fertilizer_response(ZoneCrop_of_Season(Season), N_input, 5)
If Crop2_Is_Harvested(Season) Then Biomass2 = Biomass2 * Yield2 / OLD_Y


'Call N-balance
NBalance = N_P_balance(CroppingSystem_Range, Season, Yield1, Yield2, Biomass1, Biomass2, "N_balance")
'Call P-balance
PBalance = N_P_balance(CroppingSystem_Range, Season, Yield1, Yield2, Biomass1, Biomass2, "P_balance")


'Read commodity prices
Price_Grain1 = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropA_of_Season(Season), CommPriceR.Columns(1), 0), 2)
If Is_Intercrop(Season) Then Price_Grain2 = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropB_of_Season(Season), CommPriceR.Columns(1), 0), 2)
Price_Biom1 = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropA_of_Season(Season), CommPriceR.Columns(1), 0), 3)
If Is_Intercrop(Season) Then Price_Biom2 = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropB_of_Season(Season), CommPriceR.Columns(1), 0), 3)


'INCOME--------------------------------------------------------------------------------------------------
Income = Yield1 * Price_Grain1 + Yield2 * Price_Grain2 _
        + Res_removal1 * Biomass1 * Price_Biom1 + Res_removal2 * Biomass2 * Price_Biom2



'EXPENSES

'Read labor hours, partly depending on whether activity was performed or not (see If...Then...);
'last number is column number of the Labor-hour range
If Crop1_Is_Harvested(Season - 1) Or Season = 1 Then
    LH_plough = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 2)
    LH_plant = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 3)
End If
If Fert_amount1 > 0 Then LH_fert = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 4)
If Fert_amount2 > 0 Then LH_fert = LH_fert + WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 4)
If Manure_amount > 0 Then LH_manure = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 5)
If Res_removal1 + Res_removal2 > 0 Then LH_residue = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 6)

LH_weed = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 7)
'reduce or increase time for weeding based on herbicide use; count hours for herbicide spraying
LH_spray = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 8)
LH_spray_accum = LH_spray

Select Case Herbicide_appl
'None, ½ of Average, 2x Average, 4 x Average
Case "None"
    LH_weed = LH_weed * 2
    LH_spray_accum = 0
Case "½ of Average"
    LH_weed = LH_weed * 1.5
Case "2x Average"
    LH_weed = LH_weed / 1.5
    LH_spray_accum = LH_spray * 2
Case "4x Average"
    LH_weed = LH_weed / 2
    LH_spray_accum = LH_spray * 4
End Select

'labor hours for pesticide spraying
Select Case Pesticide_appl
'None, ½ of Average, 2x Average, 4 x Average
Case "None"
    If Herbicide_appl = "None" Then LH_spray = 0
Case "½ of Average"
    LH_spray_accum = LH_spray_accum + LH_spray
Case "Average"
    LH_spray_accum = LH_spray_accum + LH_spray
Case "2x Average"
    LH_spray_accum = LH_spray_accum + LH_spray * 2
Case "4x Average"
    LH_spray_accum = LH_spray_accum + LH_spray * 4
End Select


If Crop1_Is_Harvested(Season) Or Crop2_Is_Harvested(Season) Then LH_harvest = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 9)
If Crop1_Is_Harvested(Season) Or Crop2_Is_Harvested(Season) Then LH_postharv = WorksheetFunction.Index(LaborR, WorksheetFunction.Match(Crop_of_Season(Season), LaborR.Columns(1), 0), 10)


If Is_Intercrop(Season) Then
    If Not (Crop1_Is_Harvested(Season) And Crop2_Is_Harvested(Season)) Then
    'Check whether crop is indeed intercrop, and if so, if both crops were harvested.
    'If not, then harvesting time must be reduced
        LH_harvest = LH_harvest / 2
        LH_postharv = LH_postharv / 2
    End If
End If

Sum_LH = LH_plough + LH_plant + LH_fert + LH_manure + LH_residue + LH_weed + LH_spray_accum + LH_harvest + LH_postharv
Price_Labor = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropA_of_Season(Season), CommPriceR.Columns(1), 0), 4)
'Costs for Labor
Labor_costs = Sum_LH * Price_Labor


'Read prices for seeds, fertilizer, manure, herbicides and pesticides;
If i = 0 Then Price_Seed1 = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropA_of_Season(Season), CommPriceR.Columns(1), 0), 5)
If Is_Intercrop(Season) And i = 0 Then Price_Seed2 = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropB_of_Season(Season), CommPriceR.Columns(1), 0), 5)
'Note! herbicide and pesticide only for crop-a!
Price_Pesticide = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropA_of_Season(Season), CommPriceR.Columns(1), 0), 6)
Price_Herbicide = WorksheetFunction.Index(CommPriceR, WorksheetFunction.Match(ZoneCropA_of_Season(Season), CommPriceR.Columns(1), 0), 7)
Price_Fertilizer1 = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match(Fert_type1, FertNPriceR.Columns(1), 0), 3)
Price_Fertilizer2 = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match(Fert_type2, FertNPriceR.Columns(1), 0), 3)
Price_Manure = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match("Manure", FertNPriceR.Columns(1), 0), 3)

'Amount of inputs used
Seed1_kg = WorksheetFunction.Index(InputAmountsR, WorksheetFunction.Match(Crop_of_Season(Season), InputAmountsR.Columns(1), 0), 2)
Seed2_kg = WorksheetFunction.Index(InputAmountsR, WorksheetFunction.Match(Crop_of_Season(Season), InputAmountsR.Columns(1), 0), 3)
Herbicide_litre = WorksheetFunction.Index(InputAmountsR, WorksheetFunction.Match(Crop_of_Season(Season), InputAmountsR.Columns(1), 0), 7)
Pesticide_litre = WorksheetFunction.Index(InputAmountsR, WorksheetFunction.Match(Crop_of_Season(Season), InputAmountsR.Columns(1), 0), 9)

'Reduce or increase pesticide and herbicide amount according to user-input
Select Case Pesticide_appl
'None, ½ of Average, 2x Average, 4 x Average
Case "None"
    Pesticide_litre = 0
Case "½ of Average"
    Pesticide_litre = Pesticide_litre / 2
Case "2x Average"
    Pesticide_litre = Pesticide_litre * 2
Case "4x Average"
    Pesticide_litre = Pesticide_litre * 4
End Select

Select Case Herbicide_appl
'None, ½ of Average, 2x Average, 4 x Average
Case "None"
    Herbicide_litre = 0
Case "½ of Average"
    Herbicide_litre = Herbicide_litre / 2
Case "2x Average"
    Herbicide_litre = Herbicide_litre * 2
Case "4x Average"
    Herbicide_litre = Herbicide_litre * 4
End Select


'Costs for inputs
Input_costs = Price_Seed1 * Seed1_kg + Price_Seed2 * Seed2_kg + Price_Pesticide * Pesticide_litre + Price_Herbicide * Herbicide_litre _
              + Price_Fertilizer1 * Fert_amount1 + Price_Fertilizer2 * Fert_amount2 + Price_Manure * Manure_amount

'GrossMargin
GrossMargin = Income - Labor_costs - Input_costs

'Produce final result according what return is requested
Select Case Return_is
    Case "Yield1"
        CSA_calculate = Yield1
        
    Case "Yield2"
        CSA_calculate = Yield2
        
    Case "Biomass1"
        CSA_calculate = Biomass1
    
    Case "Biomass2"
        CSA_calculate = Biomass2
    
    Case "Gross Margin"
    CSA_calculate = GrossMargin
    
    Case "Labor hours"
    CSA_calculate = Sum_LH
    
    Case "N Balance"
    CSA_calculate = NBalance
    
    Case "P Balance"
    CSA_calculate = PBalance
    Case Else
    CSA_calculate = "something wrong here"
End Select

End Function


Function Stover_Yield(Grain_yield As Double, Harvest_Index As Double) As Variant

'Grain yield [kg/ha]
'Harvest index [-]
'Original Function Jessica: =IFERROR(D2*((1-N2)/N2),0)

If (Grain_yield = 0 Or Harvest_Index = 0) Then
    Stover_Yield = "data missing"
 Else
    Stover_Yield = Grain_yield * ((1 - Harvest_Index) / Harvest_Index)
End If

End Function

Function Asymptotic_increase(a As Double, b As Double, c As Double, r As Double) As Double
'a is near maximum yield (t/ha)
'b is gain in yield due to nutrient application
'c determines the shape of the curve
'r is the nutrient application rate (kg/ha)

Asymptotic_increase = a - b * c ^ r

End Function

Function Atmosph_Deposition(Rain_amount As Double) As Double
'Atmosph_Deposition in kg N/ha
'Rainfall in mm/year
'NUTMON equation

Atmosph_Deposition = 0.14 * Sqr(Rain_amount)

End Function

Function Non_Symbiotic_Nfixation(Rain_amount As Double) As Double
'Non_Symbiotic_Nfixation in kg N/ha
'Rainfall in mm/year
'NUTMON equation

Non_Symbiotic_Nfixation = (2 + (Rain_amount - 500) * 0.005)
'RS 1 Nov2018: -1350 change to -500

End Function

Function N_leaching(Clay As Double, N_input As Double, Rain_amount As Double) As Double
'Clay in percent
'N_input in kg/ha

Select Case Clay
    Case Is < 36
        N_leaching = N_input * (0.021 * Rain_amount - 3.9) / 100
    Case Is > 54
        N_leaching = N_input * (0.0071 * Rain_amount + 5.4) / 100
    Case Else
        N_leaching = N_input * (0.014 * Rain_amount + 0.71) / 100
End Select

End Function

Function N_P_balance(CroppingSystem_Range As Range, Season As Integer, Yield_1 As Double, Yield_2 As Double, Biomass_1 As Double, _
                   Biomass_2 As Double, Return_what As String)
'Predict kg N/ha leached per season

Dim Zone As String
Dim Crop As String 'both crops combined if intercropping
Dim Crop_a As String 'crop-a of intercropping system
Dim Crop_b As String 'crop-b of intercropping system
Dim ZoneCrop As String 'concatenating Zone and Crop
Dim ZoneCrop_a As String
Dim ZoneCrop_b As String
Dim Nmin_kg As Double
Dim Nmin_ppm As Double
Dim FertN_kg As Double
Dim FertP_kg As Double
Dim ManureN_kg As Double
Dim ManureP_kg As Double
Dim Rain_amount As Double
Dim Fert_type1 As String
Dim Fert_type2 As String
Dim Fert_amount1 As Double
Dim Fert_amount2 As Double
Dim Manure_amount As Double
Dim Clay As Double 'percent clay
Dim Seasons_per_year As Single

Dim AgroEcolR As Range
Dim FertNPriceR As Range
Dim Yield1N_conc As Double
Dim Yield2N_conc As Double
Dim Biomass1N_conc As Double
Dim Biomass2N_conc As Double
Dim Yield1P_conc As Double
Dim Yield2P_conc As Double
Dim Biomass1P_conc As Double
Dim Biomass2P_conc As Double
Dim Yield1N As Double
Dim Yield2N As Double
Dim Biomass1N As Double
Dim Biomass2N As Double
Dim Yield1P As Double
Dim Yield2P As Double
Dim Biomass1P As Double
Dim Biomass2P As Double
Dim AgronR As Range 'upper case R for indicating a range

Dim Temp As Variant
Dim Is_Intercrop As Boolean
Dim N_leach As Double
Dim Res_remov1 As Variant
Dim Res_remov2 As Variant
Dim N_in As Double
Dim N_out As Double
Dim N_nonSymbiontic As Double
Dim BNF_rate(2) As Double
Dim BNF As Double
Dim N_atmo As Double
Dim P_in As Double
Dim P_out As Double
Dim P_balance As Double

'Make sure function is executed each time a value in the input sheets is changed!
Application.Volatile


'Read Cropping System settings
Zone = CroppingSystem_Range(Season, 2).Value
Crop = CroppingSystem_Range(Season, 3).Value

Res_remov1 = CroppingSystem_Range(Season, 6).Value
If Res_remov1 = "Burn" Then
    '100 % N is lost
    Res_remov1 = 1
 Else
    Res_remov1 = Res_remov1 / 100
End If

Res_remov2 = CroppingSystem_Range(Season, 7).Value
If Res_remov2 = "Burn" Then
    '100 % N is lost
    Res_remov2 = 1
 Else
    Res_remov2 = Res_remov2 / 100
End If


ZoneCrop = Zone & Crop

'Check if selected crop is intercropping
If InStr(Crop, "/") > 0 Then
    Temp = Split(Crop, "/")
    Crop_a = Temp(0)
    Crop_b = Temp(1)
    ZoneCrop_a = Zone & Crop_a
    ZoneCrop_b = Zone & Crop_b
    Is_Intercrop = True
 Else
    ZoneCrop_a = ZoneCrop
    Is_Intercrop = False
    
End If

Const BD = 1.25 'g/cm3
Const Layerthickness = 20 'cm

Set AgronR = ThisWorkbook.Sheets("Agronomy").Range("Agronomy_range")
Set AgroEcolR = ThisWorkbook.Sheets("Agroecology").Range("AgroEcol")
Set FertNPriceR = ThisWorkbook.Sheets("Costs_Prices").Range("Fertilizer_N_Price")

'Zone = CroppingSystem_Range(Season, 2).Value


'Read N concentrations; last number is column number of Agronomy range
Yield1N_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 15) / 100
Yield2N_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 16) / 100
Biomass1N_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 17) / 100
Biomass2N_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 18) / 100
BNF_rate(1) = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop_a, AgronR.Columns(1), 0), 23) / 100
If Is_Intercrop Then BNF_rate(2) = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop_b, AgronR.Columns(1), 0), 23) / 100

'Read P concentrations
Yield1P_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 24) / 100
Yield2P_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 25) / 100
Biomass1P_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 26) / 100
Biomass2P_conc = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(ZoneCrop, AgronR.Columns(1), 0), 27) / 100

'Calculate N and P uptake and removal from field
Yield1N = Yield1N_conc * Yield_1
Yield2N = Yield2N_conc * Yield_2
Biomass1N = Biomass1N_conc * Biomass_1
Biomass2N = Biomass2N_conc * Biomass_2

'calculating biological n fixation, assuming that roots make 50 % of aboveground biomass (therefore factor 1.5)
BNF = Yield1N * BNF_rate(1) + Yield2N * BNF_rate(2) + 1.5 * Biomass1N * BNF_rate(1) + 1.5 * Biomass2N * BNF_rate(2)

Yield1P = Yield1P_conc * Yield_1
Yield2P = Yield2P_conc * Yield_2
Biomass1P = Biomass1P_conc * Biomass_1
Biomass2P = Biomass2P_conc * Biomass_2
P_out = Yield1P + Yield2P + Biomass1P + Biomass2P


Nmin_ppm = WorksheetFunction.Index(AgroEcolR, WorksheetFunction.Match(Zone, AgroEcolR.Columns(1), 0), 3)
Nmin_kg = BD * Layerthickness * Nmin_ppm / 10

Fert_type1 = CroppingSystem_Range(Season, 4).Value
Fert_amount1 = CroppingSystem_Range(Season, 5).Value
Fert_type2 = CroppingSystem_Range(Season, 11).Value
Fert_amount2 = CroppingSystem_Range(Season, 12).Value

FertN_kg = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match(Fert_type1, FertNPriceR.Columns(1), 0), 2) / 100 * Fert_amount1 + _
           WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match(Fert_type2, FertNPriceR.Columns(1), 0), 2) / 100 * Fert_amount2

FertP_kg = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match(Fert_type1, FertNPriceR.Columns(1), 0), 4) / 100 * Fert_amount1 + _
           WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match(Fert_type2, FertNPriceR.Columns(1), 0), 4) / 100 * Fert_amount2


Manure_amount = CroppingSystem_Range(Season, 8).Value * 1000
ManureN_kg = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match("Manure", FertNPriceR.Columns(1), 0), 2) / 100 * Manure_amount
ManureP_kg = WorksheetFunction.Index(FertNPriceR, WorksheetFunction.Match("Manure", FertNPriceR.Columns(1), 0), 4) / 100 * Manure_amount

P_in = FertP_kg + ManureP_kg

Clay = WorksheetFunction.Index(AgroEcolR, WorksheetFunction.Match(Zone, AgroEcolR.Columns(1), 0), 5)

'Taking care that leaching is calculated per season and not per year, if the year has more than one season
Seasons_per_year = WorksheetFunction.Index(AgroEcolR, WorksheetFunction.Match(Zone, AgroEcolR.Columns(1), 0), 6)
Rain_amount = WorksheetFunction.Index(AgroEcolR, WorksheetFunction.Match(Zone, AgroEcolR.Columns(1), 0), 2) / Seasons_per_year

'Non-symbiontic N
N_nonSymbiontic = Non_Symbiotic_Nfixation(Rain_amount)

'N-atmospheric input
N_atmo = Atmosph_Deposition(Rain_amount)

'all N inputs
N_in = FertN_kg + ManureN_kg + N_nonSymbiontic + N_atmo + BNF

'N-leaching
N_leach = N_leaching(Clay, N_in, Rain_amount)


N_out = Yield1N + Yield2N + Biomass1N * Res_remov1 + Biomass2N * Res_remov2 + N_leach

Select Case Return_what
    Case Is = "N_balance"
        N_P_balance = N_in - N_out
    Case Is = "N_input"
        N_P_balance = N_in
    Case Is = "P_balance"
        N_P_balance = P_in - P_out
End Select


End Function

Function Fertilizer_response(Zone_Crop As String, N_kg As Double, P_kg As Double) As Double

Dim FertResponseR As Range
Dim a(2) As Double 'N and P
Dim b(2) As Double
Dim c(2) As Double
Dim r As Double

Set FertResponseR = ThisWorkbook.Sheets("Fert_resp").Range("Fert_response")

'Nitrogen
b(1) = WorksheetFunction.Index(FertResponseR, WorksheetFunction.Match(Zone_Crop, FertResponseR.Columns(1), 0), 7)
c(1) = WorksheetFunction.Index(FertResponseR, WorksheetFunction.Match(Zone_Crop, FertResponseR.Columns(1), 0), 8)
a(1) = WorksheetFunction.Index(FertResponseR, WorksheetFunction.Match(Zone_Crop, FertResponseR.Columns(1), 0), 10)
'Phosphate
b(2) = WorksheetFunction.Index(FertResponseR, WorksheetFunction.Match(Zone_Crop, FertResponseR.Columns(1), 0), 14)
c(2) = WorksheetFunction.Index(FertResponseR, WorksheetFunction.Match(Zone_Crop, FertResponseR.Columns(1), 0), 15)
a(2) = WorksheetFunction.Index(FertResponseR, WorksheetFunction.Match(Zone_Crop, FertResponseR.Columns(1), 0), 17)

Fertilizer_response = ((a(1) - b(1) * c(1) ^ N_kg) + (a(2) - b(2) * c(2) ^ P_kg)) * 1000


End Function

Function SOM_response(CroppingSystem_Range As Range, Season As Integer)

Dim Zone As String
Dim Crop_of_Season(7) As String
Dim SOM_Score(7) As Integer
Dim Final_Score As Integer
Dim AgronR As Range
Dim AgronSubseqR As Range
Dim i As Integer

'Assign worksheet name-ranges to VBA ranges
Set AgronR = ThisWorkbook.Sheets("Agronomy").Range("Agronomy_range")
Set AgronSubseqR = ThisWorkbook.Sheets("Agronomy_subseq ").Range("Agron_subsequent")

Zone = CroppingSystem_Range(Season, 2).Value


For i = Season To 1 Step -1
    Crop_of_Season(i) = CroppingSystem_Range(i, 3).Value
    SOM_Score(i) = WorksheetFunction.Index(AgronR, WorksheetFunction.Match(Zone & Crop_of_Season(i), AgronR.Columns(1), 0), 34)
    Final_Score = Final_Score + SOM_Score(i)
Next i


SOM_response = Final_Score


End Function


