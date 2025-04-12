# CalculationHPC-turbine

A program for calculating the high pressure cylinder of a steam turbine.

# Dependencies

Interface - tkinter

Working with exel - openpyxl

Properties of water and steam - iapws (IAPWS97)

# Name of the modules

## CalPPT - Сalculation of process parameters in a turbine

Сalculation function for module:

- DIP -> Determining the initial parameters - (Индекс 0 и Индекс kt'')
- DHT -> Disposable heat transfer
- NFR -> Nominal flow rate
- EPC -> Electrical power calculation
- APL_SIO -> Assessment of pressure losses in the steam inlet organs
- SP_CV -> Steam parameters after control valves - (Индекс 0 (штрих))
- HTofRS -> Heat transfer of the regulating stage
- SM_CS_ExLos -> Steam parameters after the control stage (excluding losses) - (Индекс 2рсt)
- COS_OF_CS -> Cost-effectiveness of the control stage
- ATD_CS -> Available thermal differential of the control stage
- PA_CS_TIAL -> Parameters after the control stage (taking into account losses) - (Индекс 2рс)
- SP_UNG_ExL -> Steam parameters at the output of an unregulated group of stages (excluding losses) - (Индекс kt)
- DHT_GNoNREG -> Disposable heat transfer of a group of non-regulating steps

## SaveExel - Saving data to main.py in exel (DB)
