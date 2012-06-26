Attribute VB_Name = "PISO813"

'/************ define PISO813 relative address **********************/
Global Const PISO813_AD_LO = &HD0        '// Analog to Digital, Low Byte
Global Const PISO813_AD_HI = &HD4        '// Analog to Digital, High Byte
Global Const PISO813_SET_CH = &HE0       '// channel selecting
Global Const PISO813_SET_GAIN = &HE4     '// PGA gain code
Global Const PISO813_SOFT_TRIG = &HF0    '// A/D trigger control register

'/****** define the gain mode ********/
Global Const PISO813_BI_1 = &H0
Global Const PISO813_BI_2 = &H1
Global Const PISO813_BI_4 = &H2
Global Const PISO813_BI_8 = &H3
Global Const PISO813_BI_16 = &H4

Global Const PISO813_UNI_1 = &H0
Global Const PISO813_UNI_2 = &H1
Global Const PISO813_UNI_4 = &H2
Global Const PISO813_UNI_8 = &H3
Global Const PISO813_UNI_16 = &H4

