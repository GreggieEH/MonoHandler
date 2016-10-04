// dispids.h
// dispatch ids

#pragma once

// properties
#define				DISPID_CurrentWavelength		0x0100
#define				DISPID_AutoGrating				0x0101
#define				DISPID_CurrentGrating			0x0102
#define				DISPID_MonoProgID				0x0103
#define				DISPID_AmBusy					0x0104
#define				DISPID_NumberOfGratings			0x0105
#define				DISPID_AmInitialized			0x0106
#define				DISPID_MonoParams				0x0107
#define				DISPID_GratingParams			0x0108

// methods
#define				DISPID_IsValidPosition			0x0120
#define				DISPID_Home						0x0121
#define				DISPID_ConvertStepsToNM			0x0122
#define				DISPID_DisplayConfigValues		0x0123
#define				DISPID_DoSetup					0x0124
#define				DISPID_ReadConfig				0x0125
#define				DISPID_WriteConfig				0x0126
#define				DISPID_ReadIni					0x0127
#define				DISPID_WriteIni					0x0128
#define				DISPID_SetGratingParams			0x0129
#define				DISPID_GetGratingDispersion		0x012a
#define				DISPID_DisplayWaveControl		0x012b
#define				DISPID_StartScan				0x012c

// events
#define				DISPID_GratingChanged			0x0140
#define				DISPID_BeforeMoveChange			0x0141
#define				DISPID_HaveAborted				0x0142
#define				DISPID_MoveCompleted			0x0143
#define				DISPID_MoveError				0x0144
#define				DISPID_QueryAllowChangeZeroOffset	0x0145
#define				DISPID_QueryScanningStatus		0x0146
#define				DISPID_RequestChangeGrating		0x0147
#define				DISPID_RequestConfigFile		0x0148
#define				DISPID_RequestIniFile			0x0149
#define				DISPID_RequestParentWindow		0x014a
#define				DISPID_StatusMessage			0x014b
#define				DISPID_RequestCurrentSignal		0x014c